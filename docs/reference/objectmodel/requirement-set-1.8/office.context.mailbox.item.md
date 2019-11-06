---
title: "\"Context\"-\"邮箱\"。项目-要求集1。8"
description: ''
ms.date: 11/05/2019
localization_priority: Normal
ms.openlocfilehash: 065ea3c74580555c0df1af7b495127a25493b612
ms.sourcegitcommit: 21aa084875c9e07a300b3bbe8852b3e5dd163e1d
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 11/06/2019
ms.locfileid: "38001570"
---
# <a name="item"></a><span data-ttu-id="1676f-102">item</span><span class="sxs-lookup"><span data-stu-id="1676f-102">item</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmditem"></a><span data-ttu-id="1676f-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span><span class="sxs-lookup"><span data-stu-id="1676f-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span></span>

<span data-ttu-id="1676f-p101">`item` 命名空间用于访问当前选定的邮件、会议请求或约会。可以通过使用 [itemType](#itemtype-officemailboxenumsitemtype) 属性确定 `item` 的类型。</span><span class="sxs-lookup"><span data-stu-id="1676f-p101">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="1676f-106">Requirements</span><span class="sxs-lookup"><span data-stu-id="1676f-106">Requirements</span></span>

|<span data-ttu-id="1676f-107">要求</span><span class="sxs-lookup"><span data-stu-id="1676f-107">Requirement</span></span>|<span data-ttu-id="1676f-108">值</span><span class="sxs-lookup"><span data-stu-id="1676f-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="1676f-109">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="1676f-109">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="1676f-110">1.0</span><span class="sxs-lookup"><span data-stu-id="1676f-110">1.0</span></span>|
|[<span data-ttu-id="1676f-111">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="1676f-111">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="1676f-112">受限</span><span class="sxs-lookup"><span data-stu-id="1676f-112">Restricted</span></span>|
|[<span data-ttu-id="1676f-113">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="1676f-113">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="1676f-114">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="1676f-114">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="1676f-115">成员和方法</span><span class="sxs-lookup"><span data-stu-id="1676f-115">Members and methods</span></span>

| <span data-ttu-id="1676f-116">成员</span><span class="sxs-lookup"><span data-stu-id="1676f-116">Member</span></span> | <span data-ttu-id="1676f-117">类型</span><span class="sxs-lookup"><span data-stu-id="1676f-117">Type</span></span> |
|--------|------|
| [<span data-ttu-id="1676f-118">attachments</span><span class="sxs-lookup"><span data-stu-id="1676f-118">attachments</span></span>](#attachments-arrayattachmentdetails) | <span data-ttu-id="1676f-119">成员</span><span class="sxs-lookup"><span data-stu-id="1676f-119">Member</span></span> |
| [<span data-ttu-id="1676f-120">bcc</span><span class="sxs-lookup"><span data-stu-id="1676f-120">bcc</span></span>](#bcc-recipients) | <span data-ttu-id="1676f-121">成员</span><span class="sxs-lookup"><span data-stu-id="1676f-121">Member</span></span> |
| [<span data-ttu-id="1676f-122">body</span><span class="sxs-lookup"><span data-stu-id="1676f-122">body</span></span>](#body-body) | <span data-ttu-id="1676f-123">成员</span><span class="sxs-lookup"><span data-stu-id="1676f-123">Member</span></span> |
| [<span data-ttu-id="1676f-124">categories</span><span class="sxs-lookup"><span data-stu-id="1676f-124">categories</span></span>](#categories-categories) | <span data-ttu-id="1676f-125">成员</span><span class="sxs-lookup"><span data-stu-id="1676f-125">Member</span></span> |
| [<span data-ttu-id="1676f-126">cc</span><span class="sxs-lookup"><span data-stu-id="1676f-126">cc</span></span>](#cc-arrayemailaddressdetailsrecipients) | <span data-ttu-id="1676f-127">成员</span><span class="sxs-lookup"><span data-stu-id="1676f-127">Member</span></span> |
| [<span data-ttu-id="1676f-128">conversationId</span><span class="sxs-lookup"><span data-stu-id="1676f-128">conversationId</span></span>](#nullable-conversationid-string) | <span data-ttu-id="1676f-129">成员</span><span class="sxs-lookup"><span data-stu-id="1676f-129">Member</span></span> |
| [<span data-ttu-id="1676f-130">dateTimeCreated</span><span class="sxs-lookup"><span data-stu-id="1676f-130">dateTimeCreated</span></span>](#datetimecreated-date) | <span data-ttu-id="1676f-131">成员</span><span class="sxs-lookup"><span data-stu-id="1676f-131">Member</span></span> |
| [<span data-ttu-id="1676f-132">dateTimeModified</span><span class="sxs-lookup"><span data-stu-id="1676f-132">dateTimeModified</span></span>](#datetimemodified-date) | <span data-ttu-id="1676f-133">成员</span><span class="sxs-lookup"><span data-stu-id="1676f-133">Member</span></span> |
| [<span data-ttu-id="1676f-134">end</span><span class="sxs-lookup"><span data-stu-id="1676f-134">end</span></span>](#end-datetime) | <span data-ttu-id="1676f-135">成员</span><span class="sxs-lookup"><span data-stu-id="1676f-135">Member</span></span> |
| [<span data-ttu-id="1676f-136">enhancedLocation</span><span class="sxs-lookup"><span data-stu-id="1676f-136">enhancedLocation</span></span>](#enhancedlocation-enhancedlocation) | <span data-ttu-id="1676f-137">成员</span><span class="sxs-lookup"><span data-stu-id="1676f-137">Member</span></span> |
| [<span data-ttu-id="1676f-138">from</span><span class="sxs-lookup"><span data-stu-id="1676f-138">from</span></span>](#from-emailaddressdetailsfrom) | <span data-ttu-id="1676f-139">成员</span><span class="sxs-lookup"><span data-stu-id="1676f-139">Member</span></span> |
| [<span data-ttu-id="1676f-140">internetHeaders</span><span class="sxs-lookup"><span data-stu-id="1676f-140">internetHeaders</span></span>](#internetheaders-internetheaders) | <span data-ttu-id="1676f-141">成员</span><span class="sxs-lookup"><span data-stu-id="1676f-141">Member</span></span> |
| [<span data-ttu-id="1676f-142">internetMessageId</span><span class="sxs-lookup"><span data-stu-id="1676f-142">internetMessageId</span></span>](#internetmessageid-string) | <span data-ttu-id="1676f-143">成员</span><span class="sxs-lookup"><span data-stu-id="1676f-143">Member</span></span> |
| [<span data-ttu-id="1676f-144">itemClass</span><span class="sxs-lookup"><span data-stu-id="1676f-144">itemClass</span></span>](#itemclass-string) | <span data-ttu-id="1676f-145">成员</span><span class="sxs-lookup"><span data-stu-id="1676f-145">Member</span></span> |
| [<span data-ttu-id="1676f-146">itemId</span><span class="sxs-lookup"><span data-stu-id="1676f-146">itemId</span></span>](#nullable-itemid-string) | <span data-ttu-id="1676f-147">成员</span><span class="sxs-lookup"><span data-stu-id="1676f-147">Member</span></span> |
| [<span data-ttu-id="1676f-148">itemType</span><span class="sxs-lookup"><span data-stu-id="1676f-148">itemType</span></span>](#itemtype-officemailboxenumsitemtype) | <span data-ttu-id="1676f-149">成员</span><span class="sxs-lookup"><span data-stu-id="1676f-149">Member</span></span> |
| [<span data-ttu-id="1676f-150">location</span><span class="sxs-lookup"><span data-stu-id="1676f-150">location</span></span>](#location-stringlocation) | <span data-ttu-id="1676f-151">成员</span><span class="sxs-lookup"><span data-stu-id="1676f-151">Member</span></span> |
| [<span data-ttu-id="1676f-152">normalizedSubject</span><span class="sxs-lookup"><span data-stu-id="1676f-152">normalizedSubject</span></span>](#normalizedsubject-string) | <span data-ttu-id="1676f-153">成员</span><span class="sxs-lookup"><span data-stu-id="1676f-153">Member</span></span> |
| [<span data-ttu-id="1676f-154">notificationMessages</span><span class="sxs-lookup"><span data-stu-id="1676f-154">notificationMessages</span></span>](#notificationmessages-notificationmessages) | <span data-ttu-id="1676f-155">Member</span><span class="sxs-lookup"><span data-stu-id="1676f-155">Member</span></span> |
| [<span data-ttu-id="1676f-156">optionalAttendees</span><span class="sxs-lookup"><span data-stu-id="1676f-156">optionalAttendees</span></span>](#optionalattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="1676f-157">成员</span><span class="sxs-lookup"><span data-stu-id="1676f-157">Member</span></span> |
| [<span data-ttu-id="1676f-158">organizer</span><span class="sxs-lookup"><span data-stu-id="1676f-158">organizer</span></span>](#organizer-emailaddressdetailsorganizer) | <span data-ttu-id="1676f-159">成员</span><span class="sxs-lookup"><span data-stu-id="1676f-159">Member</span></span> |
| [<span data-ttu-id="1676f-160">recurrence</span><span class="sxs-lookup"><span data-stu-id="1676f-160">recurrence</span></span>](#nullable-recurrence-recurrence) | <span data-ttu-id="1676f-161">成员</span><span class="sxs-lookup"><span data-stu-id="1676f-161">Member</span></span> |
| [<span data-ttu-id="1676f-162">requiredAttendees</span><span class="sxs-lookup"><span data-stu-id="1676f-162">requiredAttendees</span></span>](#requiredattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="1676f-163">成员</span><span class="sxs-lookup"><span data-stu-id="1676f-163">Member</span></span> |
| [<span data-ttu-id="1676f-164">sender</span><span class="sxs-lookup"><span data-stu-id="1676f-164">sender</span></span>](#sender-emailaddressdetails) | <span data-ttu-id="1676f-165">Member</span><span class="sxs-lookup"><span data-stu-id="1676f-165">Member</span></span> |
| [<span data-ttu-id="1676f-166">Webcasts&seriesid</span><span class="sxs-lookup"><span data-stu-id="1676f-166">seriesId</span></span>](#nullable-seriesid-string) | <span data-ttu-id="1676f-167">Member</span><span class="sxs-lookup"><span data-stu-id="1676f-167">Member</span></span> |
| [<span data-ttu-id="1676f-168">start</span><span class="sxs-lookup"><span data-stu-id="1676f-168">start</span></span>](#start-datetime) | <span data-ttu-id="1676f-169">Member</span><span class="sxs-lookup"><span data-stu-id="1676f-169">Member</span></span> |
| [<span data-ttu-id="1676f-170">subject</span><span class="sxs-lookup"><span data-stu-id="1676f-170">subject</span></span>](#subject-stringsubject) | <span data-ttu-id="1676f-171">成员</span><span class="sxs-lookup"><span data-stu-id="1676f-171">Member</span></span> |
| [<span data-ttu-id="1676f-172">to</span><span class="sxs-lookup"><span data-stu-id="1676f-172">to</span></span>](#to-arrayemailaddressdetailsrecipients) | <span data-ttu-id="1676f-173">成员</span><span class="sxs-lookup"><span data-stu-id="1676f-173">Member</span></span> |
| [<span data-ttu-id="1676f-174">addFileAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="1676f-174">addFileAttachmentAsync</span></span>](#addfileattachmentasyncuri-attachmentname-options-callback) | <span data-ttu-id="1676f-175">方法</span><span class="sxs-lookup"><span data-stu-id="1676f-175">Method</span></span> |
| [<span data-ttu-id="1676f-176">addFileAttachmentFromBase64Async</span><span class="sxs-lookup"><span data-stu-id="1676f-176">addFileAttachmentFromBase64Async</span></span>](#addfileattachmentfrombase64asyncbase64file-attachmentname-options-callback) | <span data-ttu-id="1676f-177">方法</span><span class="sxs-lookup"><span data-stu-id="1676f-177">Method</span></span> |
| [<span data-ttu-id="1676f-178">addHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="1676f-178">addHandlerAsync</span></span>](#addhandlerasynceventtype-handler-options-callback) | <span data-ttu-id="1676f-179">方法</span><span class="sxs-lookup"><span data-stu-id="1676f-179">Method</span></span> |
| [<span data-ttu-id="1676f-180">addItemAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="1676f-180">addItemAttachmentAsync</span></span>](#additemattachmentasyncitemid-attachmentname-options-callback) | <span data-ttu-id="1676f-181">方法</span><span class="sxs-lookup"><span data-stu-id="1676f-181">Method</span></span> |
| [<span data-ttu-id="1676f-182">close</span><span class="sxs-lookup"><span data-stu-id="1676f-182">close</span></span>](#close) | <span data-ttu-id="1676f-183">方法</span><span class="sxs-lookup"><span data-stu-id="1676f-183">Method</span></span> |
| [<span data-ttu-id="1676f-184">displayReplyAllForm</span><span class="sxs-lookup"><span data-stu-id="1676f-184">displayReplyAllForm</span></span>](#displayreplyallformformdata-callback) | <span data-ttu-id="1676f-185">方法</span><span class="sxs-lookup"><span data-stu-id="1676f-185">Method</span></span> |
| [<span data-ttu-id="1676f-186">displayReplyForm</span><span class="sxs-lookup"><span data-stu-id="1676f-186">displayReplyForm</span></span>](#displayreplyformformdata-callback) | <span data-ttu-id="1676f-187">方法</span><span class="sxs-lookup"><span data-stu-id="1676f-187">Method</span></span> |
| [<span data-ttu-id="1676f-188">getAllInternetHeadersAsync</span><span class="sxs-lookup"><span data-stu-id="1676f-188">getAllInternetHeadersAsync</span></span>](#getallinternetheadersasyncoptions-callback) | <span data-ttu-id="1676f-189">方法</span><span class="sxs-lookup"><span data-stu-id="1676f-189">Method</span></span> |
| [<span data-ttu-id="1676f-190">getAttachmentContentAsync</span><span class="sxs-lookup"><span data-stu-id="1676f-190">getAttachmentContentAsync</span></span>](#getattachmentcontentasyncattachmentid-options-callback--attachmentcontent) | <span data-ttu-id="1676f-191">方法</span><span class="sxs-lookup"><span data-stu-id="1676f-191">Method</span></span> |
| [<span data-ttu-id="1676f-192">getAttachmentsAsync</span><span class="sxs-lookup"><span data-stu-id="1676f-192">getAttachmentsAsync</span></span>](#getattachmentsasyncoptions-callback--arrayattachmentdetails) | <span data-ttu-id="1676f-193">方法</span><span class="sxs-lookup"><span data-stu-id="1676f-193">Method</span></span> |
| [<span data-ttu-id="1676f-194">getEntities</span><span class="sxs-lookup"><span data-stu-id="1676f-194">getEntities</span></span>](#getentities--entities) | <span data-ttu-id="1676f-195">Method</span><span class="sxs-lookup"><span data-stu-id="1676f-195">Method</span></span> |
| [<span data-ttu-id="1676f-196">getEntitiesByType</span><span class="sxs-lookup"><span data-stu-id="1676f-196">getEntitiesByType</span></span>](#getentitiesbytypeentitytype--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="1676f-197">Method</span><span class="sxs-lookup"><span data-stu-id="1676f-197">Method</span></span> |
| [<span data-ttu-id="1676f-198">getFilteredEntitiesByName</span><span class="sxs-lookup"><span data-stu-id="1676f-198">getFilteredEntitiesByName</span></span>](#getfilteredentitiesbynamename--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="1676f-199">Method</span><span class="sxs-lookup"><span data-stu-id="1676f-199">Method</span></span> |
| [<span data-ttu-id="1676f-200">getItemIdAsync</span><span class="sxs-lookup"><span data-stu-id="1676f-200">getItemIdAsync</span></span>](#getitemidasyncoptions-callback) | <span data-ttu-id="1676f-201">Method</span><span class="sxs-lookup"><span data-stu-id="1676f-201">Method</span></span> |
| [<span data-ttu-id="1676f-202">getRegExMatches</span><span class="sxs-lookup"><span data-stu-id="1676f-202">getRegExMatches</span></span>](#getregexmatches--object) | <span data-ttu-id="1676f-203">Method</span><span class="sxs-lookup"><span data-stu-id="1676f-203">Method</span></span> |
| [<span data-ttu-id="1676f-204">getRegExMatchesByName</span><span class="sxs-lookup"><span data-stu-id="1676f-204">getRegExMatchesByName</span></span>](#getregexmatchesbynamename--nullable-array-string-) | <span data-ttu-id="1676f-205">Method</span><span class="sxs-lookup"><span data-stu-id="1676f-205">Method</span></span> |
| [<span data-ttu-id="1676f-206">getSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="1676f-206">getSelectedDataAsync</span></span>](#getselecteddataasynccoerciontype-options-callback--string) | <span data-ttu-id="1676f-207">Method</span><span class="sxs-lookup"><span data-stu-id="1676f-207">Method</span></span> |
| [<span data-ttu-id="1676f-208">Office.context.mailbox.item.getselectedentities</span><span class="sxs-lookup"><span data-stu-id="1676f-208">getSelectedEntities</span></span>](#getselectedentities--entities) | <span data-ttu-id="1676f-209">Method</span><span class="sxs-lookup"><span data-stu-id="1676f-209">Method</span></span> |
| [<span data-ttu-id="1676f-210">Office.context.mailbox.item.getselectedregexmatches</span><span class="sxs-lookup"><span data-stu-id="1676f-210">getSelectedRegExMatches</span></span>](#getselectedregexmatches--object) | <span data-ttu-id="1676f-211">Method</span><span class="sxs-lookup"><span data-stu-id="1676f-211">Method</span></span> |
| [<span data-ttu-id="1676f-212">getSharedPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="1676f-212">getSharedPropertiesAsync</span></span>](#getsharedpropertiesasyncoptions-callback) | <span data-ttu-id="1676f-213">Method</span><span class="sxs-lookup"><span data-stu-id="1676f-213">Method</span></span> |
| [<span data-ttu-id="1676f-214">loadCustomPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="1676f-214">loadCustomPropertiesAsync</span></span>](#loadcustompropertiesasynccallback-usercontext) | <span data-ttu-id="1676f-215">Method</span><span class="sxs-lookup"><span data-stu-id="1676f-215">Method</span></span> |
| [<span data-ttu-id="1676f-216">removeAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="1676f-216">removeAttachmentAsync</span></span>](#removeattachmentasyncattachmentid-options-callback) | <span data-ttu-id="1676f-217">方法</span><span class="sxs-lookup"><span data-stu-id="1676f-217">Method</span></span> |
| [<span data-ttu-id="1676f-218">removeHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="1676f-218">removeHandlerAsync</span></span>](#removehandlerasynceventtype-options-callback) | <span data-ttu-id="1676f-219">方法</span><span class="sxs-lookup"><span data-stu-id="1676f-219">Method</span></span> |
| [<span data-ttu-id="1676f-220">saveAsync</span><span class="sxs-lookup"><span data-stu-id="1676f-220">saveAsync</span></span>](#saveasyncoptions-callback) | <span data-ttu-id="1676f-221">Method</span><span class="sxs-lookup"><span data-stu-id="1676f-221">Method</span></span> |
| [<span data-ttu-id="1676f-222">setSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="1676f-222">setSelectedDataAsync</span></span>](#setselecteddataasyncdata-options-callback) | <span data-ttu-id="1676f-223">方法</span><span class="sxs-lookup"><span data-stu-id="1676f-223">Method</span></span> |

### <a name="example"></a><span data-ttu-id="1676f-224">示例</span><span class="sxs-lookup"><span data-stu-id="1676f-224">Example</span></span>

<span data-ttu-id="1676f-225">以下 JavaScript 代码示例显示了如何访问 Outlook 中当前项目的 `subject` 属性。</span><span class="sxs-lookup"><span data-stu-id="1676f-225">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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

### <a name="members"></a><span data-ttu-id="1676f-226">Members</span><span class="sxs-lookup"><span data-stu-id="1676f-226">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetailsviewoutlook-js-18"></a><span data-ttu-id="1676f-227">attachments: Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.8)></span><span class="sxs-lookup"><span data-stu-id="1676f-227">attachments: Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.8)></span></span>

<span data-ttu-id="1676f-228">以数组的形式获取项目的附件。</span><span class="sxs-lookup"><span data-stu-id="1676f-228">Gets the item's attachments as an array.</span></span> <span data-ttu-id="1676f-229">仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="1676f-229">Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="1676f-230">由于潜在的安全问题，某些类型的文件会受到 Outlook 阻止，并且不会返回。</span><span class="sxs-lookup"><span data-stu-id="1676f-230">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="1676f-231">如需了解更多信息，请参阅 [Outlook 中阻止的附件](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519)。</span><span class="sxs-lookup"><span data-stu-id="1676f-231">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="1676f-232">类型</span><span class="sxs-lookup"><span data-stu-id="1676f-232">Type</span></span>

*   <span data-ttu-id="1676f-233">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.8)></span><span class="sxs-lookup"><span data-stu-id="1676f-233">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.8)></span></span>

##### <a name="requirements"></a><span data-ttu-id="1676f-234">要求</span><span class="sxs-lookup"><span data-stu-id="1676f-234">Requirements</span></span>

|<span data-ttu-id="1676f-235">要求</span><span class="sxs-lookup"><span data-stu-id="1676f-235">Requirement</span></span>|<span data-ttu-id="1676f-236">值</span><span class="sxs-lookup"><span data-stu-id="1676f-236">Value</span></span>|
|---|---|
|[<span data-ttu-id="1676f-237">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="1676f-237">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="1676f-238">1.0</span><span class="sxs-lookup"><span data-stu-id="1676f-238">1.0</span></span>|
|[<span data-ttu-id="1676f-239">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="1676f-239">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="1676f-240">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1676f-240">ReadItem</span></span>|
|[<span data-ttu-id="1676f-241">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="1676f-241">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="1676f-242">阅读</span><span class="sxs-lookup"><span data-stu-id="1676f-242">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="1676f-243">示例</span><span class="sxs-lookup"><span data-stu-id="1676f-243">Example</span></span>

<span data-ttu-id="1676f-244">以下代码使用当前项目上所有附件的详细信息构成 HTML 字符串。</span><span class="sxs-lookup"><span data-stu-id="1676f-244">The following code builds an HTML string with details of all attachments on the current item.</span></span>

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

#### <a name="bcc-recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-18"></a><span data-ttu-id="1676f-245">bcc: [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.8)</span><span class="sxs-lookup"><span data-stu-id="1676f-245">bcc: [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.8)</span></span>

<span data-ttu-id="1676f-246">获取对象，该对象提供用于获取或更新邮件的密件抄送 (Bcc) 行上的收件人的方法。</span><span class="sxs-lookup"><span data-stu-id="1676f-246">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="1676f-247">仅限撰写模式。</span><span class="sxs-lookup"><span data-stu-id="1676f-247">Compose mode only.</span></span>

<span data-ttu-id="1676f-248">默认情况下，集合上限为 100 个成员。</span><span class="sxs-lookup"><span data-stu-id="1676f-248">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="1676f-249">但是，在 Windows 和 Mac 上，以下限制适用。</span><span class="sxs-lookup"><span data-stu-id="1676f-249">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="1676f-250">最多包含 500 个成员。</span><span class="sxs-lookup"><span data-stu-id="1676f-250">Get 500 members maximum.</span></span>
- <span data-ttu-id="1676f-251">为每个呼叫最多设置 100 个成员，总共多达 500 个成员。</span><span class="sxs-lookup"><span data-stu-id="1676f-251">Set a maximum of 100 members per call, up to 500 members total.</span></span>

##### <a name="type"></a><span data-ttu-id="1676f-252">类型</span><span class="sxs-lookup"><span data-stu-id="1676f-252">Type</span></span>

*   [<span data-ttu-id="1676f-253">收件人</span><span class="sxs-lookup"><span data-stu-id="1676f-253">Recipients</span></span>](/javascript/api/outlook/office.recipients?view=outlook-js-1.8)

##### <a name="requirements"></a><span data-ttu-id="1676f-254">Requirements</span><span class="sxs-lookup"><span data-stu-id="1676f-254">Requirements</span></span>

|<span data-ttu-id="1676f-255">要求</span><span class="sxs-lookup"><span data-stu-id="1676f-255">Requirement</span></span>|<span data-ttu-id="1676f-256">值</span><span class="sxs-lookup"><span data-stu-id="1676f-256">Value</span></span>|
|---|---|
|[<span data-ttu-id="1676f-257">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="1676f-257">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="1676f-258">1.1</span><span class="sxs-lookup"><span data-stu-id="1676f-258">1.1</span></span>|
|[<span data-ttu-id="1676f-259">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="1676f-259">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="1676f-260">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1676f-260">ReadItem</span></span>|
|[<span data-ttu-id="1676f-261">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="1676f-261">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="1676f-262">撰写</span><span class="sxs-lookup"><span data-stu-id="1676f-262">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="1676f-263">示例</span><span class="sxs-lookup"><span data-stu-id="1676f-263">Example</span></span>

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

#### <a name="body-bodyjavascriptapioutlookofficebodyviewoutlook-js-18"></a><span data-ttu-id="1676f-264">body: [Body](/javascript/api/outlook/office.body?view=outlook-js-1.8)</span><span class="sxs-lookup"><span data-stu-id="1676f-264">body: [Body](/javascript/api/outlook/office.body?view=outlook-js-1.8)</span></span>

<span data-ttu-id="1676f-265">获取一个提供用于处理项目正文的方法的对象。</span><span class="sxs-lookup"><span data-stu-id="1676f-265">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="1676f-266">类型</span><span class="sxs-lookup"><span data-stu-id="1676f-266">Type</span></span>

*   [<span data-ttu-id="1676f-267">Body</span><span class="sxs-lookup"><span data-stu-id="1676f-267">Body</span></span>](/javascript/api/outlook/office.body?view=outlook-js-1.8)

##### <a name="requirements"></a><span data-ttu-id="1676f-268">Requirements</span><span class="sxs-lookup"><span data-stu-id="1676f-268">Requirements</span></span>

|<span data-ttu-id="1676f-269">要求</span><span class="sxs-lookup"><span data-stu-id="1676f-269">Requirement</span></span>|<span data-ttu-id="1676f-270">值</span><span class="sxs-lookup"><span data-stu-id="1676f-270">Value</span></span>|
|---|---|
|[<span data-ttu-id="1676f-271">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="1676f-271">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="1676f-272">1.1</span><span class="sxs-lookup"><span data-stu-id="1676f-272">1.1</span></span>|
|[<span data-ttu-id="1676f-273">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="1676f-273">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="1676f-274">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1676f-274">ReadItem</span></span>|
|[<span data-ttu-id="1676f-275">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="1676f-275">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="1676f-276">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="1676f-276">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="1676f-277">示例</span><span class="sxs-lookup"><span data-stu-id="1676f-277">Example</span></span>

<span data-ttu-id="1676f-278">本示例获取纯文本格式的邮件正文。</span><span class="sxs-lookup"><span data-stu-id="1676f-278">This example gets the body of the message in plain text.</span></span>

```js
Office.context.mailbox.item.body.getAsync(
  "text",
  { asyncContext: "This is passed to the callback" },
  function callback(result) {
    // Do something with the result.
  });

```

<span data-ttu-id="1676f-279">以下是传递到回调函数的结果参数的示例。</span><span class="sxs-lookup"><span data-stu-id="1676f-279">The following is an example of the result parameter passed to the callback function.</span></span>

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

#### <a name="categories-categoriesjavascriptapioutlookofficecategoriesviewoutlook-js-18"></a><span data-ttu-id="1676f-280">类别：[类别](/javascript/api/outlook/office.categories?view=outlook-js-1.8)</span><span class="sxs-lookup"><span data-stu-id="1676f-280">categories: [Categories](/javascript/api/outlook/office.categories?view=outlook-js-1.8)</span></span>

<span data-ttu-id="1676f-281">获取一个对象，该对象提供用于管理项的类别的方法。</span><span class="sxs-lookup"><span data-stu-id="1676f-281">Gets an object that provides methods for managing the item's categories.</span></span>

> [!NOTE]
> <span data-ttu-id="1676f-282">iOS 版 Outlook 或 Android 版 Outlook 不支持此成员。</span><span class="sxs-lookup"><span data-stu-id="1676f-282">This member is not supported in Outlook on iOS or Android.</span></span>

##### <a name="type"></a><span data-ttu-id="1676f-283">类型</span><span class="sxs-lookup"><span data-stu-id="1676f-283">Type</span></span>

*   [<span data-ttu-id="1676f-284">Categories</span><span class="sxs-lookup"><span data-stu-id="1676f-284">Categories</span></span>](/javascript/api/outlook/office.categories?view=outlook-js-1.8)

##### <a name="requirements"></a><span data-ttu-id="1676f-285">Requirements</span><span class="sxs-lookup"><span data-stu-id="1676f-285">Requirements</span></span>

|<span data-ttu-id="1676f-286">要求</span><span class="sxs-lookup"><span data-stu-id="1676f-286">Requirement</span></span>|<span data-ttu-id="1676f-287">值</span><span class="sxs-lookup"><span data-stu-id="1676f-287">Value</span></span>|
|---|---|
|[<span data-ttu-id="1676f-288">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="1676f-288">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="1676f-289">1.8</span><span class="sxs-lookup"><span data-stu-id="1676f-289">1.8</span></span>|
|[<span data-ttu-id="1676f-290">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="1676f-290">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="1676f-291">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1676f-291">ReadItem</span></span>|
|[<span data-ttu-id="1676f-292">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="1676f-292">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="1676f-293">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="1676f-293">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="1676f-294">示例</span><span class="sxs-lookup"><span data-stu-id="1676f-294">Example</span></span>

<span data-ttu-id="1676f-295">此示例获取项的类别。</span><span class="sxs-lookup"><span data-stu-id="1676f-295">This example gets the item's categories.</span></span>

```js
Office.context.mailbox.item.categories.getAsync(function (asyncResult) {
  if (asyncResult.status === Office.AsyncResultStatus.Failed) {
    console.log("Action failed with error: " + asyncResult.error.message);
  } else {
    console.log("Categories: " + JSON.stringify(asyncResult.value));
  }
});
```

<br>

---
---

#### <a name="cc-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-18recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-18"></a><span data-ttu-id="1676f-296">cc: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.8)</span><span class="sxs-lookup"><span data-stu-id="1676f-296">cc: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.8)</span></span>

<span data-ttu-id="1676f-297">提供对邮件的抄送 (Cc) 收件人的访问权限。</span><span class="sxs-lookup"><span data-stu-id="1676f-297">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="1676f-298">对象的类型和访问级别取决于当前项目的模式。</span><span class="sxs-lookup"><span data-stu-id="1676f-298">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="1676f-299">阅读模式</span><span class="sxs-lookup"><span data-stu-id="1676f-299">Read mode</span></span>

<span data-ttu-id="1676f-300">`cc` 属性返回包含邮件的**抄送**行上所列的每个收件人的 `EmailAddressDetails` 对象的数组。</span><span class="sxs-lookup"><span data-stu-id="1676f-300">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message.</span></span> <span data-ttu-id="1676f-301">默认情况下，集合上限为 100 个成员。</span><span class="sxs-lookup"><span data-stu-id="1676f-301">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="1676f-302">但是，在 Windows 和 Mac 上，最多可包含 500 个成员。</span><span class="sxs-lookup"><span data-stu-id="1676f-302">However, on Windows and Mac, you can get 500 members maximum.</span></span>

```js
console.log(JSON.stringify(Office.context.mailbox.item.cc));
```

##### <a name="compose-mode"></a><span data-ttu-id="1676f-303">撰写模式</span><span class="sxs-lookup"><span data-stu-id="1676f-303">Compose mode</span></span>

<span data-ttu-id="1676f-304">`cc` 属性返回一个 `Recipients` 对象，该对象提供用于获取或更新邮件的“**抄送**”行上收件人的方法。</span><span class="sxs-lookup"><span data-stu-id="1676f-304">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span> <span data-ttu-id="1676f-305">默认情况下，集合上限为 100 个成员。</span><span class="sxs-lookup"><span data-stu-id="1676f-305">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="1676f-306">但是，在 Windows 和 Mac 上，以下限制适用。</span><span class="sxs-lookup"><span data-stu-id="1676f-306">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="1676f-307">最多包含 500 个成员。</span><span class="sxs-lookup"><span data-stu-id="1676f-307">Get 500 members maximum.</span></span>
- <span data-ttu-id="1676f-308">为每个呼叫最多设置 100 个成员，总共多达 500 个成员。</span><span class="sxs-lookup"><span data-stu-id="1676f-308">Set a maximum of 100 members per call, up to 500 members total.</span></span>

```js
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="1676f-309">类型</span><span class="sxs-lookup"><span data-stu-id="1676f-309">Type</span></span>

*   <span data-ttu-id="1676f-310">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.8)</span><span class="sxs-lookup"><span data-stu-id="1676f-310">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.8)</span></span>

##### <a name="requirements"></a><span data-ttu-id="1676f-311">Requirements</span><span class="sxs-lookup"><span data-stu-id="1676f-311">Requirements</span></span>

|<span data-ttu-id="1676f-312">要求</span><span class="sxs-lookup"><span data-stu-id="1676f-312">Requirement</span></span>|<span data-ttu-id="1676f-313">值</span><span class="sxs-lookup"><span data-stu-id="1676f-313">Value</span></span>|
|---|---|
|[<span data-ttu-id="1676f-314">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="1676f-314">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="1676f-315">1.0</span><span class="sxs-lookup"><span data-stu-id="1676f-315">1.0</span></span>|
|[<span data-ttu-id="1676f-316">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="1676f-316">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="1676f-317">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1676f-317">ReadItem</span></span>|
|[<span data-ttu-id="1676f-318">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="1676f-318">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="1676f-319">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="1676f-319">Compose or Read</span></span>|

<br>

---
---

#### <a name="nullable-conversationid-string"></a><span data-ttu-id="1676f-320">(nullable) conversationId: String</span><span class="sxs-lookup"><span data-stu-id="1676f-320">(nullable) conversationId: String</span></span>

<span data-ttu-id="1676f-321">获取包含特定消息的电子邮件会话的标识符。</span><span class="sxs-lookup"><span data-stu-id="1676f-321">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="1676f-p109">如果在阅读窗体或撰写窗体的回复中激活邮件应用程序，则此属性可以获得一个整数值。如果用户随后更改了回复邮件的主题（若发送回复），则该邮件的对话 ID 将改变且之前获取的值将不适用。</span><span class="sxs-lookup"><span data-stu-id="1676f-p109">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="1676f-p110">对于撰写窗体的新项目，此属性获得一个 null 值。如果用户设置一个主题并保存该项目，`conversationId` 属性将返回一个值。</span><span class="sxs-lookup"><span data-stu-id="1676f-p110">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="1676f-326">Type</span><span class="sxs-lookup"><span data-stu-id="1676f-326">Type</span></span>

*   <span data-ttu-id="1676f-327">String</span><span class="sxs-lookup"><span data-stu-id="1676f-327">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="1676f-328">Requirements</span><span class="sxs-lookup"><span data-stu-id="1676f-328">Requirements</span></span>

|<span data-ttu-id="1676f-329">要求</span><span class="sxs-lookup"><span data-stu-id="1676f-329">Requirement</span></span>|<span data-ttu-id="1676f-330">值</span><span class="sxs-lookup"><span data-stu-id="1676f-330">Value</span></span>|
|---|---|
|[<span data-ttu-id="1676f-331">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="1676f-331">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="1676f-332">1.0</span><span class="sxs-lookup"><span data-stu-id="1676f-332">1.0</span></span>|
|[<span data-ttu-id="1676f-333">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="1676f-333">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="1676f-334">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1676f-334">ReadItem</span></span>|
|[<span data-ttu-id="1676f-335">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="1676f-335">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="1676f-336">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="1676f-336">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="1676f-337">示例</span><span class="sxs-lookup"><span data-stu-id="1676f-337">Example</span></span>

```js
var conversationId = Office.context.mailbox.item.conversationId;
console.log("conversationId: " + conversationId);
```

<br>

---
---

#### <a name="datetimecreated-date"></a><span data-ttu-id="1676f-338">dateTimeCreated: Date</span><span class="sxs-lookup"><span data-stu-id="1676f-338">dateTimeCreated: Date</span></span>

<span data-ttu-id="1676f-p111">获取项目创建的日期和时间。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="1676f-p111">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="1676f-341">类型</span><span class="sxs-lookup"><span data-stu-id="1676f-341">Type</span></span>

*   <span data-ttu-id="1676f-342">日期</span><span class="sxs-lookup"><span data-stu-id="1676f-342">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="1676f-343">要求</span><span class="sxs-lookup"><span data-stu-id="1676f-343">Requirements</span></span>

|<span data-ttu-id="1676f-344">要求</span><span class="sxs-lookup"><span data-stu-id="1676f-344">Requirement</span></span>|<span data-ttu-id="1676f-345">值</span><span class="sxs-lookup"><span data-stu-id="1676f-345">Value</span></span>|
|---|---|
|[<span data-ttu-id="1676f-346">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="1676f-346">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="1676f-347">1.0</span><span class="sxs-lookup"><span data-stu-id="1676f-347">1.0</span></span>|
|[<span data-ttu-id="1676f-348">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="1676f-348">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="1676f-349">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1676f-349">ReadItem</span></span>|
|[<span data-ttu-id="1676f-350">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="1676f-350">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="1676f-351">阅读</span><span class="sxs-lookup"><span data-stu-id="1676f-351">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="1676f-352">示例</span><span class="sxs-lookup"><span data-stu-id="1676f-352">Example</span></span>

```js
var dateTimeCreated = Office.context.mailbox.item.dateTimeCreated;
console.log("Date and time created: " + dateTimeCreated);
```

<br>

---
---

#### <a name="datetimemodified-date"></a><span data-ttu-id="1676f-353">dateTimeModified: Date</span><span class="sxs-lookup"><span data-stu-id="1676f-353">dateTimeModified: Date</span></span>

<span data-ttu-id="1676f-p112">获取项目最近一次修改的日期和时间。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="1676f-p112">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="1676f-356">iOS 版 Outlook 或 Android 版 Outlook 不支持此成员。</span><span class="sxs-lookup"><span data-stu-id="1676f-356">This member is not supported in Outlook on iOS or Android.</span></span>

##### <a name="type"></a><span data-ttu-id="1676f-357">类型</span><span class="sxs-lookup"><span data-stu-id="1676f-357">Type</span></span>

*   <span data-ttu-id="1676f-358">日期</span><span class="sxs-lookup"><span data-stu-id="1676f-358">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="1676f-359">要求</span><span class="sxs-lookup"><span data-stu-id="1676f-359">Requirements</span></span>

|<span data-ttu-id="1676f-360">要求</span><span class="sxs-lookup"><span data-stu-id="1676f-360">Requirement</span></span>|<span data-ttu-id="1676f-361">值</span><span class="sxs-lookup"><span data-stu-id="1676f-361">Value</span></span>|
|---|---|
|[<span data-ttu-id="1676f-362">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="1676f-362">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="1676f-363">1.0</span><span class="sxs-lookup"><span data-stu-id="1676f-363">1.0</span></span>|
|[<span data-ttu-id="1676f-364">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="1676f-364">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="1676f-365">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1676f-365">ReadItem</span></span>|
|[<span data-ttu-id="1676f-366">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="1676f-366">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="1676f-367">阅读</span><span class="sxs-lookup"><span data-stu-id="1676f-367">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="1676f-368">示例</span><span class="sxs-lookup"><span data-stu-id="1676f-368">Example</span></span>

```js
var dateTimeModified = Office.context.mailbox.item.dateTimeModified;
console.log("Date and time modified: " + dateTimeModified);
```

<br>

---
---

#### <a name="end-datetimejavascriptapioutlookofficetimeviewoutlook-js-18"></a><span data-ttu-id="1676f-369">end: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.8)</span><span class="sxs-lookup"><span data-stu-id="1676f-369">end: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.8)</span></span>

<span data-ttu-id="1676f-370">获取或设置约会结束的日期和时间。</span><span class="sxs-lookup"><span data-stu-id="1676f-370">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="1676f-p113">将 `end` 属性表示为协调世界时 (UTC) 的日期和时间值。可使用 [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) 方法将 end 属性值转换为客户端的本地日期和时间。</span><span class="sxs-lookup"><span data-stu-id="1676f-p113">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="1676f-373">阅读模式</span><span class="sxs-lookup"><span data-stu-id="1676f-373">Read mode</span></span>

<span data-ttu-id="1676f-374">`end` 属性返回 `Date` 对象。</span><span class="sxs-lookup"><span data-stu-id="1676f-374">The `end` property returns a `Date` object.</span></span>

```js
var end = Office.context.mailbox.item.end;
console.log("Appointment end: " + end);
```

##### <a name="compose-mode"></a><span data-ttu-id="1676f-375">撰写模式</span><span class="sxs-lookup"><span data-stu-id="1676f-375">Compose mode</span></span>

<span data-ttu-id="1676f-376">`end` 属性返回 `Time` 对象。</span><span class="sxs-lookup"><span data-stu-id="1676f-376">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="1676f-377">使用 [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.8#setasync-datetime--options--callback-) 方法设置结束时间时，应使用 [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) 方法将客户端的本地时间转换为服务器的 UTC。</span><span class="sxs-lookup"><span data-stu-id="1676f-377">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.8#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="1676f-378">以下示例使用 `Time` 对象的 [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.8#setasync-datetime--options--callback-) 方法设置约会的结束时间。</span><span class="sxs-lookup"><span data-stu-id="1676f-378">The following example sets the end time of an appointment by using the [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.8#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="1676f-379">类型</span><span class="sxs-lookup"><span data-stu-id="1676f-379">Type</span></span>

*   <span data-ttu-id="1676f-380">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.8)</span><span class="sxs-lookup"><span data-stu-id="1676f-380">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.8)</span></span>

##### <a name="requirements"></a><span data-ttu-id="1676f-381">要求</span><span class="sxs-lookup"><span data-stu-id="1676f-381">Requirements</span></span>

|<span data-ttu-id="1676f-382">要求</span><span class="sxs-lookup"><span data-stu-id="1676f-382">Requirement</span></span>|<span data-ttu-id="1676f-383">值</span><span class="sxs-lookup"><span data-stu-id="1676f-383">Value</span></span>|
|---|---|
|[<span data-ttu-id="1676f-384">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="1676f-384">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="1676f-385">1.0</span><span class="sxs-lookup"><span data-stu-id="1676f-385">1.0</span></span>|
|[<span data-ttu-id="1676f-386">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="1676f-386">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="1676f-387">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1676f-387">ReadItem</span></span>|
|[<span data-ttu-id="1676f-388">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="1676f-388">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="1676f-389">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="1676f-389">Compose or Read</span></span>|

<br>

---
---

#### <a name="enhancedlocation-enhancedlocationjavascriptapioutlookofficeenhancedlocationviewoutlook-js-18"></a><span data-ttu-id="1676f-390">enhancedLocation： [enhancedLocation](/javascript/api/outlook/office.enhancedlocation?view=outlook-js-1.8)</span><span class="sxs-lookup"><span data-stu-id="1676f-390">enhancedLocation: [EnhancedLocation](/javascript/api/outlook/office.enhancedlocation?view=outlook-js-1.8)</span></span>

<span data-ttu-id="1676f-391">获取或设置约会的位置。</span><span class="sxs-lookup"><span data-stu-id="1676f-391">Gets or sets the locations of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="1676f-392">阅读模式</span><span class="sxs-lookup"><span data-stu-id="1676f-392">Read mode</span></span>

<span data-ttu-id="1676f-393">该`enhancedLocation`属性返回一个[EnhancedLocation](/javascript/api/outlook/office.enhancedlocation?view=outlook-js-1.8)对象，该对象允许您获取与约会关联的一组位置（每个由[LocationDetails](/javascript/api/outlook/office.locationdetails?view=outlook-js-1.8)对象表示）。</span><span class="sxs-lookup"><span data-stu-id="1676f-393">The `enhancedLocation` property returns an [EnhancedLocation](/javascript/api/outlook/office.enhancedlocation?view=outlook-js-1.8) object that allows you to get the set of locations (each represented by a [LocationDetails](/javascript/api/outlook/office.locationdetails?view=outlook-js-1.8) object) associated with the appointment.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="1676f-394">撰写模式</span><span class="sxs-lookup"><span data-stu-id="1676f-394">Compose mode</span></span>

<span data-ttu-id="1676f-395">该`enhancedLocation`属性返回一个[EnhancedLocation](/javascript/api/outlook/office.enhancedlocation?view=outlook-js-1.8)对象，该对象提供用于获取、删除或添加约会位置的方法。</span><span class="sxs-lookup"><span data-stu-id="1676f-395">The `enhancedLocation` property returns an [EnhancedLocation](/javascript/api/outlook/office.enhancedlocation?view=outlook-js-1.8) object that provides methods to get, remove, or add locations on an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="1676f-396">类型</span><span class="sxs-lookup"><span data-stu-id="1676f-396">Type</span></span>

*   [<span data-ttu-id="1676f-397">EnhancedLocation</span><span class="sxs-lookup"><span data-stu-id="1676f-397">EnhancedLocation</span></span>](/javascript/api/outlook/office.enhancedlocation?view=outlook-js-1.8)

##### <a name="requirements"></a><span data-ttu-id="1676f-398">Requirements</span><span class="sxs-lookup"><span data-stu-id="1676f-398">Requirements</span></span>

|<span data-ttu-id="1676f-399">要求</span><span class="sxs-lookup"><span data-stu-id="1676f-399">Requirement</span></span>|<span data-ttu-id="1676f-400">值</span><span class="sxs-lookup"><span data-stu-id="1676f-400">Value</span></span>|
|---|---|
|[<span data-ttu-id="1676f-401">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="1676f-401">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="1676f-402">1.8</span><span class="sxs-lookup"><span data-stu-id="1676f-402">1.8</span></span>|
|[<span data-ttu-id="1676f-403">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="1676f-403">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="1676f-404">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1676f-404">ReadItem</span></span>|
|[<span data-ttu-id="1676f-405">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="1676f-405">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="1676f-406">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="1676f-406">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="1676f-407">示例</span><span class="sxs-lookup"><span data-stu-id="1676f-407">Example</span></span>

<span data-ttu-id="1676f-408">下面的示例将获取与约会相关联的当前位置。</span><span class="sxs-lookup"><span data-stu-id="1676f-408">The following example gets the current locations associated with the appointment.</span></span>

```js
Office.context.mailbox.item.enhancedLocation.getAsync(callbackFunction);

function callbackFunction(asyncResult) {
  asyncResult.value.forEach(function (place) {
    console.log("Display name: " + place.displayName);
    console.log("Type: " + place.locationIdentifier.type);
    if (place.locationIdentifier.type === Office.MailboxEnums.LocationType.Room) {
      console.log("Email address: " + place.emailAddress);
    }
  });
}
```

<br>

---
---

#### <a name="from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-18fromjavascriptapioutlookofficefromviewoutlook-js-18"></a><span data-ttu-id="1676f-409">发件人： [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8)|[From](/javascript/api/outlook/office.from?view=outlook-js-1.8)</span><span class="sxs-lookup"><span data-stu-id="1676f-409">from: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8)|[From](/javascript/api/outlook/office.from?view=outlook-js-1.8)</span></span>

<span data-ttu-id="1676f-410">获取邮件发件人的电子邮件地址。</span><span class="sxs-lookup"><span data-stu-id="1676f-410">Gets the email address of the sender of a message.</span></span>

<span data-ttu-id="1676f-p114">`from` 和 [`sender`](#sender-emailaddressdetails) 属性表示同一个人，邮件由代理人发送的除外。在此情况下，`from` 属性表示代理程序，而 sender 属性表示代理人。</span><span class="sxs-lookup"><span data-stu-id="1676f-p114">The `from` and [`sender`](#sender-emailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="1676f-413">`from` 属性中 `EmailAddressDetails` 对象的 `recipientType` 属性为 `undefined`。</span><span class="sxs-lookup"><span data-stu-id="1676f-413">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="1676f-414">阅读模式</span><span class="sxs-lookup"><span data-stu-id="1676f-414">Read mode</span></span>

<span data-ttu-id="1676f-415">`from`属性返回一个`EmailAddressDetails`对象。</span><span class="sxs-lookup"><span data-stu-id="1676f-415">The `from` property returns an `EmailAddressDetails` object.</span></span>

```js
var from = Office.context.mailbox.item.from;
console.log("From " + from);
```

##### <a name="compose-mode"></a><span data-ttu-id="1676f-416">撰写模式</span><span class="sxs-lookup"><span data-stu-id="1676f-416">Compose mode</span></span>

<span data-ttu-id="1676f-417">`from`属性返回一个`From`对象，该对象提供用于获取 "起始" 值的方法。</span><span class="sxs-lookup"><span data-stu-id="1676f-417">The `from` property returns a `From` object that provides a method to get the from value.</span></span>

```js
Office.context.mailbox.item.from.getAsync(callback);

function callback(asyncResult) {
  var from = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="1676f-418">类型</span><span class="sxs-lookup"><span data-stu-id="1676f-418">Type</span></span>

*   <span data-ttu-id="1676f-419">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8) | [](/javascript/api/outlook/office.from?view=outlook-js-1.8)</span><span class="sxs-lookup"><span data-stu-id="1676f-419">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8) | [From](/javascript/api/outlook/office.from?view=outlook-js-1.8)</span></span>

##### <a name="requirements"></a><span data-ttu-id="1676f-420">Requirements</span><span class="sxs-lookup"><span data-stu-id="1676f-420">Requirements</span></span>

|<span data-ttu-id="1676f-421">要求</span><span class="sxs-lookup"><span data-stu-id="1676f-421">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="1676f-422">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="1676f-422">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="1676f-423">1.0</span><span class="sxs-lookup"><span data-stu-id="1676f-423">1.0</span></span>|<span data-ttu-id="1676f-424">1.7</span><span class="sxs-lookup"><span data-stu-id="1676f-424">1.7</span></span>|
|[<span data-ttu-id="1676f-425">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="1676f-425">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="1676f-426">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1676f-426">ReadItem</span></span>|<span data-ttu-id="1676f-427">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="1676f-427">ReadWriteItem</span></span>|
|[<span data-ttu-id="1676f-428">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="1676f-428">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="1676f-429">阅读</span><span class="sxs-lookup"><span data-stu-id="1676f-429">Read</span></span>|<span data-ttu-id="1676f-430">撰写</span><span class="sxs-lookup"><span data-stu-id="1676f-430">Compose</span></span>|

<br>

---
---

#### <a name="internetheaders-internetheadersjavascriptapioutlookofficeinternetheadersviewoutlook-js-18"></a><span data-ttu-id="1676f-431">internetHeaders： [internetHeaders](/javascript/api/outlook/office.internetheaders?view=outlook-js-1.8)</span><span class="sxs-lookup"><span data-stu-id="1676f-431">internetHeaders: [InternetHeaders](/javascript/api/outlook/office.internetheaders?view=outlook-js-1.8)</span></span>

<span data-ttu-id="1676f-432">获取或设置邮件的自定义 internet 邮件头。</span><span class="sxs-lookup"><span data-stu-id="1676f-432">Gets or sets custom internet headers on a message.</span></span> <span data-ttu-id="1676f-433">仅限撰写模式。</span><span class="sxs-lookup"><span data-stu-id="1676f-433">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="1676f-434">类型</span><span class="sxs-lookup"><span data-stu-id="1676f-434">Type</span></span>

*   [<span data-ttu-id="1676f-435">InternetHeaders</span><span class="sxs-lookup"><span data-stu-id="1676f-435">InternetHeaders</span></span>](/javascript/api/outlook/office.internetheaders?view=outlook-js-1.8)

##### <a name="requirements"></a><span data-ttu-id="1676f-436">Requirements</span><span class="sxs-lookup"><span data-stu-id="1676f-436">Requirements</span></span>

|<span data-ttu-id="1676f-437">要求</span><span class="sxs-lookup"><span data-stu-id="1676f-437">Requirement</span></span>|<span data-ttu-id="1676f-438">值</span><span class="sxs-lookup"><span data-stu-id="1676f-438">Value</span></span>|
|---|---|
|[<span data-ttu-id="1676f-439">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="1676f-439">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="1676f-440">1.8</span><span class="sxs-lookup"><span data-stu-id="1676f-440">1.8</span></span>|
|[<span data-ttu-id="1676f-441">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="1676f-441">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="1676f-442">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1676f-442">ReadItem</span></span>|
|[<span data-ttu-id="1676f-443">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="1676f-443">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="1676f-444">撰写</span><span class="sxs-lookup"><span data-stu-id="1676f-444">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="1676f-445">示例</span><span class="sxs-lookup"><span data-stu-id="1676f-445">Example</span></span>

```js
Office.context.mailbox.item.internetHeaders.getAsync(["header1", "header2"], callback);

function callback(asyncResult) {
  var dictionary = asyncResult.value;
  var header1_value = dictionary["header1"];
}
```

<br>

---
---

#### <a name="internetmessageid-string"></a><span data-ttu-id="1676f-446">internetMessageId: String</span><span class="sxs-lookup"><span data-stu-id="1676f-446">internetMessageId: String</span></span>

<span data-ttu-id="1676f-p116">获取电子邮件的 Internet 消息标识符。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="1676f-p116">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="1676f-449">Type</span><span class="sxs-lookup"><span data-stu-id="1676f-449">Type</span></span>

*   <span data-ttu-id="1676f-450">String</span><span class="sxs-lookup"><span data-stu-id="1676f-450">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="1676f-451">Requirements</span><span class="sxs-lookup"><span data-stu-id="1676f-451">Requirements</span></span>

|<span data-ttu-id="1676f-452">要求</span><span class="sxs-lookup"><span data-stu-id="1676f-452">Requirement</span></span>|<span data-ttu-id="1676f-453">值</span><span class="sxs-lookup"><span data-stu-id="1676f-453">Value</span></span>|
|---|---|
|[<span data-ttu-id="1676f-454">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="1676f-454">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="1676f-455">1.0</span><span class="sxs-lookup"><span data-stu-id="1676f-455">1.0</span></span>|
|[<span data-ttu-id="1676f-456">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="1676f-456">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="1676f-457">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1676f-457">ReadItem</span></span>|
|[<span data-ttu-id="1676f-458">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="1676f-458">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="1676f-459">阅读</span><span class="sxs-lookup"><span data-stu-id="1676f-459">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="1676f-460">示例</span><span class="sxs-lookup"><span data-stu-id="1676f-460">Example</span></span>

```js
var internetMessageId = Office.context.mailbox.item.internetMessageId;
console.log("internetMessageId: " + internetMessageId);
```

<br>

---
---

#### <a name="itemclass-string"></a><span data-ttu-id="1676f-461">itemClass: String</span><span class="sxs-lookup"><span data-stu-id="1676f-461">itemClass: String</span></span>

<span data-ttu-id="1676f-p117">获取选定项目的 Exchange Web 服务项目类。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="1676f-p117">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="1676f-p118">`itemClass` 属性指定所选项目的邮件类别。以下是邮件或约会项目的默认邮件类别。</span><span class="sxs-lookup"><span data-stu-id="1676f-p118">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

|<span data-ttu-id="1676f-466">类型</span><span class="sxs-lookup"><span data-stu-id="1676f-466">Type</span></span>|<span data-ttu-id="1676f-467">说明</span><span class="sxs-lookup"><span data-stu-id="1676f-467">Description</span></span>|<span data-ttu-id="1676f-468">项目类</span><span class="sxs-lookup"><span data-stu-id="1676f-468">item class</span></span>|
|---|---|---|
|<span data-ttu-id="1676f-469">约会项目</span><span class="sxs-lookup"><span data-stu-id="1676f-469">Appointment items</span></span>|<span data-ttu-id="1676f-470">这些是项目类 `IPM.Appointment` 或 `IPM.Appointment.Occurrence` 的日历项目。</span><span class="sxs-lookup"><span data-stu-id="1676f-470">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurrence`.</span></span>|`IPM.Appointment`<br />`IPM.Appointment.Occurrence`|
|<span data-ttu-id="1676f-471">邮件项目</span><span class="sxs-lookup"><span data-stu-id="1676f-471">Message items</span></span>|<span data-ttu-id="1676f-472">这些项目包括具有默认邮件类别 `IPM.Note` 的电子邮件，以及使用 `IPM.Schedule.Meeting` 作为基础邮件类别的会议请求、响应和取消。</span><span class="sxs-lookup"><span data-stu-id="1676f-472">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span>|`IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled`|

<span data-ttu-id="1676f-473">你可以创建用于扩展默认邮件类别的自定义邮件类别，例如，自定义约会邮件类别 `IPM.Appointment.Contoso`。</span><span class="sxs-lookup"><span data-stu-id="1676f-473">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="1676f-474">Type</span><span class="sxs-lookup"><span data-stu-id="1676f-474">Type</span></span>

*   <span data-ttu-id="1676f-475">String</span><span class="sxs-lookup"><span data-stu-id="1676f-475">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="1676f-476">Requirements</span><span class="sxs-lookup"><span data-stu-id="1676f-476">Requirements</span></span>

|<span data-ttu-id="1676f-477">要求</span><span class="sxs-lookup"><span data-stu-id="1676f-477">Requirement</span></span>|<span data-ttu-id="1676f-478">值</span><span class="sxs-lookup"><span data-stu-id="1676f-478">Value</span></span>|
|---|---|
|[<span data-ttu-id="1676f-479">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="1676f-479">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="1676f-480">1.0</span><span class="sxs-lookup"><span data-stu-id="1676f-480">1.0</span></span>|
|[<span data-ttu-id="1676f-481">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="1676f-481">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="1676f-482">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1676f-482">ReadItem</span></span>|
|[<span data-ttu-id="1676f-483">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="1676f-483">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="1676f-484">阅读</span><span class="sxs-lookup"><span data-stu-id="1676f-484">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="1676f-485">示例</span><span class="sxs-lookup"><span data-stu-id="1676f-485">Example</span></span>

```js
var itemClass = Office.context.mailbox.item.itemClass;
console.log("Item class: " + itemClass);
```

<br>

---
---

#### <a name="nullable-itemid-string"></a><span data-ttu-id="1676f-486">(nullable) itemId: String</span><span class="sxs-lookup"><span data-stu-id="1676f-486">(nullable) itemId: String</span></span>

<span data-ttu-id="1676f-487">获取当前项的[Exchange Web 服务项标识符](/exchange/client-developer/exchange-web-services/ews-identifiers-in-exchange)。</span><span class="sxs-lookup"><span data-stu-id="1676f-487">Gets the [Exchange Web Services item identifier](/exchange/client-developer/exchange-web-services/ews-identifiers-in-exchange) for the current item.</span></span> <span data-ttu-id="1676f-488">仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="1676f-488">Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="1676f-489">`itemId`属性返回的标识符与[Exchange Web 服务项目标识符](/exchange/client-developer/exchange-web-services/ews-identifiers-in-exchange)相同。</span><span class="sxs-lookup"><span data-stu-id="1676f-489">The identifier returned by the `itemId` property is the same as the [Exchange Web Services item identifier](/exchange/client-developer/exchange-web-services/ews-identifiers-in-exchange).</span></span> <span data-ttu-id="1676f-490">`itemId` 属性与 Outlook 条目 ID 或 Outlook REST API 使用的 ID 不同。</span><span class="sxs-lookup"><span data-stu-id="1676f-490">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="1676f-491">使用此值进行 REST API 调用前，应使用 [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string) 对它进行转换。</span><span class="sxs-lookup"><span data-stu-id="1676f-491">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="1676f-492">有关详细信息，请参阅[从 Outlook 加载项使用 Outlook REST API](/outlook/add-ins/use-rest-api#get-the-item-id)。</span><span class="sxs-lookup"><span data-stu-id="1676f-492">For more details, see [Use the Outlook REST APIs from an Outlook add-in](/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

<span data-ttu-id="1676f-p121">`itemId` 属性在撰写模式下不可用。如果需要项目标识符，[`saveAsync`](#saveasyncoptions-callback) 方法可用于将项目保存到存储，这将在回调函数的 [`AsyncResult.value`](/javascript/api/office/office.asyncresult) 参数中返回项目标识符。</span><span class="sxs-lookup"><span data-stu-id="1676f-p121">The `itemId` property is not available in compose mode. If an item identifier is required, the [`saveAsync`](#saveasyncoptions-callback) method can be used to save the item to the store, which will return the item identifier in the [`AsyncResult.value`](/javascript/api/office/office.asyncresult) parameter in the callback function.</span></span>

##### <a name="type"></a><span data-ttu-id="1676f-495">Type</span><span class="sxs-lookup"><span data-stu-id="1676f-495">Type</span></span>

*   <span data-ttu-id="1676f-496">String</span><span class="sxs-lookup"><span data-stu-id="1676f-496">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="1676f-497">Requirements</span><span class="sxs-lookup"><span data-stu-id="1676f-497">Requirements</span></span>

|<span data-ttu-id="1676f-498">要求</span><span class="sxs-lookup"><span data-stu-id="1676f-498">Requirement</span></span>|<span data-ttu-id="1676f-499">值</span><span class="sxs-lookup"><span data-stu-id="1676f-499">Value</span></span>|
|---|---|
|[<span data-ttu-id="1676f-500">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="1676f-500">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="1676f-501">1.0</span><span class="sxs-lookup"><span data-stu-id="1676f-501">1.0</span></span>|
|[<span data-ttu-id="1676f-502">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="1676f-502">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="1676f-503">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1676f-503">ReadItem</span></span>|
|[<span data-ttu-id="1676f-504">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="1676f-504">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="1676f-505">阅读</span><span class="sxs-lookup"><span data-stu-id="1676f-505">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="1676f-506">示例</span><span class="sxs-lookup"><span data-stu-id="1676f-506">Example</span></span>

<span data-ttu-id="1676f-p122">以下代码检查项目标识符是否存在。如果 `itemId` 属性返回 `null` 或 `undefined`，则将项目保存到存储，并从异步结果中获取项目标识符。</span><span class="sxs-lookup"><span data-stu-id="1676f-p122">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

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

#### <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtypeviewoutlook-js-18"></a><span data-ttu-id="1676f-509">itemType: [Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.8)</span><span class="sxs-lookup"><span data-stu-id="1676f-509">itemType: [Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.8)</span></span>

<span data-ttu-id="1676f-510">获取实例表示的项的类型。</span><span class="sxs-lookup"><span data-stu-id="1676f-510">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="1676f-511">`itemType` 属性返回其中一个 `ItemType` 枚举值，指示 `item` 对象实例是邮件还是约会。</span><span class="sxs-lookup"><span data-stu-id="1676f-511">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="1676f-512">类型</span><span class="sxs-lookup"><span data-stu-id="1676f-512">Type</span></span>

*   [<span data-ttu-id="1676f-513">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="1676f-513">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.8)

##### <a name="requirements"></a><span data-ttu-id="1676f-514">Requirements</span><span class="sxs-lookup"><span data-stu-id="1676f-514">Requirements</span></span>

|<span data-ttu-id="1676f-515">要求</span><span class="sxs-lookup"><span data-stu-id="1676f-515">Requirement</span></span>|<span data-ttu-id="1676f-516">值</span><span class="sxs-lookup"><span data-stu-id="1676f-516">Value</span></span>|
|---|---|
|[<span data-ttu-id="1676f-517">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="1676f-517">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="1676f-518">1.0</span><span class="sxs-lookup"><span data-stu-id="1676f-518">1.0</span></span>|
|[<span data-ttu-id="1676f-519">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="1676f-519">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="1676f-520">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1676f-520">ReadItem</span></span>|
|[<span data-ttu-id="1676f-521">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="1676f-521">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="1676f-522">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="1676f-522">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="1676f-523">示例</span><span class="sxs-lookup"><span data-stu-id="1676f-523">Example</span></span>

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

#### <a name="location-stringlocationjavascriptapioutlookofficelocationviewoutlook-js-18"></a><span data-ttu-id="1676f-524">location: String|[Location](/javascript/api/outlook/office.location?view=outlook-js-1.8)</span><span class="sxs-lookup"><span data-stu-id="1676f-524">location: String|[Location](/javascript/api/outlook/office.location?view=outlook-js-1.8)</span></span>

<span data-ttu-id="1676f-525">获取或设置约会的位置。</span><span class="sxs-lookup"><span data-stu-id="1676f-525">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="1676f-526">阅读模式</span><span class="sxs-lookup"><span data-stu-id="1676f-526">Read mode</span></span>

<span data-ttu-id="1676f-527">`location` 属性返回一个包含约会位置的字符串。</span><span class="sxs-lookup"><span data-stu-id="1676f-527">The `location` property returns a string that contains the location of the appointment.</span></span>

```js
var location = Office.context.mailbox.item.location;
console.log("location: " + location);
```

##### <a name="compose-mode"></a><span data-ttu-id="1676f-528">撰写模式</span><span class="sxs-lookup"><span data-stu-id="1676f-528">Compose mode</span></span>

<span data-ttu-id="1676f-529">`location` 属性返回一个 `Location` 对象，该对象提供用于获取和设置约会位置的方法。</span><span class="sxs-lookup"><span data-stu-id="1676f-529">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

```js
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="1676f-530">类型</span><span class="sxs-lookup"><span data-stu-id="1676f-530">Type</span></span>

*   <span data-ttu-id="1676f-531">String | [Location](/javascript/api/outlook/office.location?view=outlook-js-1.8)</span><span class="sxs-lookup"><span data-stu-id="1676f-531">String | [Location](/javascript/api/outlook/office.location?view=outlook-js-1.8)</span></span>

##### <a name="requirements"></a><span data-ttu-id="1676f-532">Requirements</span><span class="sxs-lookup"><span data-stu-id="1676f-532">Requirements</span></span>

|<span data-ttu-id="1676f-533">要求</span><span class="sxs-lookup"><span data-stu-id="1676f-533">Requirement</span></span>|<span data-ttu-id="1676f-534">值</span><span class="sxs-lookup"><span data-stu-id="1676f-534">Value</span></span>|
|---|---|
|[<span data-ttu-id="1676f-535">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="1676f-535">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="1676f-536">1.0</span><span class="sxs-lookup"><span data-stu-id="1676f-536">1.0</span></span>|
|[<span data-ttu-id="1676f-537">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="1676f-537">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="1676f-538">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1676f-538">ReadItem</span></span>|
|[<span data-ttu-id="1676f-539">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="1676f-539">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="1676f-540">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="1676f-540">Compose or Read</span></span>|

<br>

---
---

#### <a name="normalizedsubject-string"></a><span data-ttu-id="1676f-541">normalizedSubject: String</span><span class="sxs-lookup"><span data-stu-id="1676f-541">normalizedSubject: String</span></span>

<span data-ttu-id="1676f-p123">获取删除了所有前缀（包括 `RE:` 和 `FWD:`）的项目主题。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="1676f-p123">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="1676f-p124">normalizedSubject 属性获取包含由电子邮件程序添加的任何标准前缀（如 `RE:` 和 `FW:`）的项目主题。若要获取包含完整前缀的项目主题，请使用 [`subject`](#subject-stringsubject) 属性。</span><span class="sxs-lookup"><span data-stu-id="1676f-p124">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="1676f-546">Type</span><span class="sxs-lookup"><span data-stu-id="1676f-546">Type</span></span>

*   <span data-ttu-id="1676f-547">String</span><span class="sxs-lookup"><span data-stu-id="1676f-547">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="1676f-548">Requirements</span><span class="sxs-lookup"><span data-stu-id="1676f-548">Requirements</span></span>

|<span data-ttu-id="1676f-549">要求</span><span class="sxs-lookup"><span data-stu-id="1676f-549">Requirement</span></span>|<span data-ttu-id="1676f-550">值</span><span class="sxs-lookup"><span data-stu-id="1676f-550">Value</span></span>|
|---|---|
|[<span data-ttu-id="1676f-551">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="1676f-551">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="1676f-552">1.0</span><span class="sxs-lookup"><span data-stu-id="1676f-552">1.0</span></span>|
|[<span data-ttu-id="1676f-553">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="1676f-553">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="1676f-554">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1676f-554">ReadItem</span></span>|
|[<span data-ttu-id="1676f-555">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="1676f-555">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="1676f-556">阅读</span><span class="sxs-lookup"><span data-stu-id="1676f-556">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="1676f-557">示例</span><span class="sxs-lookup"><span data-stu-id="1676f-557">Example</span></span>

```js
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
console.log("Normalized subject: " + normalizedSubject);
```

<br>

---
---

#### <a name="notificationmessages-notificationmessagesjavascriptapioutlookofficenotificationmessagesviewoutlook-js-18"></a><span data-ttu-id="1676f-558">notificationMessages: [NotificationMessages](/javascript/api/outlook/office.notificationmessages?view=outlook-js-1.8)</span><span class="sxs-lookup"><span data-stu-id="1676f-558">notificationMessages: [NotificationMessages](/javascript/api/outlook/office.notificationmessages?view=outlook-js-1.8)</span></span>

<span data-ttu-id="1676f-559">获取项目的通知邮件。</span><span class="sxs-lookup"><span data-stu-id="1676f-559">Gets the notification messages for an item.</span></span>

##### <a name="type"></a><span data-ttu-id="1676f-560">类型</span><span class="sxs-lookup"><span data-stu-id="1676f-560">Type</span></span>

*   [<span data-ttu-id="1676f-561">NotificationMessages</span><span class="sxs-lookup"><span data-stu-id="1676f-561">NotificationMessages</span></span>](/javascript/api/outlook/office.notificationmessages?view=outlook-js-1.8)

##### <a name="requirements"></a><span data-ttu-id="1676f-562">要求</span><span class="sxs-lookup"><span data-stu-id="1676f-562">Requirements</span></span>

|<span data-ttu-id="1676f-563">要求</span><span class="sxs-lookup"><span data-stu-id="1676f-563">Requirement</span></span>|<span data-ttu-id="1676f-564">值</span><span class="sxs-lookup"><span data-stu-id="1676f-564">Value</span></span>|
|---|---|
|[<span data-ttu-id="1676f-565">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="1676f-565">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="1676f-566">1.3</span><span class="sxs-lookup"><span data-stu-id="1676f-566">1.3</span></span>|
|[<span data-ttu-id="1676f-567">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="1676f-567">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="1676f-568">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1676f-568">ReadItem</span></span>|
|[<span data-ttu-id="1676f-569">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="1676f-569">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="1676f-570">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="1676f-570">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="1676f-571">示例</span><span class="sxs-lookup"><span data-stu-id="1676f-571">Example</span></span>

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

#### <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-18recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-18"></a><span data-ttu-id="1676f-572">optionalAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.8)</span><span class="sxs-lookup"><span data-stu-id="1676f-572">optionalAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.8)</span></span>

<span data-ttu-id="1676f-573">提供对事件的可选与会者的访问权限。</span><span class="sxs-lookup"><span data-stu-id="1676f-573">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="1676f-574">对象的类型和访问级别取决于当前项目的模式。</span><span class="sxs-lookup"><span data-stu-id="1676f-574">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="1676f-575">阅读模式</span><span class="sxs-lookup"><span data-stu-id="1676f-575">Read mode</span></span>

<span data-ttu-id="1676f-576">`optionalAttendees` 属性返回一个数组，其中包含每个可选与会者的 `EmailAddressDetails` 对象。</span><span class="sxs-lookup"><span data-stu-id="1676f-576">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span> <span data-ttu-id="1676f-577">默认情况下，集合上限为 100 个成员。</span><span class="sxs-lookup"><span data-stu-id="1676f-577">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="1676f-578">但是，在 Windows 和 Mac 上，最多可包含 500 个成员。</span><span class="sxs-lookup"><span data-stu-id="1676f-578">However, on Windows and Mac, you can get 500 members maximum.</span></span>

```js
var optionalAttendees = Office.context.mailbox.item.optionalAttendees;
console.log("Optional attendees: " + JSON.stringify(optionalAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="1676f-579">撰写模式</span><span class="sxs-lookup"><span data-stu-id="1676f-579">Compose mode</span></span>

<span data-ttu-id="1676f-580">`optionalAttendees` 属性返回一个 `Recipients` 对象，该对象提供用于获取或更新可选与会者的方法。</span><span class="sxs-lookup"><span data-stu-id="1676f-580">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span> <span data-ttu-id="1676f-581">默认情况下，集合上限为 100 个成员。</span><span class="sxs-lookup"><span data-stu-id="1676f-581">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="1676f-582">但是，在 Windows 和 Mac 上，以下限制适用。</span><span class="sxs-lookup"><span data-stu-id="1676f-582">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="1676f-583">最多包含 500 个成员。</span><span class="sxs-lookup"><span data-stu-id="1676f-583">Get 500 members maximum.</span></span>
- <span data-ttu-id="1676f-584">为每个呼叫最多设置 100 个成员，总共多达 500 个成员。</span><span class="sxs-lookup"><span data-stu-id="1676f-584">Set a maximum of 100 members per call, up to 500 members total.</span></span>

```js
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="1676f-585">类型</span><span class="sxs-lookup"><span data-stu-id="1676f-585">Type</span></span>

*   <span data-ttu-id="1676f-586">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.8)</span><span class="sxs-lookup"><span data-stu-id="1676f-586">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.8)</span></span>

##### <a name="requirements"></a><span data-ttu-id="1676f-587">Requirements</span><span class="sxs-lookup"><span data-stu-id="1676f-587">Requirements</span></span>

|<span data-ttu-id="1676f-588">要求</span><span class="sxs-lookup"><span data-stu-id="1676f-588">Requirement</span></span>|<span data-ttu-id="1676f-589">值</span><span class="sxs-lookup"><span data-stu-id="1676f-589">Value</span></span>|
|---|---|
|[<span data-ttu-id="1676f-590">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="1676f-590">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="1676f-591">1.0</span><span class="sxs-lookup"><span data-stu-id="1676f-591">1.0</span></span>|
|[<span data-ttu-id="1676f-592">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="1676f-592">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="1676f-593">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1676f-593">ReadItem</span></span>|
|[<span data-ttu-id="1676f-594">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="1676f-594">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="1676f-595">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="1676f-595">Compose or Read</span></span>|

<br>

---
---

#### <a name="organizer-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-18organizerjavascriptapioutlookofficeorganizerviewoutlook-js-18"></a><span data-ttu-id="1676f-596">组织者： [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8)|[组织者](/javascript/api/outlook/office.organizer?view=outlook-js-1.8)</span><span class="sxs-lookup"><span data-stu-id="1676f-596">organizer: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8)|[Organizer](/javascript/api/outlook/office.organizer?view=outlook-js-1.8)</span></span>

<span data-ttu-id="1676f-597">获取指定会议的组织者的电子邮件地址。</span><span class="sxs-lookup"><span data-stu-id="1676f-597">Gets the email address of the organizer for a specified meeting.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="1676f-598">阅读模式</span><span class="sxs-lookup"><span data-stu-id="1676f-598">Read mode</span></span>

<span data-ttu-id="1676f-599">该`organizer`属性返回一个[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8)对象，该对象代表会议组织者。</span><span class="sxs-lookup"><span data-stu-id="1676f-599">The `organizer` property returns an [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8) object that represents the meeting organizer.</span></span>

```js
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
console.log("Organizer: " + organizerName + " (" + organizerAddress + ")");
```

##### <a name="compose-mode"></a><span data-ttu-id="1676f-600">撰写模式</span><span class="sxs-lookup"><span data-stu-id="1676f-600">Compose mode</span></span>

<span data-ttu-id="1676f-601">该`organizer`属性返回一个[管理](/javascript/api/outlook/office.organizer?view=outlook-js-1.8)器对象，该对象提供获取组织者值的方法。</span><span class="sxs-lookup"><span data-stu-id="1676f-601">The `organizer` property returns an [Organizer](/javascript/api/outlook/office.organizer?view=outlook-js-1.8) object that provides a method to get the organizer value.</span></span>

```js
Office.context.mailbox.item.organizer.getAsync(
  function(asyncResult) {
    console.log(JSON.stringify(asyncResult));
  }
);
```

##### <a name="type"></a><span data-ttu-id="1676f-602">类型</span><span class="sxs-lookup"><span data-stu-id="1676f-602">Type</span></span>

*   <span data-ttu-id="1676f-603">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8) | [组织者](/javascript/api/outlook/office.organizer?view=outlook-js-1.8)</span><span class="sxs-lookup"><span data-stu-id="1676f-603">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8) | [Organizer](/javascript/api/outlook/office.organizer?view=outlook-js-1.8)</span></span>

##### <a name="requirements"></a><span data-ttu-id="1676f-604">Requirements</span><span class="sxs-lookup"><span data-stu-id="1676f-604">Requirements</span></span>

|<span data-ttu-id="1676f-605">要求</span><span class="sxs-lookup"><span data-stu-id="1676f-605">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="1676f-606">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="1676f-606">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="1676f-607">1.0</span><span class="sxs-lookup"><span data-stu-id="1676f-607">1.0</span></span>|<span data-ttu-id="1676f-608">1.7</span><span class="sxs-lookup"><span data-stu-id="1676f-608">1.7</span></span>|
|[<span data-ttu-id="1676f-609">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="1676f-609">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="1676f-610">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1676f-610">ReadItem</span></span>|<span data-ttu-id="1676f-611">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="1676f-611">ReadWriteItem</span></span>|
|[<span data-ttu-id="1676f-612">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="1676f-612">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="1676f-613">阅读</span><span class="sxs-lookup"><span data-stu-id="1676f-613">Read</span></span>|<span data-ttu-id="1676f-614">撰写</span><span class="sxs-lookup"><span data-stu-id="1676f-614">Compose</span></span>|

<br>

---
---

#### <a name="nullable-recurrence-recurrencejavascriptapioutlookofficerecurrenceviewoutlook-js-18"></a><span data-ttu-id="1676f-615">（可以为 null）定期：[定期](/javascript/api/outlook/office.recurrence?view=outlook-js-1.8)</span><span class="sxs-lookup"><span data-stu-id="1676f-615">(nullable) recurrence: [Recurrence](/javascript/api/outlook/office.recurrence?view=outlook-js-1.8)</span></span>

<span data-ttu-id="1676f-616">获取或设置约会的定期模式。</span><span class="sxs-lookup"><span data-stu-id="1676f-616">Gets or sets the recurrence pattern of an appointment.</span></span> <span data-ttu-id="1676f-617">获取会议请求的定期模式。</span><span class="sxs-lookup"><span data-stu-id="1676f-617">Gets the recurrence pattern of a meeting request.</span></span> <span data-ttu-id="1676f-618">约会项目的阅读和撰写模式。</span><span class="sxs-lookup"><span data-stu-id="1676f-618">Read and compose modes for appointment items.</span></span> <span data-ttu-id="1676f-619">会议请求项目的阅读模式。</span><span class="sxs-lookup"><span data-stu-id="1676f-619">Read mode for meeting request items.</span></span>

<span data-ttu-id="1676f-620">如果`recurrence`项目是系列中的一个系列或一个实例，则该属性返回定期约会或会议请求的[定期](/javascript/api/outlook/office.recurrence?view=outlook-js-1.8)对象。</span><span class="sxs-lookup"><span data-stu-id="1676f-620">The `recurrence` property returns a [recurrence](/javascript/api/outlook/office.recurrence?view=outlook-js-1.8) object for recurring appointments or meetings requests if an item is a series or an instance in a series.</span></span> <span data-ttu-id="1676f-621">`null`返回单个约会的单个约会和会议请求。</span><span class="sxs-lookup"><span data-stu-id="1676f-621">`null` is returned for single appointments and meeting requests of single appointments.</span></span> <span data-ttu-id="1676f-622">`undefined`对于不是会议请求的邮件，将返回。</span><span class="sxs-lookup"><span data-stu-id="1676f-622">`undefined` is returned for messages that are not meeting requests.</span></span>

> <span data-ttu-id="1676f-623">注意：会议请求的`itemClass`值为 IPM。Schedule. 会议请求。</span><span class="sxs-lookup"><span data-stu-id="1676f-623">Note: Meeting requests have an `itemClass` value of IPM.Schedule.Meeting.Request.</span></span>

> <span data-ttu-id="1676f-624">注意：如果定期对象为`null`，则表示该对象是单个约会的单个约会或会议请求，而不是某个系列的一部分。</span><span class="sxs-lookup"><span data-stu-id="1676f-624">Note: If the recurrence object is `null`, this indicates that the object is a single appointment or a meeting request of a single appointment and NOT a part of a series.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="1676f-625">阅读模式</span><span class="sxs-lookup"><span data-stu-id="1676f-625">Read mode</span></span>

<span data-ttu-id="1676f-626">该`recurrence`属性返回一个代表约会定期的[定期](/javascript/api/outlook/office.recurrence?view=outlook-js-1.8)对象。</span><span class="sxs-lookup"><span data-stu-id="1676f-626">The `recurrence` property returns a [Recurrence](/javascript/api/outlook/office.recurrence?view=outlook-js-1.8) object that represents the appointment recurrence.</span></span> <span data-ttu-id="1676f-627">此功能适用于约会和会议请求。</span><span class="sxs-lookup"><span data-stu-id="1676f-627">This is available for appointments and meeting requests.</span></span>

```js
var recurrence = Office.context.mailbox.item.recurrence;
console.log("Recurrence: " + JSON.stringify(recurrence));
```

##### <a name="compose-mode"></a><span data-ttu-id="1676f-628">撰写模式</span><span class="sxs-lookup"><span data-stu-id="1676f-628">Compose mode</span></span>

<span data-ttu-id="1676f-629">该`recurrence`属性返回一个[定期](/javascript/api/outlook/office.recurrence?view=outlook-js-1.8)对象，该对象提供用于管理约会周期的方法。</span><span class="sxs-lookup"><span data-stu-id="1676f-629">The `recurrence` property returns a [Recurrence](/javascript/api/outlook/office.recurrence?view=outlook-js-1.8) object that provides methods to manage the appointment recurrence.</span></span> <span data-ttu-id="1676f-630">这可用于约会。</span><span class="sxs-lookup"><span data-stu-id="1676f-630">This is available for appointments.</span></span>

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

##### <a name="type"></a><span data-ttu-id="1676f-631">类型</span><span class="sxs-lookup"><span data-stu-id="1676f-631">Type</span></span>

* [<span data-ttu-id="1676f-632">循环</span><span class="sxs-lookup"><span data-stu-id="1676f-632">Recurrence</span></span>](/javascript/api/outlook/office.recurrence?view=outlook-js-1.8)

|<span data-ttu-id="1676f-633">要求</span><span class="sxs-lookup"><span data-stu-id="1676f-633">Requirement</span></span>|<span data-ttu-id="1676f-634">值</span><span class="sxs-lookup"><span data-stu-id="1676f-634">Value</span></span>|
|---|---|
|[<span data-ttu-id="1676f-635">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="1676f-635">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="1676f-636">1.7</span><span class="sxs-lookup"><span data-stu-id="1676f-636">1.7</span></span>|
|[<span data-ttu-id="1676f-637">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="1676f-637">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="1676f-638">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1676f-638">ReadItem</span></span>|
|[<span data-ttu-id="1676f-639">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="1676f-639">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="1676f-640">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="1676f-640">Compose or Read</span></span>|

<br>

---
---

#### <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-18recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-18"></a><span data-ttu-id="1676f-641">requiredAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.8)</span><span class="sxs-lookup"><span data-stu-id="1676f-641">requiredAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.8)</span></span>

<span data-ttu-id="1676f-642">提供对事件的必需与会者的访问权限。</span><span class="sxs-lookup"><span data-stu-id="1676f-642">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="1676f-643">对象的类型和访问级别取决于当前项目的模式。</span><span class="sxs-lookup"><span data-stu-id="1676f-643">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="1676f-644">阅读模式</span><span class="sxs-lookup"><span data-stu-id="1676f-644">Read mode</span></span>

<span data-ttu-id="1676f-645">`requiredAttendees` 属性返回一个数组，其中包含每个必需与会者的 `EmailAddressDetails` 对象。</span><span class="sxs-lookup"><span data-stu-id="1676f-645">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span> <span data-ttu-id="1676f-646">默认情况下，集合上限为 100 个成员。</span><span class="sxs-lookup"><span data-stu-id="1676f-646">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="1676f-647">但是，在 Windows 和 Mac 上，最多可包含 500 个成员。</span><span class="sxs-lookup"><span data-stu-id="1676f-647">However, on Windows and Mac, you can get 500 members maximum.</span></span>

```js
var requiredAttendees = Office.context.mailbox.item.requiredAttendees;
console.log("Required attendees: " + JSON.stringify(requiredAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="1676f-648">撰写模式</span><span class="sxs-lookup"><span data-stu-id="1676f-648">Compose mode</span></span>

<span data-ttu-id="1676f-649">`requiredAttendees` 属性返回一个 `Recipients` 对象，该对象提供用于获取或更新必需与会者的方法。</span><span class="sxs-lookup"><span data-stu-id="1676f-649">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span> <span data-ttu-id="1676f-650">默认情况下，集合上限为 100 个成员。</span><span class="sxs-lookup"><span data-stu-id="1676f-650">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="1676f-651">但是，在 Windows 和 Mac 上，以下限制适用。</span><span class="sxs-lookup"><span data-stu-id="1676f-651">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="1676f-652">最多包含 500 个成员。</span><span class="sxs-lookup"><span data-stu-id="1676f-652">Get 500 members maximum.</span></span>
- <span data-ttu-id="1676f-653">为每个呼叫最多设置 100 个成员，总共多达 500 个成员。</span><span class="sxs-lookup"><span data-stu-id="1676f-653">Set a maximum of 100 members per call, up to 500 members total.</span></span>

```js
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
  console.log(JSON.stringify(arrayOfRequiredAttendeesRecipients));
}
```

##### <a name="type"></a><span data-ttu-id="1676f-654">类型</span><span class="sxs-lookup"><span data-stu-id="1676f-654">Type</span></span>

*   <span data-ttu-id="1676f-655">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.8)</span><span class="sxs-lookup"><span data-stu-id="1676f-655">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.8)</span></span>

##### <a name="requirements"></a><span data-ttu-id="1676f-656">Requirements</span><span class="sxs-lookup"><span data-stu-id="1676f-656">Requirements</span></span>

|<span data-ttu-id="1676f-657">要求</span><span class="sxs-lookup"><span data-stu-id="1676f-657">Requirement</span></span>|<span data-ttu-id="1676f-658">值</span><span class="sxs-lookup"><span data-stu-id="1676f-658">Value</span></span>|
|---|---|
|[<span data-ttu-id="1676f-659">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="1676f-659">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="1676f-660">1.0</span><span class="sxs-lookup"><span data-stu-id="1676f-660">1.0</span></span>|
|[<span data-ttu-id="1676f-661">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="1676f-661">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="1676f-662">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1676f-662">ReadItem</span></span>|
|[<span data-ttu-id="1676f-663">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="1676f-663">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="1676f-664">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="1676f-664">Compose or Read</span></span>|

<br>

---
---

#### <a name="sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-18"></a><span data-ttu-id="1676f-665">sender: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8)</span><span class="sxs-lookup"><span data-stu-id="1676f-665">sender: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8)</span></span>

<span data-ttu-id="1676f-p135">获取电子邮件发件人的电子邮件地址。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="1676f-p135">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="1676f-p136">[`from`](#from-emailaddressdetailsfrom) 和 `sender` 属性表示同一个人，邮件由代理人发送的除外。在此情况下，`from` 属性表示代理程序，而 sender 属性表示代理人。</span><span class="sxs-lookup"><span data-stu-id="1676f-p136">The [`from`](#from-emailaddressdetailsfrom) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="1676f-670">`sender` 属性中 `EmailAddressDetails` 对象的 `recipientType` 属性为 `undefined`。</span><span class="sxs-lookup"><span data-stu-id="1676f-670">The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="1676f-671">类型</span><span class="sxs-lookup"><span data-stu-id="1676f-671">Type</span></span>

*   [<span data-ttu-id="1676f-672">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="1676f-672">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8)

##### <a name="requirements"></a><span data-ttu-id="1676f-673">要求</span><span class="sxs-lookup"><span data-stu-id="1676f-673">Requirements</span></span>

|<span data-ttu-id="1676f-674">要求</span><span class="sxs-lookup"><span data-stu-id="1676f-674">Requirement</span></span>|<span data-ttu-id="1676f-675">值</span><span class="sxs-lookup"><span data-stu-id="1676f-675">Value</span></span>|
|---|---|
|[<span data-ttu-id="1676f-676">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="1676f-676">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="1676f-677">1.0</span><span class="sxs-lookup"><span data-stu-id="1676f-677">1.0</span></span>|
|[<span data-ttu-id="1676f-678">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="1676f-678">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="1676f-679">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1676f-679">ReadItem</span></span>|
|[<span data-ttu-id="1676f-680">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="1676f-680">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="1676f-681">阅读</span><span class="sxs-lookup"><span data-stu-id="1676f-681">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="1676f-682">示例</span><span class="sxs-lookup"><span data-stu-id="1676f-682">Example</span></span>

```js
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
console.log("Sender: " + senderName + " (" + senderAddress + ")");
```

<br>

---
---

#### <a name="nullable-seriesid-string"></a><span data-ttu-id="1676f-683">（可以为 null） Webcasts&seriesid： String</span><span class="sxs-lookup"><span data-stu-id="1676f-683">(nullable) seriesId: String</span></span>

<span data-ttu-id="1676f-684">获取实例所属的系列的 id。</span><span class="sxs-lookup"><span data-stu-id="1676f-684">Gets the id of the series that an instance belongs to.</span></span>

<span data-ttu-id="1676f-685">在 web 上的 Outlook 和桌面客户端中`seriesId` ，返回此项所属的父（系列）项的 Exchange web 服务（EWS） ID。</span><span class="sxs-lookup"><span data-stu-id="1676f-685">In Outlook on the web and desktop clients, the `seriesId` returns the Exchange Web Services (EWS) ID of the parent (series) item that this item belongs to.</span></span> <span data-ttu-id="1676f-686">但是，在 iOS 和 Android 中， `seriesId`将返回父项的 REST ID。</span><span class="sxs-lookup"><span data-stu-id="1676f-686">However, in iOS and Android, the `seriesId` returns the REST ID of the parent item.</span></span>

> [!NOTE]
> <span data-ttu-id="1676f-687">`seriesId` 属性返回的标识符与 Exchange Web 服务项目标识符相同。</span><span class="sxs-lookup"><span data-stu-id="1676f-687">The identifier returned by the `seriesId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="1676f-688">`seriesId`属性与 OUTLOOK REST API 使用的 outlook id 不相同。</span><span class="sxs-lookup"><span data-stu-id="1676f-688">The `seriesId` property is not identical to the Outlook IDs used by the Outlook REST API.</span></span> <span data-ttu-id="1676f-689">使用此值进行 REST API 调用前，应使用 [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string) 对它进行转换。</span><span class="sxs-lookup"><span data-stu-id="1676f-689">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="1676f-690">有关详细信息，请参阅[从 Outlook 加载项使用 Outlook REST API](/outlook/add-ins/use-rest-api)。</span><span class="sxs-lookup"><span data-stu-id="1676f-690">For more details, see [Use the Outlook REST APIs from an Outlook add-in](/outlook/add-ins/use-rest-api).</span></span>

<span data-ttu-id="1676f-691">对于`seriesId`不包含`null`父项（如单个约会、系列项或会议请求）的项，该属性将返回， `undefined`对于不是会议请求的任何其他项，该属性返回。</span><span class="sxs-lookup"><span data-stu-id="1676f-691">The `seriesId` property returns `null` for items that do not have parent items such as single appointments, series items, or meeting requests and returns `undefined` for any other items that are not meeting requests.</span></span>

##### <a name="type"></a><span data-ttu-id="1676f-692">Type</span><span class="sxs-lookup"><span data-stu-id="1676f-692">Type</span></span>

* <span data-ttu-id="1676f-693">String</span><span class="sxs-lookup"><span data-stu-id="1676f-693">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="1676f-694">Requirements</span><span class="sxs-lookup"><span data-stu-id="1676f-694">Requirements</span></span>

|<span data-ttu-id="1676f-695">要求</span><span class="sxs-lookup"><span data-stu-id="1676f-695">Requirement</span></span>|<span data-ttu-id="1676f-696">值</span><span class="sxs-lookup"><span data-stu-id="1676f-696">Value</span></span>|
|---|---|
|[<span data-ttu-id="1676f-697">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="1676f-697">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="1676f-698">1.7</span><span class="sxs-lookup"><span data-stu-id="1676f-698">1.7</span></span>|
|[<span data-ttu-id="1676f-699">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="1676f-699">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="1676f-700">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1676f-700">ReadItem</span></span>|
|[<span data-ttu-id="1676f-701">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="1676f-701">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="1676f-702">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="1676f-702">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="1676f-703">示例</span><span class="sxs-lookup"><span data-stu-id="1676f-703">Example</span></span>

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

#### <a name="start-datetimejavascriptapioutlookofficetimeviewoutlook-js-18"></a><span data-ttu-id="1676f-704">start: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.8)</span><span class="sxs-lookup"><span data-stu-id="1676f-704">start: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.8)</span></span>

<span data-ttu-id="1676f-705">获取或设置约会开始的日期和时间。</span><span class="sxs-lookup"><span data-stu-id="1676f-705">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="1676f-p139">将 `start` 属性表示为协调世界时 (UTC) 的日期和时间值。可使用 [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) 方法将该值转换为客户端的本地日期和时间。</span><span class="sxs-lookup"><span data-stu-id="1676f-p139">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="1676f-708">阅读模式</span><span class="sxs-lookup"><span data-stu-id="1676f-708">Read mode</span></span>

<span data-ttu-id="1676f-709">`start` 属性返回 `Date` 对象。</span><span class="sxs-lookup"><span data-stu-id="1676f-709">The `start` property returns a `Date` object.</span></span>

```js
var start = Office.context.mailbox.item.start;
console.log("Appointment start: " + JSON.stringify(start));
```

##### <a name="compose-mode"></a><span data-ttu-id="1676f-710">撰写模式</span><span class="sxs-lookup"><span data-stu-id="1676f-710">Compose mode</span></span>

<span data-ttu-id="1676f-711">`start` 属性返回 `Time` 对象。</span><span class="sxs-lookup"><span data-stu-id="1676f-711">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="1676f-712">使用 [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.8#setasync-datetime--options--callback-) 方法设置开始时间时，应使用 [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) 方法将客户端的本地时间转换为服务器的 UTC。</span><span class="sxs-lookup"><span data-stu-id="1676f-712">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.8#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="1676f-713">以下示例通过使用 `Time` 对象的 [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.8#setasync-datetime--options--callback-) 方法，设置撰写模式下约会的开始时间。</span><span class="sxs-lookup"><span data-stu-id="1676f-713">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.8#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="1676f-714">类型</span><span class="sxs-lookup"><span data-stu-id="1676f-714">Type</span></span>

*   <span data-ttu-id="1676f-715">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.8)</span><span class="sxs-lookup"><span data-stu-id="1676f-715">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.8)</span></span>

##### <a name="requirements"></a><span data-ttu-id="1676f-716">Requirements</span><span class="sxs-lookup"><span data-stu-id="1676f-716">Requirements</span></span>

|<span data-ttu-id="1676f-717">要求</span><span class="sxs-lookup"><span data-stu-id="1676f-717">Requirement</span></span>|<span data-ttu-id="1676f-718">值</span><span class="sxs-lookup"><span data-stu-id="1676f-718">Value</span></span>|
|---|---|
|[<span data-ttu-id="1676f-719">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="1676f-719">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="1676f-720">1.0</span><span class="sxs-lookup"><span data-stu-id="1676f-720">1.0</span></span>|
|[<span data-ttu-id="1676f-721">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="1676f-721">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="1676f-722">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1676f-722">ReadItem</span></span>|
|[<span data-ttu-id="1676f-723">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="1676f-723">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="1676f-724">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="1676f-724">Compose or Read</span></span>|

<br>

---
---

#### <a name="subject-stringsubjectjavascriptapioutlookofficesubjectviewoutlook-js-18"></a><span data-ttu-id="1676f-725">subject: String|[Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.8)</span><span class="sxs-lookup"><span data-stu-id="1676f-725">subject: String|[Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.8)</span></span>

<span data-ttu-id="1676f-726">获取或设置显示在项目的主题字段中的说明。</span><span class="sxs-lookup"><span data-stu-id="1676f-726">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="1676f-727">`subject` 属性获取或设置由电子邮件服务器发送项目时的整个主题。</span><span class="sxs-lookup"><span data-stu-id="1676f-727">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="1676f-728">阅读模式</span><span class="sxs-lookup"><span data-stu-id="1676f-728">Read mode</span></span>

<span data-ttu-id="1676f-p140">`subject` 属性返回一个字符串。使用 [`normalizedSubject`](#normalizedsubject-string) 属性获取不带任何前导前缀（如 `RE:` 和 `FW:`）的主题。</span><span class="sxs-lookup"><span data-stu-id="1676f-p140">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

<span data-ttu-id="1676f-731">以下 JavaScript 代码示例显示了如何访问 Outlook 中当前项目的 `subject` 属性。</span><span class="sxs-lookup"><span data-stu-id="1676f-731">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

```js
var subject = Office.context.mailbox.item.subject;
console.log(subject);
```

##### <a name="compose-mode"></a><span data-ttu-id="1676f-732">撰写模式</span><span class="sxs-lookup"><span data-stu-id="1676f-732">Compose mode</span></span>
<span data-ttu-id="1676f-733">`subject` 属性返回一个 `Subject` 对象，该对象提供用于获取和设置主题的方法。</span><span class="sxs-lookup"><span data-stu-id="1676f-733">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```js
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
  console.log(subject);
}
```

##### <a name="type"></a><span data-ttu-id="1676f-734">类型</span><span class="sxs-lookup"><span data-stu-id="1676f-734">Type</span></span>

*   <span data-ttu-id="1676f-735">String | [Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.8)</span><span class="sxs-lookup"><span data-stu-id="1676f-735">String | [Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.8)</span></span>

##### <a name="requirements"></a><span data-ttu-id="1676f-736">要求</span><span class="sxs-lookup"><span data-stu-id="1676f-736">Requirements</span></span>

|<span data-ttu-id="1676f-737">要求</span><span class="sxs-lookup"><span data-stu-id="1676f-737">Requirement</span></span>|<span data-ttu-id="1676f-738">值</span><span class="sxs-lookup"><span data-stu-id="1676f-738">Value</span></span>|
|---|---|
|[<span data-ttu-id="1676f-739">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="1676f-739">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="1676f-740">1.0</span><span class="sxs-lookup"><span data-stu-id="1676f-740">1.0</span></span>|
|[<span data-ttu-id="1676f-741">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="1676f-741">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="1676f-742">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1676f-742">ReadItem</span></span>|
|[<span data-ttu-id="1676f-743">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="1676f-743">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="1676f-744">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="1676f-744">Compose or Read</span></span>|

<br>

---
---

#### <a name="to-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-18recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-18"></a><span data-ttu-id="1676f-745">to: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.8)</span><span class="sxs-lookup"><span data-stu-id="1676f-745">to: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.8)</span></span>

<span data-ttu-id="1676f-746">提供对邮件的“**收件人**”行上的收件人的访问权限。</span><span class="sxs-lookup"><span data-stu-id="1676f-746">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="1676f-747">对象的类型和访问级别取决于当前项目的模式。</span><span class="sxs-lookup"><span data-stu-id="1676f-747">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="1676f-748">阅读模式</span><span class="sxs-lookup"><span data-stu-id="1676f-748">Read mode</span></span>

<span data-ttu-id="1676f-749">`to` 属性返回包含邮件的**收件人**行上所列的每个收件人的 `EmailAddressDetails` 对象的数组。</span><span class="sxs-lookup"><span data-stu-id="1676f-749">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message.</span></span> <span data-ttu-id="1676f-750">默认情况下，集合上限为 100 个成员。</span><span class="sxs-lookup"><span data-stu-id="1676f-750">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="1676f-751">但是，在 Windows 和 Mac 上，最多可包含 500 个成员。</span><span class="sxs-lookup"><span data-stu-id="1676f-751">However, on Windows and Mac, you can get 500 members maximum.</span></span>

```js
console.log(JSON.stringify(Office.context.mailbox.item.to));
```

##### <a name="compose-mode"></a><span data-ttu-id="1676f-752">撰写模式</span><span class="sxs-lookup"><span data-stu-id="1676f-752">Compose mode</span></span>

<span data-ttu-id="1676f-753">`to` 属性返回一个 `Recipients` 对象，该对象提供用于获取或更新邮件的“**收件人**”行上收件人的方法。</span><span class="sxs-lookup"><span data-stu-id="1676f-753">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span> <span data-ttu-id="1676f-754">默认情况下，集合上限为 100 个成员。</span><span class="sxs-lookup"><span data-stu-id="1676f-754">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="1676f-755">但是，在 Windows 和 Mac 上，以下限制适用。</span><span class="sxs-lookup"><span data-stu-id="1676f-755">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="1676f-756">最多包含 500 个成员。</span><span class="sxs-lookup"><span data-stu-id="1676f-756">Get 500 members maximum.</span></span>
- <span data-ttu-id="1676f-757">为每个呼叫最多设置 100 个成员，总共多达 500 个成员。</span><span class="sxs-lookup"><span data-stu-id="1676f-757">Set a maximum of 100 members per call, up to 500 members total.</span></span>

```js
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="1676f-758">类型</span><span class="sxs-lookup"><span data-stu-id="1676f-758">Type</span></span>

*   <span data-ttu-id="1676f-759">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.8)</span><span class="sxs-lookup"><span data-stu-id="1676f-759">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.8)</span></span>

##### <a name="requirements"></a><span data-ttu-id="1676f-760">Requirements</span><span class="sxs-lookup"><span data-stu-id="1676f-760">Requirements</span></span>

|<span data-ttu-id="1676f-761">要求</span><span class="sxs-lookup"><span data-stu-id="1676f-761">Requirement</span></span>|<span data-ttu-id="1676f-762">值</span><span class="sxs-lookup"><span data-stu-id="1676f-762">Value</span></span>|
|---|---|
|[<span data-ttu-id="1676f-763">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="1676f-763">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="1676f-764">1.0</span><span class="sxs-lookup"><span data-stu-id="1676f-764">1.0</span></span>|
|[<span data-ttu-id="1676f-765">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="1676f-765">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="1676f-766">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1676f-766">ReadItem</span></span>|
|[<span data-ttu-id="1676f-767">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="1676f-767">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="1676f-768">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="1676f-768">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="1676f-769">方法</span><span class="sxs-lookup"><span data-stu-id="1676f-769">Methods</span></span>

#### <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="1676f-770">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="1676f-770">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="1676f-771">将文件作为附件添加到邮件或约会。</span><span class="sxs-lookup"><span data-stu-id="1676f-771">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="1676f-772">`addFileAttachmentAsync` 方法在指定的 URI 上载文件并将其附加到撰写窗体中的项目。</span><span class="sxs-lookup"><span data-stu-id="1676f-772">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="1676f-773">随后可以将该标识符与 [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) 方法一同使用，以删除同一个会话中的附件。</span><span class="sxs-lookup"><span data-stu-id="1676f-773">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="1676f-774">参数</span><span class="sxs-lookup"><span data-stu-id="1676f-774">Parameters</span></span>
|<span data-ttu-id="1676f-775">名称</span><span class="sxs-lookup"><span data-stu-id="1676f-775">Name</span></span>|<span data-ttu-id="1676f-776">类型</span><span class="sxs-lookup"><span data-stu-id="1676f-776">Type</span></span>|<span data-ttu-id="1676f-777">属性</span><span class="sxs-lookup"><span data-stu-id="1676f-777">Attributes</span></span>|<span data-ttu-id="1676f-778">说明</span><span class="sxs-lookup"><span data-stu-id="1676f-778">Description</span></span>|
|---|---|---|---|
|`uri`|<span data-ttu-id="1676f-779">String</span><span class="sxs-lookup"><span data-stu-id="1676f-779">String</span></span>||<span data-ttu-id="1676f-p144">提供附加到邮件或约会的文件的位置的 URI。最大长度为 2048 个字符。</span><span class="sxs-lookup"><span data-stu-id="1676f-p144">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`|<span data-ttu-id="1676f-782">字符串</span><span class="sxs-lookup"><span data-stu-id="1676f-782">String</span></span>||<span data-ttu-id="1676f-p145">在附件上载过程中显示的附件名称。最大长度为 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="1676f-p145">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="1676f-785">Object</span><span class="sxs-lookup"><span data-stu-id="1676f-785">Object</span></span>|<span data-ttu-id="1676f-786">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="1676f-786">&lt;optional&gt;</span></span>|<span data-ttu-id="1676f-787">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="1676f-787">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="1676f-788">对象</span><span class="sxs-lookup"><span data-stu-id="1676f-788">Object</span></span>|<span data-ttu-id="1676f-789">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="1676f-789">&lt;optional&gt;</span></span>|<span data-ttu-id="1676f-790">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="1676f-790">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.isInline`|<span data-ttu-id="1676f-791">布尔值</span><span class="sxs-lookup"><span data-stu-id="1676f-791">Boolean</span></span>|<span data-ttu-id="1676f-792">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="1676f-792">&lt;optional&gt;</span></span>|<span data-ttu-id="1676f-793">如果为 `true`，则表示附件将在邮件正文中内联显示，并且不应显示在附件列表中。</span><span class="sxs-lookup"><span data-stu-id="1676f-793">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`callback`|<span data-ttu-id="1676f-794">函数</span><span class="sxs-lookup"><span data-stu-id="1676f-794">function</span></span>|<span data-ttu-id="1676f-795">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="1676f-795">&lt;optional&gt;</span></span>|<span data-ttu-id="1676f-796">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="1676f-796">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="1676f-797">如果成功，附件标识符将在 `asyncResult.value` 属性中提供。</span><span class="sxs-lookup"><span data-stu-id="1676f-797">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="1676f-798">如果上传附件失败，`asyncResult` 对象将包含一个提供错误说明的 `Error` 对象。</span><span class="sxs-lookup"><span data-stu-id="1676f-798">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="1676f-799">错误</span><span class="sxs-lookup"><span data-stu-id="1676f-799">Errors</span></span>

|<span data-ttu-id="1676f-800">错误代码</span><span class="sxs-lookup"><span data-stu-id="1676f-800">Error code</span></span>|<span data-ttu-id="1676f-801">说明</span><span class="sxs-lookup"><span data-stu-id="1676f-801">Description</span></span>|
|------------|-------------|
|`AttachmentSizeExceeded`|<span data-ttu-id="1676f-802">附件大小超过了允许的大小。</span><span class="sxs-lookup"><span data-stu-id="1676f-802">The attachment is larger than allowed.</span></span>|
|`FileTypeNotSupported`|<span data-ttu-id="1676f-803">该附件的扩展名不是允许的扩展名。</span><span class="sxs-lookup"><span data-stu-id="1676f-803">The attachment has an extension that is not allowed.</span></span>|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="1676f-804">邮件或约会具有的附件过多。</span><span class="sxs-lookup"><span data-stu-id="1676f-804">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="1676f-805">要求</span><span class="sxs-lookup"><span data-stu-id="1676f-805">Requirements</span></span>

|<span data-ttu-id="1676f-806">要求</span><span class="sxs-lookup"><span data-stu-id="1676f-806">Requirement</span></span>|<span data-ttu-id="1676f-807">值</span><span class="sxs-lookup"><span data-stu-id="1676f-807">Value</span></span>|
|---|---|
|[<span data-ttu-id="1676f-808">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="1676f-808">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="1676f-809">1.1</span><span class="sxs-lookup"><span data-stu-id="1676f-809">1.1</span></span>|
|[<span data-ttu-id="1676f-810">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="1676f-810">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="1676f-811">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="1676f-811">ReadWriteItem</span></span>|
|[<span data-ttu-id="1676f-812">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="1676f-812">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="1676f-813">撰写</span><span class="sxs-lookup"><span data-stu-id="1676f-813">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="1676f-814">示例</span><span class="sxs-lookup"><span data-stu-id="1676f-814">Examples</span></span>

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

<span data-ttu-id="1676f-815">以下示例将图像文件添加为内联附件，并在邮件正文中引用该附件。</span><span class="sxs-lookup"><span data-stu-id="1676f-815">The following example adds an image file as an inline attachment and references the attachment in the message body.</span></span>

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

#### <a name="addfileattachmentfrombase64asyncbase64file-attachmentname-options-callback"></a><span data-ttu-id="1676f-816">addFileAttachmentFromBase64Async （base64File，attachmentName，[options]，[callback]）</span><span class="sxs-lookup"><span data-stu-id="1676f-816">addFileAttachmentFromBase64Async(base64File, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="1676f-817">将 base64 编码中的文件作为附件添加到邮件或约会中。</span><span class="sxs-lookup"><span data-stu-id="1676f-817">Adds a file from the base64 encoding to a message or appointment as an attachment.</span></span>

<span data-ttu-id="1676f-818">该`addFileAttachmentFromBase64Async`方法从 base64 编码中上载文件，并将其附加到撰写窗体中的项目。</span><span class="sxs-lookup"><span data-stu-id="1676f-818">The `addFileAttachmentFromBase64Async` method uploads the file from the base64 encoding and attaches it to the item in the compose form.</span></span> <span data-ttu-id="1676f-819">此方法返回 AsyncResult 对象中的附件标识符。</span><span class="sxs-lookup"><span data-stu-id="1676f-819">This method returns the attachment identifier in the AsyncResult.value object.</span></span>

<span data-ttu-id="1676f-820">随后可以将该标识符与 [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) 方法一同使用，以删除同一个会话中的附件。</span><span class="sxs-lookup"><span data-stu-id="1676f-820">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="1676f-821">参数</span><span class="sxs-lookup"><span data-stu-id="1676f-821">Parameters</span></span>

|<span data-ttu-id="1676f-822">名称</span><span class="sxs-lookup"><span data-stu-id="1676f-822">Name</span></span>|<span data-ttu-id="1676f-823">类型</span><span class="sxs-lookup"><span data-stu-id="1676f-823">Type</span></span>|<span data-ttu-id="1676f-824">属性</span><span class="sxs-lookup"><span data-stu-id="1676f-824">Attributes</span></span>|<span data-ttu-id="1676f-825">说明</span><span class="sxs-lookup"><span data-stu-id="1676f-825">Description</span></span>|
|---|---|---|---|
|`base64File`|<span data-ttu-id="1676f-826">String</span><span class="sxs-lookup"><span data-stu-id="1676f-826">String</span></span>||<span data-ttu-id="1676f-827">要添加到电子邮件或事件的图像或文件的 base64 编码内容。</span><span class="sxs-lookup"><span data-stu-id="1676f-827">The base64 encoded content of an image or file to be added to an email or event.</span></span>|
|`attachmentName`|<span data-ttu-id="1676f-828">字符串</span><span class="sxs-lookup"><span data-stu-id="1676f-828">String</span></span>||<span data-ttu-id="1676f-p147">在附件上载过程中显示的附件名称。最大长度为 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="1676f-p147">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="1676f-831">对象</span><span class="sxs-lookup"><span data-stu-id="1676f-831">Object</span></span>|<span data-ttu-id="1676f-832">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="1676f-832">&lt;optional&gt;</span></span>|<span data-ttu-id="1676f-833">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="1676f-833">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="1676f-834">对象</span><span class="sxs-lookup"><span data-stu-id="1676f-834">Object</span></span>|<span data-ttu-id="1676f-835">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="1676f-835">&lt;optional&gt;</span></span>|<span data-ttu-id="1676f-836">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="1676f-836">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.isInline`|<span data-ttu-id="1676f-837">布尔值</span><span class="sxs-lookup"><span data-stu-id="1676f-837">Boolean</span></span>|<span data-ttu-id="1676f-838">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="1676f-838">&lt;optional&gt;</span></span>|<span data-ttu-id="1676f-839">如果为 `true`，则表示附件将在邮件正文中内联显示，并且不应显示在附件列表中。</span><span class="sxs-lookup"><span data-stu-id="1676f-839">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`callback`|<span data-ttu-id="1676f-840">函数</span><span class="sxs-lookup"><span data-stu-id="1676f-840">function</span></span>|<span data-ttu-id="1676f-841">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="1676f-841">&lt;optional&gt;</span></span>|<span data-ttu-id="1676f-842">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="1676f-842">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="1676f-843">如果成功，附件标识符将在 `asyncResult.value` 属性中提供。</span><span class="sxs-lookup"><span data-stu-id="1676f-843">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="1676f-844">如果上传附件失败，`asyncResult` 对象将包含一个提供错误说明的 `Error` 对象。</span><span class="sxs-lookup"><span data-stu-id="1676f-844">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="1676f-845">错误</span><span class="sxs-lookup"><span data-stu-id="1676f-845">Errors</span></span>

|<span data-ttu-id="1676f-846">错误代码</span><span class="sxs-lookup"><span data-stu-id="1676f-846">Error code</span></span>|<span data-ttu-id="1676f-847">说明</span><span class="sxs-lookup"><span data-stu-id="1676f-847">Description</span></span>|
|------------|-------------|
|`AttachmentSizeExceeded`|<span data-ttu-id="1676f-848">附件大小超过了允许的大小。</span><span class="sxs-lookup"><span data-stu-id="1676f-848">The attachment is larger than allowed.</span></span>|
|`FileTypeNotSupported`|<span data-ttu-id="1676f-849">该附件的扩展名不是允许的扩展名。</span><span class="sxs-lookup"><span data-stu-id="1676f-849">The attachment has an extension that is not allowed.</span></span>|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="1676f-850">邮件或约会具有的附件过多。</span><span class="sxs-lookup"><span data-stu-id="1676f-850">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="1676f-851">要求</span><span class="sxs-lookup"><span data-stu-id="1676f-851">Requirements</span></span>

|<span data-ttu-id="1676f-852">要求</span><span class="sxs-lookup"><span data-stu-id="1676f-852">Requirement</span></span>|<span data-ttu-id="1676f-853">值</span><span class="sxs-lookup"><span data-stu-id="1676f-853">Value</span></span>|
|---|---|
|[<span data-ttu-id="1676f-854">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="1676f-854">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="1676f-855">1.8</span><span class="sxs-lookup"><span data-stu-id="1676f-855">1.8</span></span>|
|[<span data-ttu-id="1676f-856">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="1676f-856">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="1676f-857">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="1676f-857">ReadWriteItem</span></span>|
|[<span data-ttu-id="1676f-858">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="1676f-858">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="1676f-859">撰写</span><span class="sxs-lookup"><span data-stu-id="1676f-859">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="1676f-860">示例</span><span class="sxs-lookup"><span data-stu-id="1676f-860">Examples</span></span>

```js
Office.context.mailbox.item.addFileAttachmentFromBase64Async(
  base64String,
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

#### <a name="addhandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="1676f-861">addHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="1676f-861">addHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="1676f-862">添加支持事件的事件处理程序。</span><span class="sxs-lookup"><span data-stu-id="1676f-862">Adds an event handler for a supported event.</span></span>

<span data-ttu-id="1676f-863">目前，受支持的事件`Office.EventType.AttachmentsChanged`类型`Office.EventType.AppointmentTimeChanged`是`Office.EventType.EnhancedLocationsChanged`、 `Office.EventType.RecipientsChanged`、、 `Office.EventType.RecurrenceChanged`和。</span><span class="sxs-lookup"><span data-stu-id="1676f-863">Currently the supported event types are `Office.EventType.AttachmentsChanged`, `Office.EventType.AppointmentTimeChanged`, `Office.EventType.EnhancedLocationsChanged`, `Office.EventType.RecipientsChanged`, and `Office.EventType.RecurrenceChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="1676f-864">Parameters</span><span class="sxs-lookup"><span data-stu-id="1676f-864">Parameters</span></span>

| <span data-ttu-id="1676f-865">名称</span><span class="sxs-lookup"><span data-stu-id="1676f-865">Name</span></span> | <span data-ttu-id="1676f-866">类型</span><span class="sxs-lookup"><span data-stu-id="1676f-866">Type</span></span> | <span data-ttu-id="1676f-867">属性</span><span class="sxs-lookup"><span data-stu-id="1676f-867">Attributes</span></span> | <span data-ttu-id="1676f-868">说明</span><span class="sxs-lookup"><span data-stu-id="1676f-868">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="1676f-869">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="1676f-869">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="1676f-870">应调用处理程序的事件。</span><span class="sxs-lookup"><span data-stu-id="1676f-870">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="1676f-871">函数</span><span class="sxs-lookup"><span data-stu-id="1676f-871">Function</span></span> || <span data-ttu-id="1676f-p148">用于处理事件的函数。此函数必须接受一个参数，即对象文本。参数上的 `type` 属性将匹配传递给 `addHandlerAsync` 的 `eventType` 参数。</span><span class="sxs-lookup"><span data-stu-id="1676f-p148">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `addHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="1676f-875">Object</span><span class="sxs-lookup"><span data-stu-id="1676f-875">Object</span></span> | <span data-ttu-id="1676f-876">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="1676f-876">&lt;optional&gt;</span></span> | <span data-ttu-id="1676f-877">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="1676f-877">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="1676f-878">对象</span><span class="sxs-lookup"><span data-stu-id="1676f-878">Object</span></span> | <span data-ttu-id="1676f-879">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="1676f-879">&lt;optional&gt;</span></span> | <span data-ttu-id="1676f-880">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="1676f-880">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="1676f-881">函数</span><span class="sxs-lookup"><span data-stu-id="1676f-881">function</span></span>| <span data-ttu-id="1676f-882">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="1676f-882">&lt;optional&gt;</span></span>|<span data-ttu-id="1676f-883">方法完成后，使用单个参数 `callback`（一个 [`asyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `AsyncResult` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="1676f-883">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="1676f-884">Requirements</span><span class="sxs-lookup"><span data-stu-id="1676f-884">Requirements</span></span>

|<span data-ttu-id="1676f-885">要求</span><span class="sxs-lookup"><span data-stu-id="1676f-885">Requirement</span></span>| <span data-ttu-id="1676f-886">值</span><span class="sxs-lookup"><span data-stu-id="1676f-886">Value</span></span>|
|---|---|
|[<span data-ttu-id="1676f-887">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="1676f-887">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1676f-888">1.7</span><span class="sxs-lookup"><span data-stu-id="1676f-888">1.7</span></span> |
|[<span data-ttu-id="1676f-889">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="1676f-889">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1676f-890">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1676f-890">ReadItem</span></span> |
|[<span data-ttu-id="1676f-891">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="1676f-891">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="1676f-892">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="1676f-892">Compose or Read</span></span> |

##### <a name="example"></a><span data-ttu-id="1676f-893">示例</span><span class="sxs-lookup"><span data-stu-id="1676f-893">Example</span></span>

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

#### <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="1676f-894">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="1676f-894">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="1676f-895">将 Exchange 项目（如邮件）作为附件添加到邮件或约会。</span><span class="sxs-lookup"><span data-stu-id="1676f-895">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="1676f-p149">`addItemAttachmentAsync` 方法将包含指定 Exchange 标识符的项目附加到撰写窗体中的项目。如果指定一个回调方法，此方法使用 `asyncResult` 参数调用，该参数包含一个附件标识符或代码，指示附加项目过程中出现的任何错误。可以使用 `options` 参数将状态信息传递给回调方法（如果需要）。</span><span class="sxs-lookup"><span data-stu-id="1676f-p149">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="1676f-899">随后可以将该标识符与 [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) 方法一同使用，以删除同一个会话中的附件。</span><span class="sxs-lookup"><span data-stu-id="1676f-899">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="1676f-900">如果 Office 加载项是在 Outlook 网页版中运行，`addItemAttachmentAsync` 方法可以将项附加到除正在编辑的项外的项；但既不支持也不建议这样做。</span><span class="sxs-lookup"><span data-stu-id="1676f-900">If your Office Add-in is running in Outlook on the web, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="1676f-901">Parameters</span><span class="sxs-lookup"><span data-stu-id="1676f-901">Parameters</span></span>

|<span data-ttu-id="1676f-902">名称</span><span class="sxs-lookup"><span data-stu-id="1676f-902">Name</span></span>|<span data-ttu-id="1676f-903">类型</span><span class="sxs-lookup"><span data-stu-id="1676f-903">Type</span></span>|<span data-ttu-id="1676f-904">属性</span><span class="sxs-lookup"><span data-stu-id="1676f-904">Attributes</span></span>|<span data-ttu-id="1676f-905">说明</span><span class="sxs-lookup"><span data-stu-id="1676f-905">Description</span></span>|
|---|---|---|---|
|`itemId`|<span data-ttu-id="1676f-906">String</span><span class="sxs-lookup"><span data-stu-id="1676f-906">String</span></span>||<span data-ttu-id="1676f-p150">要附加的项目的 Exchange 标识符。最大长度为 100 个字符。</span><span class="sxs-lookup"><span data-stu-id="1676f-p150">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`|<span data-ttu-id="1676f-909">String</span><span class="sxs-lookup"><span data-stu-id="1676f-909">String</span></span>||<span data-ttu-id="1676f-910">要附加的项目的主题。</span><span class="sxs-lookup"><span data-stu-id="1676f-910">The subject of the item to be attached.</span></span> <span data-ttu-id="1676f-911">最大长度为 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="1676f-911">The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="1676f-912">对象</span><span class="sxs-lookup"><span data-stu-id="1676f-912">Object</span></span>|<span data-ttu-id="1676f-913">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="1676f-913">&lt;optional&gt;</span></span>|<span data-ttu-id="1676f-914">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="1676f-914">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="1676f-915">对象</span><span class="sxs-lookup"><span data-stu-id="1676f-915">Object</span></span>|<span data-ttu-id="1676f-916">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="1676f-916">&lt;optional&gt;</span></span>|<span data-ttu-id="1676f-917">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="1676f-917">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="1676f-918">函数</span><span class="sxs-lookup"><span data-stu-id="1676f-918">function</span></span>|<span data-ttu-id="1676f-919">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="1676f-919">&lt;optional&gt;</span></span>|<span data-ttu-id="1676f-920">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="1676f-920">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="1676f-921">如果成功，附件标识符将在 `asyncResult.value` 属性中提供。</span><span class="sxs-lookup"><span data-stu-id="1676f-921">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="1676f-922">如果添加附件失败，`asyncResult` 对象将包含一个提供错误说明的 `Error` 对象。</span><span class="sxs-lookup"><span data-stu-id="1676f-922">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="1676f-923">错误</span><span class="sxs-lookup"><span data-stu-id="1676f-923">Errors</span></span>

|<span data-ttu-id="1676f-924">错误代码</span><span class="sxs-lookup"><span data-stu-id="1676f-924">Error code</span></span>|<span data-ttu-id="1676f-925">说明</span><span class="sxs-lookup"><span data-stu-id="1676f-925">Description</span></span>|
|------------|-------------|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="1676f-926">邮件或约会具有的附件过多。</span><span class="sxs-lookup"><span data-stu-id="1676f-926">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="1676f-927">Requirements</span><span class="sxs-lookup"><span data-stu-id="1676f-927">Requirements</span></span>

|<span data-ttu-id="1676f-928">要求</span><span class="sxs-lookup"><span data-stu-id="1676f-928">Requirement</span></span>|<span data-ttu-id="1676f-929">值</span><span class="sxs-lookup"><span data-stu-id="1676f-929">Value</span></span>|
|---|---|
|[<span data-ttu-id="1676f-930">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="1676f-930">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="1676f-931">1.1</span><span class="sxs-lookup"><span data-stu-id="1676f-931">1.1</span></span>|
|[<span data-ttu-id="1676f-932">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="1676f-932">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="1676f-933">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="1676f-933">ReadWriteItem</span></span>|
|[<span data-ttu-id="1676f-934">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="1676f-934">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="1676f-935">撰写</span><span class="sxs-lookup"><span data-stu-id="1676f-935">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="1676f-936">示例</span><span class="sxs-lookup"><span data-stu-id="1676f-936">Example</span></span>

<span data-ttu-id="1676f-937">以下示例将现有的 Outlook 项目添加为名为 `My Attachment` 的附件。</span><span class="sxs-lookup"><span data-stu-id="1676f-937">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

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

#### <a name="close"></a><span data-ttu-id="1676f-938">close()</span><span class="sxs-lookup"><span data-stu-id="1676f-938">close()</span></span>

<span data-ttu-id="1676f-939">关闭当前正在撰写的项目。</span><span class="sxs-lookup"><span data-stu-id="1676f-939">Closes the current item that is being composed.</span></span>

<span data-ttu-id="1676f-p152">
            \`close\` 方法的行为取决于要撰写的项目的当前状态。如果项目具有未保存的更改，客户端将提示用户保存、放弃或取消关闭操作。</span><span class="sxs-lookup"><span data-stu-id="1676f-p152">The behavior of the `close` method depends on the current state of the item being composed. If the item has unsaved changes, the client prompts the user to save, discard, or cancel the close action.</span></span>

> [!NOTE]
> <span data-ttu-id="1676f-942">在 Outlook 网页版中，如果该项目是约会并且之前已使用 `saveAsync` 保存，则即使自上次保存项目后未发生任何更改，也会提示用户保存、放弃或取消。</span><span class="sxs-lookup"><span data-stu-id="1676f-942">In Outlook on the web, if the item is an appointment and it has previously been saved using `saveAsync`, the user is prompted to save, discard, or cancel even if no changes have occurred since the item was last saved.</span></span>

<span data-ttu-id="1676f-943">在 Outlook 桌面客户端中，如果邮件是内联答复，`close` 方法不起作用。</span><span class="sxs-lookup"><span data-stu-id="1676f-943">In the Outlook desktop client, if the message is an inline reply, the `close` method has no effect.</span></span>

##### <a name="requirements"></a><span data-ttu-id="1676f-944">Requirements</span><span class="sxs-lookup"><span data-stu-id="1676f-944">Requirements</span></span>

|<span data-ttu-id="1676f-945">要求</span><span class="sxs-lookup"><span data-stu-id="1676f-945">Requirement</span></span>|<span data-ttu-id="1676f-946">值</span><span class="sxs-lookup"><span data-stu-id="1676f-946">Value</span></span>|
|---|---|
|[<span data-ttu-id="1676f-947">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="1676f-947">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="1676f-948">1.3</span><span class="sxs-lookup"><span data-stu-id="1676f-948">1.3</span></span>|
|[<span data-ttu-id="1676f-949">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="1676f-949">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="1676f-950">受限</span><span class="sxs-lookup"><span data-stu-id="1676f-950">Restricted</span></span>|
|[<span data-ttu-id="1676f-951">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="1676f-951">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="1676f-952">撰写</span><span class="sxs-lookup"><span data-stu-id="1676f-952">Compose</span></span>|

<br>

---
---

#### <a name="displayreplyallformformdata-callback"></a><span data-ttu-id="1676f-953">displayReplyAllForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="1676f-953">displayReplyAllForm(formData, [callback])</span></span>

<span data-ttu-id="1676f-954">显示答复窗体，其中包括所选邮件的发件人和所有收件人或所选约会的组织者和所有与会者。</span><span class="sxs-lookup"><span data-stu-id="1676f-954">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="1676f-955">iOS 版 Outlook 或 Android 版 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="1676f-955">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="1676f-956">在 Outlook 网页版中，答复窗体显示为包含 3 列视图的弹出式窗体，以及包含 2 列或 1 列视图的弹出式窗体。</span><span class="sxs-lookup"><span data-stu-id="1676f-956">In Outlook on the web, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="1676f-957">如果任意字符串参数超出其限制，`displayReplyAllForm` 将引发异常。</span><span class="sxs-lookup"><span data-stu-id="1676f-957">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="1676f-p153">如果附件已在 `formData.attachments` 参数中指定，Outlook 网页版和 Outlook 桌面版客户端会尝试下载所有附件，并将它们附加到答复窗体。如果无法添加任何附件，窗体 UI 中会显示错误。如果此操作是不可能完成的，系统不会抛出任何错误消息。</span><span class="sxs-lookup"><span data-stu-id="1676f-p153">When attachments are specified in the `formData.attachments` parameter, Outlook on the web and desktop clients attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="1676f-961">Parameters</span><span class="sxs-lookup"><span data-stu-id="1676f-961">Parameters</span></span>

|<span data-ttu-id="1676f-962">名称</span><span class="sxs-lookup"><span data-stu-id="1676f-962">Name</span></span>|<span data-ttu-id="1676f-963">类型</span><span class="sxs-lookup"><span data-stu-id="1676f-963">Type</span></span>|<span data-ttu-id="1676f-964">属性</span><span class="sxs-lookup"><span data-stu-id="1676f-964">Attributes</span></span>|<span data-ttu-id="1676f-965">说明</span><span class="sxs-lookup"><span data-stu-id="1676f-965">Description</span></span>|
|---|---|---|---|
|`formData`|<span data-ttu-id="1676f-966">字符串 &#124; 对象</span><span class="sxs-lookup"><span data-stu-id="1676f-966">String &#124; Object</span></span>||<span data-ttu-id="1676f-p154">一个包含文本和 HTML 且表示答复窗体的正文的字符串。字符串限制为 32 KB。</span><span class="sxs-lookup"><span data-stu-id="1676f-p154">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="1676f-969">**或**</span><span class="sxs-lookup"><span data-stu-id="1676f-969">**OR**</span></span><br/><span data-ttu-id="1676f-p155">包含正文或附件数据和回调函数的对象。对象定义如下。</span><span class="sxs-lookup"><span data-stu-id="1676f-p155">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span>|
|`formData.htmlBody`|<span data-ttu-id="1676f-972">String</span><span class="sxs-lookup"><span data-stu-id="1676f-972">String</span></span>|<span data-ttu-id="1676f-973">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="1676f-973">&lt;optional&gt;</span></span>|<span data-ttu-id="1676f-p156">一个包含文本和 HTML 且表示答复窗体的正文的字符串。字符串限制为 32 KB。</span><span class="sxs-lookup"><span data-stu-id="1676f-p156">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
|`formData.attachments`|<span data-ttu-id="1676f-976">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="1676f-976">Array.&lt;Object&gt;</span></span>|<span data-ttu-id="1676f-977">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="1676f-977">&lt;optional&gt;</span></span>|<span data-ttu-id="1676f-978">JSON 对象（是文件或项目附件）的数组。</span><span class="sxs-lookup"><span data-stu-id="1676f-978">An array of JSON objects that are either file or item attachments.</span></span>|
|`formData.attachments.type`|<span data-ttu-id="1676f-979">String</span><span class="sxs-lookup"><span data-stu-id="1676f-979">String</span></span>||<span data-ttu-id="1676f-p157">指示附件的类型。必须是文件附件的 `file` 或项目附件的 `item`。</span><span class="sxs-lookup"><span data-stu-id="1676f-p157">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span>|
|`formData.attachments.name`|<span data-ttu-id="1676f-982">字符串</span><span class="sxs-lookup"><span data-stu-id="1676f-982">String</span></span>||<span data-ttu-id="1676f-983">一个包含附件的名称的字符串，最多包含 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="1676f-983">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
|`formData.attachments.url`|<span data-ttu-id="1676f-984">字符串</span><span class="sxs-lookup"><span data-stu-id="1676f-984">String</span></span>||<span data-ttu-id="1676f-p158">仅在将 `type` 设置为 `file` 时使用。文件的位置的 URI。</span><span class="sxs-lookup"><span data-stu-id="1676f-p158">Only used if `type` is set to `file`. The URI of the location for the file.</span></span>|
|`formData.attachments.isInline`|<span data-ttu-id="1676f-987">布尔</span><span class="sxs-lookup"><span data-stu-id="1676f-987">Boolean</span></span>||<span data-ttu-id="1676f-p159">仅在将 `type` 设置为 `file` 时使用。如果为 `true`，则表示附件将在邮件正文中内联显示，并且不应显示在附件列表中。</span><span class="sxs-lookup"><span data-stu-id="1676f-p159">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`formData.attachments.itemId`|<span data-ttu-id="1676f-990">String</span><span class="sxs-lookup"><span data-stu-id="1676f-990">String</span></span>||<span data-ttu-id="1676f-p160">仅在将 `type` 设置为 `item` 时使用。附件的 EWS 项目 ID。字符串最长为 100 个字符。</span><span class="sxs-lookup"><span data-stu-id="1676f-p160">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span>|
|`callback`|<span data-ttu-id="1676f-994">函数</span><span class="sxs-lookup"><span data-stu-id="1676f-994">function</span></span>|<span data-ttu-id="1676f-995">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="1676f-995">&lt;optional&gt;</span></span>|<span data-ttu-id="1676f-996">方法完成后，使用单个参数 `asyncResult`（一个 [AsyncResult](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="1676f-996">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="1676f-997">要求</span><span class="sxs-lookup"><span data-stu-id="1676f-997">Requirements</span></span>

|<span data-ttu-id="1676f-998">要求</span><span class="sxs-lookup"><span data-stu-id="1676f-998">Requirement</span></span>|<span data-ttu-id="1676f-999">值</span><span class="sxs-lookup"><span data-stu-id="1676f-999">Value</span></span>|
|---|---|
|[<span data-ttu-id="1676f-1000">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="1676f-1000">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="1676f-1001">1.0</span><span class="sxs-lookup"><span data-stu-id="1676f-1001">1.0</span></span>|
|[<span data-ttu-id="1676f-1002">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="1676f-1002">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="1676f-1003">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1676f-1003">ReadItem</span></span>|
|[<span data-ttu-id="1676f-1004">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="1676f-1004">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="1676f-1005">阅读</span><span class="sxs-lookup"><span data-stu-id="1676f-1005">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="1676f-1006">示例</span><span class="sxs-lookup"><span data-stu-id="1676f-1006">Examples</span></span>

<span data-ttu-id="1676f-1007">以下代码将一个字符串传递到 `displayReplyAllForm` 函数。</span><span class="sxs-lookup"><span data-stu-id="1676f-1007">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="1676f-1008">使用空白正文答复。</span><span class="sxs-lookup"><span data-stu-id="1676f-1008">Reply with an empty body.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="1676f-1009">仅使用正文答复。</span><span class="sxs-lookup"><span data-stu-id="1676f-1009">Reply with just a body.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="1676f-1010">使用正文和文件附件答复。</span><span class="sxs-lookup"><span data-stu-id="1676f-1010">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="1676f-1011">使用正文和项目附件答复。</span><span class="sxs-lookup"><span data-stu-id="1676f-1011">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="1676f-1012">使用正文、文件附件、项目附件和回调答复。</span><span class="sxs-lookup"><span data-stu-id="1676f-1012">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="displayreplyformformdata-callback"></a><span data-ttu-id="1676f-1013">displayReplyForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="1676f-1013">displayReplyForm(formData, [callback])</span></span>

<span data-ttu-id="1676f-1014">显示答复窗体，其中仅包括所选邮件的发件人或所选约会的组织者。</span><span class="sxs-lookup"><span data-stu-id="1676f-1014">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="1676f-1015">iOS 版 Outlook 或 Android 版 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="1676f-1015">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="1676f-1016">在 Outlook 网页版中，答复窗体显示为包含 3 列视图的弹出式窗体，以及包含 2 列或 1 列视图的弹出式窗体。</span><span class="sxs-lookup"><span data-stu-id="1676f-1016">In Outlook on the web, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="1676f-1017">如果任意字符串参数超出其限制，`displayReplyForm` 将引发异常。</span><span class="sxs-lookup"><span data-stu-id="1676f-1017">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="1676f-p161">如果附件已在 `formData.attachments` 参数中指定，Outlook 网页版和 Outlook 桌面版客户端会尝试下载所有附件，并将它们附加到答复窗体。如果无法添加任何附件，窗体 UI 中会显示错误。如果此操作是不可能完成的，系统不会抛出任何错误消息。</span><span class="sxs-lookup"><span data-stu-id="1676f-p161">When attachments are specified in the `formData.attachments` parameter, Outlook on the web and desktop clients attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="1676f-1021">Parameters</span><span class="sxs-lookup"><span data-stu-id="1676f-1021">Parameters</span></span>

|<span data-ttu-id="1676f-1022">名称</span><span class="sxs-lookup"><span data-stu-id="1676f-1022">Name</span></span>|<span data-ttu-id="1676f-1023">类型</span><span class="sxs-lookup"><span data-stu-id="1676f-1023">Type</span></span>|<span data-ttu-id="1676f-1024">属性</span><span class="sxs-lookup"><span data-stu-id="1676f-1024">Attributes</span></span>|<span data-ttu-id="1676f-1025">说明</span><span class="sxs-lookup"><span data-stu-id="1676f-1025">Description</span></span>|
|---|---|---|---|
|`formData`|<span data-ttu-id="1676f-1026">字符串 &#124; 对象</span><span class="sxs-lookup"><span data-stu-id="1676f-1026">String &#124; Object</span></span>||<span data-ttu-id="1676f-p162">一个包含文本和 HTML 且表示答复窗体的正文的字符串。字符串限制为 32 KB。</span><span class="sxs-lookup"><span data-stu-id="1676f-p162">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="1676f-1029">**或**</span><span class="sxs-lookup"><span data-stu-id="1676f-1029">**OR**</span></span><br/><span data-ttu-id="1676f-p163">包含正文或附件数据和回调函数的对象。对象定义如下。</span><span class="sxs-lookup"><span data-stu-id="1676f-p163">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span>|
|`formData.htmlBody`|<span data-ttu-id="1676f-1032">String</span><span class="sxs-lookup"><span data-stu-id="1676f-1032">String</span></span>|<span data-ttu-id="1676f-1033">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="1676f-1033">&lt;optional&gt;</span></span>|<span data-ttu-id="1676f-p164">一个包含文本和 HTML 且表示答复窗体的正文的字符串。字符串限制为 32 KB。</span><span class="sxs-lookup"><span data-stu-id="1676f-p164">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
|`formData.attachments`|<span data-ttu-id="1676f-1036">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="1676f-1036">Array.&lt;Object&gt;</span></span>|<span data-ttu-id="1676f-1037">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="1676f-1037">&lt;optional&gt;</span></span>|<span data-ttu-id="1676f-1038">JSON 对象（是文件或项目附件）的数组。</span><span class="sxs-lookup"><span data-stu-id="1676f-1038">An array of JSON objects that are either file or item attachments.</span></span>|
|`formData.attachments.type`|<span data-ttu-id="1676f-1039">String</span><span class="sxs-lookup"><span data-stu-id="1676f-1039">String</span></span>||<span data-ttu-id="1676f-p165">指示附件的类型。必须是文件附件的 `file` 或项目附件的 `item`。</span><span class="sxs-lookup"><span data-stu-id="1676f-p165">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span>|
|`formData.attachments.name`|<span data-ttu-id="1676f-1042">字符串</span><span class="sxs-lookup"><span data-stu-id="1676f-1042">String</span></span>||<span data-ttu-id="1676f-1043">一个包含附件的名称的字符串，最多包含 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="1676f-1043">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
|`formData.attachments.url`|<span data-ttu-id="1676f-1044">字符串</span><span class="sxs-lookup"><span data-stu-id="1676f-1044">String</span></span>||<span data-ttu-id="1676f-p166">仅在将 `type` 设置为 `file` 时使用。文件的位置的 URI。</span><span class="sxs-lookup"><span data-stu-id="1676f-p166">Only used if `type` is set to `file`. The URI of the location for the file.</span></span>|
|`formData.attachments.isInline`|<span data-ttu-id="1676f-1047">布尔</span><span class="sxs-lookup"><span data-stu-id="1676f-1047">Boolean</span></span>||<span data-ttu-id="1676f-p167">仅在将 `type` 设置为 `file` 时使用。如果为 `true`，则表示附件将在邮件正文中内联显示，并且不应显示在附件列表中。</span><span class="sxs-lookup"><span data-stu-id="1676f-p167">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`formData.attachments.itemId`|<span data-ttu-id="1676f-1050">String</span><span class="sxs-lookup"><span data-stu-id="1676f-1050">String</span></span>||<span data-ttu-id="1676f-p168">仅在将 `type` 设置为 `item` 时使用。附件的 EWS 项目 ID。字符串最长为 100 个字符。</span><span class="sxs-lookup"><span data-stu-id="1676f-p168">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span>|
|`callback`|<span data-ttu-id="1676f-1054">函数</span><span class="sxs-lookup"><span data-stu-id="1676f-1054">function</span></span>|<span data-ttu-id="1676f-1055">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="1676f-1055">&lt;optional&gt;</span></span>|<span data-ttu-id="1676f-1056">方法完成后，使用单个参数 `asyncResult`（一个 [AsyncResult](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="1676f-1056">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="1676f-1057">要求</span><span class="sxs-lookup"><span data-stu-id="1676f-1057">Requirements</span></span>

|<span data-ttu-id="1676f-1058">要求</span><span class="sxs-lookup"><span data-stu-id="1676f-1058">Requirement</span></span>|<span data-ttu-id="1676f-1059">值</span><span class="sxs-lookup"><span data-stu-id="1676f-1059">Value</span></span>|
|---|---|
|[<span data-ttu-id="1676f-1060">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="1676f-1060">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="1676f-1061">1.0</span><span class="sxs-lookup"><span data-stu-id="1676f-1061">1.0</span></span>|
|[<span data-ttu-id="1676f-1062">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="1676f-1062">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="1676f-1063">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1676f-1063">ReadItem</span></span>|
|[<span data-ttu-id="1676f-1064">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="1676f-1064">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="1676f-1065">阅读</span><span class="sxs-lookup"><span data-stu-id="1676f-1065">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="1676f-1066">示例</span><span class="sxs-lookup"><span data-stu-id="1676f-1066">Examples</span></span>

<span data-ttu-id="1676f-1067">以下代码将一个字符串传递到 `displayReplyForm` 函数。</span><span class="sxs-lookup"><span data-stu-id="1676f-1067">The following code passes a string to the `displayReplyForm` function.</span></span>

```js
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="1676f-1068">使用空白正文答复。</span><span class="sxs-lookup"><span data-stu-id="1676f-1068">Reply with an empty body.</span></span>

```js
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="1676f-1069">仅使用正文答复。</span><span class="sxs-lookup"><span data-stu-id="1676f-1069">Reply with just a body.</span></span>

```js
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="1676f-1070">使用正文和文件附件答复。</span><span class="sxs-lookup"><span data-stu-id="1676f-1070">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="1676f-1071">使用正文和项目附件答复。</span><span class="sxs-lookup"><span data-stu-id="1676f-1071">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="1676f-1072">使用正文、文件附件、项目附件和回调答复。</span><span class="sxs-lookup"><span data-stu-id="1676f-1072">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="getallinternetheadersasyncoptions-callback"></a><span data-ttu-id="1676f-1073">getAllInternetHeadersAsync （[options]，[callback]）</span><span class="sxs-lookup"><span data-stu-id="1676f-1073">getAllInternetHeadersAsync([options], [callback])</span></span>

<span data-ttu-id="1676f-1074">以字符串形式获取邮件的所有 internet 邮件头。</span><span class="sxs-lookup"><span data-stu-id="1676f-1074">Gets all the internet headers for the message as a string.</span></span> <span data-ttu-id="1676f-1075">仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="1676f-1075">Read mode only.</span></span>

##### <a name="parameters"></a><span data-ttu-id="1676f-1076">Parameters</span><span class="sxs-lookup"><span data-stu-id="1676f-1076">Parameters</span></span>

|<span data-ttu-id="1676f-1077">名称</span><span class="sxs-lookup"><span data-stu-id="1676f-1077">Name</span></span>|<span data-ttu-id="1676f-1078">类型</span><span class="sxs-lookup"><span data-stu-id="1676f-1078">Type</span></span>|<span data-ttu-id="1676f-1079">属性</span><span class="sxs-lookup"><span data-stu-id="1676f-1079">Attributes</span></span>|<span data-ttu-id="1676f-1080">说明</span><span class="sxs-lookup"><span data-stu-id="1676f-1080">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="1676f-1081">对象</span><span class="sxs-lookup"><span data-stu-id="1676f-1081">Object</span></span>|<span data-ttu-id="1676f-1082">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="1676f-1082">&lt;optional&gt;</span></span>|<span data-ttu-id="1676f-1083">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="1676f-1083">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="1676f-1084">对象</span><span class="sxs-lookup"><span data-stu-id="1676f-1084">Object</span></span>|<span data-ttu-id="1676f-1085">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="1676f-1085">&lt;optional&gt;</span></span>|<span data-ttu-id="1676f-1086">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="1676f-1086">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="1676f-1087">函数</span><span class="sxs-lookup"><span data-stu-id="1676f-1087">function</span></span>|<span data-ttu-id="1676f-1088">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="1676f-1088">&lt;optional&gt;</span></span>|<span data-ttu-id="1676f-1089">方法完成后，使用单个参数 `asyncResult`（一个 [AsyncResult](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="1676f-1089">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> <span data-ttu-id="1676f-1090">在成功的情况下，internet 标头数据在 asyncResult 属性中以字符串的形式提供。</span><span class="sxs-lookup"><span data-stu-id="1676f-1090">On success, the internet headers data is provided in the asyncResult.value property as a string.</span></span> <span data-ttu-id="1676f-1091">有关返回的字符串值的格式设置信息，请参阅[RFC 2183](https://tools.ietf.org/html/rfc2183) 。</span><span class="sxs-lookup"><span data-stu-id="1676f-1091">Refer to [RFC 2183](https://tools.ietf.org/html/rfc2183) for the formatting information of the returned string value.</span></span> <span data-ttu-id="1676f-1092">如果调用失败，asyncResult 属性将包含错误代码和失败原因。</span><span class="sxs-lookup"><span data-stu-id="1676f-1092">If the call fails, the asyncResult.error property will contain an error code with the reason for the failure.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="1676f-1093">Requirements</span><span class="sxs-lookup"><span data-stu-id="1676f-1093">Requirements</span></span>

|<span data-ttu-id="1676f-1094">要求</span><span class="sxs-lookup"><span data-stu-id="1676f-1094">Requirement</span></span>|<span data-ttu-id="1676f-1095">值</span><span class="sxs-lookup"><span data-stu-id="1676f-1095">Value</span></span>|
|---|---|
|[<span data-ttu-id="1676f-1096">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="1676f-1096">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="1676f-1097">1.8</span><span class="sxs-lookup"><span data-stu-id="1676f-1097">1.8</span></span>|
|[<span data-ttu-id="1676f-1098">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="1676f-1098">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="1676f-1099">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1676f-1099">ReadItem</span></span>|
|[<span data-ttu-id="1676f-1100">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="1676f-1100">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="1676f-1101">阅读</span><span class="sxs-lookup"><span data-stu-id="1676f-1101">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="1676f-1102">返回：</span><span class="sxs-lookup"><span data-stu-id="1676f-1102">Returns:</span></span>

<span data-ttu-id="1676f-1103">作为字符串的 internet 标头数据，根据[RFC 2183](https://tools.ietf.org/html/rfc2183)格式化。</span><span class="sxs-lookup"><span data-stu-id="1676f-1103">The internet headers data as a string formatted according to [RFC 2183](https://tools.ietf.org/html/rfc2183).</span></span>

<span data-ttu-id="1676f-1104">类型：字符串</span><span class="sxs-lookup"><span data-stu-id="1676f-1104">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="1676f-1105">示例</span><span class="sxs-lookup"><span data-stu-id="1676f-1105">Example</span></span>

```js
// Get the internet headers related to the mail.
Office.context.mailbox.item.getAllInternetHeadersAsync(
  function(asyncResult) {
    if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
      console.log(asyncResult.value);
    } else {
      if (asyncResult.error.code == 9020) {
        // GenericResponseError returned when there is no context.
        // Treat as no context.
      } else {
        // Handle the error.
      }
    }
  }
);
```

<br>

---
---

#### <a name="getattachmentcontentasyncattachmentid-options-callback--attachmentcontentjavascriptapioutlookofficeattachmentcontentviewoutlook-js-18"></a><span data-ttu-id="1676f-1106">getAttachmentContentAsync （attachmentId，[options]，[callback]）→ [AttachmentContent](/javascript/api/outlook/office.attachmentcontent?view=outlook-js-1.8)</span><span class="sxs-lookup"><span data-stu-id="1676f-1106">getAttachmentContentAsync(attachmentId, [options], [callback]) → [AttachmentContent](/javascript/api/outlook/office.attachmentcontent?view=outlook-js-1.8)</span></span>

<span data-ttu-id="1676f-1107">从邮件或约会中获取指定附件并将其作为`AttachmentContent`对象返回。</span><span class="sxs-lookup"><span data-stu-id="1676f-1107">Gets the specified attachment from a message or appointment and returns it as an `AttachmentContent` object.</span></span>

<span data-ttu-id="1676f-1108">该`getAttachmentContentAsync`方法从项目中获取具有指定标识符的附件。</span><span class="sxs-lookup"><span data-stu-id="1676f-1108">The `getAttachmentContentAsync` method gets the attachment with the specified identifier from the item.</span></span> <span data-ttu-id="1676f-1109">作为一种最佳做法，您应使用标识符在与`getAttachmentsAsync` or `item.attachments`调用一起检索到会话的同一会话中检索附件。</span><span class="sxs-lookup"><span data-stu-id="1676f-1109">As a best practice, you should use the identifier to retrieve an attachment in the same session that the attachmentIds were retrieved with the `getAttachmentsAsync` or `item.attachments` call.</span></span> <span data-ttu-id="1676f-1110">在 Outlook 网页版和移动设备上，附件标识符只在同一个会话中才有效。</span><span class="sxs-lookup"><span data-stu-id="1676f-1110">In Outlook on the web and mobile devices, the attachment identifier is valid only within the same session.</span></span> <span data-ttu-id="1676f-1111">当用户关闭应用程序时，或者如果用户开始撰写内嵌窗体，随后弹出窗体以继续在单独的窗口中。</span><span class="sxs-lookup"><span data-stu-id="1676f-1111">A session is over when the user closes the app, or if the user starts composing an inline form then subsequently pops out the form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="1676f-1112">Parameters</span><span class="sxs-lookup"><span data-stu-id="1676f-1112">Parameters</span></span>

|<span data-ttu-id="1676f-1113">名称</span><span class="sxs-lookup"><span data-stu-id="1676f-1113">Name</span></span>|<span data-ttu-id="1676f-1114">类型</span><span class="sxs-lookup"><span data-stu-id="1676f-1114">Type</span></span>|<span data-ttu-id="1676f-1115">属性</span><span class="sxs-lookup"><span data-stu-id="1676f-1115">Attributes</span></span>|<span data-ttu-id="1676f-1116">说明</span><span class="sxs-lookup"><span data-stu-id="1676f-1116">Description</span></span>|
|---|---|---|---|
|`attachmentId`|<span data-ttu-id="1676f-1117">String</span><span class="sxs-lookup"><span data-stu-id="1676f-1117">String</span></span>||<span data-ttu-id="1676f-1118">要获取的附件的标识符。</span><span class="sxs-lookup"><span data-stu-id="1676f-1118">The identifier of the attachment you want to get.</span></span>|
|`options`|<span data-ttu-id="1676f-1119">对象</span><span class="sxs-lookup"><span data-stu-id="1676f-1119">Object</span></span>|<span data-ttu-id="1676f-1120">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="1676f-1120">&lt;optional&gt;</span></span>|<span data-ttu-id="1676f-1121">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="1676f-1121">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="1676f-1122">对象</span><span class="sxs-lookup"><span data-stu-id="1676f-1122">Object</span></span>|<span data-ttu-id="1676f-1123">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="1676f-1123">&lt;optional&gt;</span></span>|<span data-ttu-id="1676f-1124">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="1676f-1124">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="1676f-1125">函数</span><span class="sxs-lookup"><span data-stu-id="1676f-1125">function</span></span>|<span data-ttu-id="1676f-1126">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="1676f-1126">&lt;optional&gt;</span></span>|<span data-ttu-id="1676f-1127">方法完成后，使用单个参数 `asyncResult`（一个 [AsyncResult](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="1676f-1127">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="1676f-1128">要求</span><span class="sxs-lookup"><span data-stu-id="1676f-1128">Requirements</span></span>

|<span data-ttu-id="1676f-1129">要求</span><span class="sxs-lookup"><span data-stu-id="1676f-1129">Requirement</span></span>|<span data-ttu-id="1676f-1130">值</span><span class="sxs-lookup"><span data-stu-id="1676f-1130">Value</span></span>|
|---|---|
|[<span data-ttu-id="1676f-1131">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="1676f-1131">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="1676f-1132">1.8</span><span class="sxs-lookup"><span data-stu-id="1676f-1132">1.8</span></span>|
|[<span data-ttu-id="1676f-1133">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="1676f-1133">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="1676f-1134">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1676f-1134">ReadItem</span></span>|
|[<span data-ttu-id="1676f-1135">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="1676f-1135">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="1676f-1136">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="1676f-1136">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="1676f-1137">返回：</span><span class="sxs-lookup"><span data-stu-id="1676f-1137">Returns:</span></span>

<span data-ttu-id="1676f-1138">类型： [AttachmentContent](/javascript/api/outlook/office.attachmentcontent?view=outlook-js-1.8)</span><span class="sxs-lookup"><span data-stu-id="1676f-1138">Type: [AttachmentContent](/javascript/api/outlook/office.attachmentcontent?view=outlook-js-1.8)</span></span>

##### <a name="example"></a><span data-ttu-id="1676f-1139">示例</span><span class="sxs-lookup"><span data-stu-id="1676f-1139">Example</span></span>

```js
var item = Office.context.mailbox.item;
var listOfAttachments = [];
var options = {asyncContext: {currentItem: item}};
item.getAttachmentsAsync(options, callback);

function callback(result) {
  if (result.value.length > 0) {
    for (i = 0 ; i < result.value.length ; i++) {
      result.asyncContext.currentItem.getAttachmentContentAsync(result.value[i].id, handleAttachmentsCallback);
    }
  }
}

function handleAttachmentsCallback(result) {
  // Parse string to be a url, an .eml file, a base64-encoded string, or an .icalendar file.
  switch (result.value.format) {
    case Office.MailboxEnums.AttachmentContentFormat.Base64:
      // Handle file attachment.
      break;
    case Office.MailboxEnums.AttachmentContentFormat.Eml:
      // Handle email item attachment.
      break;
    case Office.MailboxEnums.AttachmentContentFormat.ICalendar:
      // Handle .icalender attachment.
      break;
    case Office.MailboxEnums.AttachmentContentFormat.Url:
      // Handle cloud attachment.
      break;
    default:
      // Handle attachment formats that are not supported.
  }
}
```

<br>

---
---

#### <a name="getattachmentsasyncoptions-callback--arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetailsviewoutlook-js-18"></a><span data-ttu-id="1676f-1140">getAttachmentsAsync （[options]，[callback]）→ Array. <[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.8)></span><span class="sxs-lookup"><span data-stu-id="1676f-1140">getAttachmentsAsync([options], [callback]) → Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.8)></span></span>

<span data-ttu-id="1676f-1141">以数组的形式获取项目的附件。</span><span class="sxs-lookup"><span data-stu-id="1676f-1141">Gets the item's attachments as an array.</span></span> <span data-ttu-id="1676f-1142">仅限撰写模式。</span><span class="sxs-lookup"><span data-stu-id="1676f-1142">Compose mode only.</span></span>

##### <a name="parameters"></a><span data-ttu-id="1676f-1143">Parameters</span><span class="sxs-lookup"><span data-stu-id="1676f-1143">Parameters</span></span>

|<span data-ttu-id="1676f-1144">名称</span><span class="sxs-lookup"><span data-stu-id="1676f-1144">Name</span></span>|<span data-ttu-id="1676f-1145">类型</span><span class="sxs-lookup"><span data-stu-id="1676f-1145">Type</span></span>|<span data-ttu-id="1676f-1146">属性</span><span class="sxs-lookup"><span data-stu-id="1676f-1146">Attributes</span></span>|<span data-ttu-id="1676f-1147">说明</span><span class="sxs-lookup"><span data-stu-id="1676f-1147">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="1676f-1148">对象</span><span class="sxs-lookup"><span data-stu-id="1676f-1148">Object</span></span>|<span data-ttu-id="1676f-1149">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="1676f-1149">&lt;optional&gt;</span></span>|<span data-ttu-id="1676f-1150">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="1676f-1150">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="1676f-1151">对象</span><span class="sxs-lookup"><span data-stu-id="1676f-1151">Object</span></span>|<span data-ttu-id="1676f-1152">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="1676f-1152">&lt;optional&gt;</span></span>|<span data-ttu-id="1676f-1153">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="1676f-1153">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="1676f-1154">函数</span><span class="sxs-lookup"><span data-stu-id="1676f-1154">function</span></span>|<span data-ttu-id="1676f-1155">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="1676f-1155">&lt;optional&gt;</span></span>|<span data-ttu-id="1676f-1156">方法完成后，使用单个参数 `asyncResult`（一个 [AsyncResult](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="1676f-1156">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="1676f-1157">Requirements</span><span class="sxs-lookup"><span data-stu-id="1676f-1157">Requirements</span></span>

|<span data-ttu-id="1676f-1158">要求</span><span class="sxs-lookup"><span data-stu-id="1676f-1158">Requirement</span></span>|<span data-ttu-id="1676f-1159">值</span><span class="sxs-lookup"><span data-stu-id="1676f-1159">Value</span></span>|
|---|---|
|[<span data-ttu-id="1676f-1160">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="1676f-1160">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="1676f-1161">1.8</span><span class="sxs-lookup"><span data-stu-id="1676f-1161">1.8</span></span>|
|[<span data-ttu-id="1676f-1162">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="1676f-1162">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="1676f-1163">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1676f-1163">ReadItem</span></span>|
|[<span data-ttu-id="1676f-1164">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="1676f-1164">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="1676f-1165">撰写</span><span class="sxs-lookup"><span data-stu-id="1676f-1165">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="1676f-1166">返回：</span><span class="sxs-lookup"><span data-stu-id="1676f-1166">Returns:</span></span>

<span data-ttu-id="1676f-1167">类型： Array. <[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.8)></span><span class="sxs-lookup"><span data-stu-id="1676f-1167">Type: Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.8)></span></span>

##### <a name="example"></a><span data-ttu-id="1676f-1168">示例</span><span class="sxs-lookup"><span data-stu-id="1676f-1168">Example</span></span>

<span data-ttu-id="1676f-1169">下面的示例将生成一个 HTML 字符串，其中包含当前项目上所有附件的详细信息。</span><span class="sxs-lookup"><span data-stu-id="1676f-1169">The following example builds an HTML string with details of all attachments on the current item.</span></span>

```js
var item = Office.context.mailbox.item;
var outputString = "";
item.getAttachmentsAsync(callback);

function callback(result) {
  if (result.value.length > 0) {
    for (i = 0 ; i < result.value.length ; i++) {
      var attachment = result.value [i];
      outputString += "<BR>" + i + ". Name: ";
      outputString += attachment.name;
      outputString += "<BR>ID: " + attachment.id;
      outputString += "<BR>contentType: " + attachment.contentType;
      outputString += "<BR>size: " + attachment.size;
      outputString += "<BR>attachmentType: " + attachment.attachmentType;
      outputString += "<BR>isInline: " + attachment.isInline;
    }
  }
}
```

<br>

---
---

#### <a name="getentities--entitiesjavascriptapioutlookofficeentitiesviewoutlook-js-18"></a><span data-ttu-id="1676f-1170">getEntities() → {[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.8)}</span><span class="sxs-lookup"><span data-stu-id="1676f-1170">getEntities() → {[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.8)}</span></span>

<span data-ttu-id="1676f-1171">获取在所选项目的正文中找到的实体。</span><span class="sxs-lookup"><span data-stu-id="1676f-1171">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="1676f-1172">iOS 版 Outlook 或 Android 版 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="1676f-1172">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="1676f-1173">要求</span><span class="sxs-lookup"><span data-stu-id="1676f-1173">Requirements</span></span>

|<span data-ttu-id="1676f-1174">要求</span><span class="sxs-lookup"><span data-stu-id="1676f-1174">Requirement</span></span>|<span data-ttu-id="1676f-1175">值</span><span class="sxs-lookup"><span data-stu-id="1676f-1175">Value</span></span>|
|---|---|
|[<span data-ttu-id="1676f-1176">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="1676f-1176">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="1676f-1177">1.0</span><span class="sxs-lookup"><span data-stu-id="1676f-1177">1.0</span></span>|
|[<span data-ttu-id="1676f-1178">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="1676f-1178">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="1676f-1179">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1676f-1179">ReadItem</span></span>|
|[<span data-ttu-id="1676f-1180">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="1676f-1180">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="1676f-1181">阅读</span><span class="sxs-lookup"><span data-stu-id="1676f-1181">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="1676f-1182">返回：</span><span class="sxs-lookup"><span data-stu-id="1676f-1182">Returns:</span></span>

<span data-ttu-id="1676f-1183">类型：[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.8)</span><span class="sxs-lookup"><span data-stu-id="1676f-1183">Type: [Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.8)</span></span>

##### <a name="example"></a><span data-ttu-id="1676f-1184">示例</span><span class="sxs-lookup"><span data-stu-id="1676f-1184">Example</span></span>

<span data-ttu-id="1676f-1185">以下示例访问当前项目的正文中的联系人实体。</span><span class="sxs-lookup"><span data-stu-id="1676f-1185">The following example accesses the contacts entities in the current item's body.</span></span>

```js
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

<br>

---
---

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlookofficecontactviewoutlook-js-18meetingsuggestionjavascriptapioutlookofficemeetingsuggestionviewoutlook-js-18phonenumberjavascriptapioutlookofficephonenumberviewoutlook-js-18tasksuggestionjavascriptapioutlookofficetasksuggestionviewoutlook-js-18"></a><span data-ttu-id="1676f-1186">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.8)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.8)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.8)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.8))>}</span><span class="sxs-lookup"><span data-stu-id="1676f-1186">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.8)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.8)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.8)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.8))>}</span></span>

<span data-ttu-id="1676f-1187">获取所选项目的正文中找到的指定实体类型的所有实体的数组。</span><span class="sxs-lookup"><span data-stu-id="1676f-1187">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="1676f-1188">iOS 版 Outlook 或 Android 版 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="1676f-1188">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="1676f-1189">Parameters</span><span class="sxs-lookup"><span data-stu-id="1676f-1189">Parameters</span></span>

|<span data-ttu-id="1676f-1190">名称</span><span class="sxs-lookup"><span data-stu-id="1676f-1190">Name</span></span>|<span data-ttu-id="1676f-1191">类型</span><span class="sxs-lookup"><span data-stu-id="1676f-1191">Type</span></span>|<span data-ttu-id="1676f-1192">说明</span><span class="sxs-lookup"><span data-stu-id="1676f-1192">Description</span></span>|
|---|---|---|
|`entityType`|[<span data-ttu-id="1676f-1193">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="1676f-1193">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook/office.mailboxenums.entitytype?view=outlook-js-1.8)|<span data-ttu-id="1676f-1194">EntityType 枚举值之一。</span><span class="sxs-lookup"><span data-stu-id="1676f-1194">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="1676f-1195">Requirements</span><span class="sxs-lookup"><span data-stu-id="1676f-1195">Requirements</span></span>

|<span data-ttu-id="1676f-1196">要求</span><span class="sxs-lookup"><span data-stu-id="1676f-1196">Requirement</span></span>|<span data-ttu-id="1676f-1197">值</span><span class="sxs-lookup"><span data-stu-id="1676f-1197">Value</span></span>|
|---|---|
|[<span data-ttu-id="1676f-1198">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="1676f-1198">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="1676f-1199">1.0</span><span class="sxs-lookup"><span data-stu-id="1676f-1199">1.0</span></span>|
|[<span data-ttu-id="1676f-1200">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="1676f-1200">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="1676f-1201">受限</span><span class="sxs-lookup"><span data-stu-id="1676f-1201">Restricted</span></span>|
|[<span data-ttu-id="1676f-1202">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="1676f-1202">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="1676f-1203">阅读</span><span class="sxs-lookup"><span data-stu-id="1676f-1203">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="1676f-1204">返回：</span><span class="sxs-lookup"><span data-stu-id="1676f-1204">Returns:</span></span>

<span data-ttu-id="1676f-1205">如果在 `entityType` 中传递的值不是 `EntityType` 枚举的有效成员，该方法返回 null。</span><span class="sxs-lookup"><span data-stu-id="1676f-1205">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="1676f-1206">如果指定类型的任何实体都不存在于该项目的正文中，该方法将返回空数组。</span><span class="sxs-lookup"><span data-stu-id="1676f-1206">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="1676f-1207">否则，返回的数组中对象的类型取决于 `entityType` 参数中请求实体的类型。</span><span class="sxs-lookup"><span data-stu-id="1676f-1207">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="1676f-1208">当使用此方法的最低权限级别**受限**时，某些实体类型需要 **ReadItem** 才能进行访问，如下表中所指定。</span><span class="sxs-lookup"><span data-stu-id="1676f-1208">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

|<span data-ttu-id="1676f-1209">`entityType` 的值</span><span class="sxs-lookup"><span data-stu-id="1676f-1209">Value of `entityType`</span></span>|<span data-ttu-id="1676f-1210">返回的数组中对象的类型</span><span class="sxs-lookup"><span data-stu-id="1676f-1210">Type of objects in returned array</span></span>|<span data-ttu-id="1676f-1211">所需权限级别</span><span class="sxs-lookup"><span data-stu-id="1676f-1211">Required Permission Level</span></span>|
|---|---|---|
|`Address`|<span data-ttu-id="1676f-1212">String</span><span class="sxs-lookup"><span data-stu-id="1676f-1212">String</span></span>|<span data-ttu-id="1676f-1213">**受限**</span><span class="sxs-lookup"><span data-stu-id="1676f-1213">**Restricted**</span></span>|
|`Contact`|<span data-ttu-id="1676f-1214">Contact</span><span class="sxs-lookup"><span data-stu-id="1676f-1214">Contact</span></span>|<span data-ttu-id="1676f-1215">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="1676f-1215">**ReadItem**</span></span>|
|`EmailAddress`|<span data-ttu-id="1676f-1216">String</span><span class="sxs-lookup"><span data-stu-id="1676f-1216">String</span></span>|<span data-ttu-id="1676f-1217">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="1676f-1217">**ReadItem**</span></span>|
|`MeetingSuggestion`|<span data-ttu-id="1676f-1218">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="1676f-1218">MeetingSuggestion</span></span>|<span data-ttu-id="1676f-1219">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="1676f-1219">**ReadItem**</span></span>|
|`PhoneNumber`|<span data-ttu-id="1676f-1220">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="1676f-1220">PhoneNumber</span></span>|<span data-ttu-id="1676f-1221">**受限**</span><span class="sxs-lookup"><span data-stu-id="1676f-1221">**Restricted**</span></span>|
|`TaskSuggestion`|<span data-ttu-id="1676f-1222">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="1676f-1222">TaskSuggestion</span></span>|<span data-ttu-id="1676f-1223">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="1676f-1223">**ReadItem**</span></span>|
|`URL`|<span data-ttu-id="1676f-1224">String</span><span class="sxs-lookup"><span data-stu-id="1676f-1224">String</span></span>|<span data-ttu-id="1676f-1225">**受限**</span><span class="sxs-lookup"><span data-stu-id="1676f-1225">**Restricted**</span></span>|

<span data-ttu-id="1676f-1226">类型：Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.8)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.8)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.8)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.8))></span><span class="sxs-lookup"><span data-stu-id="1676f-1226">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.8)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.8)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.8)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.8))></span></span>

##### <a name="example"></a><span data-ttu-id="1676f-1227">示例</span><span class="sxs-lookup"><span data-stu-id="1676f-1227">Example</span></span>

<span data-ttu-id="1676f-1228">以下示例显示了如何访问表示当前项目的正文中的邮政地址的字符串数组。</span><span class="sxs-lookup"><span data-stu-id="1676f-1228">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

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
};
```

<br>

---
---

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlookofficecontactviewoutlook-js-18meetingsuggestionjavascriptapioutlookofficemeetingsuggestionviewoutlook-js-18phonenumberjavascriptapioutlookofficephonenumberviewoutlook-js-18tasksuggestionjavascriptapioutlookofficetasksuggestionviewoutlook-js-18"></a><span data-ttu-id="1676f-1229">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.8)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.8)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.8)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.8))>}</span><span class="sxs-lookup"><span data-stu-id="1676f-1229">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.8)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.8)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.8)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.8))>}</span></span>

<span data-ttu-id="1676f-1230">返回传递清单 XML 文件中定义的命名筛选器的所选项目中的已知实体。</span><span class="sxs-lookup"><span data-stu-id="1676f-1230">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="1676f-1231">iOS 版 Outlook 或 Android 版 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="1676f-1231">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="1676f-1232">`getFilteredEntitiesByName` 方法返回匹配在具有指定 `FilterName` 元素值的清单 XML 文件中的 [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) 规则元素中定义的正则表达式的实体。</span><span class="sxs-lookup"><span data-stu-id="1676f-1232">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="1676f-1233">参数</span><span class="sxs-lookup"><span data-stu-id="1676f-1233">Parameters</span></span>

|<span data-ttu-id="1676f-1234">名称</span><span class="sxs-lookup"><span data-stu-id="1676f-1234">Name</span></span>|<span data-ttu-id="1676f-1235">类型</span><span class="sxs-lookup"><span data-stu-id="1676f-1235">Type</span></span>|<span data-ttu-id="1676f-1236">说明</span><span class="sxs-lookup"><span data-stu-id="1676f-1236">Description</span></span>|
|---|---|---|
|`name`|<span data-ttu-id="1676f-1237">字符串</span><span class="sxs-lookup"><span data-stu-id="1676f-1237">String</span></span>|<span data-ttu-id="1676f-1238">定义筛选器匹配的 `ItemHasKnownEntity` 规则元素的名称。</span><span class="sxs-lookup"><span data-stu-id="1676f-1238">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="1676f-1239">要求</span><span class="sxs-lookup"><span data-stu-id="1676f-1239">Requirements</span></span>

|<span data-ttu-id="1676f-1240">要求</span><span class="sxs-lookup"><span data-stu-id="1676f-1240">Requirement</span></span>|<span data-ttu-id="1676f-1241">值</span><span class="sxs-lookup"><span data-stu-id="1676f-1241">Value</span></span>|
|---|---|
|[<span data-ttu-id="1676f-1242">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="1676f-1242">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="1676f-1243">1.0</span><span class="sxs-lookup"><span data-stu-id="1676f-1243">1.0</span></span>|
|[<span data-ttu-id="1676f-1244">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="1676f-1244">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="1676f-1245">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1676f-1245">ReadItem</span></span>|
|[<span data-ttu-id="1676f-1246">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="1676f-1246">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="1676f-1247">阅读</span><span class="sxs-lookup"><span data-stu-id="1676f-1247">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="1676f-1248">返回：</span><span class="sxs-lookup"><span data-stu-id="1676f-1248">Returns:</span></span>

<span data-ttu-id="1676f-p174">如果具有匹配 `name` 参数的 `FilterName` 元素值的清单中没有任何 `ItemHasKnownEntity` 元素，则该方法返回 `null`。如果 `name` 参数匹配清单中的 `ItemHasKnownEntity` 元素，但在匹配的当前项目中没有实体，则该方法返回一个空数组。</span><span class="sxs-lookup"><span data-stu-id="1676f-p174">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="1676f-1251">类型：Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.8)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.8)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.8)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.8))></span><span class="sxs-lookup"><span data-stu-id="1676f-1251">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.8)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.8)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.8)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.8))></span></span>

<br>

---
---

#### <a name="getitemidasyncoptions-callback"></a><span data-ttu-id="1676f-1252">getItemIdAsync （[options]，回拨）</span><span class="sxs-lookup"><span data-stu-id="1676f-1252">getItemIdAsync([options], callback)</span></span>

<span data-ttu-id="1676f-1253">异步获取已保存项的 ID。</span><span class="sxs-lookup"><span data-stu-id="1676f-1253">Asynchronously gets the ID of a saved item.</span></span> <span data-ttu-id="1676f-1254">仅限撰写模式。</span><span class="sxs-lookup"><span data-stu-id="1676f-1254">Compose mode only.</span></span>

<span data-ttu-id="1676f-1255">调用此方法时，此方法通过回调方法返回项 ID。</span><span class="sxs-lookup"><span data-stu-id="1676f-1255">When invoked, this method returns the item ID via the callback method.</span></span>

> [!NOTE]
> <span data-ttu-id="1676f-1256">如果你的外接程序`getItemIdAsync`对撰写模式中的项（例如，要获取`itemId`使用 EWS 或 REST API 的使用）调用，请注意，当 Outlook 处于缓存模式下时，可能需要一段时间才能将项目同步到服务器。</span><span class="sxs-lookup"><span data-stu-id="1676f-1256">If your add-in calls `getItemIdAsync` on an item in compose mode (e.g., to get an `itemId` to use with EWS or the REST API), be aware that when Outlook is in cached mode, it may take some time before the item is synced to the server.</span></span> <span data-ttu-id="1676f-1257">在同步项目之前，无法识别`itemId`该项目并使用它将返回错误。</span><span class="sxs-lookup"><span data-stu-id="1676f-1257">Until the item is synced, the `itemId` is not recognized and using it returns an error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="1676f-1258">Parameters</span><span class="sxs-lookup"><span data-stu-id="1676f-1258">Parameters</span></span>

|<span data-ttu-id="1676f-1259">名称</span><span class="sxs-lookup"><span data-stu-id="1676f-1259">Name</span></span>|<span data-ttu-id="1676f-1260">类型</span><span class="sxs-lookup"><span data-stu-id="1676f-1260">Type</span></span>|<span data-ttu-id="1676f-1261">属性</span><span class="sxs-lookup"><span data-stu-id="1676f-1261">Attributes</span></span>|<span data-ttu-id="1676f-1262">说明</span><span class="sxs-lookup"><span data-stu-id="1676f-1262">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="1676f-1263">Object</span><span class="sxs-lookup"><span data-stu-id="1676f-1263">Object</span></span>|<span data-ttu-id="1676f-1264">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="1676f-1264">&lt;optional&gt;</span></span>|<span data-ttu-id="1676f-1265">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="1676f-1265">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="1676f-1266">Object</span><span class="sxs-lookup"><span data-stu-id="1676f-1266">Object</span></span>|<span data-ttu-id="1676f-1267">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="1676f-1267">&lt;optional&gt;</span></span>|<span data-ttu-id="1676f-1268">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="1676f-1268">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="1676f-1269">函数</span><span class="sxs-lookup"><span data-stu-id="1676f-1269">function</span></span>||<span data-ttu-id="1676f-1270">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="1676f-1270">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="1676f-1271">如果成功，该项目标识符将在 `asyncResult.value` 属性中提供。</span><span class="sxs-lookup"><span data-stu-id="1676f-1271">On success, the item identifier is provided in the `asyncResult.value` property.</span></span>|

##### <a name="errors"></a><span data-ttu-id="1676f-1272">错误</span><span class="sxs-lookup"><span data-stu-id="1676f-1272">Errors</span></span>

|<span data-ttu-id="1676f-1273">错误代码</span><span class="sxs-lookup"><span data-stu-id="1676f-1273">Error code</span></span>|<span data-ttu-id="1676f-1274">说明</span><span class="sxs-lookup"><span data-stu-id="1676f-1274">Description</span></span>|
|------------|-------------|
|`ItemNotSaved`|<span data-ttu-id="1676f-1275">在保存项目之前，无法检索此 id。</span><span class="sxs-lookup"><span data-stu-id="1676f-1275">The id can't be retrieved until the item is saved.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="1676f-1276">Requirements</span><span class="sxs-lookup"><span data-stu-id="1676f-1276">Requirements</span></span>

|<span data-ttu-id="1676f-1277">要求</span><span class="sxs-lookup"><span data-stu-id="1676f-1277">Requirement</span></span>|<span data-ttu-id="1676f-1278">值</span><span class="sxs-lookup"><span data-stu-id="1676f-1278">Value</span></span>|
|---|---|
|[<span data-ttu-id="1676f-1279">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="1676f-1279">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="1676f-1280">1.8</span><span class="sxs-lookup"><span data-stu-id="1676f-1280">1.8</span></span>|
|[<span data-ttu-id="1676f-1281">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="1676f-1281">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="1676f-1282">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1676f-1282">ReadItem</span></span>|
|[<span data-ttu-id="1676f-1283">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="1676f-1283">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="1676f-1284">撰写</span><span class="sxs-lookup"><span data-stu-id="1676f-1284">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="1676f-1285">示例</span><span class="sxs-lookup"><span data-stu-id="1676f-1285">Examples</span></span>

```js
Office.context.mailbox.item.getItemIdAsync(
  function callback(result) {
    // Process the result.
  });
```

<span data-ttu-id="1676f-1286">下面的示例演示传递给回调函数`result`的参数的结构。</span><span class="sxs-lookup"><span data-stu-id="1676f-1286">The following example shows the structure of the `result` parameter that's passed to the callback function.</span></span> <span data-ttu-id="1676f-1287">`value`属性包含项 ID。</span><span class="sxs-lookup"><span data-stu-id="1676f-1287">The `value` property contains the item ID.</span></span>

```json
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

<br>

---
---

#### <a name="getregexmatches--object"></a><span data-ttu-id="1676f-1288">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="1676f-1288">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="1676f-1289">返回所选项目中匹配在清单 XML 文件中定义的正则表达式的字符串值。</span><span class="sxs-lookup"><span data-stu-id="1676f-1289">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="1676f-1290">iOS 版 Outlook 或 Android 版 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="1676f-1290">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="1676f-p178">`getRegExMatches` 方法返回匹配在清单 XML 文件中的每个 `ItemHasRegularExpressionMatch` 或 `ItemHasKnownEntity` 规则元素中定义的正则表达式的字符串。对于 `ItemHasRegularExpressionMatch` 规则，匹配字符串必须发生在该规则指定的项目的属性中。`PropertyName` 简单类型定义支持的属性。</span><span class="sxs-lookup"><span data-stu-id="1676f-p178">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="1676f-1294">例如，考虑一个外接程序清单具有以下 `Rule` 元素：</span><span class="sxs-lookup"><span data-stu-id="1676f-1294">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="1676f-1295">从 `getRegExMatches` 返回的对象应有两个属性：`fruits` 和 `veggies`。</span><span class="sxs-lookup"><span data-stu-id="1676f-1295">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="1676f-p179">如果在项目的正文属性上指定 `ItemHasRegularExpressionMatch` 规则，则正则表达式应进一步筛选正文，不应尝试返回该项目的整个正文。使用正则表达式（如 `.*`）获取项目的整个正文并不总是返回预期的结果。而是使用 [`Body.getAsync`](/javascript/api/outlook/office.body?view=outlook-js-1.8#getasync-coerciontype--options--callback-) 方法检索整个正文。</span><span class="sxs-lookup"><span data-stu-id="1676f-p179">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook/office.body?view=outlook-js-1.8#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="1676f-1299">Requirements</span><span class="sxs-lookup"><span data-stu-id="1676f-1299">Requirements</span></span>

|<span data-ttu-id="1676f-1300">要求</span><span class="sxs-lookup"><span data-stu-id="1676f-1300">Requirement</span></span>|<span data-ttu-id="1676f-1301">值</span><span class="sxs-lookup"><span data-stu-id="1676f-1301">Value</span></span>|
|---|---|
|[<span data-ttu-id="1676f-1302">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="1676f-1302">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="1676f-1303">1.0</span><span class="sxs-lookup"><span data-stu-id="1676f-1303">1.0</span></span>|
|[<span data-ttu-id="1676f-1304">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="1676f-1304">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="1676f-1305">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1676f-1305">ReadItem</span></span>|
|[<span data-ttu-id="1676f-1306">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="1676f-1306">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="1676f-1307">阅读</span><span class="sxs-lookup"><span data-stu-id="1676f-1307">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="1676f-1308">返回：</span><span class="sxs-lookup"><span data-stu-id="1676f-1308">Returns:</span></span>

<span data-ttu-id="1676f-p180">一个包含匹配在清单 XML 文件中定义的正则表达式的字符串数组的对象。每个数组的名称等于匹配 `ItemHasRegularExpressionMatch` 规则的 `RegExName` 属性或匹配 `ItemHasKnownEntity` 规则的 `FilterName` 属性的相应值。</span><span class="sxs-lookup"><span data-stu-id="1676f-p180">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<dl class="param-type"><span data-ttu-id="1676f-1311">

<dt>
类型</dt>


</span><span class="sxs-lookup"><span data-stu-id="1676f-1311">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="1676f-1312">对象</span><span class="sxs-lookup"><span data-stu-id="1676f-1312">Object</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="1676f-1313">示例</span><span class="sxs-lookup"><span data-stu-id="1676f-1313">Example</span></span>

<span data-ttu-id="1676f-1314">以下示例显示了如何访问正则表达式规则元素 `fruits` 和 `veggies` 的匹配项的数组，这些元素在清单中指定。</span><span class="sxs-lookup"><span data-stu-id="1676f-1314">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```js
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veggies = allMatches.veggies;
```

<br>

---
---

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="1676f-1315">getRegExMatchesByName(name) → (nullable) {Array.<String>}</span><span class="sxs-lookup"><span data-stu-id="1676f-1315">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="1676f-1316">返回匹配在清单 XML 文件中定义的命名正则表达式的所选项目中的字符串值。</span><span class="sxs-lookup"><span data-stu-id="1676f-1316">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="1676f-1317">iOS 版 Outlook 或 Android 版 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="1676f-1317">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="1676f-1318">`getRegExMatchesByName` 方法返回匹配在具有指定 `RegExName` 元素值的清单 XML 文件中的 `ItemHasRegularExpressionMatch` 规则元素中定义的正则表达式的字符串。</span><span class="sxs-lookup"><span data-stu-id="1676f-1318">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="1676f-p181">如果在项目的正文属性上指定 `ItemHasRegularExpressionMatch` 规则，则正则表达式应进一步筛选正文，不应尝试返回该项目的整个正文。使用正则表达式（如 `.*`）获取项目的整个正文并不总是返回预期的结果。</span><span class="sxs-lookup"><span data-stu-id="1676f-p181">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="1676f-1321">参数</span><span class="sxs-lookup"><span data-stu-id="1676f-1321">Parameters</span></span>

|<span data-ttu-id="1676f-1322">名称</span><span class="sxs-lookup"><span data-stu-id="1676f-1322">Name</span></span>|<span data-ttu-id="1676f-1323">类型</span><span class="sxs-lookup"><span data-stu-id="1676f-1323">Type</span></span>|<span data-ttu-id="1676f-1324">说明</span><span class="sxs-lookup"><span data-stu-id="1676f-1324">Description</span></span>|
|---|---|---|
|`name`|<span data-ttu-id="1676f-1325">字符串</span><span class="sxs-lookup"><span data-stu-id="1676f-1325">String</span></span>|<span data-ttu-id="1676f-1326">定义筛选器匹配的 `ItemHasRegularExpressionMatch` 规则元素的名称。</span><span class="sxs-lookup"><span data-stu-id="1676f-1326">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="1676f-1327">要求</span><span class="sxs-lookup"><span data-stu-id="1676f-1327">Requirements</span></span>

|<span data-ttu-id="1676f-1328">要求</span><span class="sxs-lookup"><span data-stu-id="1676f-1328">Requirement</span></span>|<span data-ttu-id="1676f-1329">值</span><span class="sxs-lookup"><span data-stu-id="1676f-1329">Value</span></span>|
|---|---|
|[<span data-ttu-id="1676f-1330">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="1676f-1330">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="1676f-1331">1.0</span><span class="sxs-lookup"><span data-stu-id="1676f-1331">1.0</span></span>|
|[<span data-ttu-id="1676f-1332">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="1676f-1332">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="1676f-1333">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1676f-1333">ReadItem</span></span>|
|[<span data-ttu-id="1676f-1334">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="1676f-1334">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="1676f-1335">阅读</span><span class="sxs-lookup"><span data-stu-id="1676f-1335">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="1676f-1336">返回：</span><span class="sxs-lookup"><span data-stu-id="1676f-1336">Returns:</span></span>

<span data-ttu-id="1676f-1337">一个包含匹配在清单 XML 文件中定义的正则表达式的字符串的数组。</span><span class="sxs-lookup"><span data-stu-id="1676f-1337">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<span data-ttu-id="1676f-1338">类型：Array.< String ></span><span class="sxs-lookup"><span data-stu-id="1676f-1338">Type: Array.< String ></span></span>

##### <a name="example"></a><span data-ttu-id="1676f-1339">示例</span><span class="sxs-lookup"><span data-stu-id="1676f-1339">Example</span></span>

```js
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

<br>

---
---

#### <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="1676f-1340">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="1676f-1340">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="1676f-1341">以异步方式返回邮件的主题或正文中选定的数据。</span><span class="sxs-lookup"><span data-stu-id="1676f-1341">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="1676f-p182">如果没有选定内容，但光标位于正文或主题中，此方法将会为所选数据返回 null。如果选定的是字段，而不是正文或主题，则此方法返回 `InvalidSelection` 错误。</span><span class="sxs-lookup"><span data-stu-id="1676f-p182">If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

> [!NOTE]
> <span data-ttu-id="1676f-1344">在 web 上的 Outlook 中，如果未选择任何文本，但光标在正文中，则该方法将返回字符串 "null"。</span><span class="sxs-lookup"><span data-stu-id="1676f-1344">In Outlook on the web, the method returns the string "null" if no text is selected but the cursor is in the body.</span></span> <span data-ttu-id="1676f-1345">若要检查此情况，请包含与以下内容类似的代码：</span><span class="sxs-lookup"><span data-stu-id="1676f-1345">To check for this situation, include code similar to the following:</span></span>
>
> `var selectedText = (asyncResult.value.endPosition === asyncResult.value.startPosition) ? "" : asyncResult.value.data;`

##### <a name="parameters"></a><span data-ttu-id="1676f-1346">Parameters</span><span class="sxs-lookup"><span data-stu-id="1676f-1346">Parameters</span></span>

|<span data-ttu-id="1676f-1347">名称</span><span class="sxs-lookup"><span data-stu-id="1676f-1347">Name</span></span>|<span data-ttu-id="1676f-1348">类型</span><span class="sxs-lookup"><span data-stu-id="1676f-1348">Type</span></span>|<span data-ttu-id="1676f-1349">属性</span><span class="sxs-lookup"><span data-stu-id="1676f-1349">Attributes</span></span>|<span data-ttu-id="1676f-1350">说明</span><span class="sxs-lookup"><span data-stu-id="1676f-1350">Description</span></span>|
|---|---|---|---|
|`coercionType`|[<span data-ttu-id="1676f-1351">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="1676f-1351">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="1676f-p184">请求数据的格式。如果为文本，则此方法返回纯文本作为字符串，删除任何显示的 HTML 标记。如果为 HTML，则此方法返回所选文本，不论是纯文本还是 HTML。</span><span class="sxs-lookup"><span data-stu-id="1676f-p184">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`|<span data-ttu-id="1676f-1355">Object</span><span class="sxs-lookup"><span data-stu-id="1676f-1355">Object</span></span>|<span data-ttu-id="1676f-1356">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="1676f-1356">&lt;optional&gt;</span></span>|<span data-ttu-id="1676f-1357">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="1676f-1357">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="1676f-1358">对象</span><span class="sxs-lookup"><span data-stu-id="1676f-1358">Object</span></span>|<span data-ttu-id="1676f-1359">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="1676f-1359">&lt;optional&gt;</span></span>|<span data-ttu-id="1676f-1360">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="1676f-1360">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="1676f-1361">function</span><span class="sxs-lookup"><span data-stu-id="1676f-1361">function</span></span>||<span data-ttu-id="1676f-1362">方法完成后，使用单个参数 `callback`（一个 [`asyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `AsyncResult` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="1676f-1362">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="1676f-1363">若要从回调方法访问所选数据，请调用 `asyncResult.value.data`。</span><span class="sxs-lookup"><span data-stu-id="1676f-1363">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="1676f-1364">若要访问选定内容的源属性，请调用 `asyncResult.value.sourceProperty`，这将为 `body` 或 `subject`。</span><span class="sxs-lookup"><span data-stu-id="1676f-1364">To access the source property that the selection comes from, call `asyncResult.value.sourceProperty`, which will be either `body` or `subject`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="1676f-1365">Requirements</span><span class="sxs-lookup"><span data-stu-id="1676f-1365">Requirements</span></span>

|<span data-ttu-id="1676f-1366">要求</span><span class="sxs-lookup"><span data-stu-id="1676f-1366">Requirement</span></span>|<span data-ttu-id="1676f-1367">值</span><span class="sxs-lookup"><span data-stu-id="1676f-1367">Value</span></span>|
|---|---|
|[<span data-ttu-id="1676f-1368">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="1676f-1368">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="1676f-1369">1.2</span><span class="sxs-lookup"><span data-stu-id="1676f-1369">1.2</span></span>|
|[<span data-ttu-id="1676f-1370">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="1676f-1370">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="1676f-1371">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1676f-1371">ReadItem</span></span>|
|[<span data-ttu-id="1676f-1372">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="1676f-1372">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="1676f-1373">撰写</span><span class="sxs-lookup"><span data-stu-id="1676f-1373">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="1676f-1374">返回：</span><span class="sxs-lookup"><span data-stu-id="1676f-1374">Returns:</span></span>

<span data-ttu-id="1676f-1375">作为字符串的所选数据的格式由 `coercionType` 确定。</span><span class="sxs-lookup"><span data-stu-id="1676f-1375">The selected data as a string with format determined by `coercionType`.</span></span>

<span data-ttu-id="1676f-1376">类型：字符串</span><span class="sxs-lookup"><span data-stu-id="1676f-1376">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="1676f-1377">示例</span><span class="sxs-lookup"><span data-stu-id="1676f-1377">Example</span></span>

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

#### <a name="getselectedentities--entitiesjavascriptapioutlookofficeentitiesviewoutlook-js-18"></a><span data-ttu-id="1676f-1378">getSelectedEntities() → {[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.8)}</span><span class="sxs-lookup"><span data-stu-id="1676f-1378">getSelectedEntities() → {[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.8)}</span></span>

<span data-ttu-id="1676f-1379">获取在用户已选择的突出显示匹配项中找到的实体。</span><span class="sxs-lookup"><span data-stu-id="1676f-1379">Gets the entities found in a highlighted match a user has selected.</span></span> <span data-ttu-id="1676f-1380">突出显示匹配项适用于[上下文外接程序](/outlook/add-ins/contextual-outlook-add-ins)。</span><span class="sxs-lookup"><span data-stu-id="1676f-1380">Highlighted matches apply to [contextual add-ins](/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="1676f-1381">iOS 版 Outlook 或 Android 版 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="1676f-1381">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="1676f-1382">要求</span><span class="sxs-lookup"><span data-stu-id="1676f-1382">Requirements</span></span>

|<span data-ttu-id="1676f-1383">要求</span><span class="sxs-lookup"><span data-stu-id="1676f-1383">Requirement</span></span>|<span data-ttu-id="1676f-1384">值</span><span class="sxs-lookup"><span data-stu-id="1676f-1384">Value</span></span>|
|---|---|
|[<span data-ttu-id="1676f-1385">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="1676f-1385">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="1676f-1386">1.6</span><span class="sxs-lookup"><span data-stu-id="1676f-1386">1.6</span></span>|
|[<span data-ttu-id="1676f-1387">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="1676f-1387">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="1676f-1388">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1676f-1388">ReadItem</span></span>|
|[<span data-ttu-id="1676f-1389">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="1676f-1389">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="1676f-1390">阅读</span><span class="sxs-lookup"><span data-stu-id="1676f-1390">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="1676f-1391">返回：</span><span class="sxs-lookup"><span data-stu-id="1676f-1391">Returns:</span></span>

<span data-ttu-id="1676f-1392">类型：[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.8)</span><span class="sxs-lookup"><span data-stu-id="1676f-1392">Type: [Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.8)</span></span>

##### <a name="example"></a><span data-ttu-id="1676f-1393">示例</span><span class="sxs-lookup"><span data-stu-id="1676f-1393">Example</span></span>

<span data-ttu-id="1676f-1394">以下示例访问用户选择的突出显示匹配项中的地址实体。</span><span class="sxs-lookup"><span data-stu-id="1676f-1394">The following example accesses the addresses entities in the highlighted match selected by the user.</span></span>

```js
var contacts = Office.context.mailbox.item.getSelectedEntities().addresses;
```

<br>

---
---

#### <a name="getselectedregexmatches--object"></a><span data-ttu-id="1676f-1395">getSelectedRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="1676f-1395">getSelectedRegExMatches() → {Object}</span></span>

<span data-ttu-id="1676f-p187">返回突出显示匹配项中匹配在清单 XML 文件中定义的正则表达式的字符串值。突出显示匹配项适用于[上下文外接程序](/outlook/add-ins/contextual-outlook-add-ins)。</span><span class="sxs-lookup"><span data-stu-id="1676f-p187">Returns string values in a highlighted match that match the regular expressions defined in the manifest XML file. Highlighted matches apply to [contextual add-ins](/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="1676f-1398">iOS 版 Outlook 或 Android 版 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="1676f-1398">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="1676f-p188">`getSelectedRegExMatches` 方法返回匹配在清单 XML 文件中的每个 `ItemHasRegularExpressionMatch` 或 `ItemHasKnownEntity` 规则元素中定义的正则表达式的字符串。对于 `ItemHasRegularExpressionMatch` 规则，匹配字符串必须发生在该规则指定的项目的属性中。`PropertyName` 简单类型定义支持的属性。</span><span class="sxs-lookup"><span data-stu-id="1676f-p188">The `getSelectedRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="1676f-1402">例如，考虑一个外接程序清单具有以下 `Rule` 元素：</span><span class="sxs-lookup"><span data-stu-id="1676f-1402">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="1676f-1403">从 `getRegExMatches` 返回的对象应有两个属性：`fruits` 和 `veggies`。</span><span class="sxs-lookup"><span data-stu-id="1676f-1403">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="1676f-p189">如果在项目的正文属性上指定 `ItemHasRegularExpressionMatch` 规则，则正则表达式应进一步筛选正文，不应尝试返回该项目的整个正文。使用正则表达式（如 `.*`）获取项目的整个正文并不总是返回预期的结果。而是使用 [`Body.getAsync`](/javascript/api/outlook/office.body?view=outlook-js-1.8#getasync-coerciontype--options--callback-) 方法检索整个正文。</span><span class="sxs-lookup"><span data-stu-id="1676f-p189">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook/office.body?view=outlook-js-1.8#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="1676f-1407">Requirements</span><span class="sxs-lookup"><span data-stu-id="1676f-1407">Requirements</span></span>

|<span data-ttu-id="1676f-1408">要求</span><span class="sxs-lookup"><span data-stu-id="1676f-1408">Requirement</span></span>|<span data-ttu-id="1676f-1409">值</span><span class="sxs-lookup"><span data-stu-id="1676f-1409">Value</span></span>|
|---|---|
|[<span data-ttu-id="1676f-1410">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="1676f-1410">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="1676f-1411">1.6</span><span class="sxs-lookup"><span data-stu-id="1676f-1411">1.6</span></span>|
|[<span data-ttu-id="1676f-1412">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="1676f-1412">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="1676f-1413">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1676f-1413">ReadItem</span></span>|
|[<span data-ttu-id="1676f-1414">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="1676f-1414">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="1676f-1415">阅读</span><span class="sxs-lookup"><span data-stu-id="1676f-1415">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="1676f-1416">返回：</span><span class="sxs-lookup"><span data-stu-id="1676f-1416">Returns:</span></span>

<span data-ttu-id="1676f-p190">一个包含匹配在清单 XML 文件中定义的正则表达式的字符串数组的对象。每个数组的名称等于匹配 `ItemHasRegularExpressionMatch` 规则的 `RegExName` 属性或匹配 `ItemHasKnownEntity` 规则的 `FilterName` 属性的相应值。</span><span class="sxs-lookup"><span data-stu-id="1676f-p190">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

##### <a name="example"></a><span data-ttu-id="1676f-1419">示例</span><span class="sxs-lookup"><span data-stu-id="1676f-1419">Example</span></span>

<span data-ttu-id="1676f-1420">以下示例显示了如何访问正则表达式规则元素 `fruits` 和 `veggies` 的匹配项的数组，这些元素在清单中指定。</span><span class="sxs-lookup"><span data-stu-id="1676f-1420">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```js
var selectedMatches = Office.context.mailbox.item.getSelectedRegExMatches();
var fruits = selectedMatches.fruits;
var veggies = selectedMatches.veggies;
```

<br>

---
---

#### <a name="getsharedpropertiesasyncoptions-callback"></a><span data-ttu-id="1676f-1421">getSharedPropertiesAsync （[options]，回拨）</span><span class="sxs-lookup"><span data-stu-id="1676f-1421">getSharedPropertiesAsync([options], callback)</span></span>

<span data-ttu-id="1676f-1422">获取共享文件夹、日历或邮箱中的所选约会或邮件的属性。</span><span class="sxs-lookup"><span data-stu-id="1676f-1422">Gets the properties of the selected appointment or message in a shared folder, calendar, or mailbox.</span></span>

##### <a name="parameters"></a><span data-ttu-id="1676f-1423">Parameters</span><span class="sxs-lookup"><span data-stu-id="1676f-1423">Parameters</span></span>

|<span data-ttu-id="1676f-1424">名称</span><span class="sxs-lookup"><span data-stu-id="1676f-1424">Name</span></span>|<span data-ttu-id="1676f-1425">类型</span><span class="sxs-lookup"><span data-stu-id="1676f-1425">Type</span></span>|<span data-ttu-id="1676f-1426">属性</span><span class="sxs-lookup"><span data-stu-id="1676f-1426">Attributes</span></span>|<span data-ttu-id="1676f-1427">说明</span><span class="sxs-lookup"><span data-stu-id="1676f-1427">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="1676f-1428">对象</span><span class="sxs-lookup"><span data-stu-id="1676f-1428">Object</span></span>|<span data-ttu-id="1676f-1429">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="1676f-1429">&lt;optional&gt;</span></span>|<span data-ttu-id="1676f-1430">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="1676f-1430">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="1676f-1431">对象</span><span class="sxs-lookup"><span data-stu-id="1676f-1431">Object</span></span>|<span data-ttu-id="1676f-1432">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="1676f-1432">&lt;optional&gt;</span></span>|<span data-ttu-id="1676f-1433">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="1676f-1433">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="1676f-1434">function</span><span class="sxs-lookup"><span data-stu-id="1676f-1434">function</span></span>||<span data-ttu-id="1676f-1435">方法完成后，使用单个参数 `callback`（一个 [`asyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `AsyncResult` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="1676f-1435">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="1676f-1436">共享属性作为[`SharedProperties`](/javascript/api/outlook/office.sharedproperties?view=outlook-js-1.8) `asyncResult.value`属性中的对象提供。</span><span class="sxs-lookup"><span data-stu-id="1676f-1436">The shared properties are provided as a [`SharedProperties`](/javascript/api/outlook/office.sharedproperties?view=outlook-js-1.8) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="1676f-1437">此对象可用于获取项目的共享属性。</span><span class="sxs-lookup"><span data-stu-id="1676f-1437">This object can be used to get the item's shared properties.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="1676f-1438">Requirements</span><span class="sxs-lookup"><span data-stu-id="1676f-1438">Requirements</span></span>

|<span data-ttu-id="1676f-1439">要求</span><span class="sxs-lookup"><span data-stu-id="1676f-1439">Requirement</span></span>|<span data-ttu-id="1676f-1440">值</span><span class="sxs-lookup"><span data-stu-id="1676f-1440">Value</span></span>|
|---|---|
|[<span data-ttu-id="1676f-1441">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="1676f-1441">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="1676f-1442">1.8</span><span class="sxs-lookup"><span data-stu-id="1676f-1442">1.8</span></span>|
|[<span data-ttu-id="1676f-1443">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="1676f-1443">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="1676f-1444">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1676f-1444">ReadItem</span></span>|
|[<span data-ttu-id="1676f-1445">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="1676f-1445">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="1676f-1446">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="1676f-1446">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="1676f-1447">示例</span><span class="sxs-lookup"><span data-stu-id="1676f-1447">Example</span></span>

```js
Office.context.mailbox.item.getSharedPropertiesAsync(callback);

function callback (asyncResult) {
  var context = asyncResult.context;
  var sharedProperties = asyncResult.value;
}
```

<br>

---
---

#### <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="1676f-1448">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="1676f-1448">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="1676f-1449">异步加载所选项目上此外接程序的自定义属性。</span><span class="sxs-lookup"><span data-stu-id="1676f-1449">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="1676f-p192">自定义属性基于每个应用、每个项目存储为键/值对。此方法在回调中返回 `CustomProperties` 对象，该回调提供访问特定于当前项目和当前外接程序的自定义属性的方法。自定义属性未在项目上加密，因此这不应用作安全存储。</span><span class="sxs-lookup"><span data-stu-id="1676f-p192">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="1676f-1453">参数</span><span class="sxs-lookup"><span data-stu-id="1676f-1453">Parameters</span></span>

|<span data-ttu-id="1676f-1454">名称</span><span class="sxs-lookup"><span data-stu-id="1676f-1454">Name</span></span>|<span data-ttu-id="1676f-1455">类型</span><span class="sxs-lookup"><span data-stu-id="1676f-1455">Type</span></span>|<span data-ttu-id="1676f-1456">属性</span><span class="sxs-lookup"><span data-stu-id="1676f-1456">Attributes</span></span>|<span data-ttu-id="1676f-1457">说明</span><span class="sxs-lookup"><span data-stu-id="1676f-1457">Description</span></span>|
|---|---|---|---|
|`callback`|<span data-ttu-id="1676f-1458">函数</span><span class="sxs-lookup"><span data-stu-id="1676f-1458">function</span></span>||<span data-ttu-id="1676f-1459">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="1676f-1459">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="1676f-1460">自定义属性作为 `asyncResult.value` 属性中的 [`CustomProperties`](/javascript/api/outlook/office.customproperties?view=outlook-js-1.8) 对象提供。</span><span class="sxs-lookup"><span data-stu-id="1676f-1460">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook/office.customproperties?view=outlook-js-1.8) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="1676f-1461">此对象可用于获取、设置以及从项目中删除自定义属性，并将自定义属性集的更改重新保存到服务器。</span><span class="sxs-lookup"><span data-stu-id="1676f-1461">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`|<span data-ttu-id="1676f-1462">对象</span><span class="sxs-lookup"><span data-stu-id="1676f-1462">Object</span></span>|<span data-ttu-id="1676f-1463">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="1676f-1463">&lt;optional&gt;</span></span>|<span data-ttu-id="1676f-1464">开发人员可以提供他们想要在回调函数中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="1676f-1464">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="1676f-1465">此对象可以通过回调函数中的 `asyncResult.asyncContext` 属性进行访问。</span><span class="sxs-lookup"><span data-stu-id="1676f-1465">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="1676f-1466">Requirements</span><span class="sxs-lookup"><span data-stu-id="1676f-1466">Requirements</span></span>

|<span data-ttu-id="1676f-1467">要求</span><span class="sxs-lookup"><span data-stu-id="1676f-1467">Requirement</span></span>|<span data-ttu-id="1676f-1468">值</span><span class="sxs-lookup"><span data-stu-id="1676f-1468">Value</span></span>|
|---|---|
|[<span data-ttu-id="1676f-1469">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="1676f-1469">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="1676f-1470">1.0</span><span class="sxs-lookup"><span data-stu-id="1676f-1470">1.0</span></span>|
|[<span data-ttu-id="1676f-1471">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="1676f-1471">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="1676f-1472">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1676f-1472">ReadItem</span></span>|
|[<span data-ttu-id="1676f-1473">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="1676f-1473">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="1676f-1474">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="1676f-1474">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="1676f-1475">示例</span><span class="sxs-lookup"><span data-stu-id="1676f-1475">Example</span></span>

<span data-ttu-id="1676f-p195">以下代码示例显示了如何使用 `loadCustomPropertiesAsync` 方法异步加载特定于当前项目的自定义属性。该示例还显示了如何使用 `CustomProperties.saveAsync` 方法将这些属性重新保存到服务器。加载自定义属性后，该代码示例将使用 `CustomProperties.get` 方法读取自定义属性 `myProp`，使用 `CustomProperties.set` 方法写入自定义属性 `otherProp`，最后调用 `saveAsync` 方法保存这些自定义属性。</span><span class="sxs-lookup"><span data-stu-id="1676f-p195">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

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

#### <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="1676f-1479">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="1676f-1479">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="1676f-1480">将附件从邮件或约会中删除。</span><span class="sxs-lookup"><span data-stu-id="1676f-1480">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="1676f-1481">`removeAttachmentAsync` 方法删除项目中带指定标识符的附件。</span><span class="sxs-lookup"><span data-stu-id="1676f-1481">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item.</span></span> <span data-ttu-id="1676f-1482">最佳做法是，仅当同一个邮件应用程序在同一会话中添加了一个附件时，你才应使用该附件标识符来删除该附件。</span><span class="sxs-lookup"><span data-stu-id="1676f-1482">As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session.</span></span> <span data-ttu-id="1676f-1483">在 Outlook 网页版和移动设备上，附件标识符只在同一个会话中才有效。</span><span class="sxs-lookup"><span data-stu-id="1676f-1483">In Outlook on the web and mobile devices, the attachment identifier is valid only within the same session.</span></span> <span data-ttu-id="1676f-1484">当用户关闭应用程序时，或者如果用户开始撰写内嵌窗体，随后弹出窗体以继续在单独的窗口中。</span><span class="sxs-lookup"><span data-stu-id="1676f-1484">A session is over when the user closes the app, or if the user starts composing an inline form then subsequently pops out the form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="1676f-1485">Parameters</span><span class="sxs-lookup"><span data-stu-id="1676f-1485">Parameters</span></span>

|<span data-ttu-id="1676f-1486">名称</span><span class="sxs-lookup"><span data-stu-id="1676f-1486">Name</span></span>|<span data-ttu-id="1676f-1487">类型</span><span class="sxs-lookup"><span data-stu-id="1676f-1487">Type</span></span>|<span data-ttu-id="1676f-1488">属性</span><span class="sxs-lookup"><span data-stu-id="1676f-1488">Attributes</span></span>|<span data-ttu-id="1676f-1489">说明</span><span class="sxs-lookup"><span data-stu-id="1676f-1489">Description</span></span>|
|---|---|---|---|
|`attachmentId`|<span data-ttu-id="1676f-1490">字符串</span><span class="sxs-lookup"><span data-stu-id="1676f-1490">String</span></span>||<span data-ttu-id="1676f-1491">要删除的附件的标识符。</span><span class="sxs-lookup"><span data-stu-id="1676f-1491">The identifier of the attachment to remove.</span></span>|
|`options`|<span data-ttu-id="1676f-1492">对象</span><span class="sxs-lookup"><span data-stu-id="1676f-1492">Object</span></span>|<span data-ttu-id="1676f-1493">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="1676f-1493">&lt;optional&gt;</span></span>|<span data-ttu-id="1676f-1494">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="1676f-1494">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="1676f-1495">对象</span><span class="sxs-lookup"><span data-stu-id="1676f-1495">Object</span></span>|<span data-ttu-id="1676f-1496">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="1676f-1496">&lt;optional&gt;</span></span>|<span data-ttu-id="1676f-1497">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="1676f-1497">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="1676f-1498">函数</span><span class="sxs-lookup"><span data-stu-id="1676f-1498">function</span></span>|<span data-ttu-id="1676f-1499">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="1676f-1499">&lt;optional&gt;</span></span>|<span data-ttu-id="1676f-1500">方法完成后，使用单个参数 `callback`（一个 [`asyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `AsyncResult` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="1676f-1500">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="1676f-1501">如果删除附件失败，`asyncResult.error` 属性将包含一个说明失败原因的错误代码。</span><span class="sxs-lookup"><span data-stu-id="1676f-1501">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="1676f-1502">错误</span><span class="sxs-lookup"><span data-stu-id="1676f-1502">Errors</span></span>

|<span data-ttu-id="1676f-1503">错误代码</span><span class="sxs-lookup"><span data-stu-id="1676f-1503">Error code</span></span>|<span data-ttu-id="1676f-1504">说明</span><span class="sxs-lookup"><span data-stu-id="1676f-1504">Description</span></span>|
|------------|-------------|
|`InvalidAttachmentId`|<span data-ttu-id="1676f-1505">附件标识符不存在。</span><span class="sxs-lookup"><span data-stu-id="1676f-1505">The attachment identifier does not exist.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="1676f-1506">Requirements</span><span class="sxs-lookup"><span data-stu-id="1676f-1506">Requirements</span></span>

|<span data-ttu-id="1676f-1507">要求</span><span class="sxs-lookup"><span data-stu-id="1676f-1507">Requirement</span></span>|<span data-ttu-id="1676f-1508">值</span><span class="sxs-lookup"><span data-stu-id="1676f-1508">Value</span></span>|
|---|---|
|[<span data-ttu-id="1676f-1509">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="1676f-1509">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="1676f-1510">1.1</span><span class="sxs-lookup"><span data-stu-id="1676f-1510">1.1</span></span>|
|[<span data-ttu-id="1676f-1511">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="1676f-1511">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="1676f-1512">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="1676f-1512">ReadWriteItem</span></span>|
|[<span data-ttu-id="1676f-1513">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="1676f-1513">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="1676f-1514">撰写</span><span class="sxs-lookup"><span data-stu-id="1676f-1514">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="1676f-1515">示例</span><span class="sxs-lookup"><span data-stu-id="1676f-1515">Example</span></span>

<span data-ttu-id="1676f-1516">以下代码删除包含标识符 '0' 的附件。</span><span class="sxs-lookup"><span data-stu-id="1676f-1516">The following code removes an attachment with an identifier of '0'.</span></span>

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

#### <a name="removehandlerasynceventtype-options-callback"></a><span data-ttu-id="1676f-1517">removeHandlerAsync(eventType, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="1676f-1517">removeHandlerAsync(eventType, [options], [callback])</span></span>

<span data-ttu-id="1676f-1518">删除受支持事件类型的事件处理程序。</span><span class="sxs-lookup"><span data-stu-id="1676f-1518">Removes the event handlers for a supported event type.</span></span>

<span data-ttu-id="1676f-1519">目前，受支持的事件`Office.EventType.AttachmentsChanged`类型`Office.EventType.AppointmentTimeChanged`是`Office.EventType.EnhancedLocationsChanged`、 `Office.EventType.RecipientsChanged`、、 `Office.EventType.RecurrenceChanged`和。</span><span class="sxs-lookup"><span data-stu-id="1676f-1519">Currently the supported event types are `Office.EventType.AttachmentsChanged`, `Office.EventType.AppointmentTimeChanged`, `Office.EventType.EnhancedLocationsChanged`, `Office.EventType.RecipientsChanged`, and `Office.EventType.RecurrenceChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="1676f-1520">Parameters</span><span class="sxs-lookup"><span data-stu-id="1676f-1520">Parameters</span></span>

| <span data-ttu-id="1676f-1521">名称</span><span class="sxs-lookup"><span data-stu-id="1676f-1521">Name</span></span> | <span data-ttu-id="1676f-1522">类型</span><span class="sxs-lookup"><span data-stu-id="1676f-1522">Type</span></span> | <span data-ttu-id="1676f-1523">属性</span><span class="sxs-lookup"><span data-stu-id="1676f-1523">Attributes</span></span> | <span data-ttu-id="1676f-1524">说明</span><span class="sxs-lookup"><span data-stu-id="1676f-1524">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="1676f-1525">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="1676f-1525">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="1676f-1526">应撤销处理程序的事件。</span><span class="sxs-lookup"><span data-stu-id="1676f-1526">The event that should revoke the handler.</span></span> |
| `options` | <span data-ttu-id="1676f-1527">对象</span><span class="sxs-lookup"><span data-stu-id="1676f-1527">Object</span></span> | <span data-ttu-id="1676f-1528">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="1676f-1528">&lt;optional&gt;</span></span> | <span data-ttu-id="1676f-1529">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="1676f-1529">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="1676f-1530">对象</span><span class="sxs-lookup"><span data-stu-id="1676f-1530">Object</span></span> | <span data-ttu-id="1676f-1531">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="1676f-1531">&lt;optional&gt;</span></span> | <span data-ttu-id="1676f-1532">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="1676f-1532">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="1676f-1533">函数</span><span class="sxs-lookup"><span data-stu-id="1676f-1533">function</span></span>| <span data-ttu-id="1676f-1534">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="1676f-1534">&lt;optional&gt;</span></span>|<span data-ttu-id="1676f-1535">方法完成后，使用单个参数 `callback`（一个 [`asyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `AsyncResult` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="1676f-1535">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="1676f-1536">Requirements</span><span class="sxs-lookup"><span data-stu-id="1676f-1536">Requirements</span></span>

|<span data-ttu-id="1676f-1537">要求</span><span class="sxs-lookup"><span data-stu-id="1676f-1537">Requirement</span></span>| <span data-ttu-id="1676f-1538">值</span><span class="sxs-lookup"><span data-stu-id="1676f-1538">Value</span></span>|
|---|---|
|[<span data-ttu-id="1676f-1539">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="1676f-1539">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1676f-1540">1.7</span><span class="sxs-lookup"><span data-stu-id="1676f-1540">1.7</span></span> |
|[<span data-ttu-id="1676f-1541">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="1676f-1541">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1676f-1542">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1676f-1542">ReadItem</span></span> |
|[<span data-ttu-id="1676f-1543">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="1676f-1543">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="1676f-1544">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="1676f-1544">Compose or Read</span></span> |

<br>

---
---

#### <a name="saveasyncoptions-callback"></a><span data-ttu-id="1676f-1545">saveAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="1676f-1545">saveAsync([options], callback)</span></span>

<span data-ttu-id="1676f-1546">异步保存项目。</span><span class="sxs-lookup"><span data-stu-id="1676f-1546">Asynchronously saves an item.</span></span>

<span data-ttu-id="1676f-1547">调用时，此方法将当前邮件保存为草稿，并通过回调方法返回项目 ID。</span><span class="sxs-lookup"><span data-stu-id="1676f-1547">When invoked, this method saves the current message as a draft and returns the item id via the callback method.</span></span> <span data-ttu-id="1676f-1548">在 Outlook 网页版或 Outlook 联机模式下，该项目被保存到服务器中。</span><span class="sxs-lookup"><span data-stu-id="1676f-1548">In Outlook on the web or Outlook in online mode, the item is saved to the server.</span></span> <span data-ttu-id="1676f-1549">在 Outlook 缓存模式下，该项目被保存到本地缓存中。</span><span class="sxs-lookup"><span data-stu-id="1676f-1549">In Outlook in cached mode, the item is saved to the local cache.</span></span>

> [!NOTE]
> <span data-ttu-id="1676f-1550">如果加载项在撰写模式下对某个项目调用 `saveAsync` 来获得 `itemId`，以便与 EWS 或 REST API 一同使用，请注意，当 Outlook 处于高速缓存模式时，可能需要一段时间项目才能真正同步到服务器。</span><span class="sxs-lookup"><span data-stu-id="1676f-1550">If your add-in calls `saveAsync` on an item in compose mode in order to get an `itemId` to use with EWS or the REST API, be aware that when Outlook is in cached mode, it may take some time before the item is actually synced to the server.</span></span> <span data-ttu-id="1676f-1551">在项目同步前，使用 `itemId` 将返回一个错误。</span><span class="sxs-lookup"><span data-stu-id="1676f-1551">Until the item is synced, using the `itemId` will return an error.</span></span>

<span data-ttu-id="1676f-p199">由于约会没有草稿状态，如果以撰写模式在约会中调用 `saveAsync`，则该项将被保存为用户日历中的正常约会。对于之前未保存过的新约会，则不会发送邀请。保存现有约会将向添加或删除的与会者发送更新。</span><span class="sxs-lookup"><span data-stu-id="1676f-p199">Since appointments have no draft state, if `saveAsync` is called on an appointment in compose mode, the item will be saved as a normal appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent. Saving an existing appointment will send an update to added or removed attendees.</span></span>

> [!NOTE]
> <span data-ttu-id="1676f-1555">以下客户端在撰写模式下对约会上的 `saveAsync` 具有不同的行为：</span><span class="sxs-lookup"><span data-stu-id="1676f-1555">The following clients have different behavior for `saveAsync` on appointments in compose mode:</span></span>
>
> - <span data-ttu-id="1676f-1556">Mac 版 Outlook 不支持保存会议。</span><span class="sxs-lookup"><span data-stu-id="1676f-1556">Outlook on Mac does not support saving a meeting.</span></span> <span data-ttu-id="1676f-1557">在撰写模式下，无法从会议调用 `saveAsync` 方法。</span><span class="sxs-lookup"><span data-stu-id="1676f-1557">The `saveAsync` method fails when called from a meeting in compose mode.</span></span> <span data-ttu-id="1676f-1558">若需解决办法，请参阅[无法在 Outlook for Mac 中使用 Office JS API 将会议另存为草稿](https://support.microsoft.com/help/4505745)。</span><span class="sxs-lookup"><span data-stu-id="1676f-1558">See [Cannot save a meeting as a draft in Outlook for Mac by using Office JS API](https://support.microsoft.com/help/4505745) for a workaround.</span></span>
> - <span data-ttu-id="1676f-1559">在撰写模式下的约会上调用 `saveAsync` 时，Outlook 网页版始终发送邀请或更新。</span><span class="sxs-lookup"><span data-stu-id="1676f-1559">Outlook on the web always sends an invitation or update when `saveAsync` is called on an appointment in compose mode.</span></span>

##### <a name="parameters"></a><span data-ttu-id="1676f-1560">参数</span><span class="sxs-lookup"><span data-stu-id="1676f-1560">Parameters</span></span>

|<span data-ttu-id="1676f-1561">名称</span><span class="sxs-lookup"><span data-stu-id="1676f-1561">Name</span></span>|<span data-ttu-id="1676f-1562">类型</span><span class="sxs-lookup"><span data-stu-id="1676f-1562">Type</span></span>|<span data-ttu-id="1676f-1563">属性</span><span class="sxs-lookup"><span data-stu-id="1676f-1563">Attributes</span></span>|<span data-ttu-id="1676f-1564">说明</span><span class="sxs-lookup"><span data-stu-id="1676f-1564">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="1676f-1565">Object</span><span class="sxs-lookup"><span data-stu-id="1676f-1565">Object</span></span>|<span data-ttu-id="1676f-1566">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="1676f-1566">&lt;optional&gt;</span></span>|<span data-ttu-id="1676f-1567">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="1676f-1567">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="1676f-1568">对象</span><span class="sxs-lookup"><span data-stu-id="1676f-1568">Object</span></span>|<span data-ttu-id="1676f-1569">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="1676f-1569">&lt;optional&gt;</span></span>|<span data-ttu-id="1676f-1570">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="1676f-1570">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="1676f-1571">函数</span><span class="sxs-lookup"><span data-stu-id="1676f-1571">function</span></span>||<span data-ttu-id="1676f-1572">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="1676f-1572">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="1676f-1573">如果成功，该项目标识符将在 `asyncResult.value` 属性中提供。</span><span class="sxs-lookup"><span data-stu-id="1676f-1573">On success, the item identifier is provided in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="1676f-1574">要求</span><span class="sxs-lookup"><span data-stu-id="1676f-1574">Requirements</span></span>

|<span data-ttu-id="1676f-1575">要求</span><span class="sxs-lookup"><span data-stu-id="1676f-1575">Requirement</span></span>|<span data-ttu-id="1676f-1576">值</span><span class="sxs-lookup"><span data-stu-id="1676f-1576">Value</span></span>|
|---|---|
|[<span data-ttu-id="1676f-1577">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="1676f-1577">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="1676f-1578">1.3</span><span class="sxs-lookup"><span data-stu-id="1676f-1578">1.3</span></span>|
|[<span data-ttu-id="1676f-1579">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="1676f-1579">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="1676f-1580">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="1676f-1580">ReadWriteItem</span></span>|
|[<span data-ttu-id="1676f-1581">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="1676f-1581">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="1676f-1582">撰写</span><span class="sxs-lookup"><span data-stu-id="1676f-1582">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="1676f-1583">示例</span><span class="sxs-lookup"><span data-stu-id="1676f-1583">Examples</span></span>

```js
Office.context.mailbox.item.saveAsync(
  function callback(result) {
    // Process the result.
  });
```

<span data-ttu-id="1676f-p201">下面是传递给回调函数的 `result` 参数的示例。`value` 属性包含的项目的项目 ID。</span><span class="sxs-lookup"><span data-stu-id="1676f-p201">The following is an example of the `result` parameter passed to the callback function. The `value` property contains the item ID of the item.</span></span>

```json
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

<br>

---
---

#### <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="1676f-1586">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="1676f-1586">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="1676f-1587">以异步方式将数据插入到邮件的正文或主题中。</span><span class="sxs-lookup"><span data-stu-id="1676f-1587">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="1676f-p202">`setSelectedDataAsync` 方法将指定字符串插入到项目主题或正文的光标位置，或者，如果在编辑器中已选择文本，则该方法将替换选择的文本。如果光标不在正文或主题字段中，则返回错误。插入之后，光标会位于插入内容的末尾。</span><span class="sxs-lookup"><span data-stu-id="1676f-p202">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="1676f-1591">参数</span><span class="sxs-lookup"><span data-stu-id="1676f-1591">Parameters</span></span>

|<span data-ttu-id="1676f-1592">名称</span><span class="sxs-lookup"><span data-stu-id="1676f-1592">Name</span></span>|<span data-ttu-id="1676f-1593">类型</span><span class="sxs-lookup"><span data-stu-id="1676f-1593">Type</span></span>|<span data-ttu-id="1676f-1594">属性</span><span class="sxs-lookup"><span data-stu-id="1676f-1594">Attributes</span></span>|<span data-ttu-id="1676f-1595">说明</span><span class="sxs-lookup"><span data-stu-id="1676f-1595">Description</span></span>|
|---|---|---|---|
|`data`|<span data-ttu-id="1676f-1596">字符串</span><span class="sxs-lookup"><span data-stu-id="1676f-1596">String</span></span>||<span data-ttu-id="1676f-p203">要插入的数据。数据不得超过 1,000,000 个字符。如果传入的数据超过 1,000,000 个字符，则会引发 `ArgumentOutOfRange` 异常。</span><span class="sxs-lookup"><span data-stu-id="1676f-p203">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`|<span data-ttu-id="1676f-1600">Object</span><span class="sxs-lookup"><span data-stu-id="1676f-1600">Object</span></span>|<span data-ttu-id="1676f-1601">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="1676f-1601">&lt;optional&gt;</span></span>|<span data-ttu-id="1676f-1602">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="1676f-1602">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="1676f-1603">对象</span><span class="sxs-lookup"><span data-stu-id="1676f-1603">Object</span></span>|<span data-ttu-id="1676f-1604">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="1676f-1604">&lt;optional&gt;</span></span>|<span data-ttu-id="1676f-1605">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="1676f-1605">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`|[<span data-ttu-id="1676f-1606">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="1676f-1606">Office.CoercionType</span></span>](office.md#coerciontype-string)|<span data-ttu-id="1676f-1607">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="1676f-1607">&lt;optional&gt;</span></span>|<span data-ttu-id="1676f-1608">如果为 `text`，系统在 Outlook 网页版和 Outlook 桌面版客户端中应用当前样式。</span><span class="sxs-lookup"><span data-stu-id="1676f-1608">If `text`, the current style is applied in Outlook on the web and desktop clients.</span></span> <span data-ttu-id="1676f-1609">如果字段是 HTML 编辑器，只会插入文本数据，即使数据为 HTML，也不例外。</span><span class="sxs-lookup"><span data-stu-id="1676f-1609">If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="1676f-1610">如果 `html` 和字段支持 HTML（主题不支持），系统在 Outlook 网页版中应用当前样式，而在 Outlook 桌面版客户端中则应用默认样式。</span><span class="sxs-lookup"><span data-stu-id="1676f-1610">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook on the web and the default style is applied in Outlook desktop clients.</span></span> <span data-ttu-id="1676f-1611">如果该字段是文本字段，则返回 `InvalidDataFormat` 错误。</span><span class="sxs-lookup"><span data-stu-id="1676f-1611">If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="1676f-1612">如果未设置 `coercionType`，则结果取决于该字段：如果该字段是 HTML，则使用 HTML；如果该字段是文本，则使用纯文本。</span><span class="sxs-lookup"><span data-stu-id="1676f-1612">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`|<span data-ttu-id="1676f-1613">function</span><span class="sxs-lookup"><span data-stu-id="1676f-1613">function</span></span>||<span data-ttu-id="1676f-1614">方法完成后，使用单个参数 `callback`（一个 [`asyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `AsyncResult` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="1676f-1614">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="1676f-1615">Requirements</span><span class="sxs-lookup"><span data-stu-id="1676f-1615">Requirements</span></span>

|<span data-ttu-id="1676f-1616">要求</span><span class="sxs-lookup"><span data-stu-id="1676f-1616">Requirement</span></span>|<span data-ttu-id="1676f-1617">值</span><span class="sxs-lookup"><span data-stu-id="1676f-1617">Value</span></span>|
|---|---|
|[<span data-ttu-id="1676f-1618">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="1676f-1618">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="1676f-1619">1.2</span><span class="sxs-lookup"><span data-stu-id="1676f-1619">1.2</span></span>|
|[<span data-ttu-id="1676f-1620">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="1676f-1620">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="1676f-1621">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="1676f-1621">ReadWriteItem</span></span>|
|[<span data-ttu-id="1676f-1622">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="1676f-1622">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="1676f-1623">撰写</span><span class="sxs-lookup"><span data-stu-id="1676f-1623">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="1676f-1624">示例</span><span class="sxs-lookup"><span data-stu-id="1676f-1624">Example</span></span>

```js
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```
