---
title: "\"Context.subname\"-\"邮箱\"-预览要求集"
description: ''
ms.date: 10/30/2019
localization_priority: Normal
ms.openlocfilehash: 3ccafccd0c84ab243572421609083f56e3f7dfb1
ms.sourcegitcommit: e989096f3d19761bf8477c585cde20b3f8e0b90d
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 10/31/2019
ms.locfileid: "37902233"
---
# <a name="item"></a><span data-ttu-id="f156e-102">item</span><span class="sxs-lookup"><span data-stu-id="f156e-102">item</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmditem"></a><span data-ttu-id="f156e-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span><span class="sxs-lookup"><span data-stu-id="f156e-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span></span>

<span data-ttu-id="f156e-p101">`item` 命名空间用于访问当前选定的邮件、会议请求或约会。可以通过使用 [itemType](#itemtype-officemailboxenumsitemtype) 属性确定 `item` 的类型。</span><span class="sxs-lookup"><span data-stu-id="f156e-p101">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="f156e-106">Requirements</span><span class="sxs-lookup"><span data-stu-id="f156e-106">Requirements</span></span>

|<span data-ttu-id="f156e-107">要求</span><span class="sxs-lookup"><span data-stu-id="f156e-107">Requirement</span></span>|<span data-ttu-id="f156e-108">值</span><span class="sxs-lookup"><span data-stu-id="f156e-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="f156e-109">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="f156e-109">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f156e-110">1.0</span><span class="sxs-lookup"><span data-stu-id="f156e-110">1.0</span></span>|
|[<span data-ttu-id="f156e-111">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="f156e-111">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f156e-112">受限</span><span class="sxs-lookup"><span data-stu-id="f156e-112">Restricted</span></span>|
|[<span data-ttu-id="f156e-113">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="f156e-113">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="f156e-114">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="f156e-114">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="f156e-115">成员和方法</span><span class="sxs-lookup"><span data-stu-id="f156e-115">Members and methods</span></span>

| <span data-ttu-id="f156e-116">成员</span><span class="sxs-lookup"><span data-stu-id="f156e-116">Member</span></span> | <span data-ttu-id="f156e-117">类型</span><span class="sxs-lookup"><span data-stu-id="f156e-117">Type</span></span> |
|--------|------|
| [<span data-ttu-id="f156e-118">attachments</span><span class="sxs-lookup"><span data-stu-id="f156e-118">attachments</span></span>](#attachments-arrayattachmentdetails) | <span data-ttu-id="f156e-119">成员</span><span class="sxs-lookup"><span data-stu-id="f156e-119">Member</span></span> |
| [<span data-ttu-id="f156e-120">bcc</span><span class="sxs-lookup"><span data-stu-id="f156e-120">bcc</span></span>](#bcc-recipients) | <span data-ttu-id="f156e-121">成员</span><span class="sxs-lookup"><span data-stu-id="f156e-121">Member</span></span> |
| [<span data-ttu-id="f156e-122">body</span><span class="sxs-lookup"><span data-stu-id="f156e-122">body</span></span>](#body-body) | <span data-ttu-id="f156e-123">成员</span><span class="sxs-lookup"><span data-stu-id="f156e-123">Member</span></span> |
| [<span data-ttu-id="f156e-124">种类</span><span class="sxs-lookup"><span data-stu-id="f156e-124">categories</span></span>](#categories-categories) | <span data-ttu-id="f156e-125">成员</span><span class="sxs-lookup"><span data-stu-id="f156e-125">Member</span></span> |
| [<span data-ttu-id="f156e-126">cc</span><span class="sxs-lookup"><span data-stu-id="f156e-126">cc</span></span>](#cc-arrayemailaddressdetailsrecipients) | <span data-ttu-id="f156e-127">成员</span><span class="sxs-lookup"><span data-stu-id="f156e-127">Member</span></span> |
| [<span data-ttu-id="f156e-128">conversationId</span><span class="sxs-lookup"><span data-stu-id="f156e-128">conversationId</span></span>](#nullable-conversationid-string) | <span data-ttu-id="f156e-129">成员</span><span class="sxs-lookup"><span data-stu-id="f156e-129">Member</span></span> |
| [<span data-ttu-id="f156e-130">dateTimeCreated</span><span class="sxs-lookup"><span data-stu-id="f156e-130">dateTimeCreated</span></span>](#datetimecreated-date) | <span data-ttu-id="f156e-131">成员</span><span class="sxs-lookup"><span data-stu-id="f156e-131">Member</span></span> |
| [<span data-ttu-id="f156e-132">dateTimeModified</span><span class="sxs-lookup"><span data-stu-id="f156e-132">dateTimeModified</span></span>](#datetimemodified-date) | <span data-ttu-id="f156e-133">成员</span><span class="sxs-lookup"><span data-stu-id="f156e-133">Member</span></span> |
| [<span data-ttu-id="f156e-134">end</span><span class="sxs-lookup"><span data-stu-id="f156e-134">end</span></span>](#end-datetime) | <span data-ttu-id="f156e-135">成员</span><span class="sxs-lookup"><span data-stu-id="f156e-135">Member</span></span> |
| [<span data-ttu-id="f156e-136">enhancedLocation</span><span class="sxs-lookup"><span data-stu-id="f156e-136">enhancedLocation</span></span>](#enhancedlocation-enhancedlocation) | <span data-ttu-id="f156e-137">成员</span><span class="sxs-lookup"><span data-stu-id="f156e-137">Member</span></span> |
| [<span data-ttu-id="f156e-138">from</span><span class="sxs-lookup"><span data-stu-id="f156e-138">from</span></span>](#from-emailaddressdetailsfrom) | <span data-ttu-id="f156e-139">成员</span><span class="sxs-lookup"><span data-stu-id="f156e-139">Member</span></span> |
| [<span data-ttu-id="f156e-140">internetHeaders</span><span class="sxs-lookup"><span data-stu-id="f156e-140">internetHeaders</span></span>](#internetheaders-internetheaders) | <span data-ttu-id="f156e-141">成员</span><span class="sxs-lookup"><span data-stu-id="f156e-141">Member</span></span> |
| [<span data-ttu-id="f156e-142">internetMessageId</span><span class="sxs-lookup"><span data-stu-id="f156e-142">internetMessageId</span></span>](#internetmessageid-string) | <span data-ttu-id="f156e-143">成员</span><span class="sxs-lookup"><span data-stu-id="f156e-143">Member</span></span> |
| [<span data-ttu-id="f156e-144">itemClass</span><span class="sxs-lookup"><span data-stu-id="f156e-144">itemClass</span></span>](#itemclass-string) | <span data-ttu-id="f156e-145">成员</span><span class="sxs-lookup"><span data-stu-id="f156e-145">Member</span></span> |
| [<span data-ttu-id="f156e-146">itemId</span><span class="sxs-lookup"><span data-stu-id="f156e-146">itemId</span></span>](#nullable-itemid-string) | <span data-ttu-id="f156e-147">成员</span><span class="sxs-lookup"><span data-stu-id="f156e-147">Member</span></span> |
| [<span data-ttu-id="f156e-148">itemType</span><span class="sxs-lookup"><span data-stu-id="f156e-148">itemType</span></span>](#itemtype-officemailboxenumsitemtype) | <span data-ttu-id="f156e-149">成员</span><span class="sxs-lookup"><span data-stu-id="f156e-149">Member</span></span> |
| [<span data-ttu-id="f156e-150">location</span><span class="sxs-lookup"><span data-stu-id="f156e-150">location</span></span>](#location-stringlocation) | <span data-ttu-id="f156e-151">成员</span><span class="sxs-lookup"><span data-stu-id="f156e-151">Member</span></span> |
| [<span data-ttu-id="f156e-152">normalizedSubject</span><span class="sxs-lookup"><span data-stu-id="f156e-152">normalizedSubject</span></span>](#normalizedsubject-string) | <span data-ttu-id="f156e-153">成员</span><span class="sxs-lookup"><span data-stu-id="f156e-153">Member</span></span> |
| [<span data-ttu-id="f156e-154">notificationMessages</span><span class="sxs-lookup"><span data-stu-id="f156e-154">notificationMessages</span></span>](#notificationmessages-notificationmessages) | <span data-ttu-id="f156e-155">Member</span><span class="sxs-lookup"><span data-stu-id="f156e-155">Member</span></span> |
| [<span data-ttu-id="f156e-156">optionalAttendees</span><span class="sxs-lookup"><span data-stu-id="f156e-156">optionalAttendees</span></span>](#optionalattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="f156e-157">成员</span><span class="sxs-lookup"><span data-stu-id="f156e-157">Member</span></span> |
| [<span data-ttu-id="f156e-158">organizer</span><span class="sxs-lookup"><span data-stu-id="f156e-158">organizer</span></span>](#organizer-emailaddressdetailsorganizer) | <span data-ttu-id="f156e-159">成员</span><span class="sxs-lookup"><span data-stu-id="f156e-159">Member</span></span> |
| [<span data-ttu-id="f156e-160">recurrence</span><span class="sxs-lookup"><span data-stu-id="f156e-160">recurrence</span></span>](#nullable-recurrence-recurrence) | <span data-ttu-id="f156e-161">成员</span><span class="sxs-lookup"><span data-stu-id="f156e-161">Member</span></span> |
| [<span data-ttu-id="f156e-162">requiredAttendees</span><span class="sxs-lookup"><span data-stu-id="f156e-162">requiredAttendees</span></span>](#requiredattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="f156e-163">成员</span><span class="sxs-lookup"><span data-stu-id="f156e-163">Member</span></span> |
| [<span data-ttu-id="f156e-164">sender</span><span class="sxs-lookup"><span data-stu-id="f156e-164">sender</span></span>](#sender-emailaddressdetails) | <span data-ttu-id="f156e-165">Member</span><span class="sxs-lookup"><span data-stu-id="f156e-165">Member</span></span> |
| [<span data-ttu-id="f156e-166">Webcasts&seriesid</span><span class="sxs-lookup"><span data-stu-id="f156e-166">seriesId</span></span>](#nullable-seriesid-string) | <span data-ttu-id="f156e-167">Member</span><span class="sxs-lookup"><span data-stu-id="f156e-167">Member</span></span> |
| [<span data-ttu-id="f156e-168">start</span><span class="sxs-lookup"><span data-stu-id="f156e-168">start</span></span>](#start-datetime) | <span data-ttu-id="f156e-169">Member</span><span class="sxs-lookup"><span data-stu-id="f156e-169">Member</span></span> |
| [<span data-ttu-id="f156e-170">subject</span><span class="sxs-lookup"><span data-stu-id="f156e-170">subject</span></span>](#subject-stringsubject) | <span data-ttu-id="f156e-171">成员</span><span class="sxs-lookup"><span data-stu-id="f156e-171">Member</span></span> |
| [<span data-ttu-id="f156e-172">to</span><span class="sxs-lookup"><span data-stu-id="f156e-172">to</span></span>](#to-arrayemailaddressdetailsrecipients) | <span data-ttu-id="f156e-173">成员</span><span class="sxs-lookup"><span data-stu-id="f156e-173">Member</span></span> |
| [<span data-ttu-id="f156e-174">addFileAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="f156e-174">addFileAttachmentAsync</span></span>](#addfileattachmentasyncuri-attachmentname-options-callback) | <span data-ttu-id="f156e-175">方法</span><span class="sxs-lookup"><span data-stu-id="f156e-175">Method</span></span> |
| [<span data-ttu-id="f156e-176">addFileAttachmentFromBase64Async</span><span class="sxs-lookup"><span data-stu-id="f156e-176">addFileAttachmentFromBase64Async</span></span>](#addfileattachmentfrombase64asyncbase64file-attachmentname-options-callback) | <span data-ttu-id="f156e-177">方法</span><span class="sxs-lookup"><span data-stu-id="f156e-177">Method</span></span> |
| [<span data-ttu-id="f156e-178">addHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="f156e-178">addHandlerAsync</span></span>](#addhandlerasynceventtype-handler-options-callback) | <span data-ttu-id="f156e-179">方法</span><span class="sxs-lookup"><span data-stu-id="f156e-179">Method</span></span> |
| [<span data-ttu-id="f156e-180">addItemAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="f156e-180">addItemAttachmentAsync</span></span>](#additemattachmentasyncitemid-attachmentname-options-callback) | <span data-ttu-id="f156e-181">方法</span><span class="sxs-lookup"><span data-stu-id="f156e-181">Method</span></span> |
| [<span data-ttu-id="f156e-182">close</span><span class="sxs-lookup"><span data-stu-id="f156e-182">close</span></span>](#close) | <span data-ttu-id="f156e-183">方法</span><span class="sxs-lookup"><span data-stu-id="f156e-183">Method</span></span> |
| [<span data-ttu-id="f156e-184">displayReplyAllForm</span><span class="sxs-lookup"><span data-stu-id="f156e-184">displayReplyAllForm</span></span>](#displayreplyallformformdata-callback) | <span data-ttu-id="f156e-185">方法</span><span class="sxs-lookup"><span data-stu-id="f156e-185">Method</span></span> |
| [<span data-ttu-id="f156e-186">displayReplyForm</span><span class="sxs-lookup"><span data-stu-id="f156e-186">displayReplyForm</span></span>](#displayreplyformformdata-callback) | <span data-ttu-id="f156e-187">方法</span><span class="sxs-lookup"><span data-stu-id="f156e-187">Method</span></span> |
| [<span data-ttu-id="f156e-188">getAllInternetHeadersAsync</span><span class="sxs-lookup"><span data-stu-id="f156e-188">getAllInternetHeadersAsync</span></span>](#getallinternetheadersasyncoptions-callback) | <span data-ttu-id="f156e-189">方法</span><span class="sxs-lookup"><span data-stu-id="f156e-189">Method</span></span> |
| [<span data-ttu-id="f156e-190">getAttachmentContentAsync</span><span class="sxs-lookup"><span data-stu-id="f156e-190">getAttachmentContentAsync</span></span>](#getattachmentcontentasyncattachmentid-options-callback--attachmentcontent) | <span data-ttu-id="f156e-191">方法</span><span class="sxs-lookup"><span data-stu-id="f156e-191">Method</span></span> |
| [<span data-ttu-id="f156e-192">getAttachmentsAsync</span><span class="sxs-lookup"><span data-stu-id="f156e-192">getAttachmentsAsync</span></span>](#getattachmentsasyncoptions-callback--arrayattachmentdetails) | <span data-ttu-id="f156e-193">方法</span><span class="sxs-lookup"><span data-stu-id="f156e-193">Method</span></span> |
| [<span data-ttu-id="f156e-194">getEntities</span><span class="sxs-lookup"><span data-stu-id="f156e-194">getEntities</span></span>](#getentities--entities) | <span data-ttu-id="f156e-195">Method</span><span class="sxs-lookup"><span data-stu-id="f156e-195">Method</span></span> |
| [<span data-ttu-id="f156e-196">getEntitiesByType</span><span class="sxs-lookup"><span data-stu-id="f156e-196">getEntitiesByType</span></span>](#getentitiesbytypeentitytype--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="f156e-197">Method</span><span class="sxs-lookup"><span data-stu-id="f156e-197">Method</span></span> |
| [<span data-ttu-id="f156e-198">getFilteredEntitiesByName</span><span class="sxs-lookup"><span data-stu-id="f156e-198">getFilteredEntitiesByName</span></span>](#getfilteredentitiesbynamename--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="f156e-199">Method</span><span class="sxs-lookup"><span data-stu-id="f156e-199">Method</span></span> |
| [<span data-ttu-id="f156e-200">Office.context.mailbox.item.getinitializationcontextasync</span><span class="sxs-lookup"><span data-stu-id="f156e-200">getInitializationContextAsync</span></span>](#getinitializationcontextasyncoptions-callback) | <span data-ttu-id="f156e-201">Method</span><span class="sxs-lookup"><span data-stu-id="f156e-201">Method</span></span> |
| [<span data-ttu-id="f156e-202">getItemIdAsync</span><span class="sxs-lookup"><span data-stu-id="f156e-202">getItemIdAsync</span></span>](#getitemidasyncoptions-callback) | <span data-ttu-id="f156e-203">Method</span><span class="sxs-lookup"><span data-stu-id="f156e-203">Method</span></span> |
| [<span data-ttu-id="f156e-204">getRegExMatches</span><span class="sxs-lookup"><span data-stu-id="f156e-204">getRegExMatches</span></span>](#getregexmatches--object) | <span data-ttu-id="f156e-205">Method</span><span class="sxs-lookup"><span data-stu-id="f156e-205">Method</span></span> |
| [<span data-ttu-id="f156e-206">getRegExMatchesByName</span><span class="sxs-lookup"><span data-stu-id="f156e-206">getRegExMatchesByName</span></span>](#getregexmatchesbynamename--nullable-array-string-) | <span data-ttu-id="f156e-207">Method</span><span class="sxs-lookup"><span data-stu-id="f156e-207">Method</span></span> |
| [<span data-ttu-id="f156e-208">getSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="f156e-208">getSelectedDataAsync</span></span>](#getselecteddataasynccoerciontype-options-callback--string) | <span data-ttu-id="f156e-209">Method</span><span class="sxs-lookup"><span data-stu-id="f156e-209">Method</span></span> |
| [<span data-ttu-id="f156e-210">Office.context.mailbox.item.getselectedentities</span><span class="sxs-lookup"><span data-stu-id="f156e-210">getSelectedEntities</span></span>](#getselectedentities--entities) | <span data-ttu-id="f156e-211">Method</span><span class="sxs-lookup"><span data-stu-id="f156e-211">Method</span></span> |
| [<span data-ttu-id="f156e-212">Office.context.mailbox.item.getselectedregexmatches</span><span class="sxs-lookup"><span data-stu-id="f156e-212">getSelectedRegExMatches</span></span>](#getselectedregexmatches--object) | <span data-ttu-id="f156e-213">Method</span><span class="sxs-lookup"><span data-stu-id="f156e-213">Method</span></span> |
| [<span data-ttu-id="f156e-214">getSharedPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="f156e-214">getSharedPropertiesAsync</span></span>](#getsharedpropertiesasyncoptions-callback) | <span data-ttu-id="f156e-215">Method</span><span class="sxs-lookup"><span data-stu-id="f156e-215">Method</span></span> |
| [<span data-ttu-id="f156e-216">loadCustomPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="f156e-216">loadCustomPropertiesAsync</span></span>](#loadcustompropertiesasynccallback-usercontext) | <span data-ttu-id="f156e-217">Method</span><span class="sxs-lookup"><span data-stu-id="f156e-217">Method</span></span> |
| [<span data-ttu-id="f156e-218">removeAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="f156e-218">removeAttachmentAsync</span></span>](#removeattachmentasyncattachmentid-options-callback) | <span data-ttu-id="f156e-219">方法</span><span class="sxs-lookup"><span data-stu-id="f156e-219">Method</span></span> |
| [<span data-ttu-id="f156e-220">removeHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="f156e-220">removeHandlerAsync</span></span>](#removehandlerasynceventtype-options-callback) | <span data-ttu-id="f156e-221">方法</span><span class="sxs-lookup"><span data-stu-id="f156e-221">Method</span></span> |
| [<span data-ttu-id="f156e-222">saveAsync</span><span class="sxs-lookup"><span data-stu-id="f156e-222">saveAsync</span></span>](#saveasyncoptions-callback) | <span data-ttu-id="f156e-223">Method</span><span class="sxs-lookup"><span data-stu-id="f156e-223">Method</span></span> |
| [<span data-ttu-id="f156e-224">setSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="f156e-224">setSelectedDataAsync</span></span>](#setselecteddataasyncdata-options-callback) | <span data-ttu-id="f156e-225">方法</span><span class="sxs-lookup"><span data-stu-id="f156e-225">Method</span></span> |

### <a name="example"></a><span data-ttu-id="f156e-226">示例</span><span class="sxs-lookup"><span data-stu-id="f156e-226">Example</span></span>

<span data-ttu-id="f156e-227">以下 JavaScript 代码示例显示了如何访问 Outlook 中当前项目的 `subject` 属性。</span><span class="sxs-lookup"><span data-stu-id="f156e-227">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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

### <a name="members"></a><span data-ttu-id="f156e-228">Members</span><span class="sxs-lookup"><span data-stu-id="f156e-228">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetails"></a><span data-ttu-id="f156e-229">attachments: Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="f156e-229">attachments: Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span></span>

<span data-ttu-id="f156e-230">以数组的形式获取项目的附件。</span><span class="sxs-lookup"><span data-stu-id="f156e-230">Gets the item's attachments as an array.</span></span> <span data-ttu-id="f156e-231">仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="f156e-231">Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="f156e-232">由于潜在的安全问题，某些类型的文件会受到 Outlook 阻止，并且不会返回。</span><span class="sxs-lookup"><span data-stu-id="f156e-232">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="f156e-233">如需了解更多信息，请参阅 [Outlook 中阻止的附件](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519)。</span><span class="sxs-lookup"><span data-stu-id="f156e-233">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="f156e-234">类型</span><span class="sxs-lookup"><span data-stu-id="f156e-234">Type</span></span>

*   <span data-ttu-id="f156e-235">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="f156e-235">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span></span>

##### <a name="requirements"></a><span data-ttu-id="f156e-236">要求</span><span class="sxs-lookup"><span data-stu-id="f156e-236">Requirements</span></span>

|<span data-ttu-id="f156e-237">要求</span><span class="sxs-lookup"><span data-stu-id="f156e-237">Requirement</span></span>|<span data-ttu-id="f156e-238">值</span><span class="sxs-lookup"><span data-stu-id="f156e-238">Value</span></span>|
|---|---|
|[<span data-ttu-id="f156e-239">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="f156e-239">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f156e-240">1.0</span><span class="sxs-lookup"><span data-stu-id="f156e-240">1.0</span></span>|
|[<span data-ttu-id="f156e-241">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="f156e-241">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f156e-242">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f156e-242">ReadItem</span></span>|
|[<span data-ttu-id="f156e-243">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="f156e-243">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="f156e-244">阅读</span><span class="sxs-lookup"><span data-stu-id="f156e-244">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="f156e-245">示例</span><span class="sxs-lookup"><span data-stu-id="f156e-245">Example</span></span>

<span data-ttu-id="f156e-246">以下代码使用当前项目上所有附件的详细信息构成 HTML 字符串。</span><span class="sxs-lookup"><span data-stu-id="f156e-246">The following code builds an HTML string with details of all attachments on the current item.</span></span>

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

#### <a name="bcc-recipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="f156e-247">bcc: [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="f156e-247">bcc: [Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="f156e-248">获取对象，该对象提供用于获取或更新邮件的密件抄送 (Bcc) 行上的收件人的方法。</span><span class="sxs-lookup"><span data-stu-id="f156e-248">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="f156e-249">仅限撰写模式。</span><span class="sxs-lookup"><span data-stu-id="f156e-249">Compose mode only.</span></span>

<span data-ttu-id="f156e-250">默认情况下，集合上限为 100 个成员。</span><span class="sxs-lookup"><span data-stu-id="f156e-250">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="f156e-251">但是，在 Windows 和 Mac 上，以下限制适用。</span><span class="sxs-lookup"><span data-stu-id="f156e-251">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="f156e-252">最多包含 500 个成员。</span><span class="sxs-lookup"><span data-stu-id="f156e-252">Get 500 members maximum.</span></span>
- <span data-ttu-id="f156e-253">为每个呼叫最多设置 100 个成员，总共多达 500 个成员。</span><span class="sxs-lookup"><span data-stu-id="f156e-253">Set a maximum of 100 members per call, up to 500 members total.</span></span>

##### <a name="type"></a><span data-ttu-id="f156e-254">类型</span><span class="sxs-lookup"><span data-stu-id="f156e-254">Type</span></span>

*   [<span data-ttu-id="f156e-255">收件人</span><span class="sxs-lookup"><span data-stu-id="f156e-255">Recipients</span></span>](/javascript/api/outlook/office.recipients)

##### <a name="requirements"></a><span data-ttu-id="f156e-256">Requirements</span><span class="sxs-lookup"><span data-stu-id="f156e-256">Requirements</span></span>

|<span data-ttu-id="f156e-257">要求</span><span class="sxs-lookup"><span data-stu-id="f156e-257">Requirement</span></span>|<span data-ttu-id="f156e-258">值</span><span class="sxs-lookup"><span data-stu-id="f156e-258">Value</span></span>|
|---|---|
|[<span data-ttu-id="f156e-259">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="f156e-259">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f156e-260">1.1</span><span class="sxs-lookup"><span data-stu-id="f156e-260">1.1</span></span>|
|[<span data-ttu-id="f156e-261">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="f156e-261">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f156e-262">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f156e-262">ReadItem</span></span>|
|[<span data-ttu-id="f156e-263">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="f156e-263">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="f156e-264">撰写</span><span class="sxs-lookup"><span data-stu-id="f156e-264">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="f156e-265">示例</span><span class="sxs-lookup"><span data-stu-id="f156e-265">Example</span></span>

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

#### <a name="body-bodyjavascriptapioutlookofficebody"></a><span data-ttu-id="f156e-266">body: [Body](/javascript/api/outlook/office.body)</span><span class="sxs-lookup"><span data-stu-id="f156e-266">body: [Body](/javascript/api/outlook/office.body)</span></span>

<span data-ttu-id="f156e-267">获取一个提供用于处理项目正文的方法的对象。</span><span class="sxs-lookup"><span data-stu-id="f156e-267">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="f156e-268">类型</span><span class="sxs-lookup"><span data-stu-id="f156e-268">Type</span></span>

*   [<span data-ttu-id="f156e-269">Body</span><span class="sxs-lookup"><span data-stu-id="f156e-269">Body</span></span>](/javascript/api/outlook/office.body)

##### <a name="requirements"></a><span data-ttu-id="f156e-270">Requirements</span><span class="sxs-lookup"><span data-stu-id="f156e-270">Requirements</span></span>

|<span data-ttu-id="f156e-271">要求</span><span class="sxs-lookup"><span data-stu-id="f156e-271">Requirement</span></span>|<span data-ttu-id="f156e-272">值</span><span class="sxs-lookup"><span data-stu-id="f156e-272">Value</span></span>|
|---|---|
|[<span data-ttu-id="f156e-273">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="f156e-273">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f156e-274">1.1</span><span class="sxs-lookup"><span data-stu-id="f156e-274">1.1</span></span>|
|[<span data-ttu-id="f156e-275">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="f156e-275">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f156e-276">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f156e-276">ReadItem</span></span>|
|[<span data-ttu-id="f156e-277">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="f156e-277">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="f156e-278">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="f156e-278">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="f156e-279">示例</span><span class="sxs-lookup"><span data-stu-id="f156e-279">Example</span></span>

<span data-ttu-id="f156e-280">本示例获取纯文本格式的邮件正文。</span><span class="sxs-lookup"><span data-stu-id="f156e-280">This example gets the body of the message in plain text.</span></span>

```js
Office.context.mailbox.item.body.getAsync(
  "text",
  { asyncContext: "This is passed to the callback" },
  function callback(result) {
    // Do something with the result.
  });

```

<span data-ttu-id="f156e-281">以下是传递到回调函数的结果参数的示例。</span><span class="sxs-lookup"><span data-stu-id="f156e-281">The following is an example of the result parameter passed to the callback function.</span></span>

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

#### <a name="categories-categoriesjavascriptapioutlookofficecategories"></a><span data-ttu-id="f156e-282">类别：[类别](/javascript/api/outlook/office.categories)</span><span class="sxs-lookup"><span data-stu-id="f156e-282">categories: [Categories](/javascript/api/outlook/office.categories)</span></span>

<span data-ttu-id="f156e-283">获取一个对象，该对象提供用于管理项的类别的方法。</span><span class="sxs-lookup"><span data-stu-id="f156e-283">Gets an object that provides methods for managing the item's categories.</span></span>

> [!NOTE]
> <span data-ttu-id="f156e-284">iOS 版 Outlook 或 Android 版 Outlook 不支持此成员。</span><span class="sxs-lookup"><span data-stu-id="f156e-284">This member is not supported in Outlook on iOS or Android.</span></span>

##### <a name="type"></a><span data-ttu-id="f156e-285">类型</span><span class="sxs-lookup"><span data-stu-id="f156e-285">Type</span></span>

*   [<span data-ttu-id="f156e-286">Categories</span><span class="sxs-lookup"><span data-stu-id="f156e-286">Categories</span></span>](/javascript/api/outlook/office.categories)

##### <a name="requirements"></a><span data-ttu-id="f156e-287">Requirements</span><span class="sxs-lookup"><span data-stu-id="f156e-287">Requirements</span></span>

|<span data-ttu-id="f156e-288">要求</span><span class="sxs-lookup"><span data-stu-id="f156e-288">Requirement</span></span>|<span data-ttu-id="f156e-289">值</span><span class="sxs-lookup"><span data-stu-id="f156e-289">Value</span></span>|
|---|---|
|[<span data-ttu-id="f156e-290">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="f156e-290">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f156e-291">1.8</span><span class="sxs-lookup"><span data-stu-id="f156e-291">1.8</span></span>|
|[<span data-ttu-id="f156e-292">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="f156e-292">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f156e-293">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f156e-293">ReadItem</span></span>|
|[<span data-ttu-id="f156e-294">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="f156e-294">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="f156e-295">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="f156e-295">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="f156e-296">示例</span><span class="sxs-lookup"><span data-stu-id="f156e-296">Example</span></span>

<span data-ttu-id="f156e-297">此示例获取项的类别。</span><span class="sxs-lookup"><span data-stu-id="f156e-297">This example gets the item's categories.</span></span>

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

#### <a name="cc-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="f156e-298">cc: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="f156e-298">cc: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="f156e-299">提供对邮件的抄送 (Cc) 收件人的访问权限。</span><span class="sxs-lookup"><span data-stu-id="f156e-299">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="f156e-300">对象的类型和访问级别取决于当前项目的模式。</span><span class="sxs-lookup"><span data-stu-id="f156e-300">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="f156e-301">阅读模式</span><span class="sxs-lookup"><span data-stu-id="f156e-301">Read mode</span></span>

<span data-ttu-id="f156e-302">`cc` 属性返回包含邮件的**抄送**行上所列的每个收件人的 `EmailAddressDetails` 对象的数组。</span><span class="sxs-lookup"><span data-stu-id="f156e-302">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message.</span></span> <span data-ttu-id="f156e-303">默认情况下，集合上限为 100 个成员。</span><span class="sxs-lookup"><span data-stu-id="f156e-303">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="f156e-304">但是，在 Windows 和 Mac 上，最多可包含 500 个成员。</span><span class="sxs-lookup"><span data-stu-id="f156e-304">However, on Windows and Mac, you can get 500 members maximum.</span></span>

```js
console.log(JSON.stringify(Office.context.mailbox.item.cc));
```

##### <a name="compose-mode"></a><span data-ttu-id="f156e-305">撰写模式</span><span class="sxs-lookup"><span data-stu-id="f156e-305">Compose mode</span></span>

<span data-ttu-id="f156e-306">`cc` 属性返回一个 `Recipients` 对象，该对象提供用于获取或更新邮件的“**抄送**”行上收件人的方法。</span><span class="sxs-lookup"><span data-stu-id="f156e-306">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span> <span data-ttu-id="f156e-307">默认情况下，集合上限为 100 个成员。</span><span class="sxs-lookup"><span data-stu-id="f156e-307">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="f156e-308">但是，在 Windows 和 Mac 上，以下限制适用。</span><span class="sxs-lookup"><span data-stu-id="f156e-308">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="f156e-309">最多包含 500 个成员。</span><span class="sxs-lookup"><span data-stu-id="f156e-309">Get 500 members maximum.</span></span>
- <span data-ttu-id="f156e-310">为每个呼叫最多设置 100 个成员，总共多达 500 个成员。</span><span class="sxs-lookup"><span data-stu-id="f156e-310">Set a maximum of 100 members per call, up to 500 members total.</span></span>

```js
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="f156e-311">类型</span><span class="sxs-lookup"><span data-stu-id="f156e-311">Type</span></span>

*   <span data-ttu-id="f156e-312">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="f156e-312">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="f156e-313">Requirements</span><span class="sxs-lookup"><span data-stu-id="f156e-313">Requirements</span></span>

|<span data-ttu-id="f156e-314">要求</span><span class="sxs-lookup"><span data-stu-id="f156e-314">Requirement</span></span>|<span data-ttu-id="f156e-315">值</span><span class="sxs-lookup"><span data-stu-id="f156e-315">Value</span></span>|
|---|---|
|[<span data-ttu-id="f156e-316">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="f156e-316">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f156e-317">1.0</span><span class="sxs-lookup"><span data-stu-id="f156e-317">1.0</span></span>|
|[<span data-ttu-id="f156e-318">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="f156e-318">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f156e-319">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f156e-319">ReadItem</span></span>|
|[<span data-ttu-id="f156e-320">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="f156e-320">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="f156e-321">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="f156e-321">Compose or Read</span></span>|

<br>

---
---

#### <a name="nullable-conversationid-string"></a><span data-ttu-id="f156e-322">(nullable) conversationId: String</span><span class="sxs-lookup"><span data-stu-id="f156e-322">(nullable) conversationId: String</span></span>

<span data-ttu-id="f156e-323">获取包含特定消息的电子邮件会话的标识符。</span><span class="sxs-lookup"><span data-stu-id="f156e-323">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="f156e-p109">如果在阅读窗体或撰写窗体的回复中激活邮件应用程序，则此属性可以获得一个整数值。如果用户随后更改了回复邮件的主题（若发送回复），则该邮件的对话 ID 将改变且之前获取的值将不适用。</span><span class="sxs-lookup"><span data-stu-id="f156e-p109">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="f156e-p110">对于撰写窗体的新项目，此属性获得一个 null 值。如果用户设置一个主题并保存该项目，`conversationId` 属性将返回一个值。</span><span class="sxs-lookup"><span data-stu-id="f156e-p110">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="f156e-328">Type</span><span class="sxs-lookup"><span data-stu-id="f156e-328">Type</span></span>

*   <span data-ttu-id="f156e-329">String</span><span class="sxs-lookup"><span data-stu-id="f156e-329">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="f156e-330">Requirements</span><span class="sxs-lookup"><span data-stu-id="f156e-330">Requirements</span></span>

|<span data-ttu-id="f156e-331">要求</span><span class="sxs-lookup"><span data-stu-id="f156e-331">Requirement</span></span>|<span data-ttu-id="f156e-332">值</span><span class="sxs-lookup"><span data-stu-id="f156e-332">Value</span></span>|
|---|---|
|[<span data-ttu-id="f156e-333">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="f156e-333">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f156e-334">1.0</span><span class="sxs-lookup"><span data-stu-id="f156e-334">1.0</span></span>|
|[<span data-ttu-id="f156e-335">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="f156e-335">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f156e-336">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f156e-336">ReadItem</span></span>|
|[<span data-ttu-id="f156e-337">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="f156e-337">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="f156e-338">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="f156e-338">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="f156e-339">示例</span><span class="sxs-lookup"><span data-stu-id="f156e-339">Example</span></span>

```js
var conversationId = Office.context.mailbox.item.conversationId;
console.log("conversationId: " + conversationId);
```

<br>

---
---

#### <a name="datetimecreated-date"></a><span data-ttu-id="f156e-340">dateTimeCreated: Date</span><span class="sxs-lookup"><span data-stu-id="f156e-340">dateTimeCreated: Date</span></span>

<span data-ttu-id="f156e-p111">获取项目创建的日期和时间。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="f156e-p111">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="f156e-343">类型</span><span class="sxs-lookup"><span data-stu-id="f156e-343">Type</span></span>

*   <span data-ttu-id="f156e-344">日期</span><span class="sxs-lookup"><span data-stu-id="f156e-344">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="f156e-345">要求</span><span class="sxs-lookup"><span data-stu-id="f156e-345">Requirements</span></span>

|<span data-ttu-id="f156e-346">要求</span><span class="sxs-lookup"><span data-stu-id="f156e-346">Requirement</span></span>|<span data-ttu-id="f156e-347">值</span><span class="sxs-lookup"><span data-stu-id="f156e-347">Value</span></span>|
|---|---|
|[<span data-ttu-id="f156e-348">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="f156e-348">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f156e-349">1.0</span><span class="sxs-lookup"><span data-stu-id="f156e-349">1.0</span></span>|
|[<span data-ttu-id="f156e-350">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="f156e-350">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f156e-351">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f156e-351">ReadItem</span></span>|
|[<span data-ttu-id="f156e-352">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="f156e-352">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="f156e-353">阅读</span><span class="sxs-lookup"><span data-stu-id="f156e-353">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="f156e-354">示例</span><span class="sxs-lookup"><span data-stu-id="f156e-354">Example</span></span>

```js
var dateTimeCreated = Office.context.mailbox.item.dateTimeCreated;
console.log("Date and time created: " + dateTimeCreated);
```

<br>

---
---

#### <a name="datetimemodified-date"></a><span data-ttu-id="f156e-355">dateTimeModified: Date</span><span class="sxs-lookup"><span data-stu-id="f156e-355">dateTimeModified: Date</span></span>

<span data-ttu-id="f156e-p112">获取项目最近一次修改的日期和时间。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="f156e-p112">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="f156e-358">iOS 版 Outlook 或 Android 版 Outlook 不支持此成员。</span><span class="sxs-lookup"><span data-stu-id="f156e-358">This member is not supported in Outlook on iOS or Android.</span></span>

##### <a name="type"></a><span data-ttu-id="f156e-359">类型</span><span class="sxs-lookup"><span data-stu-id="f156e-359">Type</span></span>

*   <span data-ttu-id="f156e-360">日期</span><span class="sxs-lookup"><span data-stu-id="f156e-360">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="f156e-361">Requirements</span><span class="sxs-lookup"><span data-stu-id="f156e-361">Requirements</span></span>

|<span data-ttu-id="f156e-362">要求</span><span class="sxs-lookup"><span data-stu-id="f156e-362">Requirement</span></span>|<span data-ttu-id="f156e-363">值</span><span class="sxs-lookup"><span data-stu-id="f156e-363">Value</span></span>|
|---|---|
|[<span data-ttu-id="f156e-364">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="f156e-364">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f156e-365">1.0</span><span class="sxs-lookup"><span data-stu-id="f156e-365">1.0</span></span>|
|[<span data-ttu-id="f156e-366">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="f156e-366">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f156e-367">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f156e-367">ReadItem</span></span>|
|[<span data-ttu-id="f156e-368">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="f156e-368">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="f156e-369">阅读</span><span class="sxs-lookup"><span data-stu-id="f156e-369">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="f156e-370">示例</span><span class="sxs-lookup"><span data-stu-id="f156e-370">Example</span></span>

```js
var dateTimeModified = Office.context.mailbox.item.dateTimeModified;
console.log("Date and time modified: " + dateTimeModified);
```

<br>

---
---

#### <a name="end-datetimejavascriptapioutlookofficetime"></a><span data-ttu-id="f156e-371">end: Date|[Time](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="f156e-371">end: Date|[Time](/javascript/api/outlook/office.time)</span></span>

<span data-ttu-id="f156e-372">获取或设置约会结束的日期和时间。</span><span class="sxs-lookup"><span data-stu-id="f156e-372">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="f156e-p113">将 `end` 属性表示为协调世界时 (UTC) 的日期和时间值。可使用 [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) 方法将 end 属性值转换为客户端的本地日期和时间。</span><span class="sxs-lookup"><span data-stu-id="f156e-p113">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="f156e-375">阅读模式</span><span class="sxs-lookup"><span data-stu-id="f156e-375">Read mode</span></span>

<span data-ttu-id="f156e-376">`end` 属性返回 `Date` 对象。</span><span class="sxs-lookup"><span data-stu-id="f156e-376">The `end` property returns a `Date` object.</span></span>

```js
var end = Office.context.mailbox.item.end;
console.log("Appointment end: " + end);
```

##### <a name="compose-mode"></a><span data-ttu-id="f156e-377">撰写模式</span><span class="sxs-lookup"><span data-stu-id="f156e-377">Compose mode</span></span>

<span data-ttu-id="f156e-378">`end` 属性返回 `Time` 对象。</span><span class="sxs-lookup"><span data-stu-id="f156e-378">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="f156e-379">使用 [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) 方法设置结束时间时，应使用 [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) 方法将客户端的本地时间转换为服务器的 UTC。</span><span class="sxs-lookup"><span data-stu-id="f156e-379">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="f156e-380">以下示例使用 `Time` 对象的 [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) 方法设置约会的结束时间。</span><span class="sxs-lookup"><span data-stu-id="f156e-380">The following example sets the end time of an appointment by using the [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="f156e-381">类型</span><span class="sxs-lookup"><span data-stu-id="f156e-381">Type</span></span>

*   <span data-ttu-id="f156e-382">Date | [Time](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="f156e-382">Date | [Time](/javascript/api/outlook/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="f156e-383">要求</span><span class="sxs-lookup"><span data-stu-id="f156e-383">Requirements</span></span>

|<span data-ttu-id="f156e-384">要求</span><span class="sxs-lookup"><span data-stu-id="f156e-384">Requirement</span></span>|<span data-ttu-id="f156e-385">值</span><span class="sxs-lookup"><span data-stu-id="f156e-385">Value</span></span>|
|---|---|
|[<span data-ttu-id="f156e-386">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="f156e-386">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f156e-387">1.0</span><span class="sxs-lookup"><span data-stu-id="f156e-387">1.0</span></span>|
|[<span data-ttu-id="f156e-388">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="f156e-388">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f156e-389">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f156e-389">ReadItem</span></span>|
|[<span data-ttu-id="f156e-390">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="f156e-390">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="f156e-391">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="f156e-391">Compose or Read</span></span>|

<br>

---
---

#### <a name="enhancedlocation-enhancedlocationjavascriptapioutlookofficeenhancedlocation"></a><span data-ttu-id="f156e-392">enhancedLocation： [enhancedLocation](/javascript/api/outlook/office.enhancedlocation)</span><span class="sxs-lookup"><span data-stu-id="f156e-392">enhancedLocation: [EnhancedLocation](/javascript/api/outlook/office.enhancedlocation)</span></span>

<span data-ttu-id="f156e-393">获取或设置约会的位置。</span><span class="sxs-lookup"><span data-stu-id="f156e-393">Gets or sets the locations of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="f156e-394">阅读模式</span><span class="sxs-lookup"><span data-stu-id="f156e-394">Read mode</span></span>

<span data-ttu-id="f156e-395">该`enhancedLocation`属性返回一个[EnhancedLocation](/javascript/api/outlook/office.enhancedlocation)对象，该对象允许您获取与约会关联的一组位置（每个由[LocationDetails](/javascript/api/outlook/office.locationdetails)对象表示）。</span><span class="sxs-lookup"><span data-stu-id="f156e-395">The `enhancedLocation` property returns an [EnhancedLocation](/javascript/api/outlook/office.enhancedlocation) object that allows you to get the set of locations (each represented by a [LocationDetails](/javascript/api/outlook/office.locationdetails) object) associated with the appointment.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="f156e-396">撰写模式</span><span class="sxs-lookup"><span data-stu-id="f156e-396">Compose mode</span></span>

<span data-ttu-id="f156e-397">该`enhancedLocation`属性返回一个[EnhancedLocation](/javascript/api/outlook/office.enhancedlocation)对象，该对象提供用于获取、删除或添加约会位置的方法。</span><span class="sxs-lookup"><span data-stu-id="f156e-397">The `enhancedLocation` property returns an [EnhancedLocation](/javascript/api/outlook/office.enhancedlocation) object that provides methods to get, remove, or add locations on an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="f156e-398">类型</span><span class="sxs-lookup"><span data-stu-id="f156e-398">Type</span></span>

*   [<span data-ttu-id="f156e-399">EnhancedLocation</span><span class="sxs-lookup"><span data-stu-id="f156e-399">EnhancedLocation</span></span>](/javascript/api/outlook/office.enhancedlocation)

##### <a name="requirements"></a><span data-ttu-id="f156e-400">Requirements</span><span class="sxs-lookup"><span data-stu-id="f156e-400">Requirements</span></span>

|<span data-ttu-id="f156e-401">要求</span><span class="sxs-lookup"><span data-stu-id="f156e-401">Requirement</span></span>|<span data-ttu-id="f156e-402">值</span><span class="sxs-lookup"><span data-stu-id="f156e-402">Value</span></span>|
|---|---|
|[<span data-ttu-id="f156e-403">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="f156e-403">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f156e-404">1.8</span><span class="sxs-lookup"><span data-stu-id="f156e-404">1.8</span></span>|
|[<span data-ttu-id="f156e-405">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="f156e-405">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f156e-406">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f156e-406">ReadItem</span></span>|
|[<span data-ttu-id="f156e-407">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="f156e-407">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="f156e-408">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="f156e-408">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="f156e-409">示例</span><span class="sxs-lookup"><span data-stu-id="f156e-409">Example</span></span>

<span data-ttu-id="f156e-410">下面的示例将获取与约会相关联的当前位置。</span><span class="sxs-lookup"><span data-stu-id="f156e-410">The following example gets the current locations associated with the appointment.</span></span>

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

#### <a name="from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsfromjavascriptapioutlookofficefrom"></a><span data-ttu-id="f156e-411">发件人： [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[From](/javascript/api/outlook/office.from)</span><span class="sxs-lookup"><span data-stu-id="f156e-411">from: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[From](/javascript/api/outlook/office.from)</span></span>

<span data-ttu-id="f156e-412">获取邮件发件人的电子邮件地址。</span><span class="sxs-lookup"><span data-stu-id="f156e-412">Gets the email address of the sender of a message.</span></span>

<span data-ttu-id="f156e-p114">`from` 和 [`sender`](#sender-emailaddressdetails) 属性表示同一个人，邮件由代理人发送的除外。在此情况下，`from` 属性表示代理程序，而 sender 属性表示代理人。</span><span class="sxs-lookup"><span data-stu-id="f156e-p114">The `from` and [`sender`](#sender-emailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="f156e-415">`from` 属性中 `EmailAddressDetails` 对象的 `recipientType` 属性为 `undefined`。</span><span class="sxs-lookup"><span data-stu-id="f156e-415">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="f156e-416">阅读模式</span><span class="sxs-lookup"><span data-stu-id="f156e-416">Read mode</span></span>

<span data-ttu-id="f156e-417">`from`属性返回一个`EmailAddressDetails`对象。</span><span class="sxs-lookup"><span data-stu-id="f156e-417">The `from` property returns an `EmailAddressDetails` object.</span></span>

```js
var from = Office.context.mailbox.item.from;
console.log("From " + from);
```

##### <a name="compose-mode"></a><span data-ttu-id="f156e-418">撰写模式</span><span class="sxs-lookup"><span data-stu-id="f156e-418">Compose mode</span></span>

<span data-ttu-id="f156e-419">`from`属性返回一个`From`对象，该对象提供用于获取 "起始" 值的方法。</span><span class="sxs-lookup"><span data-stu-id="f156e-419">The `from` property returns a `From` object that provides a method to get the from value.</span></span>

```js
Office.context.mailbox.item.from.getAsync(callback);

function callback(asyncResult) {
  var from = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="f156e-420">类型</span><span class="sxs-lookup"><span data-stu-id="f156e-420">Type</span></span>

*   <span data-ttu-id="f156e-421">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [](/javascript/api/outlook/office.from)</span><span class="sxs-lookup"><span data-stu-id="f156e-421">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [From](/javascript/api/outlook/office.from)</span></span>

##### <a name="requirements"></a><span data-ttu-id="f156e-422">Requirements</span><span class="sxs-lookup"><span data-stu-id="f156e-422">Requirements</span></span>

|<span data-ttu-id="f156e-423">要求</span><span class="sxs-lookup"><span data-stu-id="f156e-423">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="f156e-424">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="f156e-424">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f156e-425">1.0</span><span class="sxs-lookup"><span data-stu-id="f156e-425">1.0</span></span>|<span data-ttu-id="f156e-426">1.7</span><span class="sxs-lookup"><span data-stu-id="f156e-426">1.7</span></span>|
|[<span data-ttu-id="f156e-427">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="f156e-427">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f156e-428">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f156e-428">ReadItem</span></span>|<span data-ttu-id="f156e-429">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="f156e-429">ReadWriteItem</span></span>|
|[<span data-ttu-id="f156e-430">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="f156e-430">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="f156e-431">阅读</span><span class="sxs-lookup"><span data-stu-id="f156e-431">Read</span></span>|<span data-ttu-id="f156e-432">撰写</span><span class="sxs-lookup"><span data-stu-id="f156e-432">Compose</span></span>|

<br>

---
---

#### <a name="internetheaders-internetheadersjavascriptapioutlookofficeinternetheaders"></a><span data-ttu-id="f156e-433">internetHeaders： [internetHeaders](/javascript/api/outlook/office.internetheaders)</span><span class="sxs-lookup"><span data-stu-id="f156e-433">internetHeaders: [InternetHeaders](/javascript/api/outlook/office.internetheaders)</span></span>

<span data-ttu-id="f156e-434">获取或设置邮件的自定义 internet 邮件头。</span><span class="sxs-lookup"><span data-stu-id="f156e-434">Gets or sets custom internet headers on a message.</span></span> <span data-ttu-id="f156e-435">仅限撰写模式。</span><span class="sxs-lookup"><span data-stu-id="f156e-435">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="f156e-436">类型</span><span class="sxs-lookup"><span data-stu-id="f156e-436">Type</span></span>

*   [<span data-ttu-id="f156e-437">InternetHeaders</span><span class="sxs-lookup"><span data-stu-id="f156e-437">InternetHeaders</span></span>](/javascript/api/outlook/office.internetheaders)

##### <a name="requirements"></a><span data-ttu-id="f156e-438">Requirements</span><span class="sxs-lookup"><span data-stu-id="f156e-438">Requirements</span></span>

|<span data-ttu-id="f156e-439">要求</span><span class="sxs-lookup"><span data-stu-id="f156e-439">Requirement</span></span>|<span data-ttu-id="f156e-440">值</span><span class="sxs-lookup"><span data-stu-id="f156e-440">Value</span></span>|
|---|---|
|[<span data-ttu-id="f156e-441">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="f156e-441">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f156e-442">1.8</span><span class="sxs-lookup"><span data-stu-id="f156e-442">1.8</span></span>|
|[<span data-ttu-id="f156e-443">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="f156e-443">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f156e-444">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f156e-444">ReadItem</span></span>|
|[<span data-ttu-id="f156e-445">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="f156e-445">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="f156e-446">撰写</span><span class="sxs-lookup"><span data-stu-id="f156e-446">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="f156e-447">示例</span><span class="sxs-lookup"><span data-stu-id="f156e-447">Example</span></span>

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

#### <a name="internetmessageid-string"></a><span data-ttu-id="f156e-448">internetMessageId: String</span><span class="sxs-lookup"><span data-stu-id="f156e-448">internetMessageId: String</span></span>

<span data-ttu-id="f156e-p116">获取电子邮件的 Internet 消息标识符。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="f156e-p116">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="f156e-451">Type</span><span class="sxs-lookup"><span data-stu-id="f156e-451">Type</span></span>

*   <span data-ttu-id="f156e-452">String</span><span class="sxs-lookup"><span data-stu-id="f156e-452">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="f156e-453">Requirements</span><span class="sxs-lookup"><span data-stu-id="f156e-453">Requirements</span></span>

|<span data-ttu-id="f156e-454">要求</span><span class="sxs-lookup"><span data-stu-id="f156e-454">Requirement</span></span>|<span data-ttu-id="f156e-455">值</span><span class="sxs-lookup"><span data-stu-id="f156e-455">Value</span></span>|
|---|---|
|[<span data-ttu-id="f156e-456">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="f156e-456">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f156e-457">1.0</span><span class="sxs-lookup"><span data-stu-id="f156e-457">1.0</span></span>|
|[<span data-ttu-id="f156e-458">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="f156e-458">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f156e-459">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f156e-459">ReadItem</span></span>|
|[<span data-ttu-id="f156e-460">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="f156e-460">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="f156e-461">阅读</span><span class="sxs-lookup"><span data-stu-id="f156e-461">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="f156e-462">示例</span><span class="sxs-lookup"><span data-stu-id="f156e-462">Example</span></span>

```js
var internetMessageId = Office.context.mailbox.item.internetMessageId;
console.log("internetMessageId: " + internetMessageId);
```

<br>

---
---

#### <a name="itemclass-string"></a><span data-ttu-id="f156e-463">itemClass: String</span><span class="sxs-lookup"><span data-stu-id="f156e-463">itemClass: String</span></span>

<span data-ttu-id="f156e-p117">获取选定项目的 Exchange Web 服务项目类。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="f156e-p117">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="f156e-p118">`itemClass` 属性指定所选项目的邮件类别。以下是邮件或约会项目的默认邮件类别。</span><span class="sxs-lookup"><span data-stu-id="f156e-p118">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

|<span data-ttu-id="f156e-468">类型</span><span class="sxs-lookup"><span data-stu-id="f156e-468">Type</span></span>|<span data-ttu-id="f156e-469">说明</span><span class="sxs-lookup"><span data-stu-id="f156e-469">Description</span></span>|<span data-ttu-id="f156e-470">项目类</span><span class="sxs-lookup"><span data-stu-id="f156e-470">item class</span></span>|
|---|---|---|
|<span data-ttu-id="f156e-471">约会项目</span><span class="sxs-lookup"><span data-stu-id="f156e-471">Appointment items</span></span>|<span data-ttu-id="f156e-472">这些是项目类 `IPM.Appointment` 或 `IPM.Appointment.Occurrence` 的日历项目。</span><span class="sxs-lookup"><span data-stu-id="f156e-472">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurrence`.</span></span>|`IPM.Appointment`<br />`IPM.Appointment.Occurrence`|
|<span data-ttu-id="f156e-473">邮件项目</span><span class="sxs-lookup"><span data-stu-id="f156e-473">Message items</span></span>|<span data-ttu-id="f156e-474">这些项目包括具有默认邮件类别 `IPM.Note` 的电子邮件，以及使用 `IPM.Schedule.Meeting` 作为基础邮件类别的会议请求、响应和取消。</span><span class="sxs-lookup"><span data-stu-id="f156e-474">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span>|`IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled`|

<span data-ttu-id="f156e-475">你可以创建用于扩展默认邮件类别的自定义邮件类别，例如，自定义约会邮件类别 `IPM.Appointment.Contoso`。</span><span class="sxs-lookup"><span data-stu-id="f156e-475">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="f156e-476">Type</span><span class="sxs-lookup"><span data-stu-id="f156e-476">Type</span></span>

*   <span data-ttu-id="f156e-477">String</span><span class="sxs-lookup"><span data-stu-id="f156e-477">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="f156e-478">Requirements</span><span class="sxs-lookup"><span data-stu-id="f156e-478">Requirements</span></span>

|<span data-ttu-id="f156e-479">要求</span><span class="sxs-lookup"><span data-stu-id="f156e-479">Requirement</span></span>|<span data-ttu-id="f156e-480">值</span><span class="sxs-lookup"><span data-stu-id="f156e-480">Value</span></span>|
|---|---|
|[<span data-ttu-id="f156e-481">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="f156e-481">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f156e-482">1.0</span><span class="sxs-lookup"><span data-stu-id="f156e-482">1.0</span></span>|
|[<span data-ttu-id="f156e-483">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="f156e-483">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f156e-484">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f156e-484">ReadItem</span></span>|
|[<span data-ttu-id="f156e-485">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="f156e-485">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="f156e-486">阅读</span><span class="sxs-lookup"><span data-stu-id="f156e-486">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="f156e-487">示例</span><span class="sxs-lookup"><span data-stu-id="f156e-487">Example</span></span>

```js
var itemClass = Office.context.mailbox.item.itemClass;
console.log("Item class: " + itemClass);
```

<br>

---
---

#### <a name="nullable-itemid-string"></a><span data-ttu-id="f156e-488">(nullable) itemId: String</span><span class="sxs-lookup"><span data-stu-id="f156e-488">(nullable) itemId: String</span></span>

<span data-ttu-id="f156e-p119">获取当前项目的 Exchange Web 服务项目标识符。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="f156e-p119">Gets the Exchange Web Services item identifier for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="f156e-491">`itemId` 属性返回的标识符与 Exchange Web 服务项目标识符相同。</span><span class="sxs-lookup"><span data-stu-id="f156e-491">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="f156e-492">`itemId` 属性与 Outlook 条目 ID 或 Outlook REST API 使用的 ID 不同。</span><span class="sxs-lookup"><span data-stu-id="f156e-492">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="f156e-493">使用此值进行 REST API 调用前，应使用 [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string) 对它进行转换。</span><span class="sxs-lookup"><span data-stu-id="f156e-493">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="f156e-494">有关详细信息，请参阅[从 Outlook 加载项使用 Outlook REST API](/outlook/add-ins/use-rest-api#get-the-item-id)。</span><span class="sxs-lookup"><span data-stu-id="f156e-494">For more details, see [Use the Outlook REST APIs from an Outlook add-in](/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

<span data-ttu-id="f156e-p121">`itemId` 属性在撰写模式下不可用。如果需要项目标识符，[`saveAsync`](#saveasyncoptions-callback) 方法可用于将项目保存到存储，这将在回调函数的 [`AsyncResult.value`](/javascript/api/office/office.asyncresult) 参数中返回项目标识符。</span><span class="sxs-lookup"><span data-stu-id="f156e-p121">The `itemId` property is not available in compose mode. If an item identifier is required, the [`saveAsync`](#saveasyncoptions-callback) method can be used to save the item to the store, which will return the item identifier in the [`AsyncResult.value`](/javascript/api/office/office.asyncresult) parameter in the callback function.</span></span>

##### <a name="type"></a><span data-ttu-id="f156e-497">Type</span><span class="sxs-lookup"><span data-stu-id="f156e-497">Type</span></span>

*   <span data-ttu-id="f156e-498">String</span><span class="sxs-lookup"><span data-stu-id="f156e-498">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="f156e-499">Requirements</span><span class="sxs-lookup"><span data-stu-id="f156e-499">Requirements</span></span>

|<span data-ttu-id="f156e-500">要求</span><span class="sxs-lookup"><span data-stu-id="f156e-500">Requirement</span></span>|<span data-ttu-id="f156e-501">值</span><span class="sxs-lookup"><span data-stu-id="f156e-501">Value</span></span>|
|---|---|
|[<span data-ttu-id="f156e-502">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="f156e-502">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f156e-503">1.0</span><span class="sxs-lookup"><span data-stu-id="f156e-503">1.0</span></span>|
|[<span data-ttu-id="f156e-504">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="f156e-504">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f156e-505">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f156e-505">ReadItem</span></span>|
|[<span data-ttu-id="f156e-506">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="f156e-506">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="f156e-507">阅读</span><span class="sxs-lookup"><span data-stu-id="f156e-507">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="f156e-508">示例</span><span class="sxs-lookup"><span data-stu-id="f156e-508">Example</span></span>

<span data-ttu-id="f156e-p122">以下代码检查项目标识符是否存在。如果 `itemId` 属性返回 `null` 或 `undefined`，则将项目保存到存储，并从异步结果中获取项目标识符。</span><span class="sxs-lookup"><span data-stu-id="f156e-p122">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

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

#### <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtype"></a><span data-ttu-id="f156e-511">itemType: [Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype)</span><span class="sxs-lookup"><span data-stu-id="f156e-511">itemType: [Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype)</span></span>

<span data-ttu-id="f156e-512">获取实例表示的项的类型。</span><span class="sxs-lookup"><span data-stu-id="f156e-512">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="f156e-513">`itemType` 属性返回其中一个 `ItemType` 枚举值，指示 `item` 对象实例是邮件还是约会。</span><span class="sxs-lookup"><span data-stu-id="f156e-513">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="f156e-514">类型</span><span class="sxs-lookup"><span data-stu-id="f156e-514">Type</span></span>

*   [<span data-ttu-id="f156e-515">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="f156e-515">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook/office.mailboxenums.itemtype)

##### <a name="requirements"></a><span data-ttu-id="f156e-516">Requirements</span><span class="sxs-lookup"><span data-stu-id="f156e-516">Requirements</span></span>

|<span data-ttu-id="f156e-517">要求</span><span class="sxs-lookup"><span data-stu-id="f156e-517">Requirement</span></span>|<span data-ttu-id="f156e-518">值</span><span class="sxs-lookup"><span data-stu-id="f156e-518">Value</span></span>|
|---|---|
|[<span data-ttu-id="f156e-519">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="f156e-519">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f156e-520">1.0</span><span class="sxs-lookup"><span data-stu-id="f156e-520">1.0</span></span>|
|[<span data-ttu-id="f156e-521">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="f156e-521">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f156e-522">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f156e-522">ReadItem</span></span>|
|[<span data-ttu-id="f156e-523">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="f156e-523">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="f156e-524">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="f156e-524">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="f156e-525">示例</span><span class="sxs-lookup"><span data-stu-id="f156e-525">Example</span></span>

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

#### <a name="location-stringlocationjavascriptapioutlookofficelocation"></a><span data-ttu-id="f156e-526">location: String|[Location](/javascript/api/outlook/office.location)</span><span class="sxs-lookup"><span data-stu-id="f156e-526">location: String|[Location](/javascript/api/outlook/office.location)</span></span>

<span data-ttu-id="f156e-527">获取或设置约会的位置。</span><span class="sxs-lookup"><span data-stu-id="f156e-527">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="f156e-528">阅读模式</span><span class="sxs-lookup"><span data-stu-id="f156e-528">Read mode</span></span>

<span data-ttu-id="f156e-529">`location` 属性返回一个包含约会位置的字符串。</span><span class="sxs-lookup"><span data-stu-id="f156e-529">The `location` property returns a string that contains the location of the appointment.</span></span>

```js
var location = Office.context.mailbox.item.location;
console.log("location: " + location);
```

##### <a name="compose-mode"></a><span data-ttu-id="f156e-530">撰写模式</span><span class="sxs-lookup"><span data-stu-id="f156e-530">Compose mode</span></span>

<span data-ttu-id="f156e-531">`location` 属性返回一个 `Location` 对象，该对象提供用于获取和设置约会位置的方法。</span><span class="sxs-lookup"><span data-stu-id="f156e-531">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

```js
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="f156e-532">类型</span><span class="sxs-lookup"><span data-stu-id="f156e-532">Type</span></span>

*   <span data-ttu-id="f156e-533">String | [Location](/javascript/api/outlook/office.location)</span><span class="sxs-lookup"><span data-stu-id="f156e-533">String | [Location](/javascript/api/outlook/office.location)</span></span>

##### <a name="requirements"></a><span data-ttu-id="f156e-534">Requirements</span><span class="sxs-lookup"><span data-stu-id="f156e-534">Requirements</span></span>

|<span data-ttu-id="f156e-535">要求</span><span class="sxs-lookup"><span data-stu-id="f156e-535">Requirement</span></span>|<span data-ttu-id="f156e-536">值</span><span class="sxs-lookup"><span data-stu-id="f156e-536">Value</span></span>|
|---|---|
|[<span data-ttu-id="f156e-537">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="f156e-537">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f156e-538">1.0</span><span class="sxs-lookup"><span data-stu-id="f156e-538">1.0</span></span>|
|[<span data-ttu-id="f156e-539">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="f156e-539">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f156e-540">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f156e-540">ReadItem</span></span>|
|[<span data-ttu-id="f156e-541">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="f156e-541">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="f156e-542">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="f156e-542">Compose or Read</span></span>|

<br>

---
---

#### <a name="normalizedsubject-string"></a><span data-ttu-id="f156e-543">normalizedSubject: String</span><span class="sxs-lookup"><span data-stu-id="f156e-543">normalizedSubject: String</span></span>

<span data-ttu-id="f156e-p123">获取删除了所有前缀（包括 `RE:` 和 `FWD:`）的项目主题。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="f156e-p123">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="f156e-p124">normalizedSubject 属性获取包含由电子邮件程序添加的任何标准前缀（如 `RE:` 和 `FW:`）的项目主题。若要获取包含完整前缀的项目主题，请使用 [`subject`](#subject-stringsubject) 属性。</span><span class="sxs-lookup"><span data-stu-id="f156e-p124">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="f156e-548">Type</span><span class="sxs-lookup"><span data-stu-id="f156e-548">Type</span></span>

*   <span data-ttu-id="f156e-549">String</span><span class="sxs-lookup"><span data-stu-id="f156e-549">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="f156e-550">Requirements</span><span class="sxs-lookup"><span data-stu-id="f156e-550">Requirements</span></span>

|<span data-ttu-id="f156e-551">要求</span><span class="sxs-lookup"><span data-stu-id="f156e-551">Requirement</span></span>|<span data-ttu-id="f156e-552">值</span><span class="sxs-lookup"><span data-stu-id="f156e-552">Value</span></span>|
|---|---|
|[<span data-ttu-id="f156e-553">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="f156e-553">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f156e-554">1.0</span><span class="sxs-lookup"><span data-stu-id="f156e-554">1.0</span></span>|
|[<span data-ttu-id="f156e-555">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="f156e-555">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f156e-556">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f156e-556">ReadItem</span></span>|
|[<span data-ttu-id="f156e-557">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="f156e-557">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="f156e-558">阅读</span><span class="sxs-lookup"><span data-stu-id="f156e-558">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="f156e-559">示例</span><span class="sxs-lookup"><span data-stu-id="f156e-559">Example</span></span>

```js
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
console.log("Normalized subject: " + normalizedSubject);
```

<br>

---
---

#### <a name="notificationmessages-notificationmessagesjavascriptapioutlookofficenotificationmessages"></a><span data-ttu-id="f156e-560">notificationMessages: [NotificationMessages](/javascript/api/outlook/office.notificationmessages)</span><span class="sxs-lookup"><span data-stu-id="f156e-560">notificationMessages: [NotificationMessages](/javascript/api/outlook/office.notificationmessages)</span></span>

<span data-ttu-id="f156e-561">获取项目的通知邮件。</span><span class="sxs-lookup"><span data-stu-id="f156e-561">Gets the notification messages for an item.</span></span>

##### <a name="type"></a><span data-ttu-id="f156e-562">类型</span><span class="sxs-lookup"><span data-stu-id="f156e-562">Type</span></span>

*   [<span data-ttu-id="f156e-563">NotificationMessages</span><span class="sxs-lookup"><span data-stu-id="f156e-563">NotificationMessages</span></span>](/javascript/api/outlook/office.notificationmessages)

##### <a name="requirements"></a><span data-ttu-id="f156e-564">要求</span><span class="sxs-lookup"><span data-stu-id="f156e-564">Requirements</span></span>

|<span data-ttu-id="f156e-565">要求</span><span class="sxs-lookup"><span data-stu-id="f156e-565">Requirement</span></span>|<span data-ttu-id="f156e-566">值</span><span class="sxs-lookup"><span data-stu-id="f156e-566">Value</span></span>|
|---|---|
|[<span data-ttu-id="f156e-567">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="f156e-567">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f156e-568">1.3</span><span class="sxs-lookup"><span data-stu-id="f156e-568">1.3</span></span>|
|[<span data-ttu-id="f156e-569">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="f156e-569">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f156e-570">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f156e-570">ReadItem</span></span>|
|[<span data-ttu-id="f156e-571">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="f156e-571">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="f156e-572">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="f156e-572">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="f156e-573">示例</span><span class="sxs-lookup"><span data-stu-id="f156e-573">Example</span></span>

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

#### <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="f156e-574">optionalAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="f156e-574">optionalAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="f156e-575">提供对事件的可选与会者的访问权限。</span><span class="sxs-lookup"><span data-stu-id="f156e-575">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="f156e-576">对象的类型和访问级别取决于当前项目的模式。</span><span class="sxs-lookup"><span data-stu-id="f156e-576">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="f156e-577">阅读模式</span><span class="sxs-lookup"><span data-stu-id="f156e-577">Read mode</span></span>

<span data-ttu-id="f156e-578">`optionalAttendees` 属性返回一个数组，其中包含每个可选与会者的 `EmailAddressDetails` 对象。</span><span class="sxs-lookup"><span data-stu-id="f156e-578">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span> <span data-ttu-id="f156e-579">默认情况下，集合上限为 100 个成员。</span><span class="sxs-lookup"><span data-stu-id="f156e-579">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="f156e-580">但是，在 Windows 和 Mac 上，最多可包含 500 个成员。</span><span class="sxs-lookup"><span data-stu-id="f156e-580">However, on Windows and Mac, you can get 500 members maximum.</span></span>

```js
var optionalAttendees = Office.context.mailbox.item.optionalAttendees;
console.log("Optional attendees: " + JSON.stringify(optionalAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="f156e-581">撰写模式</span><span class="sxs-lookup"><span data-stu-id="f156e-581">Compose mode</span></span>

<span data-ttu-id="f156e-582">`optionalAttendees` 属性返回一个 `Recipients` 对象，该对象提供用于获取或更新可选与会者的方法。</span><span class="sxs-lookup"><span data-stu-id="f156e-582">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span> <span data-ttu-id="f156e-583">默认情况下，集合上限为 100 个成员。</span><span class="sxs-lookup"><span data-stu-id="f156e-583">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="f156e-584">但是，在 Windows 和 Mac 上，以下限制适用。</span><span class="sxs-lookup"><span data-stu-id="f156e-584">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="f156e-585">最多包含 500 个成员。</span><span class="sxs-lookup"><span data-stu-id="f156e-585">Get 500 members maximum.</span></span>
- <span data-ttu-id="f156e-586">为每个呼叫最多设置 100 个成员，总共多达 500 个成员。</span><span class="sxs-lookup"><span data-stu-id="f156e-586">Set a maximum of 100 members per call, up to 500 members total.</span></span>

```js
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="f156e-587">类型</span><span class="sxs-lookup"><span data-stu-id="f156e-587">Type</span></span>

*   <span data-ttu-id="f156e-588">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="f156e-588">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="f156e-589">Requirements</span><span class="sxs-lookup"><span data-stu-id="f156e-589">Requirements</span></span>

|<span data-ttu-id="f156e-590">要求</span><span class="sxs-lookup"><span data-stu-id="f156e-590">Requirement</span></span>|<span data-ttu-id="f156e-591">值</span><span class="sxs-lookup"><span data-stu-id="f156e-591">Value</span></span>|
|---|---|
|[<span data-ttu-id="f156e-592">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="f156e-592">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f156e-593">1.0</span><span class="sxs-lookup"><span data-stu-id="f156e-593">1.0</span></span>|
|[<span data-ttu-id="f156e-594">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="f156e-594">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f156e-595">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f156e-595">ReadItem</span></span>|
|[<span data-ttu-id="f156e-596">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="f156e-596">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="f156e-597">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="f156e-597">Compose or Read</span></span>|

<br>

---
---

#### <a name="organizer-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsorganizerjavascriptapioutlookofficeorganizer"></a><span data-ttu-id="f156e-598">组织者： [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[组织者](/javascript/api/outlook/office.organizer)</span><span class="sxs-lookup"><span data-stu-id="f156e-598">organizer: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[Organizer](/javascript/api/outlook/office.organizer)</span></span>

<span data-ttu-id="f156e-599">获取指定会议的组织者的电子邮件地址。</span><span class="sxs-lookup"><span data-stu-id="f156e-599">Gets the email address of the organizer for a specified meeting.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="f156e-600">阅读模式</span><span class="sxs-lookup"><span data-stu-id="f156e-600">Read mode</span></span>

<span data-ttu-id="f156e-601">该`organizer`属性返回一个[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)对象，该对象代表会议组织者。</span><span class="sxs-lookup"><span data-stu-id="f156e-601">The `organizer` property returns an [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) object that represents the meeting organizer.</span></span>

```js
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
console.log("Organizer: " + organizerName + " (" + organizerAddress + ")");
```

##### <a name="compose-mode"></a><span data-ttu-id="f156e-602">撰写模式</span><span class="sxs-lookup"><span data-stu-id="f156e-602">Compose mode</span></span>

<span data-ttu-id="f156e-603">该`organizer`属性返回一个[管理](/javascript/api/outlook/office.organizer)器对象，该对象提供获取组织者值的方法。</span><span class="sxs-lookup"><span data-stu-id="f156e-603">The `organizer` property returns an [Organizer](/javascript/api/outlook/office.organizer) object that provides a method to get the organizer value.</span></span>

```js
Office.context.mailbox.item.organizer.getAsync(
  function(asyncResult) {
    console.log(JSON.stringify(asyncResult));
  }
);
```

##### <a name="type"></a><span data-ttu-id="f156e-604">类型</span><span class="sxs-lookup"><span data-stu-id="f156e-604">Type</span></span>

*   <span data-ttu-id="f156e-605">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [组织者](/javascript/api/outlook/office.organizer)</span><span class="sxs-lookup"><span data-stu-id="f156e-605">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [Organizer](/javascript/api/outlook/office.organizer)</span></span>

##### <a name="requirements"></a><span data-ttu-id="f156e-606">Requirements</span><span class="sxs-lookup"><span data-stu-id="f156e-606">Requirements</span></span>

|<span data-ttu-id="f156e-607">要求</span><span class="sxs-lookup"><span data-stu-id="f156e-607">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="f156e-608">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="f156e-608">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f156e-609">1.0</span><span class="sxs-lookup"><span data-stu-id="f156e-609">1.0</span></span>|<span data-ttu-id="f156e-610">1.7</span><span class="sxs-lookup"><span data-stu-id="f156e-610">1.7</span></span>|
|[<span data-ttu-id="f156e-611">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="f156e-611">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f156e-612">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f156e-612">ReadItem</span></span>|<span data-ttu-id="f156e-613">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="f156e-613">ReadWriteItem</span></span>|
|[<span data-ttu-id="f156e-614">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="f156e-614">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="f156e-615">阅读</span><span class="sxs-lookup"><span data-stu-id="f156e-615">Read</span></span>|<span data-ttu-id="f156e-616">撰写</span><span class="sxs-lookup"><span data-stu-id="f156e-616">Compose</span></span>|

<br>

---
---

#### <a name="nullable-recurrence-recurrencejavascriptapioutlookofficerecurrence"></a><span data-ttu-id="f156e-617">（可以为 null）定期：[定期](/javascript/api/outlook/office.recurrence)</span><span class="sxs-lookup"><span data-stu-id="f156e-617">(nullable) recurrence: [Recurrence](/javascript/api/outlook/office.recurrence)</span></span>

<span data-ttu-id="f156e-618">获取或设置约会的定期模式。</span><span class="sxs-lookup"><span data-stu-id="f156e-618">Gets or sets the recurrence pattern of an appointment.</span></span> <span data-ttu-id="f156e-619">获取会议请求的定期模式。</span><span class="sxs-lookup"><span data-stu-id="f156e-619">Gets the recurrence pattern of a meeting request.</span></span> <span data-ttu-id="f156e-620">约会项目的阅读和撰写模式。</span><span class="sxs-lookup"><span data-stu-id="f156e-620">Read and compose modes for appointment items.</span></span> <span data-ttu-id="f156e-621">会议请求项目的阅读模式。</span><span class="sxs-lookup"><span data-stu-id="f156e-621">Read mode for meeting request items.</span></span>

<span data-ttu-id="f156e-622">如果`recurrence`项目是系列中的一个系列或一个实例，则该属性返回定期约会或会议请求的[定期](/javascript/api/outlook/office.recurrence)对象。</span><span class="sxs-lookup"><span data-stu-id="f156e-622">The `recurrence` property returns a [recurrence](/javascript/api/outlook/office.recurrence) object for recurring appointments or meetings requests if an item is a series or an instance in a series.</span></span> <span data-ttu-id="f156e-623">`null`返回单个约会的单个约会和会议请求。</span><span class="sxs-lookup"><span data-stu-id="f156e-623">`null` is returned for single appointments and meeting requests of single appointments.</span></span> <span data-ttu-id="f156e-624">`undefined`对于不是会议请求的邮件，将返回。</span><span class="sxs-lookup"><span data-stu-id="f156e-624">`undefined` is returned for messages that are not meeting requests.</span></span>

> <span data-ttu-id="f156e-625">注意：会议请求的`itemClass`值为 IPM。Schedule. 会议请求。</span><span class="sxs-lookup"><span data-stu-id="f156e-625">Note: Meeting requests have an `itemClass` value of IPM.Schedule.Meeting.Request.</span></span>

> <span data-ttu-id="f156e-626">注意：如果定期对象为`null`，则表示该对象是单个约会的单个约会或会议请求，而不是某个系列的一部分。</span><span class="sxs-lookup"><span data-stu-id="f156e-626">Note: If the recurrence object is `null`, this indicates that the object is a single appointment or a meeting request of a single appointment and NOT a part of a series.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="f156e-627">阅读模式</span><span class="sxs-lookup"><span data-stu-id="f156e-627">Read mode</span></span>

<span data-ttu-id="f156e-628">该`recurrence`属性返回一个代表约会定期的[定期](/javascript/api/outlook/office.recurrence)对象。</span><span class="sxs-lookup"><span data-stu-id="f156e-628">The `recurrence` property returns a [Recurrence](/javascript/api/outlook/office.recurrence) object that represents the appointment recurrence.</span></span> <span data-ttu-id="f156e-629">此功能适用于约会和会议请求。</span><span class="sxs-lookup"><span data-stu-id="f156e-629">This is available for appointments and meeting requests.</span></span>

```js
var recurrence = Office.context.mailbox.item.recurrence;
console.log("Recurrence: " + JSON.stringify(recurrence));
```

##### <a name="compose-mode"></a><span data-ttu-id="f156e-630">撰写模式</span><span class="sxs-lookup"><span data-stu-id="f156e-630">Compose mode</span></span>

<span data-ttu-id="f156e-631">该`recurrence`属性返回一个[定期](/javascript/api/outlook/office.recurrence)对象，该对象提供用于管理约会周期的方法。</span><span class="sxs-lookup"><span data-stu-id="f156e-631">The `recurrence` property returns a [Recurrence](/javascript/api/outlook/office.recurrence) object that provides methods to manage the appointment recurrence.</span></span> <span data-ttu-id="f156e-632">这可用于约会。</span><span class="sxs-lookup"><span data-stu-id="f156e-632">This is available for appointments.</span></span>

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

##### <a name="type"></a><span data-ttu-id="f156e-633">类型</span><span class="sxs-lookup"><span data-stu-id="f156e-633">Type</span></span>

* [<span data-ttu-id="f156e-634">循环</span><span class="sxs-lookup"><span data-stu-id="f156e-634">Recurrence</span></span>](/javascript/api/outlook/office.recurrence)

|<span data-ttu-id="f156e-635">要求</span><span class="sxs-lookup"><span data-stu-id="f156e-635">Requirement</span></span>|<span data-ttu-id="f156e-636">值</span><span class="sxs-lookup"><span data-stu-id="f156e-636">Value</span></span>|
|---|---|
|[<span data-ttu-id="f156e-637">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="f156e-637">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f156e-638">1.7</span><span class="sxs-lookup"><span data-stu-id="f156e-638">1.7</span></span>|
|[<span data-ttu-id="f156e-639">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="f156e-639">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f156e-640">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f156e-640">ReadItem</span></span>|
|[<span data-ttu-id="f156e-641">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="f156e-641">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="f156e-642">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="f156e-642">Compose or Read</span></span>|

<br>

---
---

#### <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="f156e-643">requiredAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="f156e-643">requiredAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="f156e-644">提供对事件的必需与会者的访问权限。</span><span class="sxs-lookup"><span data-stu-id="f156e-644">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="f156e-645">对象的类型和访问级别取决于当前项目的模式。</span><span class="sxs-lookup"><span data-stu-id="f156e-645">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="f156e-646">阅读模式</span><span class="sxs-lookup"><span data-stu-id="f156e-646">Read mode</span></span>

<span data-ttu-id="f156e-647">`requiredAttendees` 属性返回一个数组，其中包含每个必需与会者的 `EmailAddressDetails` 对象。</span><span class="sxs-lookup"><span data-stu-id="f156e-647">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span> <span data-ttu-id="f156e-648">默认情况下，集合上限为 100 个成员。</span><span class="sxs-lookup"><span data-stu-id="f156e-648">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="f156e-649">但是，在 Windows 和 Mac 上，最多可包含 500 个成员。</span><span class="sxs-lookup"><span data-stu-id="f156e-649">However, on Windows and Mac, you can get 500 members maximum.</span></span>

```js
var requiredAttendees = Office.context.mailbox.item.requiredAttendees;
console.log("Required attendees: " + JSON.stringify(requiredAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="f156e-650">撰写模式</span><span class="sxs-lookup"><span data-stu-id="f156e-650">Compose mode</span></span>

<span data-ttu-id="f156e-651">`requiredAttendees` 属性返回一个 `Recipients` 对象，该对象提供用于获取或更新必需与会者的方法。</span><span class="sxs-lookup"><span data-stu-id="f156e-651">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span> <span data-ttu-id="f156e-652">默认情况下，集合上限为 100 个成员。</span><span class="sxs-lookup"><span data-stu-id="f156e-652">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="f156e-653">但是，在 Windows 和 Mac 上，以下限制适用。</span><span class="sxs-lookup"><span data-stu-id="f156e-653">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="f156e-654">最多包含 500 个成员。</span><span class="sxs-lookup"><span data-stu-id="f156e-654">Get 500 members maximum.</span></span>
- <span data-ttu-id="f156e-655">为每个呼叫最多设置 100 个成员，总共多达 500 个成员。</span><span class="sxs-lookup"><span data-stu-id="f156e-655">Set a maximum of 100 members per call, up to 500 members total.</span></span>

```js
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
  console.log(JSON.stringify(arrayOfRequiredAttendeesRecipients));
}
```

##### <a name="type"></a><span data-ttu-id="f156e-656">类型</span><span class="sxs-lookup"><span data-stu-id="f156e-656">Type</span></span>

*   <span data-ttu-id="f156e-657">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="f156e-657">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="f156e-658">Requirements</span><span class="sxs-lookup"><span data-stu-id="f156e-658">Requirements</span></span>

|<span data-ttu-id="f156e-659">要求</span><span class="sxs-lookup"><span data-stu-id="f156e-659">Requirement</span></span>|<span data-ttu-id="f156e-660">值</span><span class="sxs-lookup"><span data-stu-id="f156e-660">Value</span></span>|
|---|---|
|[<span data-ttu-id="f156e-661">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="f156e-661">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f156e-662">1.0</span><span class="sxs-lookup"><span data-stu-id="f156e-662">1.0</span></span>|
|[<span data-ttu-id="f156e-663">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="f156e-663">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f156e-664">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f156e-664">ReadItem</span></span>|
|[<span data-ttu-id="f156e-665">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="f156e-665">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="f156e-666">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="f156e-666">Compose or Read</span></span>|

<br>

---
---

#### <a name="sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetails"></a><span data-ttu-id="f156e-667">sender: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="f156e-667">sender: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)</span></span>

<span data-ttu-id="f156e-p135">获取电子邮件发件人的电子邮件地址。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="f156e-p135">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="f156e-p136">[`from`](#from-emailaddressdetailsfrom) 和 `sender` 属性表示同一个人，邮件由代理人发送的除外。在此情况下，`from` 属性表示代理程序，而 sender 属性表示代理人。</span><span class="sxs-lookup"><span data-stu-id="f156e-p136">The [`from`](#from-emailaddressdetailsfrom) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="f156e-672">`sender` 属性中 `EmailAddressDetails` 对象的 `recipientType` 属性为 `undefined`。</span><span class="sxs-lookup"><span data-stu-id="f156e-672">The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="f156e-673">类型</span><span class="sxs-lookup"><span data-stu-id="f156e-673">Type</span></span>

*   [<span data-ttu-id="f156e-674">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="f156e-674">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="f156e-675">要求</span><span class="sxs-lookup"><span data-stu-id="f156e-675">Requirements</span></span>

|<span data-ttu-id="f156e-676">要求</span><span class="sxs-lookup"><span data-stu-id="f156e-676">Requirement</span></span>|<span data-ttu-id="f156e-677">值</span><span class="sxs-lookup"><span data-stu-id="f156e-677">Value</span></span>|
|---|---|
|[<span data-ttu-id="f156e-678">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="f156e-678">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f156e-679">1.0</span><span class="sxs-lookup"><span data-stu-id="f156e-679">1.0</span></span>|
|[<span data-ttu-id="f156e-680">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="f156e-680">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f156e-681">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f156e-681">ReadItem</span></span>|
|[<span data-ttu-id="f156e-682">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="f156e-682">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="f156e-683">阅读</span><span class="sxs-lookup"><span data-stu-id="f156e-683">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="f156e-684">示例</span><span class="sxs-lookup"><span data-stu-id="f156e-684">Example</span></span>

```js
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
console.log("Sender: " + senderName + " (" + senderAddress + ")");
```

<br>

---
---

#### <a name="nullable-seriesid-string"></a><span data-ttu-id="f156e-685">（可以为 null） Webcasts&seriesid： String</span><span class="sxs-lookup"><span data-stu-id="f156e-685">(nullable) seriesId: String</span></span>

<span data-ttu-id="f156e-686">获取实例所属的系列的 id。</span><span class="sxs-lookup"><span data-stu-id="f156e-686">Gets the id of the series that an instance belongs to.</span></span>

<span data-ttu-id="f156e-687">在 web 上的 Outlook 和桌面客户端中`seriesId` ，返回此项所属的父（系列）项的 Exchange web 服务（EWS） ID。</span><span class="sxs-lookup"><span data-stu-id="f156e-687">In Outlook on the web and desktop clients, the `seriesId` returns the Exchange Web Services (EWS) ID of the parent (series) item that this item belongs to.</span></span> <span data-ttu-id="f156e-688">但是，在 iOS 和 Android 中， `seriesId`将返回父项的 REST ID。</span><span class="sxs-lookup"><span data-stu-id="f156e-688">However, in iOS and Android, the `seriesId` returns the REST ID of the parent item.</span></span>

> [!NOTE]
> <span data-ttu-id="f156e-689">`seriesId` 属性返回的标识符与 Exchange Web 服务项目标识符相同。</span><span class="sxs-lookup"><span data-stu-id="f156e-689">The identifier returned by the `seriesId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="f156e-690">`seriesId`属性与 OUTLOOK REST API 使用的 outlook id 不相同。</span><span class="sxs-lookup"><span data-stu-id="f156e-690">The `seriesId` property is not identical to the Outlook IDs used by the Outlook REST API.</span></span> <span data-ttu-id="f156e-691">使用此值进行 REST API 调用前，应使用 [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string) 对它进行转换。</span><span class="sxs-lookup"><span data-stu-id="f156e-691">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="f156e-692">有关详细信息，请参阅[从 Outlook 加载项使用 Outlook REST API](/outlook/add-ins/use-rest-api)。</span><span class="sxs-lookup"><span data-stu-id="f156e-692">For more details, see [Use the Outlook REST APIs from an Outlook add-in](/outlook/add-ins/use-rest-api).</span></span>

<span data-ttu-id="f156e-693">对于`seriesId`不包含`null`父项（如单个约会、系列项或会议请求）的项，该属性将返回， `undefined`对于不是会议请求的任何其他项，该属性返回。</span><span class="sxs-lookup"><span data-stu-id="f156e-693">The `seriesId` property returns `null` for items that do not have parent items such as single appointments, series items, or meeting requests and returns `undefined` for any other items that are not meeting requests.</span></span>

##### <a name="type"></a><span data-ttu-id="f156e-694">Type</span><span class="sxs-lookup"><span data-stu-id="f156e-694">Type</span></span>

* <span data-ttu-id="f156e-695">String</span><span class="sxs-lookup"><span data-stu-id="f156e-695">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="f156e-696">Requirements</span><span class="sxs-lookup"><span data-stu-id="f156e-696">Requirements</span></span>

|<span data-ttu-id="f156e-697">要求</span><span class="sxs-lookup"><span data-stu-id="f156e-697">Requirement</span></span>|<span data-ttu-id="f156e-698">值</span><span class="sxs-lookup"><span data-stu-id="f156e-698">Value</span></span>|
|---|---|
|[<span data-ttu-id="f156e-699">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="f156e-699">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f156e-700">1.7</span><span class="sxs-lookup"><span data-stu-id="f156e-700">1.7</span></span>|
|[<span data-ttu-id="f156e-701">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="f156e-701">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f156e-702">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f156e-702">ReadItem</span></span>|
|[<span data-ttu-id="f156e-703">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="f156e-703">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="f156e-704">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="f156e-704">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="f156e-705">示例</span><span class="sxs-lookup"><span data-stu-id="f156e-705">Example</span></span>

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

#### <a name="start-datetimejavascriptapioutlookofficetime"></a><span data-ttu-id="f156e-706">start: Date|[Time](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="f156e-706">start: Date|[Time](/javascript/api/outlook/office.time)</span></span>

<span data-ttu-id="f156e-707">获取或设置约会开始的日期和时间。</span><span class="sxs-lookup"><span data-stu-id="f156e-707">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="f156e-p139">将 `start` 属性表示为协调世界时 (UTC) 的日期和时间值。可使用 [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) 方法将该值转换为客户端的本地日期和时间。</span><span class="sxs-lookup"><span data-stu-id="f156e-p139">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="f156e-710">阅读模式</span><span class="sxs-lookup"><span data-stu-id="f156e-710">Read mode</span></span>

<span data-ttu-id="f156e-711">`start` 属性返回 `Date` 对象。</span><span class="sxs-lookup"><span data-stu-id="f156e-711">The `start` property returns a `Date` object.</span></span>

```js
var start = Office.context.mailbox.item.start;
console.log("Appointment start: " + JSON.stringify(start));
```

##### <a name="compose-mode"></a><span data-ttu-id="f156e-712">撰写模式</span><span class="sxs-lookup"><span data-stu-id="f156e-712">Compose mode</span></span>

<span data-ttu-id="f156e-713">`start` 属性返回 `Time` 对象。</span><span class="sxs-lookup"><span data-stu-id="f156e-713">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="f156e-714">使用 [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) 方法设置开始时间时，应使用 [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) 方法将客户端的本地时间转换为服务器的 UTC。</span><span class="sxs-lookup"><span data-stu-id="f156e-714">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="f156e-715">以下示例通过使用 `Time` 对象的 [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) 方法，设置撰写模式下约会的开始时间。</span><span class="sxs-lookup"><span data-stu-id="f156e-715">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="f156e-716">类型</span><span class="sxs-lookup"><span data-stu-id="f156e-716">Type</span></span>

*   <span data-ttu-id="f156e-717">Date | [Time](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="f156e-717">Date | [Time](/javascript/api/outlook/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="f156e-718">Requirements</span><span class="sxs-lookup"><span data-stu-id="f156e-718">Requirements</span></span>

|<span data-ttu-id="f156e-719">要求</span><span class="sxs-lookup"><span data-stu-id="f156e-719">Requirement</span></span>|<span data-ttu-id="f156e-720">值</span><span class="sxs-lookup"><span data-stu-id="f156e-720">Value</span></span>|
|---|---|
|[<span data-ttu-id="f156e-721">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="f156e-721">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f156e-722">1.0</span><span class="sxs-lookup"><span data-stu-id="f156e-722">1.0</span></span>|
|[<span data-ttu-id="f156e-723">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="f156e-723">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f156e-724">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f156e-724">ReadItem</span></span>|
|[<span data-ttu-id="f156e-725">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="f156e-725">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="f156e-726">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="f156e-726">Compose or Read</span></span>|

<br>

---
---

#### <a name="subject-stringsubjectjavascriptapioutlookofficesubject"></a><span data-ttu-id="f156e-727">subject: String|[Subject](/javascript/api/outlook/office.subject)</span><span class="sxs-lookup"><span data-stu-id="f156e-727">subject: String|[Subject](/javascript/api/outlook/office.subject)</span></span>

<span data-ttu-id="f156e-728">获取或设置显示在项目的主题字段中的说明。</span><span class="sxs-lookup"><span data-stu-id="f156e-728">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="f156e-729">`subject` 属性获取或设置由电子邮件服务器发送项目时的整个主题。</span><span class="sxs-lookup"><span data-stu-id="f156e-729">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="f156e-730">阅读模式</span><span class="sxs-lookup"><span data-stu-id="f156e-730">Read mode</span></span>

<span data-ttu-id="f156e-p140">`subject` 属性返回一个字符串。使用 [`normalizedSubject`](#normalizedsubject-string) 属性获取不带任何前导前缀（如 `RE:` 和 `FW:`）的主题。</span><span class="sxs-lookup"><span data-stu-id="f156e-p140">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

<span data-ttu-id="f156e-733">以下 JavaScript 代码示例显示了如何访问 Outlook 中当前项目的 `subject` 属性。</span><span class="sxs-lookup"><span data-stu-id="f156e-733">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

```js
var subject = Office.context.mailbox.item.subject;
console.log(subject);
```

##### <a name="compose-mode"></a><span data-ttu-id="f156e-734">撰写模式</span><span class="sxs-lookup"><span data-stu-id="f156e-734">Compose mode</span></span>
<span data-ttu-id="f156e-735">`subject` 属性返回一个 `Subject` 对象，该对象提供用于获取和设置主题的方法。</span><span class="sxs-lookup"><span data-stu-id="f156e-735">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```js
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
  console.log(subject);
}
```

##### <a name="type"></a><span data-ttu-id="f156e-736">类型</span><span class="sxs-lookup"><span data-stu-id="f156e-736">Type</span></span>

*   <span data-ttu-id="f156e-737">String | [Subject](/javascript/api/outlook/office.subject)</span><span class="sxs-lookup"><span data-stu-id="f156e-737">String | [Subject](/javascript/api/outlook/office.subject)</span></span>

##### <a name="requirements"></a><span data-ttu-id="f156e-738">要求</span><span class="sxs-lookup"><span data-stu-id="f156e-738">Requirements</span></span>

|<span data-ttu-id="f156e-739">要求</span><span class="sxs-lookup"><span data-stu-id="f156e-739">Requirement</span></span>|<span data-ttu-id="f156e-740">值</span><span class="sxs-lookup"><span data-stu-id="f156e-740">Value</span></span>|
|---|---|
|[<span data-ttu-id="f156e-741">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="f156e-741">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f156e-742">1.0</span><span class="sxs-lookup"><span data-stu-id="f156e-742">1.0</span></span>|
|[<span data-ttu-id="f156e-743">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="f156e-743">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f156e-744">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f156e-744">ReadItem</span></span>|
|[<span data-ttu-id="f156e-745">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="f156e-745">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="f156e-746">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="f156e-746">Compose or Read</span></span>|

<br>

---
---

#### <a name="to-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="f156e-747">to: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="f156e-747">to: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="f156e-748">提供对邮件的“**收件人**”行上的收件人的访问权限。</span><span class="sxs-lookup"><span data-stu-id="f156e-748">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="f156e-749">对象的类型和访问级别取决于当前项目的模式。</span><span class="sxs-lookup"><span data-stu-id="f156e-749">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="f156e-750">阅读模式</span><span class="sxs-lookup"><span data-stu-id="f156e-750">Read mode</span></span>

<span data-ttu-id="f156e-751">`to` 属性返回包含邮件的**收件人**行上所列的每个收件人的 `EmailAddressDetails` 对象的数组。</span><span class="sxs-lookup"><span data-stu-id="f156e-751">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message.</span></span> <span data-ttu-id="f156e-752">默认情况下，集合上限为 100 个成员。</span><span class="sxs-lookup"><span data-stu-id="f156e-752">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="f156e-753">但是，在 Windows 和 Mac 上，最多可包含 500 个成员。</span><span class="sxs-lookup"><span data-stu-id="f156e-753">However, on Windows and Mac, you can get 500 members maximum.</span></span>

```js
console.log(JSON.stringify(Office.context.mailbox.item.to));
```

##### <a name="compose-mode"></a><span data-ttu-id="f156e-754">撰写模式</span><span class="sxs-lookup"><span data-stu-id="f156e-754">Compose mode</span></span>

<span data-ttu-id="f156e-755">`to` 属性返回一个 `Recipients` 对象，该对象提供用于获取或更新邮件的“**收件人**”行上收件人的方法。</span><span class="sxs-lookup"><span data-stu-id="f156e-755">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span> <span data-ttu-id="f156e-756">默认情况下，集合上限为 100 个成员。</span><span class="sxs-lookup"><span data-stu-id="f156e-756">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="f156e-757">但是，在 Windows 和 Mac 上，以下限制适用。</span><span class="sxs-lookup"><span data-stu-id="f156e-757">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="f156e-758">最多包含 500 个成员。</span><span class="sxs-lookup"><span data-stu-id="f156e-758">Get 500 members maximum.</span></span>
- <span data-ttu-id="f156e-759">为每个呼叫最多设置 100 个成员，总共多达 500 个成员。</span><span class="sxs-lookup"><span data-stu-id="f156e-759">Set a maximum of 100 members per call, up to 500 members total.</span></span>

```js
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="f156e-760">类型</span><span class="sxs-lookup"><span data-stu-id="f156e-760">Type</span></span>

*   <span data-ttu-id="f156e-761">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="f156e-761">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="f156e-762">Requirements</span><span class="sxs-lookup"><span data-stu-id="f156e-762">Requirements</span></span>

|<span data-ttu-id="f156e-763">要求</span><span class="sxs-lookup"><span data-stu-id="f156e-763">Requirement</span></span>|<span data-ttu-id="f156e-764">值</span><span class="sxs-lookup"><span data-stu-id="f156e-764">Value</span></span>|
|---|---|
|[<span data-ttu-id="f156e-765">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="f156e-765">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f156e-766">1.0</span><span class="sxs-lookup"><span data-stu-id="f156e-766">1.0</span></span>|
|[<span data-ttu-id="f156e-767">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="f156e-767">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f156e-768">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f156e-768">ReadItem</span></span>|
|[<span data-ttu-id="f156e-769">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="f156e-769">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="f156e-770">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="f156e-770">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="f156e-771">方法</span><span class="sxs-lookup"><span data-stu-id="f156e-771">Methods</span></span>

#### <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="f156e-772">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="f156e-772">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="f156e-773">将文件作为附件添加到邮件或约会。</span><span class="sxs-lookup"><span data-stu-id="f156e-773">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="f156e-774">`addFileAttachmentAsync` 方法在指定的 URI 上载文件并将其附加到撰写窗体中的项目。</span><span class="sxs-lookup"><span data-stu-id="f156e-774">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="f156e-775">随后可以将该标识符与 [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) 方法一同使用，以删除同一个会话中的附件。</span><span class="sxs-lookup"><span data-stu-id="f156e-775">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="f156e-776">参数</span><span class="sxs-lookup"><span data-stu-id="f156e-776">Parameters</span></span>
|<span data-ttu-id="f156e-777">名称</span><span class="sxs-lookup"><span data-stu-id="f156e-777">Name</span></span>|<span data-ttu-id="f156e-778">类型</span><span class="sxs-lookup"><span data-stu-id="f156e-778">Type</span></span>|<span data-ttu-id="f156e-779">属性</span><span class="sxs-lookup"><span data-stu-id="f156e-779">Attributes</span></span>|<span data-ttu-id="f156e-780">说明</span><span class="sxs-lookup"><span data-stu-id="f156e-780">Description</span></span>|
|---|---|---|---|
|`uri`|<span data-ttu-id="f156e-781">字符串</span><span class="sxs-lookup"><span data-stu-id="f156e-781">String</span></span>||<span data-ttu-id="f156e-p144">提供附加到邮件或约会的文件的位置的 URI。最大长度为 2048 个字符。</span><span class="sxs-lookup"><span data-stu-id="f156e-p144">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`|<span data-ttu-id="f156e-784">字符串</span><span class="sxs-lookup"><span data-stu-id="f156e-784">String</span></span>||<span data-ttu-id="f156e-p145">在附件上载过程中显示的附件名称。最大长度为 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="f156e-p145">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="f156e-787">Object</span><span class="sxs-lookup"><span data-stu-id="f156e-787">Object</span></span>|<span data-ttu-id="f156e-788">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="f156e-788">&lt;optional&gt;</span></span>|<span data-ttu-id="f156e-789">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="f156e-789">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="f156e-790">对象</span><span class="sxs-lookup"><span data-stu-id="f156e-790">Object</span></span>|<span data-ttu-id="f156e-791">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="f156e-791">&lt;optional&gt;</span></span>|<span data-ttu-id="f156e-792">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="f156e-792">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.isInline`|<span data-ttu-id="f156e-793">布尔值</span><span class="sxs-lookup"><span data-stu-id="f156e-793">Boolean</span></span>|<span data-ttu-id="f156e-794">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="f156e-794">&lt;optional&gt;</span></span>|<span data-ttu-id="f156e-795">如果为 `true`，则表示附件将在邮件正文中内联显示，并且不应显示在附件列表中。</span><span class="sxs-lookup"><span data-stu-id="f156e-795">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`callback`|<span data-ttu-id="f156e-796">函数</span><span class="sxs-lookup"><span data-stu-id="f156e-796">function</span></span>|<span data-ttu-id="f156e-797">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="f156e-797">&lt;optional&gt;</span></span>|<span data-ttu-id="f156e-798">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="f156e-798">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="f156e-799">如果成功，附件标识符将在 `asyncResult.value` 属性中提供。</span><span class="sxs-lookup"><span data-stu-id="f156e-799">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="f156e-800">如果上传附件失败，`asyncResult` 对象将包含一个提供错误说明的 `Error` 对象。</span><span class="sxs-lookup"><span data-stu-id="f156e-800">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="f156e-801">错误</span><span class="sxs-lookup"><span data-stu-id="f156e-801">Errors</span></span>

|<span data-ttu-id="f156e-802">错误代码</span><span class="sxs-lookup"><span data-stu-id="f156e-802">Error code</span></span>|<span data-ttu-id="f156e-803">说明</span><span class="sxs-lookup"><span data-stu-id="f156e-803">Description</span></span>|
|------------|-------------|
|`AttachmentSizeExceeded`|<span data-ttu-id="f156e-804">附件大小超过了允许的大小。</span><span class="sxs-lookup"><span data-stu-id="f156e-804">The attachment is larger than allowed.</span></span>|
|`FileTypeNotSupported`|<span data-ttu-id="f156e-805">该附件的扩展名不是允许的扩展名。</span><span class="sxs-lookup"><span data-stu-id="f156e-805">The attachment has an extension that is not allowed.</span></span>|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="f156e-806">邮件或约会具有的附件过多。</span><span class="sxs-lookup"><span data-stu-id="f156e-806">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f156e-807">要求</span><span class="sxs-lookup"><span data-stu-id="f156e-807">Requirements</span></span>

|<span data-ttu-id="f156e-808">要求</span><span class="sxs-lookup"><span data-stu-id="f156e-808">Requirement</span></span>|<span data-ttu-id="f156e-809">值</span><span class="sxs-lookup"><span data-stu-id="f156e-809">Value</span></span>|
|---|---|
|[<span data-ttu-id="f156e-810">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="f156e-810">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f156e-811">1.1</span><span class="sxs-lookup"><span data-stu-id="f156e-811">1.1</span></span>|
|[<span data-ttu-id="f156e-812">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="f156e-812">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f156e-813">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="f156e-813">ReadWriteItem</span></span>|
|[<span data-ttu-id="f156e-814">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="f156e-814">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="f156e-815">撰写</span><span class="sxs-lookup"><span data-stu-id="f156e-815">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="f156e-816">示例</span><span class="sxs-lookup"><span data-stu-id="f156e-816">Examples</span></span>

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

<span data-ttu-id="f156e-817">以下示例将图像文件添加为内联附件，并在邮件正文中引用该附件。</span><span class="sxs-lookup"><span data-stu-id="f156e-817">The following example adds an image file as an inline attachment and references the attachment in the message body.</span></span>

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

#### <a name="addfileattachmentfrombase64asyncbase64file-attachmentname-options-callback"></a><span data-ttu-id="f156e-818">addFileAttachmentFromBase64Async （base64File，attachmentName，[options]，[callback]）</span><span class="sxs-lookup"><span data-stu-id="f156e-818">addFileAttachmentFromBase64Async(base64File, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="f156e-819">将 base64 编码中的文件作为附件添加到邮件或约会中。</span><span class="sxs-lookup"><span data-stu-id="f156e-819">Adds a file from the base64 encoding to a message or appointment as an attachment.</span></span>

<span data-ttu-id="f156e-820">该`addFileAttachmentFromBase64Async`方法从 base64 编码中上载文件，并将其附加到撰写窗体中的项目。</span><span class="sxs-lookup"><span data-stu-id="f156e-820">The `addFileAttachmentFromBase64Async` method uploads the file from the base64 encoding and attaches it to the item in the compose form.</span></span> <span data-ttu-id="f156e-821">此方法返回 AsyncResult 对象中的附件标识符。</span><span class="sxs-lookup"><span data-stu-id="f156e-821">This method returns the attachment identifier in the AsyncResult.value object.</span></span>

<span data-ttu-id="f156e-822">随后可以将该标识符与 [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) 方法一同使用，以删除同一个会话中的附件。</span><span class="sxs-lookup"><span data-stu-id="f156e-822">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="f156e-823">参数</span><span class="sxs-lookup"><span data-stu-id="f156e-823">Parameters</span></span>

|<span data-ttu-id="f156e-824">名称</span><span class="sxs-lookup"><span data-stu-id="f156e-824">Name</span></span>|<span data-ttu-id="f156e-825">类型</span><span class="sxs-lookup"><span data-stu-id="f156e-825">Type</span></span>|<span data-ttu-id="f156e-826">属性</span><span class="sxs-lookup"><span data-stu-id="f156e-826">Attributes</span></span>|<span data-ttu-id="f156e-827">说明</span><span class="sxs-lookup"><span data-stu-id="f156e-827">Description</span></span>|
|---|---|---|---|
|`base64File`|<span data-ttu-id="f156e-828">字符串</span><span class="sxs-lookup"><span data-stu-id="f156e-828">String</span></span>||<span data-ttu-id="f156e-829">要添加到电子邮件或事件的图像或文件的 base64 编码内容。</span><span class="sxs-lookup"><span data-stu-id="f156e-829">The base64 encoded content of an image or file to be added to an email or event.</span></span>|
|`attachmentName`|<span data-ttu-id="f156e-830">字符串</span><span class="sxs-lookup"><span data-stu-id="f156e-830">String</span></span>||<span data-ttu-id="f156e-p147">在附件上载过程中显示的附件名称。最大长度为 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="f156e-p147">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="f156e-833">对象</span><span class="sxs-lookup"><span data-stu-id="f156e-833">Object</span></span>|<span data-ttu-id="f156e-834">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="f156e-834">&lt;optional&gt;</span></span>|<span data-ttu-id="f156e-835">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="f156e-835">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="f156e-836">对象</span><span class="sxs-lookup"><span data-stu-id="f156e-836">Object</span></span>|<span data-ttu-id="f156e-837">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="f156e-837">&lt;optional&gt;</span></span>|<span data-ttu-id="f156e-838">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="f156e-838">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.isInline`|<span data-ttu-id="f156e-839">布尔值</span><span class="sxs-lookup"><span data-stu-id="f156e-839">Boolean</span></span>|<span data-ttu-id="f156e-840">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="f156e-840">&lt;optional&gt;</span></span>|<span data-ttu-id="f156e-841">如果为 `true`，则表示附件将在邮件正文中内联显示，并且不应显示在附件列表中。</span><span class="sxs-lookup"><span data-stu-id="f156e-841">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`callback`|<span data-ttu-id="f156e-842">函数</span><span class="sxs-lookup"><span data-stu-id="f156e-842">function</span></span>|<span data-ttu-id="f156e-843">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="f156e-843">&lt;optional&gt;</span></span>|<span data-ttu-id="f156e-844">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="f156e-844">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="f156e-845">如果成功，附件标识符将在 `asyncResult.value` 属性中提供。</span><span class="sxs-lookup"><span data-stu-id="f156e-845">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="f156e-846">如果上传附件失败，`asyncResult` 对象将包含一个提供错误说明的 `Error` 对象。</span><span class="sxs-lookup"><span data-stu-id="f156e-846">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="f156e-847">错误</span><span class="sxs-lookup"><span data-stu-id="f156e-847">Errors</span></span>

|<span data-ttu-id="f156e-848">错误代码</span><span class="sxs-lookup"><span data-stu-id="f156e-848">Error code</span></span>|<span data-ttu-id="f156e-849">说明</span><span class="sxs-lookup"><span data-stu-id="f156e-849">Description</span></span>|
|------------|-------------|
|`AttachmentSizeExceeded`|<span data-ttu-id="f156e-850">附件大小超过了允许的大小。</span><span class="sxs-lookup"><span data-stu-id="f156e-850">The attachment is larger than allowed.</span></span>|
|`FileTypeNotSupported`|<span data-ttu-id="f156e-851">该附件的扩展名不是允许的扩展名。</span><span class="sxs-lookup"><span data-stu-id="f156e-851">The attachment has an extension that is not allowed.</span></span>|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="f156e-852">邮件或约会具有的附件过多。</span><span class="sxs-lookup"><span data-stu-id="f156e-852">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f156e-853">要求</span><span class="sxs-lookup"><span data-stu-id="f156e-853">Requirements</span></span>

|<span data-ttu-id="f156e-854">要求</span><span class="sxs-lookup"><span data-stu-id="f156e-854">Requirement</span></span>|<span data-ttu-id="f156e-855">值</span><span class="sxs-lookup"><span data-stu-id="f156e-855">Value</span></span>|
|---|---|
|[<span data-ttu-id="f156e-856">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="f156e-856">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f156e-857">1.8</span><span class="sxs-lookup"><span data-stu-id="f156e-857">1.8</span></span>|
|[<span data-ttu-id="f156e-858">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="f156e-858">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f156e-859">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="f156e-859">ReadWriteItem</span></span>|
|[<span data-ttu-id="f156e-860">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="f156e-860">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="f156e-861">撰写</span><span class="sxs-lookup"><span data-stu-id="f156e-861">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="f156e-862">示例</span><span class="sxs-lookup"><span data-stu-id="f156e-862">Examples</span></span>

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

#### <a name="addhandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="f156e-863">addHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="f156e-863">addHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="f156e-864">添加支持事件的事件处理程序。</span><span class="sxs-lookup"><span data-stu-id="f156e-864">Adds an event handler for a supported event.</span></span>

<span data-ttu-id="f156e-865">目前，受支持的事件`Office.EventType.AttachmentsChanged`类型`Office.EventType.AppointmentTimeChanged`是`Office.EventType.EnhancedLocationsChanged`、 `Office.EventType.RecipientsChanged`、、 `Office.EventType.RecurrenceChanged`和。</span><span class="sxs-lookup"><span data-stu-id="f156e-865">Currently the supported event types are `Office.EventType.AttachmentsChanged`, `Office.EventType.AppointmentTimeChanged`, `Office.EventType.EnhancedLocationsChanged`, `Office.EventType.RecipientsChanged`, and `Office.EventType.RecurrenceChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="f156e-866">Parameters</span><span class="sxs-lookup"><span data-stu-id="f156e-866">Parameters</span></span>

| <span data-ttu-id="f156e-867">名称</span><span class="sxs-lookup"><span data-stu-id="f156e-867">Name</span></span> | <span data-ttu-id="f156e-868">类型</span><span class="sxs-lookup"><span data-stu-id="f156e-868">Type</span></span> | <span data-ttu-id="f156e-869">属性</span><span class="sxs-lookup"><span data-stu-id="f156e-869">Attributes</span></span> | <span data-ttu-id="f156e-870">说明</span><span class="sxs-lookup"><span data-stu-id="f156e-870">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="f156e-871">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="f156e-871">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="f156e-872">应调用处理程序的事件。</span><span class="sxs-lookup"><span data-stu-id="f156e-872">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="f156e-873">函数</span><span class="sxs-lookup"><span data-stu-id="f156e-873">Function</span></span> || <span data-ttu-id="f156e-p148">用于处理事件的函数。此函数必须接受一个参数，即对象文本。参数上的 `type` 属性将匹配传递给 `addHandlerAsync` 的 `eventType` 参数。</span><span class="sxs-lookup"><span data-stu-id="f156e-p148">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `addHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="f156e-877">Object</span><span class="sxs-lookup"><span data-stu-id="f156e-877">Object</span></span> | <span data-ttu-id="f156e-878">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="f156e-878">&lt;optional&gt;</span></span> | <span data-ttu-id="f156e-879">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="f156e-879">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="f156e-880">对象</span><span class="sxs-lookup"><span data-stu-id="f156e-880">Object</span></span> | <span data-ttu-id="f156e-881">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="f156e-881">&lt;optional&gt;</span></span> | <span data-ttu-id="f156e-882">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="f156e-882">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="f156e-883">函数</span><span class="sxs-lookup"><span data-stu-id="f156e-883">function</span></span>| <span data-ttu-id="f156e-884">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="f156e-884">&lt;optional&gt;</span></span>|<span data-ttu-id="f156e-885">方法完成后，使用单个参数 `callback`（一个 [`asyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `AsyncResult` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="f156e-885">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f156e-886">Requirements</span><span class="sxs-lookup"><span data-stu-id="f156e-886">Requirements</span></span>

|<span data-ttu-id="f156e-887">要求</span><span class="sxs-lookup"><span data-stu-id="f156e-887">Requirement</span></span>| <span data-ttu-id="f156e-888">值</span><span class="sxs-lookup"><span data-stu-id="f156e-888">Value</span></span>|
|---|---|
|[<span data-ttu-id="f156e-889">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="f156e-889">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f156e-890">1.7</span><span class="sxs-lookup"><span data-stu-id="f156e-890">1.7</span></span> |
|[<span data-ttu-id="f156e-891">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="f156e-891">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f156e-892">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f156e-892">ReadItem</span></span> |
|[<span data-ttu-id="f156e-893">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="f156e-893">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f156e-894">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="f156e-894">Compose or Read</span></span> |

##### <a name="example"></a><span data-ttu-id="f156e-895">示例</span><span class="sxs-lookup"><span data-stu-id="f156e-895">Example</span></span>

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

#### <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="f156e-896">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="f156e-896">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="f156e-897">将 Exchange 项目（如邮件）作为附件添加到邮件或约会。</span><span class="sxs-lookup"><span data-stu-id="f156e-897">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="f156e-p149">`addItemAttachmentAsync` 方法将包含指定 Exchange 标识符的项目附加到撰写窗体中的项目。如果指定一个回调方法，此方法使用 `asyncResult` 参数调用，该参数包含一个附件标识符或代码，指示附加项目过程中出现的任何错误。可以使用 `options` 参数将状态信息传递给回调方法（如果需要）。</span><span class="sxs-lookup"><span data-stu-id="f156e-p149">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="f156e-901">随后可以将该标识符与 [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) 方法一同使用，以删除同一个会话中的附件。</span><span class="sxs-lookup"><span data-stu-id="f156e-901">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="f156e-902">如果 Office 加载项是在 Outlook 网页版中运行，`addItemAttachmentAsync` 方法可以将项附加到除正在编辑的项外的项；但既不支持也不建议这样做。</span><span class="sxs-lookup"><span data-stu-id="f156e-902">If your Office Add-in is running in Outlook on the web, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="f156e-903">Parameters</span><span class="sxs-lookup"><span data-stu-id="f156e-903">Parameters</span></span>

|<span data-ttu-id="f156e-904">名称</span><span class="sxs-lookup"><span data-stu-id="f156e-904">Name</span></span>|<span data-ttu-id="f156e-905">类型</span><span class="sxs-lookup"><span data-stu-id="f156e-905">Type</span></span>|<span data-ttu-id="f156e-906">属性</span><span class="sxs-lookup"><span data-stu-id="f156e-906">Attributes</span></span>|<span data-ttu-id="f156e-907">说明</span><span class="sxs-lookup"><span data-stu-id="f156e-907">Description</span></span>|
|---|---|---|---|
|`itemId`|<span data-ttu-id="f156e-908">字符串</span><span class="sxs-lookup"><span data-stu-id="f156e-908">String</span></span>||<span data-ttu-id="f156e-p150">要附加的项目的 Exchange 标识符。最大长度为 100 个字符。</span><span class="sxs-lookup"><span data-stu-id="f156e-p150">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`|<span data-ttu-id="f156e-911">String</span><span class="sxs-lookup"><span data-stu-id="f156e-911">String</span></span>||<span data-ttu-id="f156e-912">要附加的项目的主题。</span><span class="sxs-lookup"><span data-stu-id="f156e-912">The subject of the item to be attached.</span></span> <span data-ttu-id="f156e-913">最大长度为 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="f156e-913">The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="f156e-914">对象</span><span class="sxs-lookup"><span data-stu-id="f156e-914">Object</span></span>|<span data-ttu-id="f156e-915">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="f156e-915">&lt;optional&gt;</span></span>|<span data-ttu-id="f156e-916">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="f156e-916">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="f156e-917">对象</span><span class="sxs-lookup"><span data-stu-id="f156e-917">Object</span></span>|<span data-ttu-id="f156e-918">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="f156e-918">&lt;optional&gt;</span></span>|<span data-ttu-id="f156e-919">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="f156e-919">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="f156e-920">函数</span><span class="sxs-lookup"><span data-stu-id="f156e-920">function</span></span>|<span data-ttu-id="f156e-921">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="f156e-921">&lt;optional&gt;</span></span>|<span data-ttu-id="f156e-922">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="f156e-922">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="f156e-923">如果成功，附件标识符将在 `asyncResult.value` 属性中提供。</span><span class="sxs-lookup"><span data-stu-id="f156e-923">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="f156e-924">如果添加附件失败，`asyncResult` 对象将包含一个提供错误说明的 `Error` 对象。</span><span class="sxs-lookup"><span data-stu-id="f156e-924">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="f156e-925">错误</span><span class="sxs-lookup"><span data-stu-id="f156e-925">Errors</span></span>

|<span data-ttu-id="f156e-926">错误代码</span><span class="sxs-lookup"><span data-stu-id="f156e-926">Error code</span></span>|<span data-ttu-id="f156e-927">说明</span><span class="sxs-lookup"><span data-stu-id="f156e-927">Description</span></span>|
|------------|-------------|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="f156e-928">邮件或约会具有的附件过多。</span><span class="sxs-lookup"><span data-stu-id="f156e-928">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f156e-929">Requirements</span><span class="sxs-lookup"><span data-stu-id="f156e-929">Requirements</span></span>

|<span data-ttu-id="f156e-930">要求</span><span class="sxs-lookup"><span data-stu-id="f156e-930">Requirement</span></span>|<span data-ttu-id="f156e-931">值</span><span class="sxs-lookup"><span data-stu-id="f156e-931">Value</span></span>|
|---|---|
|[<span data-ttu-id="f156e-932">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="f156e-932">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f156e-933">1.1</span><span class="sxs-lookup"><span data-stu-id="f156e-933">1.1</span></span>|
|[<span data-ttu-id="f156e-934">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="f156e-934">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f156e-935">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="f156e-935">ReadWriteItem</span></span>|
|[<span data-ttu-id="f156e-936">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="f156e-936">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="f156e-937">撰写</span><span class="sxs-lookup"><span data-stu-id="f156e-937">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="f156e-938">示例</span><span class="sxs-lookup"><span data-stu-id="f156e-938">Example</span></span>

<span data-ttu-id="f156e-939">以下示例将现有的 Outlook 项目添加为名为 `My Attachment` 的附件。</span><span class="sxs-lookup"><span data-stu-id="f156e-939">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

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

#### <a name="close"></a><span data-ttu-id="f156e-940">close()</span><span class="sxs-lookup"><span data-stu-id="f156e-940">close()</span></span>

<span data-ttu-id="f156e-941">关闭当前正在撰写的项目。</span><span class="sxs-lookup"><span data-stu-id="f156e-941">Closes the current item that is being composed.</span></span>

<span data-ttu-id="f156e-p152">
            \`close\` 方法的行为取决于要撰写的项目的当前状态。如果项目具有未保存的更改，客户端将提示用户保存、放弃或取消关闭操作。</span><span class="sxs-lookup"><span data-stu-id="f156e-p152">The behavior of the `close` method depends on the current state of the item being composed. If the item has unsaved changes, the client prompts the user to save, discard, or cancel the close action.</span></span>

> [!NOTE]
> <span data-ttu-id="f156e-944">在 Outlook 网页版中，如果该项目是约会并且之前已使用 `saveAsync` 保存，则即使自上次保存项目后未发生任何更改，也会提示用户保存、放弃或取消。</span><span class="sxs-lookup"><span data-stu-id="f156e-944">In Outlook on the web, if the item is an appointment and it has previously been saved using `saveAsync`, the user is prompted to save, discard, or cancel even if no changes have occurred since the item was last saved.</span></span>

<span data-ttu-id="f156e-945">在 Outlook 桌面客户端中，如果邮件是内联答复，`close` 方法不起作用。</span><span class="sxs-lookup"><span data-stu-id="f156e-945">In the Outlook desktop client, if the message is an inline reply, the `close` method has no effect.</span></span>

##### <a name="requirements"></a><span data-ttu-id="f156e-946">Requirements</span><span class="sxs-lookup"><span data-stu-id="f156e-946">Requirements</span></span>

|<span data-ttu-id="f156e-947">要求</span><span class="sxs-lookup"><span data-stu-id="f156e-947">Requirement</span></span>|<span data-ttu-id="f156e-948">值</span><span class="sxs-lookup"><span data-stu-id="f156e-948">Value</span></span>|
|---|---|
|[<span data-ttu-id="f156e-949">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="f156e-949">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f156e-950">1.3</span><span class="sxs-lookup"><span data-stu-id="f156e-950">1.3</span></span>|
|[<span data-ttu-id="f156e-951">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="f156e-951">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f156e-952">受限</span><span class="sxs-lookup"><span data-stu-id="f156e-952">Restricted</span></span>|
|[<span data-ttu-id="f156e-953">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="f156e-953">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="f156e-954">撰写</span><span class="sxs-lookup"><span data-stu-id="f156e-954">Compose</span></span>|

<br>

---
---

#### <a name="displayreplyallformformdata-callback"></a><span data-ttu-id="f156e-955">displayReplyAllForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="f156e-955">displayReplyAllForm(formData, [callback])</span></span>

<span data-ttu-id="f156e-956">显示答复窗体，其中包括所选邮件的发件人和所有收件人或所选约会的组织者和所有与会者。</span><span class="sxs-lookup"><span data-stu-id="f156e-956">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="f156e-957">iOS 版 Outlook 或 Android 版 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="f156e-957">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="f156e-958">在 Outlook 网页版中，答复窗体显示为包含 3 列视图的弹出式窗体，以及包含 2 列或 1 列视图的弹出式窗体。</span><span class="sxs-lookup"><span data-stu-id="f156e-958">In Outlook on the web, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="f156e-959">如果任意字符串参数超出其限制，`displayReplyAllForm` 将引发异常。</span><span class="sxs-lookup"><span data-stu-id="f156e-959">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="f156e-p153">如果附件已在 `formData.attachments` 参数中指定，Outlook 网页版和 Outlook 桌面版客户端会尝试下载所有附件，并将它们附加到答复窗体。如果无法添加任何附件，窗体 UI 中会显示错误。如果此操作是不可能完成的，系统不会抛出任何错误消息。</span><span class="sxs-lookup"><span data-stu-id="f156e-p153">When attachments are specified in the `formData.attachments` parameter, Outlook on the web and desktop clients attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="f156e-963">Parameters</span><span class="sxs-lookup"><span data-stu-id="f156e-963">Parameters</span></span>

|<span data-ttu-id="f156e-964">名称</span><span class="sxs-lookup"><span data-stu-id="f156e-964">Name</span></span>|<span data-ttu-id="f156e-965">类型</span><span class="sxs-lookup"><span data-stu-id="f156e-965">Type</span></span>|<span data-ttu-id="f156e-966">属性</span><span class="sxs-lookup"><span data-stu-id="f156e-966">Attributes</span></span>|<span data-ttu-id="f156e-967">说明</span><span class="sxs-lookup"><span data-stu-id="f156e-967">Description</span></span>|
|---|---|---|---|
|`formData`|<span data-ttu-id="f156e-968">字符串 &#124; 对象</span><span class="sxs-lookup"><span data-stu-id="f156e-968">String &#124; Object</span></span>||<span data-ttu-id="f156e-p154">一个包含文本和 HTML 且表示答复窗体的正文的字符串。字符串限制为 32 KB。</span><span class="sxs-lookup"><span data-stu-id="f156e-p154">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="f156e-971">**或**</span><span class="sxs-lookup"><span data-stu-id="f156e-971">**OR**</span></span><br/><span data-ttu-id="f156e-p155">包含正文或附件数据和回调函数的对象。对象定义如下。</span><span class="sxs-lookup"><span data-stu-id="f156e-p155">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span>|
|`formData.htmlBody`|<span data-ttu-id="f156e-974">字符串</span><span class="sxs-lookup"><span data-stu-id="f156e-974">String</span></span>|<span data-ttu-id="f156e-975">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="f156e-975">&lt;optional&gt;</span></span>|<span data-ttu-id="f156e-p156">一个包含文本和 HTML 且表示答复窗体的正文的字符串。字符串限制为 32 KB。</span><span class="sxs-lookup"><span data-stu-id="f156e-p156">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
|`formData.attachments`|<span data-ttu-id="f156e-978">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="f156e-978">Array.&lt;Object&gt;</span></span>|<span data-ttu-id="f156e-979">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="f156e-979">&lt;optional&gt;</span></span>|<span data-ttu-id="f156e-980">JSON 对象（是文件或项目附件）的数组。</span><span class="sxs-lookup"><span data-stu-id="f156e-980">An array of JSON objects that are either file or item attachments.</span></span>|
|`formData.attachments.type`|<span data-ttu-id="f156e-981">字符串</span><span class="sxs-lookup"><span data-stu-id="f156e-981">String</span></span>||<span data-ttu-id="f156e-p157">指示附件的类型。必须是文件附件的 `file` 或项目附件的 `item`。</span><span class="sxs-lookup"><span data-stu-id="f156e-p157">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span>|
|`formData.attachments.name`|<span data-ttu-id="f156e-984">字符串</span><span class="sxs-lookup"><span data-stu-id="f156e-984">String</span></span>||<span data-ttu-id="f156e-985">一个包含附件的名称的字符串，最多包含 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="f156e-985">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
|`formData.attachments.url`|<span data-ttu-id="f156e-986">字符串</span><span class="sxs-lookup"><span data-stu-id="f156e-986">String</span></span>||<span data-ttu-id="f156e-p158">仅在将 `type` 设置为 `file` 时使用。文件的位置的 URI。</span><span class="sxs-lookup"><span data-stu-id="f156e-p158">Only used if `type` is set to `file`. The URI of the location for the file.</span></span>|
|`formData.attachments.isInline`|<span data-ttu-id="f156e-989">布尔</span><span class="sxs-lookup"><span data-stu-id="f156e-989">Boolean</span></span>||<span data-ttu-id="f156e-p159">仅在将 `type` 设置为 `file` 时使用。如果为 `true`，则表示附件将在邮件正文中内联显示，并且不应显示在附件列表中。</span><span class="sxs-lookup"><span data-stu-id="f156e-p159">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`formData.attachments.itemId`|<span data-ttu-id="f156e-992">String</span><span class="sxs-lookup"><span data-stu-id="f156e-992">String</span></span>||<span data-ttu-id="f156e-p160">仅在将 `type` 设置为 `item` 时使用。附件的 EWS 项目 ID。字符串最长为 100 个字符。</span><span class="sxs-lookup"><span data-stu-id="f156e-p160">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span>|
|`callback`|<span data-ttu-id="f156e-996">函数</span><span class="sxs-lookup"><span data-stu-id="f156e-996">function</span></span>|<span data-ttu-id="f156e-997">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="f156e-997">&lt;optional&gt;</span></span>|<span data-ttu-id="f156e-998">方法完成后，使用单个参数 `asyncResult`（一个 [AsyncResult](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="f156e-998">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f156e-999">要求</span><span class="sxs-lookup"><span data-stu-id="f156e-999">Requirements</span></span>

|<span data-ttu-id="f156e-1000">要求</span><span class="sxs-lookup"><span data-stu-id="f156e-1000">Requirement</span></span>|<span data-ttu-id="f156e-1001">值</span><span class="sxs-lookup"><span data-stu-id="f156e-1001">Value</span></span>|
|---|---|
|[<span data-ttu-id="f156e-1002">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="f156e-1002">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f156e-1003">1.0</span><span class="sxs-lookup"><span data-stu-id="f156e-1003">1.0</span></span>|
|[<span data-ttu-id="f156e-1004">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="f156e-1004">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f156e-1005">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f156e-1005">ReadItem</span></span>|
|[<span data-ttu-id="f156e-1006">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="f156e-1006">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="f156e-1007">阅读</span><span class="sxs-lookup"><span data-stu-id="f156e-1007">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="f156e-1008">示例</span><span class="sxs-lookup"><span data-stu-id="f156e-1008">Examples</span></span>

<span data-ttu-id="f156e-1009">以下代码将一个字符串传递到 `displayReplyAllForm` 函数。</span><span class="sxs-lookup"><span data-stu-id="f156e-1009">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="f156e-1010">使用空白正文答复。</span><span class="sxs-lookup"><span data-stu-id="f156e-1010">Reply with an empty body.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="f156e-1011">仅使用正文答复。</span><span class="sxs-lookup"><span data-stu-id="f156e-1011">Reply with just a body.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="f156e-1012">使用正文和文件附件答复。</span><span class="sxs-lookup"><span data-stu-id="f156e-1012">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="f156e-1013">使用正文和项目附件答复。</span><span class="sxs-lookup"><span data-stu-id="f156e-1013">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="f156e-1014">使用正文、文件附件、项目附件和回调答复。</span><span class="sxs-lookup"><span data-stu-id="f156e-1014">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="displayreplyformformdata-callback"></a><span data-ttu-id="f156e-1015">displayReplyForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="f156e-1015">displayReplyForm(formData, [callback])</span></span>

<span data-ttu-id="f156e-1016">显示答复窗体，其中仅包括所选邮件的发件人或所选约会的组织者。</span><span class="sxs-lookup"><span data-stu-id="f156e-1016">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="f156e-1017">iOS 版 Outlook 或 Android 版 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="f156e-1017">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="f156e-1018">在 Outlook 网页版中，答复窗体显示为包含 3 列视图的弹出式窗体，以及包含 2 列或 1 列视图的弹出式窗体。</span><span class="sxs-lookup"><span data-stu-id="f156e-1018">In Outlook on the web, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="f156e-1019">如果任意字符串参数超出其限制，`displayReplyForm` 将引发异常。</span><span class="sxs-lookup"><span data-stu-id="f156e-1019">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="f156e-p161">如果附件已在 `formData.attachments` 参数中指定，Outlook 网页版和 Outlook 桌面版客户端会尝试下载所有附件，并将它们附加到答复窗体。如果无法添加任何附件，窗体 UI 中会显示错误。如果此操作是不可能完成的，系统不会抛出任何错误消息。</span><span class="sxs-lookup"><span data-stu-id="f156e-p161">When attachments are specified in the `formData.attachments` parameter, Outlook on the web and desktop clients attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="f156e-1023">Parameters</span><span class="sxs-lookup"><span data-stu-id="f156e-1023">Parameters</span></span>

|<span data-ttu-id="f156e-1024">名称</span><span class="sxs-lookup"><span data-stu-id="f156e-1024">Name</span></span>|<span data-ttu-id="f156e-1025">类型</span><span class="sxs-lookup"><span data-stu-id="f156e-1025">Type</span></span>|<span data-ttu-id="f156e-1026">属性</span><span class="sxs-lookup"><span data-stu-id="f156e-1026">Attributes</span></span>|<span data-ttu-id="f156e-1027">说明</span><span class="sxs-lookup"><span data-stu-id="f156e-1027">Description</span></span>|
|---|---|---|---|
|`formData`|<span data-ttu-id="f156e-1028">字符串 &#124; 对象</span><span class="sxs-lookup"><span data-stu-id="f156e-1028">String &#124; Object</span></span>||<span data-ttu-id="f156e-p162">一个包含文本和 HTML 且表示答复窗体的正文的字符串。字符串限制为 32 KB。</span><span class="sxs-lookup"><span data-stu-id="f156e-p162">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="f156e-1031">**或**</span><span class="sxs-lookup"><span data-stu-id="f156e-1031">**OR**</span></span><br/><span data-ttu-id="f156e-p163">包含正文或附件数据和回调函数的对象。对象定义如下。</span><span class="sxs-lookup"><span data-stu-id="f156e-p163">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span>|
|`formData.htmlBody`|<span data-ttu-id="f156e-1034">字符串</span><span class="sxs-lookup"><span data-stu-id="f156e-1034">String</span></span>|<span data-ttu-id="f156e-1035">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="f156e-1035">&lt;optional&gt;</span></span>|<span data-ttu-id="f156e-p164">一个包含文本和 HTML 且表示答复窗体的正文的字符串。字符串限制为 32 KB。</span><span class="sxs-lookup"><span data-stu-id="f156e-p164">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
|`formData.attachments`|<span data-ttu-id="f156e-1038">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="f156e-1038">Array.&lt;Object&gt;</span></span>|<span data-ttu-id="f156e-1039">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="f156e-1039">&lt;optional&gt;</span></span>|<span data-ttu-id="f156e-1040">JSON 对象（是文件或项目附件）的数组。</span><span class="sxs-lookup"><span data-stu-id="f156e-1040">An array of JSON objects that are either file or item attachments.</span></span>|
|`formData.attachments.type`|<span data-ttu-id="f156e-1041">字符串</span><span class="sxs-lookup"><span data-stu-id="f156e-1041">String</span></span>||<span data-ttu-id="f156e-p165">指示附件的类型。必须是文件附件的 `file` 或项目附件的 `item`。</span><span class="sxs-lookup"><span data-stu-id="f156e-p165">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span>|
|`formData.attachments.name`|<span data-ttu-id="f156e-1044">字符串</span><span class="sxs-lookup"><span data-stu-id="f156e-1044">String</span></span>||<span data-ttu-id="f156e-1045">一个包含附件的名称的字符串，最多包含 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="f156e-1045">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
|`formData.attachments.url`|<span data-ttu-id="f156e-1046">字符串</span><span class="sxs-lookup"><span data-stu-id="f156e-1046">String</span></span>||<span data-ttu-id="f156e-p166">仅在将 `type` 设置为 `file` 时使用。文件的位置的 URI。</span><span class="sxs-lookup"><span data-stu-id="f156e-p166">Only used if `type` is set to `file`. The URI of the location for the file.</span></span>|
|`formData.attachments.isInline`|<span data-ttu-id="f156e-1049">布尔</span><span class="sxs-lookup"><span data-stu-id="f156e-1049">Boolean</span></span>||<span data-ttu-id="f156e-p167">仅在将 `type` 设置为 `file` 时使用。如果为 `true`，则表示附件将在邮件正文中内联显示，并且不应显示在附件列表中。</span><span class="sxs-lookup"><span data-stu-id="f156e-p167">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`formData.attachments.itemId`|<span data-ttu-id="f156e-1052">String</span><span class="sxs-lookup"><span data-stu-id="f156e-1052">String</span></span>||<span data-ttu-id="f156e-p168">仅在将 `type` 设置为 `item` 时使用。附件的 EWS 项目 ID。字符串最长为 100 个字符。</span><span class="sxs-lookup"><span data-stu-id="f156e-p168">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span>|
|`callback`|<span data-ttu-id="f156e-1056">函数</span><span class="sxs-lookup"><span data-stu-id="f156e-1056">function</span></span>|<span data-ttu-id="f156e-1057">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="f156e-1057">&lt;optional&gt;</span></span>|<span data-ttu-id="f156e-1058">方法完成后，使用单个参数 `asyncResult`（一个 [AsyncResult](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="f156e-1058">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f156e-1059">要求</span><span class="sxs-lookup"><span data-stu-id="f156e-1059">Requirements</span></span>

|<span data-ttu-id="f156e-1060">要求</span><span class="sxs-lookup"><span data-stu-id="f156e-1060">Requirement</span></span>|<span data-ttu-id="f156e-1061">值</span><span class="sxs-lookup"><span data-stu-id="f156e-1061">Value</span></span>|
|---|---|
|[<span data-ttu-id="f156e-1062">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="f156e-1062">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f156e-1063">1.0</span><span class="sxs-lookup"><span data-stu-id="f156e-1063">1.0</span></span>|
|[<span data-ttu-id="f156e-1064">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="f156e-1064">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f156e-1065">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f156e-1065">ReadItem</span></span>|
|[<span data-ttu-id="f156e-1066">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="f156e-1066">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="f156e-1067">阅读</span><span class="sxs-lookup"><span data-stu-id="f156e-1067">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="f156e-1068">示例</span><span class="sxs-lookup"><span data-stu-id="f156e-1068">Examples</span></span>

<span data-ttu-id="f156e-1069">以下代码将一个字符串传递到 `displayReplyForm` 函数。</span><span class="sxs-lookup"><span data-stu-id="f156e-1069">The following code passes a string to the `displayReplyForm` function.</span></span>

```js
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="f156e-1070">使用空白正文答复。</span><span class="sxs-lookup"><span data-stu-id="f156e-1070">Reply with an empty body.</span></span>

```js
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="f156e-1071">仅使用正文答复。</span><span class="sxs-lookup"><span data-stu-id="f156e-1071">Reply with just a body.</span></span>

```js
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="f156e-1072">使用正文和文件附件答复。</span><span class="sxs-lookup"><span data-stu-id="f156e-1072">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="f156e-1073">使用正文和项目附件答复。</span><span class="sxs-lookup"><span data-stu-id="f156e-1073">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="f156e-1074">使用正文、文件附件、项目附件和回调答复。</span><span class="sxs-lookup"><span data-stu-id="f156e-1074">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="getallinternetheadersasyncoptions-callback"></a><span data-ttu-id="f156e-1075">getAllInternetHeadersAsync （[options]，[callback]）</span><span class="sxs-lookup"><span data-stu-id="f156e-1075">getAllInternetHeadersAsync([options], [callback])</span></span>

<span data-ttu-id="f156e-1076">以字符串形式获取邮件的所有 internet 邮件头。</span><span class="sxs-lookup"><span data-stu-id="f156e-1076">Gets all the internet headers for the message as a string.</span></span> <span data-ttu-id="f156e-1077">仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="f156e-1077">Read mode only.</span></span>

##### <a name="parameters"></a><span data-ttu-id="f156e-1078">Parameters</span><span class="sxs-lookup"><span data-stu-id="f156e-1078">Parameters</span></span>

|<span data-ttu-id="f156e-1079">名称</span><span class="sxs-lookup"><span data-stu-id="f156e-1079">Name</span></span>|<span data-ttu-id="f156e-1080">类型</span><span class="sxs-lookup"><span data-stu-id="f156e-1080">Type</span></span>|<span data-ttu-id="f156e-1081">属性</span><span class="sxs-lookup"><span data-stu-id="f156e-1081">Attributes</span></span>|<span data-ttu-id="f156e-1082">说明</span><span class="sxs-lookup"><span data-stu-id="f156e-1082">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="f156e-1083">对象</span><span class="sxs-lookup"><span data-stu-id="f156e-1083">Object</span></span>|<span data-ttu-id="f156e-1084">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="f156e-1084">&lt;optional&gt;</span></span>|<span data-ttu-id="f156e-1085">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="f156e-1085">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="f156e-1086">对象</span><span class="sxs-lookup"><span data-stu-id="f156e-1086">Object</span></span>|<span data-ttu-id="f156e-1087">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="f156e-1087">&lt;optional&gt;</span></span>|<span data-ttu-id="f156e-1088">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="f156e-1088">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="f156e-1089">函数</span><span class="sxs-lookup"><span data-stu-id="f156e-1089">function</span></span>|<span data-ttu-id="f156e-1090">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="f156e-1090">&lt;optional&gt;</span></span>|<span data-ttu-id="f156e-1091">方法完成后，使用单个参数 `asyncResult`（一个 [AsyncResult](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="f156e-1091">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> <span data-ttu-id="f156e-1092">在成功的情况下，internet 标头数据在 asyncResult 属性中以字符串的形式提供。</span><span class="sxs-lookup"><span data-stu-id="f156e-1092">On success, the internet headers data is provided in the asyncResult.value property as a string.</span></span> <span data-ttu-id="f156e-1093">有关返回的字符串值的格式设置信息，请参阅[RFC 2183](https://tools.ietf.org/html/rfc2183) 。</span><span class="sxs-lookup"><span data-stu-id="f156e-1093">Refer to [RFC 2183](https://tools.ietf.org/html/rfc2183) for the formatting information of the returned string value.</span></span> <span data-ttu-id="f156e-1094">如果调用失败，asyncResult 属性将包含错误代码和失败原因。</span><span class="sxs-lookup"><span data-stu-id="f156e-1094">If the call fails, the asyncResult.error property will contain an error code with the reason for the failure.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f156e-1095">Requirements</span><span class="sxs-lookup"><span data-stu-id="f156e-1095">Requirements</span></span>

|<span data-ttu-id="f156e-1096">要求</span><span class="sxs-lookup"><span data-stu-id="f156e-1096">Requirement</span></span>|<span data-ttu-id="f156e-1097">值</span><span class="sxs-lookup"><span data-stu-id="f156e-1097">Value</span></span>|
|---|---|
|[<span data-ttu-id="f156e-1098">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="f156e-1098">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f156e-1099">1.8</span><span class="sxs-lookup"><span data-stu-id="f156e-1099">1.8</span></span>|
|[<span data-ttu-id="f156e-1100">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="f156e-1100">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f156e-1101">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f156e-1101">ReadItem</span></span>|
|[<span data-ttu-id="f156e-1102">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="f156e-1102">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="f156e-1103">阅读</span><span class="sxs-lookup"><span data-stu-id="f156e-1103">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="f156e-1104">返回：</span><span class="sxs-lookup"><span data-stu-id="f156e-1104">Returns:</span></span>

<span data-ttu-id="f156e-1105">作为字符串的 internet 标头数据，根据[RFC 2183](https://tools.ietf.org/html/rfc2183)格式化。</span><span class="sxs-lookup"><span data-stu-id="f156e-1105">The internet headers data as a string formatted according to [RFC 2183](https://tools.ietf.org/html/rfc2183).</span></span>

<span data-ttu-id="f156e-1106">类型：字符串</span><span class="sxs-lookup"><span data-stu-id="f156e-1106">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="f156e-1107">示例</span><span class="sxs-lookup"><span data-stu-id="f156e-1107">Example</span></span>

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

#### <a name="getattachmentcontentasyncattachmentid-options-callback--attachmentcontentjavascriptapioutlookofficeattachmentcontent"></a><span data-ttu-id="f156e-1108">getAttachmentContentAsync （attachmentId，[options]，[callback]）→ [AttachmentContent](/javascript/api/outlook/office.attachmentcontent)</span><span class="sxs-lookup"><span data-stu-id="f156e-1108">getAttachmentContentAsync(attachmentId, [options], [callback]) → [AttachmentContent](/javascript/api/outlook/office.attachmentcontent)</span></span>

<span data-ttu-id="f156e-1109">从邮件或约会中获取指定附件并将其作为`AttachmentContent`对象返回。</span><span class="sxs-lookup"><span data-stu-id="f156e-1109">Gets the specified attachment from a message or appointment and returns it as an `AttachmentContent` object.</span></span>

<span data-ttu-id="f156e-1110">该`getAttachmentContentAsync`方法从项目中获取具有指定标识符的附件。</span><span class="sxs-lookup"><span data-stu-id="f156e-1110">The `getAttachmentContentAsync` method gets the attachment with the specified identifier from the item.</span></span> <span data-ttu-id="f156e-1111">作为一种最佳做法，您应使用标识符在与`getAttachmentsAsync` or `item.attachments`调用一起检索到会话的同一会话中检索附件。</span><span class="sxs-lookup"><span data-stu-id="f156e-1111">As a best practice, you should use the identifier to retrieve an attachment in the same session that the attachmentIds were retrieved with the `getAttachmentsAsync` or `item.attachments` call.</span></span> <span data-ttu-id="f156e-1112">在 Outlook 网页版和移动设备上，附件标识符只在同一个会话中才有效。</span><span class="sxs-lookup"><span data-stu-id="f156e-1112">In Outlook on the web and mobile devices, the attachment identifier is valid only within the same session.</span></span> <span data-ttu-id="f156e-1113">当用户关闭应用程序时，或者如果用户开始撰写内嵌窗体，随后弹出窗体以继续在单独的窗口中。</span><span class="sxs-lookup"><span data-stu-id="f156e-1113">A session is over when the user closes the app, or if the user starts composing an inline form then subsequently pops out the form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="f156e-1114">Parameters</span><span class="sxs-lookup"><span data-stu-id="f156e-1114">Parameters</span></span>

|<span data-ttu-id="f156e-1115">名称</span><span class="sxs-lookup"><span data-stu-id="f156e-1115">Name</span></span>|<span data-ttu-id="f156e-1116">类型</span><span class="sxs-lookup"><span data-stu-id="f156e-1116">Type</span></span>|<span data-ttu-id="f156e-1117">属性</span><span class="sxs-lookup"><span data-stu-id="f156e-1117">Attributes</span></span>|<span data-ttu-id="f156e-1118">说明</span><span class="sxs-lookup"><span data-stu-id="f156e-1118">Description</span></span>|
|---|---|---|---|
|`attachmentId`|<span data-ttu-id="f156e-1119">字符串</span><span class="sxs-lookup"><span data-stu-id="f156e-1119">String</span></span>||<span data-ttu-id="f156e-1120">要获取的附件的标识符。</span><span class="sxs-lookup"><span data-stu-id="f156e-1120">The identifier of the attachment you want to get.</span></span>|
|`options`|<span data-ttu-id="f156e-1121">对象</span><span class="sxs-lookup"><span data-stu-id="f156e-1121">Object</span></span>|<span data-ttu-id="f156e-1122">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="f156e-1122">&lt;optional&gt;</span></span>|<span data-ttu-id="f156e-1123">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="f156e-1123">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="f156e-1124">对象</span><span class="sxs-lookup"><span data-stu-id="f156e-1124">Object</span></span>|<span data-ttu-id="f156e-1125">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="f156e-1125">&lt;optional&gt;</span></span>|<span data-ttu-id="f156e-1126">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="f156e-1126">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="f156e-1127">函数</span><span class="sxs-lookup"><span data-stu-id="f156e-1127">function</span></span>|<span data-ttu-id="f156e-1128">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="f156e-1128">&lt;optional&gt;</span></span>|<span data-ttu-id="f156e-1129">方法完成后，使用单个参数 `asyncResult`（一个 [AsyncResult](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="f156e-1129">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f156e-1130">要求</span><span class="sxs-lookup"><span data-stu-id="f156e-1130">Requirements</span></span>

|<span data-ttu-id="f156e-1131">要求</span><span class="sxs-lookup"><span data-stu-id="f156e-1131">Requirement</span></span>|<span data-ttu-id="f156e-1132">值</span><span class="sxs-lookup"><span data-stu-id="f156e-1132">Value</span></span>|
|---|---|
|[<span data-ttu-id="f156e-1133">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="f156e-1133">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f156e-1134">1.8</span><span class="sxs-lookup"><span data-stu-id="f156e-1134">1.8</span></span>|
|[<span data-ttu-id="f156e-1135">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="f156e-1135">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f156e-1136">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f156e-1136">ReadItem</span></span>|
|[<span data-ttu-id="f156e-1137">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="f156e-1137">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="f156e-1138">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="f156e-1138">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="f156e-1139">返回：</span><span class="sxs-lookup"><span data-stu-id="f156e-1139">Returns:</span></span>

<span data-ttu-id="f156e-1140">类型： [AttachmentContent](/javascript/api/outlook/office.attachmentcontent)</span><span class="sxs-lookup"><span data-stu-id="f156e-1140">Type: [AttachmentContent](/javascript/api/outlook/office.attachmentcontent)</span></span>

##### <a name="example"></a><span data-ttu-id="f156e-1141">示例</span><span class="sxs-lookup"><span data-stu-id="f156e-1141">Example</span></span>

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

#### <a name="getattachmentsasyncoptions-callback--arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetails"></a><span data-ttu-id="f156e-1142">getAttachmentsAsync （[options]，[callback]）→ Array. <[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="f156e-1142">getAttachmentsAsync([options], [callback]) → Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span></span>

<span data-ttu-id="f156e-1143">以数组的形式获取项目的附件。</span><span class="sxs-lookup"><span data-stu-id="f156e-1143">Gets the item's attachments as an array.</span></span> <span data-ttu-id="f156e-1144">仅限撰写模式。</span><span class="sxs-lookup"><span data-stu-id="f156e-1144">Compose mode only.</span></span>

##### <a name="parameters"></a><span data-ttu-id="f156e-1145">Parameters</span><span class="sxs-lookup"><span data-stu-id="f156e-1145">Parameters</span></span>

|<span data-ttu-id="f156e-1146">名称</span><span class="sxs-lookup"><span data-stu-id="f156e-1146">Name</span></span>|<span data-ttu-id="f156e-1147">类型</span><span class="sxs-lookup"><span data-stu-id="f156e-1147">Type</span></span>|<span data-ttu-id="f156e-1148">属性</span><span class="sxs-lookup"><span data-stu-id="f156e-1148">Attributes</span></span>|<span data-ttu-id="f156e-1149">说明</span><span class="sxs-lookup"><span data-stu-id="f156e-1149">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="f156e-1150">对象</span><span class="sxs-lookup"><span data-stu-id="f156e-1150">Object</span></span>|<span data-ttu-id="f156e-1151">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="f156e-1151">&lt;optional&gt;</span></span>|<span data-ttu-id="f156e-1152">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="f156e-1152">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="f156e-1153">对象</span><span class="sxs-lookup"><span data-stu-id="f156e-1153">Object</span></span>|<span data-ttu-id="f156e-1154">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="f156e-1154">&lt;optional&gt;</span></span>|<span data-ttu-id="f156e-1155">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="f156e-1155">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="f156e-1156">函数</span><span class="sxs-lookup"><span data-stu-id="f156e-1156">function</span></span>|<span data-ttu-id="f156e-1157">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="f156e-1157">&lt;optional&gt;</span></span>|<span data-ttu-id="f156e-1158">方法完成后，使用单个参数 `asyncResult`（一个 [AsyncResult](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="f156e-1158">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f156e-1159">Requirements</span><span class="sxs-lookup"><span data-stu-id="f156e-1159">Requirements</span></span>

|<span data-ttu-id="f156e-1160">要求</span><span class="sxs-lookup"><span data-stu-id="f156e-1160">Requirement</span></span>|<span data-ttu-id="f156e-1161">值</span><span class="sxs-lookup"><span data-stu-id="f156e-1161">Value</span></span>|
|---|---|
|[<span data-ttu-id="f156e-1162">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="f156e-1162">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f156e-1163">1.8</span><span class="sxs-lookup"><span data-stu-id="f156e-1163">1.8</span></span>|
|[<span data-ttu-id="f156e-1164">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="f156e-1164">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f156e-1165">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f156e-1165">ReadItem</span></span>|
|[<span data-ttu-id="f156e-1166">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="f156e-1166">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="f156e-1167">撰写</span><span class="sxs-lookup"><span data-stu-id="f156e-1167">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="f156e-1168">返回：</span><span class="sxs-lookup"><span data-stu-id="f156e-1168">Returns:</span></span>

<span data-ttu-id="f156e-1169">类型： Array. <[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="f156e-1169">Type: Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span></span>

##### <a name="example"></a><span data-ttu-id="f156e-1170">示例</span><span class="sxs-lookup"><span data-stu-id="f156e-1170">Example</span></span>

<span data-ttu-id="f156e-1171">下面的示例将生成一个 HTML 字符串，其中包含当前项目上所有附件的详细信息。</span><span class="sxs-lookup"><span data-stu-id="f156e-1171">The following example builds an HTML string with details of all attachments on the current item.</span></span>

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

#### <a name="getentities--entitiesjavascriptapioutlookofficeentities"></a><span data-ttu-id="f156e-1172">getEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="f156e-1172">getEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span></span>

<span data-ttu-id="f156e-1173">获取在所选项目的正文中找到的实体。</span><span class="sxs-lookup"><span data-stu-id="f156e-1173">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="f156e-1174">iOS 版 Outlook 或 Android 版 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="f156e-1174">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="f156e-1175">要求</span><span class="sxs-lookup"><span data-stu-id="f156e-1175">Requirements</span></span>

|<span data-ttu-id="f156e-1176">要求</span><span class="sxs-lookup"><span data-stu-id="f156e-1176">Requirement</span></span>|<span data-ttu-id="f156e-1177">值</span><span class="sxs-lookup"><span data-stu-id="f156e-1177">Value</span></span>|
|---|---|
|[<span data-ttu-id="f156e-1178">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="f156e-1178">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f156e-1179">1.0</span><span class="sxs-lookup"><span data-stu-id="f156e-1179">1.0</span></span>|
|[<span data-ttu-id="f156e-1180">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="f156e-1180">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f156e-1181">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f156e-1181">ReadItem</span></span>|
|[<span data-ttu-id="f156e-1182">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="f156e-1182">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="f156e-1183">阅读</span><span class="sxs-lookup"><span data-stu-id="f156e-1183">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="f156e-1184">返回：</span><span class="sxs-lookup"><span data-stu-id="f156e-1184">Returns:</span></span>

<span data-ttu-id="f156e-1185">类型：[Entities](/javascript/api/outlook/office.entities)</span><span class="sxs-lookup"><span data-stu-id="f156e-1185">Type: [Entities](/javascript/api/outlook/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="f156e-1186">示例</span><span class="sxs-lookup"><span data-stu-id="f156e-1186">Example</span></span>

<span data-ttu-id="f156e-1187">以下示例访问当前项目的正文中的联系人实体。</span><span class="sxs-lookup"><span data-stu-id="f156e-1187">The following example accesses the contacts entities in the current item's body.</span></span>

```js
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

<br>

---
---

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlookofficecontactmeetingsuggestionjavascriptapioutlookofficemeetingsuggestionphonenumberjavascriptapioutlookofficephonenumbertasksuggestionjavascriptapioutlookofficetasksuggestion"></a><span data-ttu-id="f156e-1188">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="f156e-1188">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span></span>

<span data-ttu-id="f156e-1189">获取所选项目的正文中找到的指定实体类型的所有实体的数组。</span><span class="sxs-lookup"><span data-stu-id="f156e-1189">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="f156e-1190">iOS 版 Outlook 或 Android 版 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="f156e-1190">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="f156e-1191">Parameters</span><span class="sxs-lookup"><span data-stu-id="f156e-1191">Parameters</span></span>

|<span data-ttu-id="f156e-1192">名称</span><span class="sxs-lookup"><span data-stu-id="f156e-1192">Name</span></span>|<span data-ttu-id="f156e-1193">类型</span><span class="sxs-lookup"><span data-stu-id="f156e-1193">Type</span></span>|<span data-ttu-id="f156e-1194">说明</span><span class="sxs-lookup"><span data-stu-id="f156e-1194">Description</span></span>|
|---|---|---|
|`entityType`|[<span data-ttu-id="f156e-1195">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="f156e-1195">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook/office.mailboxenums.entitytype)|<span data-ttu-id="f156e-1196">EntityType 枚举值之一。</span><span class="sxs-lookup"><span data-stu-id="f156e-1196">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f156e-1197">Requirements</span><span class="sxs-lookup"><span data-stu-id="f156e-1197">Requirements</span></span>

|<span data-ttu-id="f156e-1198">要求</span><span class="sxs-lookup"><span data-stu-id="f156e-1198">Requirement</span></span>|<span data-ttu-id="f156e-1199">值</span><span class="sxs-lookup"><span data-stu-id="f156e-1199">Value</span></span>|
|---|---|
|[<span data-ttu-id="f156e-1200">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="f156e-1200">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f156e-1201">1.0</span><span class="sxs-lookup"><span data-stu-id="f156e-1201">1.0</span></span>|
|[<span data-ttu-id="f156e-1202">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="f156e-1202">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f156e-1203">受限</span><span class="sxs-lookup"><span data-stu-id="f156e-1203">Restricted</span></span>|
|[<span data-ttu-id="f156e-1204">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="f156e-1204">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="f156e-1205">阅读</span><span class="sxs-lookup"><span data-stu-id="f156e-1205">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="f156e-1206">返回：</span><span class="sxs-lookup"><span data-stu-id="f156e-1206">Returns:</span></span>

<span data-ttu-id="f156e-1207">如果在 `entityType` 中传递的值不是 `EntityType` 枚举的有效成员，该方法返回 null。</span><span class="sxs-lookup"><span data-stu-id="f156e-1207">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="f156e-1208">如果指定类型的任何实体都不存在于该项目的正文中，该方法将返回空数组。</span><span class="sxs-lookup"><span data-stu-id="f156e-1208">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="f156e-1209">否则，返回的数组中对象的类型取决于 `entityType` 参数中请求实体的类型。</span><span class="sxs-lookup"><span data-stu-id="f156e-1209">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="f156e-1210">当使用此方法的最低权限级别**受限**时，某些实体类型需要 **ReadItem** 才能进行访问，如下表中所指定。</span><span class="sxs-lookup"><span data-stu-id="f156e-1210">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

|<span data-ttu-id="f156e-1211">`entityType` 的值</span><span class="sxs-lookup"><span data-stu-id="f156e-1211">Value of `entityType`</span></span>|<span data-ttu-id="f156e-1212">返回的数组中对象的类型</span><span class="sxs-lookup"><span data-stu-id="f156e-1212">Type of objects in returned array</span></span>|<span data-ttu-id="f156e-1213">所需权限级别</span><span class="sxs-lookup"><span data-stu-id="f156e-1213">Required Permission Level</span></span>|
|---|---|---|
|`Address`|<span data-ttu-id="f156e-1214">String</span><span class="sxs-lookup"><span data-stu-id="f156e-1214">String</span></span>|<span data-ttu-id="f156e-1215">**受限**</span><span class="sxs-lookup"><span data-stu-id="f156e-1215">**Restricted**</span></span>|
|`Contact`|<span data-ttu-id="f156e-1216">Contact</span><span class="sxs-lookup"><span data-stu-id="f156e-1216">Contact</span></span>|<span data-ttu-id="f156e-1217">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="f156e-1217">**ReadItem**</span></span>|
|`EmailAddress`|<span data-ttu-id="f156e-1218">String</span><span class="sxs-lookup"><span data-stu-id="f156e-1218">String</span></span>|<span data-ttu-id="f156e-1219">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="f156e-1219">**ReadItem**</span></span>|
|`MeetingSuggestion`|<span data-ttu-id="f156e-1220">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="f156e-1220">MeetingSuggestion</span></span>|<span data-ttu-id="f156e-1221">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="f156e-1221">**ReadItem**</span></span>|
|`PhoneNumber`|<span data-ttu-id="f156e-1222">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="f156e-1222">PhoneNumber</span></span>|<span data-ttu-id="f156e-1223">**受限**</span><span class="sxs-lookup"><span data-stu-id="f156e-1223">**Restricted**</span></span>|
|`TaskSuggestion`|<span data-ttu-id="f156e-1224">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="f156e-1224">TaskSuggestion</span></span>|<span data-ttu-id="f156e-1225">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="f156e-1225">**ReadItem**</span></span>|
|`URL`|<span data-ttu-id="f156e-1226">String</span><span class="sxs-lookup"><span data-stu-id="f156e-1226">String</span></span>|<span data-ttu-id="f156e-1227">**受限**</span><span class="sxs-lookup"><span data-stu-id="f156e-1227">**Restricted**</span></span>|

<span data-ttu-id="f156e-1228">类型：Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="f156e-1228">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span></span>

##### <a name="example"></a><span data-ttu-id="f156e-1229">示例</span><span class="sxs-lookup"><span data-stu-id="f156e-1229">Example</span></span>

<span data-ttu-id="f156e-1230">以下示例显示了如何访问表示当前项目的正文中的邮政地址的字符串数组。</span><span class="sxs-lookup"><span data-stu-id="f156e-1230">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlookofficecontactmeetingsuggestionjavascriptapioutlookofficemeetingsuggestionphonenumberjavascriptapioutlookofficephonenumbertasksuggestionjavascriptapioutlookofficetasksuggestion"></a><span data-ttu-id="f156e-1231">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="f156e-1231">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span></span>

<span data-ttu-id="f156e-1232">返回传递清单 XML 文件中定义的命名筛选器的所选项目中的已知实体。</span><span class="sxs-lookup"><span data-stu-id="f156e-1232">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="f156e-1233">iOS 版 Outlook 或 Android 版 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="f156e-1233">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="f156e-1234">`getFilteredEntitiesByName` 方法返回匹配在具有指定 `FilterName` 元素值的清单 XML 文件中的 [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) 规则元素中定义的正则表达式的实体。</span><span class="sxs-lookup"><span data-stu-id="f156e-1234">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="f156e-1235">参数</span><span class="sxs-lookup"><span data-stu-id="f156e-1235">Parameters</span></span>

|<span data-ttu-id="f156e-1236">名称</span><span class="sxs-lookup"><span data-stu-id="f156e-1236">Name</span></span>|<span data-ttu-id="f156e-1237">类型</span><span class="sxs-lookup"><span data-stu-id="f156e-1237">Type</span></span>|<span data-ttu-id="f156e-1238">说明</span><span class="sxs-lookup"><span data-stu-id="f156e-1238">Description</span></span>|
|---|---|---|
|`name`|<span data-ttu-id="f156e-1239">字符串</span><span class="sxs-lookup"><span data-stu-id="f156e-1239">String</span></span>|<span data-ttu-id="f156e-1240">定义筛选器匹配的 `ItemHasKnownEntity` 规则元素的名称。</span><span class="sxs-lookup"><span data-stu-id="f156e-1240">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f156e-1241">要求</span><span class="sxs-lookup"><span data-stu-id="f156e-1241">Requirements</span></span>

|<span data-ttu-id="f156e-1242">要求</span><span class="sxs-lookup"><span data-stu-id="f156e-1242">Requirement</span></span>|<span data-ttu-id="f156e-1243">值</span><span class="sxs-lookup"><span data-stu-id="f156e-1243">Value</span></span>|
|---|---|
|[<span data-ttu-id="f156e-1244">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="f156e-1244">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f156e-1245">1.0</span><span class="sxs-lookup"><span data-stu-id="f156e-1245">1.0</span></span>|
|[<span data-ttu-id="f156e-1246">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="f156e-1246">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f156e-1247">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f156e-1247">ReadItem</span></span>|
|[<span data-ttu-id="f156e-1248">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="f156e-1248">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="f156e-1249">阅读</span><span class="sxs-lookup"><span data-stu-id="f156e-1249">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="f156e-1250">返回：</span><span class="sxs-lookup"><span data-stu-id="f156e-1250">Returns:</span></span>

<span data-ttu-id="f156e-p174">如果具有匹配 `name` 参数的 `FilterName` 元素值的清单中没有任何 `ItemHasKnownEntity` 元素，则该方法返回 `null`。如果 `name` 参数匹配清单中的 `ItemHasKnownEntity` 元素，但在匹配的当前项目中没有实体，则该方法返回一个空数组。</span><span class="sxs-lookup"><span data-stu-id="f156e-p174">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="f156e-1253">类型：Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="f156e-1253">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span></span>

<br>

---
---

#### <a name="getinitializationcontextasyncoptions-callback"></a><span data-ttu-id="f156e-1254">Office.context.mailbox.item.getinitializationcontextasync （[options]，[callback]）</span><span class="sxs-lookup"><span data-stu-id="f156e-1254">getInitializationContextAsync([options], [callback])</span></span>

<span data-ttu-id="f156e-1255">获取[通过可操作邮件激活](/outlook/actionable-messages/invoke-add-in-from-actionable-message)加载项时传递的初始化数据。</span><span class="sxs-lookup"><span data-stu-id="f156e-1255">Gets initialization data passed when the add-in is [activated by an actionable message](/outlook/actionable-messages/invoke-add-in-from-actionable-message).</span></span>

> [!NOTE]
> <span data-ttu-id="f156e-1256">仅 Outlook 2016 或更高版本（高于16.0.8413.1000 的即点即用版本）和适用于 Office 365 的 Outlook 网页版支持此方法。</span><span class="sxs-lookup"><span data-stu-id="f156e-1256">This method is only supported by Outlook 2016 or later on Windows (Click-to-Run versions later than 16.0.8413.1000) and Outlook on the web for Office 365.</span></span>

##### <a name="parameters"></a><span data-ttu-id="f156e-1257">Parameters</span><span class="sxs-lookup"><span data-stu-id="f156e-1257">Parameters</span></span>

|<span data-ttu-id="f156e-1258">名称</span><span class="sxs-lookup"><span data-stu-id="f156e-1258">Name</span></span>|<span data-ttu-id="f156e-1259">类型</span><span class="sxs-lookup"><span data-stu-id="f156e-1259">Type</span></span>|<span data-ttu-id="f156e-1260">属性</span><span class="sxs-lookup"><span data-stu-id="f156e-1260">Attributes</span></span>|<span data-ttu-id="f156e-1261">说明</span><span class="sxs-lookup"><span data-stu-id="f156e-1261">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="f156e-1262">Object</span><span class="sxs-lookup"><span data-stu-id="f156e-1262">Object</span></span>|<span data-ttu-id="f156e-1263">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="f156e-1263">&lt;optional&gt;</span></span>|<span data-ttu-id="f156e-1264">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="f156e-1264">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="f156e-1265">Object</span><span class="sxs-lookup"><span data-stu-id="f156e-1265">Object</span></span>|<span data-ttu-id="f156e-1266">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="f156e-1266">&lt;optional&gt;</span></span>|<span data-ttu-id="f156e-1267">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="f156e-1267">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="f156e-1268">函数</span><span class="sxs-lookup"><span data-stu-id="f156e-1268">function</span></span>|<span data-ttu-id="f156e-1269">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="f156e-1269">&lt;optional&gt;</span></span>|<span data-ttu-id="f156e-1270">方法完成后，使用单个参数 `callback`（一个 [`asyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `AsyncResult` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="f156e-1270">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="f156e-1271">如果成功，初始化数据在`asyncResult.value`属性中提供为字符串。</span><span class="sxs-lookup"><span data-stu-id="f156e-1271">On success, the initialization data is provided in the `asyncResult.value` property as a string.</span></span><br/><span data-ttu-id="f156e-1272">如果没有初始化上下文，该`asyncResult`对象将包含其`Error` `code`属性设置为`9020`的对象及其`name`属性设置为。 `GenericResponseError`</span><span class="sxs-lookup"><span data-stu-id="f156e-1272">If there is no initialization context, the `asyncResult` object will contain an `Error` object with its `code` property set to `9020` and its `name` property set to `GenericResponseError`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f156e-1273">Requirements</span><span class="sxs-lookup"><span data-stu-id="f156e-1273">Requirements</span></span>

|<span data-ttu-id="f156e-1274">要求</span><span class="sxs-lookup"><span data-stu-id="f156e-1274">Requirement</span></span>|<span data-ttu-id="f156e-1275">值</span><span class="sxs-lookup"><span data-stu-id="f156e-1275">Value</span></span>|
|---|---|
|[<span data-ttu-id="f156e-1276">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="f156e-1276">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f156e-1277">预览</span><span class="sxs-lookup"><span data-stu-id="f156e-1277">Preview</span></span>|
|[<span data-ttu-id="f156e-1278">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="f156e-1278">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f156e-1279">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f156e-1279">ReadItem</span></span>|
|[<span data-ttu-id="f156e-1280">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="f156e-1280">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="f156e-1281">阅读</span><span class="sxs-lookup"><span data-stu-id="f156e-1281">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="f156e-1282">示例</span><span class="sxs-lookup"><span data-stu-id="f156e-1282">Example</span></span>

```js
// Get the initialization context (if present).
Office.context.mailbox.item.getInitializationContextAsync(
  function(asyncResult) {
    if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
      if (asyncResult.value != null && asyncResult.value.length > 0) {
        // The value is a string, parse to an object.
        var context = JSON.parse(asyncResult.value);
        // Do something with context.
      } else {
        // Empty context, treat as no context.
      }
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

#### <a name="getitemidasyncoptions-callback"></a><span data-ttu-id="f156e-1283">getItemIdAsync （[options]，回拨）</span><span class="sxs-lookup"><span data-stu-id="f156e-1283">getItemIdAsync([options], callback)</span></span>

<span data-ttu-id="f156e-1284">异步获取已保存项的 ID。</span><span class="sxs-lookup"><span data-stu-id="f156e-1284">Asynchronously gets the ID of a saved item.</span></span> <span data-ttu-id="f156e-1285">仅限撰写模式。</span><span class="sxs-lookup"><span data-stu-id="f156e-1285">Compose mode only.</span></span>

<span data-ttu-id="f156e-1286">调用此方法时，此方法通过回调方法返回项 ID。</span><span class="sxs-lookup"><span data-stu-id="f156e-1286">When invoked, this method returns the item ID via the callback method.</span></span>

> [!NOTE]
> <span data-ttu-id="f156e-1287">如果你的外接程序`getItemIdAsync`对撰写模式中的项（例如，要获取`itemId`使用 EWS 或 REST API 的使用）调用，请注意，当 Outlook 处于缓存模式下时，可能需要一段时间才能将项目同步到服务器。</span><span class="sxs-lookup"><span data-stu-id="f156e-1287">If your add-in calls `getItemIdAsync` on an item in compose mode (e.g., to get an `itemId` to use with EWS or the REST API), be aware that when Outlook is in cached mode, it may take some time before the item is synced to the server.</span></span> <span data-ttu-id="f156e-1288">在同步项目之前，无法识别`itemId`该项目并使用它将返回错误。</span><span class="sxs-lookup"><span data-stu-id="f156e-1288">Until the item is synced, the `itemId` is not recognized and using it returns an error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="f156e-1289">Parameters</span><span class="sxs-lookup"><span data-stu-id="f156e-1289">Parameters</span></span>

|<span data-ttu-id="f156e-1290">名称</span><span class="sxs-lookup"><span data-stu-id="f156e-1290">Name</span></span>|<span data-ttu-id="f156e-1291">类型</span><span class="sxs-lookup"><span data-stu-id="f156e-1291">Type</span></span>|<span data-ttu-id="f156e-1292">属性</span><span class="sxs-lookup"><span data-stu-id="f156e-1292">Attributes</span></span>|<span data-ttu-id="f156e-1293">说明</span><span class="sxs-lookup"><span data-stu-id="f156e-1293">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="f156e-1294">Object</span><span class="sxs-lookup"><span data-stu-id="f156e-1294">Object</span></span>|<span data-ttu-id="f156e-1295">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="f156e-1295">&lt;optional&gt;</span></span>|<span data-ttu-id="f156e-1296">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="f156e-1296">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="f156e-1297">Object</span><span class="sxs-lookup"><span data-stu-id="f156e-1297">Object</span></span>|<span data-ttu-id="f156e-1298">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="f156e-1298">&lt;optional&gt;</span></span>|<span data-ttu-id="f156e-1299">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="f156e-1299">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="f156e-1300">函数</span><span class="sxs-lookup"><span data-stu-id="f156e-1300">function</span></span>||<span data-ttu-id="f156e-1301">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="f156e-1301">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="f156e-1302">如果成功，该项目标识符将在 `asyncResult.value` 属性中提供。</span><span class="sxs-lookup"><span data-stu-id="f156e-1302">On success, the item identifier is provided in the `asyncResult.value` property.</span></span>|

##### <a name="errors"></a><span data-ttu-id="f156e-1303">错误</span><span class="sxs-lookup"><span data-stu-id="f156e-1303">Errors</span></span>

|<span data-ttu-id="f156e-1304">错误代码</span><span class="sxs-lookup"><span data-stu-id="f156e-1304">Error code</span></span>|<span data-ttu-id="f156e-1305">说明</span><span class="sxs-lookup"><span data-stu-id="f156e-1305">Description</span></span>|
|------------|-------------|
|`ItemNotSaved`|<span data-ttu-id="f156e-1306">在保存项目之前，无法检索此 id。</span><span class="sxs-lookup"><span data-stu-id="f156e-1306">The id can't be retrieved until the item is saved.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f156e-1307">Requirements</span><span class="sxs-lookup"><span data-stu-id="f156e-1307">Requirements</span></span>

|<span data-ttu-id="f156e-1308">要求</span><span class="sxs-lookup"><span data-stu-id="f156e-1308">Requirement</span></span>|<span data-ttu-id="f156e-1309">值</span><span class="sxs-lookup"><span data-stu-id="f156e-1309">Value</span></span>|
|---|---|
|[<span data-ttu-id="f156e-1310">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="f156e-1310">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f156e-1311">1.8</span><span class="sxs-lookup"><span data-stu-id="f156e-1311">1.8</span></span>|
|[<span data-ttu-id="f156e-1312">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="f156e-1312">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f156e-1313">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f156e-1313">ReadItem</span></span>|
|[<span data-ttu-id="f156e-1314">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="f156e-1314">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="f156e-1315">撰写</span><span class="sxs-lookup"><span data-stu-id="f156e-1315">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="f156e-1316">示例</span><span class="sxs-lookup"><span data-stu-id="f156e-1316">Examples</span></span>

```js
Office.context.mailbox.item.getItemIdAsync(
  function callback(result) {
    // Process the result.
  });
```

<span data-ttu-id="f156e-1317">下面的示例演示传递给回调函数`result`的参数的结构。</span><span class="sxs-lookup"><span data-stu-id="f156e-1317">The following example shows the structure of the `result` parameter that's passed to the callback function.</span></span> <span data-ttu-id="f156e-1318">`value`属性包含项 ID。</span><span class="sxs-lookup"><span data-stu-id="f156e-1318">The `value` property contains the item ID.</span></span>

```json
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

<br>

---
---

#### <a name="getregexmatches--object"></a><span data-ttu-id="f156e-1319">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="f156e-1319">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="f156e-1320">返回所选项目中匹配在清单 XML 文件中定义的正则表达式的字符串值。</span><span class="sxs-lookup"><span data-stu-id="f156e-1320">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="f156e-1321">iOS 版 Outlook 或 Android 版 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="f156e-1321">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="f156e-p178">`getRegExMatches` 方法返回匹配在清单 XML 文件中的每个 `ItemHasRegularExpressionMatch` 或 `ItemHasKnownEntity` 规则元素中定义的正则表达式的字符串。对于 `ItemHasRegularExpressionMatch` 规则，匹配字符串必须发生在该规则指定的项目的属性中。`PropertyName` 简单类型定义支持的属性。</span><span class="sxs-lookup"><span data-stu-id="f156e-p178">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="f156e-1325">例如，考虑一个外接程序清单具有以下 `Rule` 元素：</span><span class="sxs-lookup"><span data-stu-id="f156e-1325">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="f156e-1326">从 `getRegExMatches` 返回的对象应有两个属性：`fruits` 和 `veggies`。</span><span class="sxs-lookup"><span data-stu-id="f156e-1326">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="f156e-p179">如果在项目的正文属性上指定 `ItemHasRegularExpressionMatch` 规则，则正则表达式应进一步筛选正文，不应尝试返回该项目的整个正文。使用正则表达式（如 `.*`）获取项目的整个正文并不总是返回预期的结果。而是使用 [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) 方法检索整个正文。</span><span class="sxs-lookup"><span data-stu-id="f156e-p179">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="f156e-1330">Requirements</span><span class="sxs-lookup"><span data-stu-id="f156e-1330">Requirements</span></span>

|<span data-ttu-id="f156e-1331">要求</span><span class="sxs-lookup"><span data-stu-id="f156e-1331">Requirement</span></span>|<span data-ttu-id="f156e-1332">值</span><span class="sxs-lookup"><span data-stu-id="f156e-1332">Value</span></span>|
|---|---|
|[<span data-ttu-id="f156e-1333">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="f156e-1333">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f156e-1334">1.0</span><span class="sxs-lookup"><span data-stu-id="f156e-1334">1.0</span></span>|
|[<span data-ttu-id="f156e-1335">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="f156e-1335">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f156e-1336">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f156e-1336">ReadItem</span></span>|
|[<span data-ttu-id="f156e-1337">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="f156e-1337">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="f156e-1338">阅读</span><span class="sxs-lookup"><span data-stu-id="f156e-1338">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="f156e-1339">返回：</span><span class="sxs-lookup"><span data-stu-id="f156e-1339">Returns:</span></span>

<span data-ttu-id="f156e-p180">一个包含匹配在清单 XML 文件中定义的正则表达式的字符串数组的对象。每个数组的名称等于匹配 `ItemHasRegularExpressionMatch` 规则的 `RegExName` 属性或匹配 `ItemHasKnownEntity` 规则的 `FilterName` 属性的相应值。</span><span class="sxs-lookup"><span data-stu-id="f156e-p180">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<dl class="param-type"><span data-ttu-id="f156e-1342">

<dt>
类型</dt>


</span><span class="sxs-lookup"><span data-stu-id="f156e-1342">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="f156e-1343">对象</span><span class="sxs-lookup"><span data-stu-id="f156e-1343">Object</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="f156e-1344">示例</span><span class="sxs-lookup"><span data-stu-id="f156e-1344">Example</span></span>

<span data-ttu-id="f156e-1345">以下示例显示了如何访问正则表达式规则元素 `fruits` 和 `veggies` 的匹配项的数组，这些元素在清单中指定。</span><span class="sxs-lookup"><span data-stu-id="f156e-1345">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```js
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veggies = allMatches.veggies;
```

<br>

---
---

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="f156e-1346">getRegExMatchesByName(name) → (nullable) {Array.<String>}</span><span class="sxs-lookup"><span data-stu-id="f156e-1346">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="f156e-1347">返回匹配在清单 XML 文件中定义的命名正则表达式的所选项目中的字符串值。</span><span class="sxs-lookup"><span data-stu-id="f156e-1347">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="f156e-1348">iOS 版 Outlook 或 Android 版 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="f156e-1348">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="f156e-1349">`getRegExMatchesByName` 方法返回匹配在具有指定 `RegExName` 元素值的清单 XML 文件中的 `ItemHasRegularExpressionMatch` 规则元素中定义的正则表达式的字符串。</span><span class="sxs-lookup"><span data-stu-id="f156e-1349">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="f156e-p181">如果在项目的正文属性上指定 `ItemHasRegularExpressionMatch` 规则，则正则表达式应进一步筛选正文，不应尝试返回该项目的整个正文。使用正则表达式（如 `.*`）获取项目的整个正文并不总是返回预期的结果。</span><span class="sxs-lookup"><span data-stu-id="f156e-p181">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="f156e-1352">参数</span><span class="sxs-lookup"><span data-stu-id="f156e-1352">Parameters</span></span>

|<span data-ttu-id="f156e-1353">名称</span><span class="sxs-lookup"><span data-stu-id="f156e-1353">Name</span></span>|<span data-ttu-id="f156e-1354">类型</span><span class="sxs-lookup"><span data-stu-id="f156e-1354">Type</span></span>|<span data-ttu-id="f156e-1355">说明</span><span class="sxs-lookup"><span data-stu-id="f156e-1355">Description</span></span>|
|---|---|---|
|`name`|<span data-ttu-id="f156e-1356">字符串</span><span class="sxs-lookup"><span data-stu-id="f156e-1356">String</span></span>|<span data-ttu-id="f156e-1357">定义筛选器匹配的 `ItemHasRegularExpressionMatch` 规则元素的名称。</span><span class="sxs-lookup"><span data-stu-id="f156e-1357">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f156e-1358">Requirements</span><span class="sxs-lookup"><span data-stu-id="f156e-1358">Requirements</span></span>

|<span data-ttu-id="f156e-1359">要求</span><span class="sxs-lookup"><span data-stu-id="f156e-1359">Requirement</span></span>|<span data-ttu-id="f156e-1360">值</span><span class="sxs-lookup"><span data-stu-id="f156e-1360">Value</span></span>|
|---|---|
|[<span data-ttu-id="f156e-1361">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="f156e-1361">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f156e-1362">1.0</span><span class="sxs-lookup"><span data-stu-id="f156e-1362">1.0</span></span>|
|[<span data-ttu-id="f156e-1363">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="f156e-1363">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f156e-1364">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f156e-1364">ReadItem</span></span>|
|[<span data-ttu-id="f156e-1365">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="f156e-1365">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="f156e-1366">阅读</span><span class="sxs-lookup"><span data-stu-id="f156e-1366">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="f156e-1367">返回：</span><span class="sxs-lookup"><span data-stu-id="f156e-1367">Returns:</span></span>

<span data-ttu-id="f156e-1368">一个包含匹配在清单 XML 文件中定义的正则表达式的字符串的数组。</span><span class="sxs-lookup"><span data-stu-id="f156e-1368">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<span data-ttu-id="f156e-1369">类型：Array.< String ></span><span class="sxs-lookup"><span data-stu-id="f156e-1369">Type: Array.< String ></span></span>

##### <a name="example"></a><span data-ttu-id="f156e-1370">示例</span><span class="sxs-lookup"><span data-stu-id="f156e-1370">Example</span></span>

```js
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

<br>

---
---

#### <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="f156e-1371">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="f156e-1371">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="f156e-1372">以异步方式返回邮件的主题或正文中选定的数据。</span><span class="sxs-lookup"><span data-stu-id="f156e-1372">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="f156e-p182">如果没有选定内容，但光标位于正文或主题中，此方法将会为所选数据返回 null。如果选定的是字段，而不是正文或主题，则此方法返回 `InvalidSelection` 错误。</span><span class="sxs-lookup"><span data-stu-id="f156e-p182">If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="f156e-1375">参数</span><span class="sxs-lookup"><span data-stu-id="f156e-1375">Parameters</span></span>

|<span data-ttu-id="f156e-1376">名称</span><span class="sxs-lookup"><span data-stu-id="f156e-1376">Name</span></span>|<span data-ttu-id="f156e-1377">类型</span><span class="sxs-lookup"><span data-stu-id="f156e-1377">Type</span></span>|<span data-ttu-id="f156e-1378">属性</span><span class="sxs-lookup"><span data-stu-id="f156e-1378">Attributes</span></span>|<span data-ttu-id="f156e-1379">说明</span><span class="sxs-lookup"><span data-stu-id="f156e-1379">Description</span></span>|
|---|---|---|---|
|`coercionType`|[<span data-ttu-id="f156e-1380">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="f156e-1380">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="f156e-p183">请求数据的格式。如果为文本，则此方法返回纯文本作为字符串，删除任何显示的 HTML 标记。如果为 HTML，则此方法返回所选文本，不论是纯文本还是 HTML。</span><span class="sxs-lookup"><span data-stu-id="f156e-p183">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`|<span data-ttu-id="f156e-1384">对象</span><span class="sxs-lookup"><span data-stu-id="f156e-1384">Object</span></span>|<span data-ttu-id="f156e-1385">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="f156e-1385">&lt;optional&gt;</span></span>|<span data-ttu-id="f156e-1386">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="f156e-1386">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="f156e-1387">对象</span><span class="sxs-lookup"><span data-stu-id="f156e-1387">Object</span></span>|<span data-ttu-id="f156e-1388">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="f156e-1388">&lt;optional&gt;</span></span>|<span data-ttu-id="f156e-1389">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="f156e-1389">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="f156e-1390">函数</span><span class="sxs-lookup"><span data-stu-id="f156e-1390">function</span></span>||<span data-ttu-id="f156e-1391">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="f156e-1391">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="f156e-1392">若要从回调方法访问所选数据，请调用 `asyncResult.value.data`。</span><span class="sxs-lookup"><span data-stu-id="f156e-1392">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="f156e-1393">若要访问选定内容的源属性，请调用 `asyncResult.value.sourceProperty`，这将为 `body` 或 `subject`。</span><span class="sxs-lookup"><span data-stu-id="f156e-1393">To access the source property that the selection comes from, call `asyncResult.value.sourceProperty`, which will be either `body` or `subject`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f156e-1394">Requirements</span><span class="sxs-lookup"><span data-stu-id="f156e-1394">Requirements</span></span>

|<span data-ttu-id="f156e-1395">要求</span><span class="sxs-lookup"><span data-stu-id="f156e-1395">Requirement</span></span>|<span data-ttu-id="f156e-1396">值</span><span class="sxs-lookup"><span data-stu-id="f156e-1396">Value</span></span>|
|---|---|
|[<span data-ttu-id="f156e-1397">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="f156e-1397">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f156e-1398">1.2</span><span class="sxs-lookup"><span data-stu-id="f156e-1398">1.2</span></span>|
|[<span data-ttu-id="f156e-1399">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="f156e-1399">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f156e-1400">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f156e-1400">ReadItem</span></span>|
|[<span data-ttu-id="f156e-1401">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="f156e-1401">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="f156e-1402">撰写</span><span class="sxs-lookup"><span data-stu-id="f156e-1402">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="f156e-1403">返回：</span><span class="sxs-lookup"><span data-stu-id="f156e-1403">Returns:</span></span>

<span data-ttu-id="f156e-1404">作为字符串的所选数据的格式由 `coercionType` 确定。</span><span class="sxs-lookup"><span data-stu-id="f156e-1404">The selected data as a string with format determined by `coercionType`.</span></span>

<span data-ttu-id="f156e-1405">类型：字符串</span><span class="sxs-lookup"><span data-stu-id="f156e-1405">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="f156e-1406">示例</span><span class="sxs-lookup"><span data-stu-id="f156e-1406">Example</span></span>

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

#### <a name="getselectedentities--entitiesjavascriptapioutlookofficeentities"></a><span data-ttu-id="f156e-1407">getSelectedEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="f156e-1407">getSelectedEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span></span>

<span data-ttu-id="f156e-1408">获取在用户已选择的突出显示匹配项中找到的实体。</span><span class="sxs-lookup"><span data-stu-id="f156e-1408">Gets the entities found in a highlighted match a user has selected.</span></span> <span data-ttu-id="f156e-1409">突出显示匹配项适用于[上下文外接程序](/outlook/add-ins/contextual-outlook-add-ins)。</span><span class="sxs-lookup"><span data-stu-id="f156e-1409">Highlighted matches apply to [contextual add-ins](/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="f156e-1410">iOS 版 Outlook 或 Android 版 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="f156e-1410">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="f156e-1411">要求</span><span class="sxs-lookup"><span data-stu-id="f156e-1411">Requirements</span></span>

|<span data-ttu-id="f156e-1412">要求</span><span class="sxs-lookup"><span data-stu-id="f156e-1412">Requirement</span></span>|<span data-ttu-id="f156e-1413">值</span><span class="sxs-lookup"><span data-stu-id="f156e-1413">Value</span></span>|
|---|---|
|[<span data-ttu-id="f156e-1414">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="f156e-1414">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f156e-1415">1.6</span><span class="sxs-lookup"><span data-stu-id="f156e-1415">1.6</span></span>|
|[<span data-ttu-id="f156e-1416">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="f156e-1416">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f156e-1417">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f156e-1417">ReadItem</span></span>|
|[<span data-ttu-id="f156e-1418">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="f156e-1418">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="f156e-1419">阅读</span><span class="sxs-lookup"><span data-stu-id="f156e-1419">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="f156e-1420">返回：</span><span class="sxs-lookup"><span data-stu-id="f156e-1420">Returns:</span></span>

<span data-ttu-id="f156e-1421">类型：[Entities](/javascript/api/outlook/office.entities)</span><span class="sxs-lookup"><span data-stu-id="f156e-1421">Type: [Entities](/javascript/api/outlook/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="f156e-1422">示例</span><span class="sxs-lookup"><span data-stu-id="f156e-1422">Example</span></span>

<span data-ttu-id="f156e-1423">以下示例访问用户选择的突出显示匹配项中的地址实体。</span><span class="sxs-lookup"><span data-stu-id="f156e-1423">The following example accesses the addresses entities in the highlighted match selected by the user.</span></span>

```js
var contacts = Office.context.mailbox.item.getSelectedEntities().addresses;
```

<br>

---
---

#### <a name="getselectedregexmatches--object"></a><span data-ttu-id="f156e-1424">getSelectedRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="f156e-1424">getSelectedRegExMatches() → {Object}</span></span>

<span data-ttu-id="f156e-p186">返回突出显示匹配项中匹配在清单 XML 文件中定义的正则表达式的字符串值。突出显示匹配项适用于[上下文外接程序](/outlook/add-ins/contextual-outlook-add-ins)。</span><span class="sxs-lookup"><span data-stu-id="f156e-p186">Returns string values in a highlighted match that match the regular expressions defined in the manifest XML file. Highlighted matches apply to [contextual add-ins](/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="f156e-1427">iOS 版 Outlook 或 Android 版 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="f156e-1427">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="f156e-p187">`getSelectedRegExMatches` 方法返回匹配在清单 XML 文件中的每个 `ItemHasRegularExpressionMatch` 或 `ItemHasKnownEntity` 规则元素中定义的正则表达式的字符串。对于 `ItemHasRegularExpressionMatch` 规则，匹配字符串必须发生在该规则指定的项目的属性中。`PropertyName` 简单类型定义支持的属性。</span><span class="sxs-lookup"><span data-stu-id="f156e-p187">The `getSelectedRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="f156e-1431">例如，考虑一个外接程序清单具有以下 `Rule` 元素：</span><span class="sxs-lookup"><span data-stu-id="f156e-1431">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="f156e-1432">从 `getRegExMatches` 返回的对象应有两个属性：`fruits` 和 `veggies`。</span><span class="sxs-lookup"><span data-stu-id="f156e-1432">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="f156e-p188">如果在项目的正文属性上指定 `ItemHasRegularExpressionMatch` 规则，则正则表达式应进一步筛选正文，不应尝试返回该项目的整个正文。使用正则表达式（如 `.*`）获取项目的整个正文并不总是返回预期的结果。而是使用 [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) 方法检索整个正文。</span><span class="sxs-lookup"><span data-stu-id="f156e-p188">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="f156e-1436">Requirements</span><span class="sxs-lookup"><span data-stu-id="f156e-1436">Requirements</span></span>

|<span data-ttu-id="f156e-1437">要求</span><span class="sxs-lookup"><span data-stu-id="f156e-1437">Requirement</span></span>|<span data-ttu-id="f156e-1438">值</span><span class="sxs-lookup"><span data-stu-id="f156e-1438">Value</span></span>|
|---|---|
|[<span data-ttu-id="f156e-1439">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="f156e-1439">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f156e-1440">1.6</span><span class="sxs-lookup"><span data-stu-id="f156e-1440">1.6</span></span>|
|[<span data-ttu-id="f156e-1441">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="f156e-1441">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f156e-1442">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f156e-1442">ReadItem</span></span>|
|[<span data-ttu-id="f156e-1443">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="f156e-1443">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="f156e-1444">阅读</span><span class="sxs-lookup"><span data-stu-id="f156e-1444">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="f156e-1445">返回：</span><span class="sxs-lookup"><span data-stu-id="f156e-1445">Returns:</span></span>

<span data-ttu-id="f156e-p189">一个包含匹配在清单 XML 文件中定义的正则表达式的字符串数组的对象。每个数组的名称等于匹配 `ItemHasRegularExpressionMatch` 规则的 `RegExName` 属性或匹配 `ItemHasKnownEntity` 规则的 `FilterName` 属性的相应值。</span><span class="sxs-lookup"><span data-stu-id="f156e-p189">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

##### <a name="example"></a><span data-ttu-id="f156e-1448">示例</span><span class="sxs-lookup"><span data-stu-id="f156e-1448">Example</span></span>

<span data-ttu-id="f156e-1449">以下示例显示了如何访问正则表达式规则元素 `fruits` 和 `veggies` 的匹配项的数组，这些元素在清单中指定。</span><span class="sxs-lookup"><span data-stu-id="f156e-1449">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```js
var selectedMatches = Office.context.mailbox.item.getSelectedRegExMatches();
var fruits = selectedMatches.fruits;
var veggies = selectedMatches.veggies;
```

<br>

---
---

#### <a name="getsharedpropertiesasyncoptions-callback"></a><span data-ttu-id="f156e-1450">getSharedPropertiesAsync （[options]，回拨）</span><span class="sxs-lookup"><span data-stu-id="f156e-1450">getSharedPropertiesAsync([options], callback)</span></span>

<span data-ttu-id="f156e-1451">获取共享文件夹、日历或邮箱中的所选约会或邮件的属性。</span><span class="sxs-lookup"><span data-stu-id="f156e-1451">Gets the properties of the selected appointment or message in a shared folder, calendar, or mailbox.</span></span>

##### <a name="parameters"></a><span data-ttu-id="f156e-1452">Parameters</span><span class="sxs-lookup"><span data-stu-id="f156e-1452">Parameters</span></span>

|<span data-ttu-id="f156e-1453">名称</span><span class="sxs-lookup"><span data-stu-id="f156e-1453">Name</span></span>|<span data-ttu-id="f156e-1454">类型</span><span class="sxs-lookup"><span data-stu-id="f156e-1454">Type</span></span>|<span data-ttu-id="f156e-1455">属性</span><span class="sxs-lookup"><span data-stu-id="f156e-1455">Attributes</span></span>|<span data-ttu-id="f156e-1456">说明</span><span class="sxs-lookup"><span data-stu-id="f156e-1456">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="f156e-1457">对象</span><span class="sxs-lookup"><span data-stu-id="f156e-1457">Object</span></span>|<span data-ttu-id="f156e-1458">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="f156e-1458">&lt;optional&gt;</span></span>|<span data-ttu-id="f156e-1459">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="f156e-1459">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="f156e-1460">对象</span><span class="sxs-lookup"><span data-stu-id="f156e-1460">Object</span></span>|<span data-ttu-id="f156e-1461">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="f156e-1461">&lt;optional&gt;</span></span>|<span data-ttu-id="f156e-1462">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="f156e-1462">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="f156e-1463">函数</span><span class="sxs-lookup"><span data-stu-id="f156e-1463">function</span></span>||<span data-ttu-id="f156e-1464">方法完成后，使用单个参数 `callback`（一个 [`asyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `AsyncResult` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="f156e-1464">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="f156e-1465">共享属性作为[`SharedProperties`](/javascript/api/outlook/office.sharedproperties) `asyncResult.value`属性中的对象提供。</span><span class="sxs-lookup"><span data-stu-id="f156e-1465">The shared properties are provided as a [`SharedProperties`](/javascript/api/outlook/office.sharedproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="f156e-1466">此对象可用于获取项目的共享属性。</span><span class="sxs-lookup"><span data-stu-id="f156e-1466">This object can be used to get the item's shared properties.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f156e-1467">Requirements</span><span class="sxs-lookup"><span data-stu-id="f156e-1467">Requirements</span></span>

|<span data-ttu-id="f156e-1468">要求</span><span class="sxs-lookup"><span data-stu-id="f156e-1468">Requirement</span></span>|<span data-ttu-id="f156e-1469">值</span><span class="sxs-lookup"><span data-stu-id="f156e-1469">Value</span></span>|
|---|---|
|[<span data-ttu-id="f156e-1470">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="f156e-1470">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f156e-1471">1.8</span><span class="sxs-lookup"><span data-stu-id="f156e-1471">1.8</span></span>|
|[<span data-ttu-id="f156e-1472">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="f156e-1472">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f156e-1473">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f156e-1473">ReadItem</span></span>|
|[<span data-ttu-id="f156e-1474">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="f156e-1474">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="f156e-1475">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="f156e-1475">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="f156e-1476">示例</span><span class="sxs-lookup"><span data-stu-id="f156e-1476">Example</span></span>

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

#### <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="f156e-1477">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="f156e-1477">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="f156e-1478">异步加载所选项目上此外接程序的自定义属性。</span><span class="sxs-lookup"><span data-stu-id="f156e-1478">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="f156e-p191">自定义属性基于每个应用、每个项目存储为键/值对。此方法在回调中返回 `CustomProperties` 对象，该回调提供访问特定于当前项目和当前外接程序的自定义属性的方法。自定义属性未在项目上加密，因此这不应用作安全存储。</span><span class="sxs-lookup"><span data-stu-id="f156e-p191">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="f156e-1482">参数</span><span class="sxs-lookup"><span data-stu-id="f156e-1482">Parameters</span></span>

|<span data-ttu-id="f156e-1483">名称</span><span class="sxs-lookup"><span data-stu-id="f156e-1483">Name</span></span>|<span data-ttu-id="f156e-1484">类型</span><span class="sxs-lookup"><span data-stu-id="f156e-1484">Type</span></span>|<span data-ttu-id="f156e-1485">属性</span><span class="sxs-lookup"><span data-stu-id="f156e-1485">Attributes</span></span>|<span data-ttu-id="f156e-1486">说明</span><span class="sxs-lookup"><span data-stu-id="f156e-1486">Description</span></span>|
|---|---|---|---|
|`callback`|<span data-ttu-id="f156e-1487">函数</span><span class="sxs-lookup"><span data-stu-id="f156e-1487">function</span></span>||<span data-ttu-id="f156e-1488">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="f156e-1488">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="f156e-1489">自定义属性作为 `asyncResult.value` 属性中的 [`CustomProperties`](/javascript/api/outlook/office.customproperties) 对象提供。</span><span class="sxs-lookup"><span data-stu-id="f156e-1489">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook/office.customproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="f156e-1490">此对象可用于获取、设置以及从项目中删除自定义属性，并将自定义属性集的更改重新保存到服务器。</span><span class="sxs-lookup"><span data-stu-id="f156e-1490">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`|<span data-ttu-id="f156e-1491">对象</span><span class="sxs-lookup"><span data-stu-id="f156e-1491">Object</span></span>|<span data-ttu-id="f156e-1492">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="f156e-1492">&lt;optional&gt;</span></span>|<span data-ttu-id="f156e-1493">开发人员可以提供他们想要在回调函数中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="f156e-1493">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="f156e-1494">此对象可以通过回调函数中的 `asyncResult.asyncContext` 属性进行访问。</span><span class="sxs-lookup"><span data-stu-id="f156e-1494">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f156e-1495">Requirements</span><span class="sxs-lookup"><span data-stu-id="f156e-1495">Requirements</span></span>

|<span data-ttu-id="f156e-1496">要求</span><span class="sxs-lookup"><span data-stu-id="f156e-1496">Requirement</span></span>|<span data-ttu-id="f156e-1497">值</span><span class="sxs-lookup"><span data-stu-id="f156e-1497">Value</span></span>|
|---|---|
|[<span data-ttu-id="f156e-1498">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="f156e-1498">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f156e-1499">1.0</span><span class="sxs-lookup"><span data-stu-id="f156e-1499">1.0</span></span>|
|[<span data-ttu-id="f156e-1500">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="f156e-1500">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f156e-1501">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f156e-1501">ReadItem</span></span>|
|[<span data-ttu-id="f156e-1502">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="f156e-1502">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="f156e-1503">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="f156e-1503">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="f156e-1504">示例</span><span class="sxs-lookup"><span data-stu-id="f156e-1504">Example</span></span>

<span data-ttu-id="f156e-p194">以下代码示例显示了如何使用 `loadCustomPropertiesAsync` 方法异步加载特定于当前项目的自定义属性。该示例还显示了如何使用 `CustomProperties.saveAsync` 方法将这些属性重新保存到服务器。加载自定义属性后，该代码示例将使用 `CustomProperties.get` 方法读取自定义属性 `myProp`，使用 `CustomProperties.set` 方法写入自定义属性 `otherProp`，最后调用 `saveAsync` 方法保存这些自定义属性。</span><span class="sxs-lookup"><span data-stu-id="f156e-p194">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

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

#### <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="f156e-1508">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="f156e-1508">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="f156e-1509">将附件从邮件或约会中删除。</span><span class="sxs-lookup"><span data-stu-id="f156e-1509">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="f156e-1510">`removeAttachmentAsync` 方法删除项目中带指定标识符的附件。</span><span class="sxs-lookup"><span data-stu-id="f156e-1510">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item.</span></span> <span data-ttu-id="f156e-1511">最佳做法是，仅当同一个邮件应用程序在同一会话中添加了一个附件时，你才应使用该附件标识符来删除该附件。</span><span class="sxs-lookup"><span data-stu-id="f156e-1511">As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session.</span></span> <span data-ttu-id="f156e-1512">在 Outlook 网页版和移动设备上，附件标识符只在同一个会话中才有效。</span><span class="sxs-lookup"><span data-stu-id="f156e-1512">In Outlook on the web and mobile devices, the attachment identifier is valid only within the same session.</span></span> <span data-ttu-id="f156e-1513">当用户关闭应用程序时，或者如果用户开始撰写内嵌窗体，随后弹出窗体以继续在单独的窗口中。</span><span class="sxs-lookup"><span data-stu-id="f156e-1513">A session is over when the user closes the app, or if the user starts composing an inline form then subsequently pops out the form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="f156e-1514">Parameters</span><span class="sxs-lookup"><span data-stu-id="f156e-1514">Parameters</span></span>

|<span data-ttu-id="f156e-1515">名称</span><span class="sxs-lookup"><span data-stu-id="f156e-1515">Name</span></span>|<span data-ttu-id="f156e-1516">类型</span><span class="sxs-lookup"><span data-stu-id="f156e-1516">Type</span></span>|<span data-ttu-id="f156e-1517">属性</span><span class="sxs-lookup"><span data-stu-id="f156e-1517">Attributes</span></span>|<span data-ttu-id="f156e-1518">说明</span><span class="sxs-lookup"><span data-stu-id="f156e-1518">Description</span></span>|
|---|---|---|---|
|`attachmentId`|<span data-ttu-id="f156e-1519">字符串</span><span class="sxs-lookup"><span data-stu-id="f156e-1519">String</span></span>||<span data-ttu-id="f156e-1520">要删除的附件的标识符。</span><span class="sxs-lookup"><span data-stu-id="f156e-1520">The identifier of the attachment to remove.</span></span>|
|`options`|<span data-ttu-id="f156e-1521">对象</span><span class="sxs-lookup"><span data-stu-id="f156e-1521">Object</span></span>|<span data-ttu-id="f156e-1522">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="f156e-1522">&lt;optional&gt;</span></span>|<span data-ttu-id="f156e-1523">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="f156e-1523">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="f156e-1524">对象</span><span class="sxs-lookup"><span data-stu-id="f156e-1524">Object</span></span>|<span data-ttu-id="f156e-1525">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="f156e-1525">&lt;optional&gt;</span></span>|<span data-ttu-id="f156e-1526">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="f156e-1526">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="f156e-1527">函数</span><span class="sxs-lookup"><span data-stu-id="f156e-1527">function</span></span>|<span data-ttu-id="f156e-1528">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="f156e-1528">&lt;optional&gt;</span></span>|<span data-ttu-id="f156e-1529">方法完成后，使用单个参数 `callback`（一个 [`asyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `AsyncResult` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="f156e-1529">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="f156e-1530">如果删除附件失败，`asyncResult.error` 属性将包含一个说明失败原因的错误代码。</span><span class="sxs-lookup"><span data-stu-id="f156e-1530">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="f156e-1531">错误</span><span class="sxs-lookup"><span data-stu-id="f156e-1531">Errors</span></span>

|<span data-ttu-id="f156e-1532">错误代码</span><span class="sxs-lookup"><span data-stu-id="f156e-1532">Error code</span></span>|<span data-ttu-id="f156e-1533">说明</span><span class="sxs-lookup"><span data-stu-id="f156e-1533">Description</span></span>|
|------------|-------------|
|`InvalidAttachmentId`|<span data-ttu-id="f156e-1534">附件标识符不存在。</span><span class="sxs-lookup"><span data-stu-id="f156e-1534">The attachment identifier does not exist.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f156e-1535">要求</span><span class="sxs-lookup"><span data-stu-id="f156e-1535">Requirements</span></span>

|<span data-ttu-id="f156e-1536">要求</span><span class="sxs-lookup"><span data-stu-id="f156e-1536">Requirement</span></span>|<span data-ttu-id="f156e-1537">值</span><span class="sxs-lookup"><span data-stu-id="f156e-1537">Value</span></span>|
|---|---|
|[<span data-ttu-id="f156e-1538">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="f156e-1538">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f156e-1539">1.1</span><span class="sxs-lookup"><span data-stu-id="f156e-1539">1.1</span></span>|
|[<span data-ttu-id="f156e-1540">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="f156e-1540">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f156e-1541">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="f156e-1541">ReadWriteItem</span></span>|
|[<span data-ttu-id="f156e-1542">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="f156e-1542">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="f156e-1543">撰写</span><span class="sxs-lookup"><span data-stu-id="f156e-1543">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="f156e-1544">示例</span><span class="sxs-lookup"><span data-stu-id="f156e-1544">Example</span></span>

<span data-ttu-id="f156e-1545">以下代码删除包含标识符 '0' 的附件。</span><span class="sxs-lookup"><span data-stu-id="f156e-1545">The following code removes an attachment with an identifier of '0'.</span></span>

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

#### <a name="removehandlerasynceventtype-options-callback"></a><span data-ttu-id="f156e-1546">removeHandlerAsync(eventType, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="f156e-1546">removeHandlerAsync(eventType, [options], [callback])</span></span>

<span data-ttu-id="f156e-1547">删除受支持事件类型的事件处理程序。</span><span class="sxs-lookup"><span data-stu-id="f156e-1547">Removes the event handlers for a supported event type.</span></span>

<span data-ttu-id="f156e-1548">目前，受支持的事件`Office.EventType.AttachmentsChanged`类型`Office.EventType.AppointmentTimeChanged`是`Office.EventType.EnhancedLocationsChanged`、 `Office.EventType.RecipientsChanged`、、 `Office.EventType.RecurrenceChanged`和。</span><span class="sxs-lookup"><span data-stu-id="f156e-1548">Currently the supported event types are `Office.EventType.AttachmentsChanged`, `Office.EventType.AppointmentTimeChanged`, `Office.EventType.EnhancedLocationsChanged`, `Office.EventType.RecipientsChanged`, and `Office.EventType.RecurrenceChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="f156e-1549">Parameters</span><span class="sxs-lookup"><span data-stu-id="f156e-1549">Parameters</span></span>

| <span data-ttu-id="f156e-1550">名称</span><span class="sxs-lookup"><span data-stu-id="f156e-1550">Name</span></span> | <span data-ttu-id="f156e-1551">类型</span><span class="sxs-lookup"><span data-stu-id="f156e-1551">Type</span></span> | <span data-ttu-id="f156e-1552">属性</span><span class="sxs-lookup"><span data-stu-id="f156e-1552">Attributes</span></span> | <span data-ttu-id="f156e-1553">说明</span><span class="sxs-lookup"><span data-stu-id="f156e-1553">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="f156e-1554">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="f156e-1554">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="f156e-1555">应撤销处理程序的事件。</span><span class="sxs-lookup"><span data-stu-id="f156e-1555">The event that should revoke the handler.</span></span> |
| `options` | <span data-ttu-id="f156e-1556">对象</span><span class="sxs-lookup"><span data-stu-id="f156e-1556">Object</span></span> | <span data-ttu-id="f156e-1557">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="f156e-1557">&lt;optional&gt;</span></span> | <span data-ttu-id="f156e-1558">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="f156e-1558">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="f156e-1559">对象</span><span class="sxs-lookup"><span data-stu-id="f156e-1559">Object</span></span> | <span data-ttu-id="f156e-1560">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="f156e-1560">&lt;optional&gt;</span></span> | <span data-ttu-id="f156e-1561">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="f156e-1561">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="f156e-1562">函数</span><span class="sxs-lookup"><span data-stu-id="f156e-1562">function</span></span>| <span data-ttu-id="f156e-1563">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="f156e-1563">&lt;optional&gt;</span></span>|<span data-ttu-id="f156e-1564">方法完成后，使用单个参数 `callback`（一个 [`asyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `AsyncResult` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="f156e-1564">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f156e-1565">Requirements</span><span class="sxs-lookup"><span data-stu-id="f156e-1565">Requirements</span></span>

|<span data-ttu-id="f156e-1566">要求</span><span class="sxs-lookup"><span data-stu-id="f156e-1566">Requirement</span></span>| <span data-ttu-id="f156e-1567">值</span><span class="sxs-lookup"><span data-stu-id="f156e-1567">Value</span></span>|
|---|---|
|[<span data-ttu-id="f156e-1568">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="f156e-1568">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f156e-1569">1.7</span><span class="sxs-lookup"><span data-stu-id="f156e-1569">1.7</span></span> |
|[<span data-ttu-id="f156e-1570">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="f156e-1570">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f156e-1571">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f156e-1571">ReadItem</span></span> |
|[<span data-ttu-id="f156e-1572">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="f156e-1572">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f156e-1573">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="f156e-1573">Compose or Read</span></span> |

<br>

---
---

#### <a name="saveasyncoptions-callback"></a><span data-ttu-id="f156e-1574">saveAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="f156e-1574">saveAsync([options], callback)</span></span>

<span data-ttu-id="f156e-1575">异步保存项目。</span><span class="sxs-lookup"><span data-stu-id="f156e-1575">Asynchronously saves an item.</span></span>

<span data-ttu-id="f156e-1576">调用时，此方法将当前邮件保存为草稿，并通过回调方法返回项目 ID。</span><span class="sxs-lookup"><span data-stu-id="f156e-1576">When invoked, this method saves the current message as a draft and returns the item id via the callback method.</span></span> <span data-ttu-id="f156e-1577">在 Outlook 网页版或 Outlook 联机模式下，该项目被保存到服务器中。</span><span class="sxs-lookup"><span data-stu-id="f156e-1577">In Outlook on the web or Outlook in online mode, the item is saved to the server.</span></span> <span data-ttu-id="f156e-1578">在 Outlook 缓存模式下，该项目被保存到本地缓存中。</span><span class="sxs-lookup"><span data-stu-id="f156e-1578">In Outlook in cached mode, the item is saved to the local cache.</span></span>

> [!NOTE]
> <span data-ttu-id="f156e-1579">如果加载项在撰写模式下对某个项目调用 `saveAsync` 来获得 `itemId`，以便与 EWS 或 REST API 一同使用，请注意，当 Outlook 处于高速缓存模式时，可能需要一段时间项目才能真正同步到服务器。</span><span class="sxs-lookup"><span data-stu-id="f156e-1579">If your add-in calls `saveAsync` on an item in compose mode in order to get an `itemId` to use with EWS or the REST API, be aware that when Outlook is in cached mode, it may take some time before the item is actually synced to the server.</span></span> <span data-ttu-id="f156e-1580">在项目同步前，使用 `itemId` 将返回一个错误。</span><span class="sxs-lookup"><span data-stu-id="f156e-1580">Until the item is synced, using the `itemId` will return an error.</span></span>

<span data-ttu-id="f156e-p198">由于约会没有草稿状态，如果以撰写模式在约会中调用 `saveAsync`，则该项将被保存为用户日历中的正常约会。对于之前未保存过的新约会，则不会发送邀请。保存现有约会将向添加或删除的与会者发送更新。</span><span class="sxs-lookup"><span data-stu-id="f156e-p198">Since appointments have no draft state, if `saveAsync` is called on an appointment in compose mode, the item will be saved as a normal appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent. Saving an existing appointment will send an update to added or removed attendees.</span></span>

> [!NOTE]
> <span data-ttu-id="f156e-1584">以下客户端在撰写模式下对约会上的 `saveAsync` 具有不同的行为：</span><span class="sxs-lookup"><span data-stu-id="f156e-1584">The following clients have different behavior for `saveAsync` on appointments in compose mode:</span></span>
>
> - <span data-ttu-id="f156e-1585">Mac 版 Outlook 不支持保存会议。</span><span class="sxs-lookup"><span data-stu-id="f156e-1585">Outlook on Mac does not support saving a meeting.</span></span> <span data-ttu-id="f156e-1586">在撰写模式下，无法从会议调用 `saveAsync` 方法。</span><span class="sxs-lookup"><span data-stu-id="f156e-1586">The `saveAsync` method fails when called from a meeting in compose mode.</span></span> <span data-ttu-id="f156e-1587">若需解决办法，请参阅[无法在 Outlook for Mac 中使用 Office JS API 将会议另存为草稿](https://support.microsoft.com/help/4505745)。</span><span class="sxs-lookup"><span data-stu-id="f156e-1587">See [Cannot save a meeting as a draft in Outlook for Mac by using Office JS API](https://support.microsoft.com/help/4505745) for a workaround.</span></span>
> - <span data-ttu-id="f156e-1588">在撰写模式下的约会上调用 `saveAsync` 时，Outlook 网页版始终发送邀请或更新。</span><span class="sxs-lookup"><span data-stu-id="f156e-1588">Outlook on the web always sends an invitation or update when `saveAsync` is called on an appointment in compose mode.</span></span>

##### <a name="parameters"></a><span data-ttu-id="f156e-1589">Parameters</span><span class="sxs-lookup"><span data-stu-id="f156e-1589">Parameters</span></span>

|<span data-ttu-id="f156e-1590">名称</span><span class="sxs-lookup"><span data-stu-id="f156e-1590">Name</span></span>|<span data-ttu-id="f156e-1591">类型</span><span class="sxs-lookup"><span data-stu-id="f156e-1591">Type</span></span>|<span data-ttu-id="f156e-1592">属性</span><span class="sxs-lookup"><span data-stu-id="f156e-1592">Attributes</span></span>|<span data-ttu-id="f156e-1593">说明</span><span class="sxs-lookup"><span data-stu-id="f156e-1593">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="f156e-1594">对象</span><span class="sxs-lookup"><span data-stu-id="f156e-1594">Object</span></span>|<span data-ttu-id="f156e-1595">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="f156e-1595">&lt;optional&gt;</span></span>|<span data-ttu-id="f156e-1596">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="f156e-1596">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="f156e-1597">对象</span><span class="sxs-lookup"><span data-stu-id="f156e-1597">Object</span></span>|<span data-ttu-id="f156e-1598">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="f156e-1598">&lt;optional&gt;</span></span>|<span data-ttu-id="f156e-1599">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="f156e-1599">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="f156e-1600">函数</span><span class="sxs-lookup"><span data-stu-id="f156e-1600">function</span></span>||<span data-ttu-id="f156e-1601">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="f156e-1601">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="f156e-1602">如果成功，该项目标识符将在 `asyncResult.value` 属性中提供。</span><span class="sxs-lookup"><span data-stu-id="f156e-1602">On success, the item identifier is provided in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f156e-1603">要求</span><span class="sxs-lookup"><span data-stu-id="f156e-1603">Requirements</span></span>

|<span data-ttu-id="f156e-1604">要求</span><span class="sxs-lookup"><span data-stu-id="f156e-1604">Requirement</span></span>|<span data-ttu-id="f156e-1605">值</span><span class="sxs-lookup"><span data-stu-id="f156e-1605">Value</span></span>|
|---|---|
|[<span data-ttu-id="f156e-1606">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="f156e-1606">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f156e-1607">1.3</span><span class="sxs-lookup"><span data-stu-id="f156e-1607">1.3</span></span>|
|[<span data-ttu-id="f156e-1608">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="f156e-1608">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f156e-1609">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="f156e-1609">ReadWriteItem</span></span>|
|[<span data-ttu-id="f156e-1610">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="f156e-1610">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="f156e-1611">撰写</span><span class="sxs-lookup"><span data-stu-id="f156e-1611">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="f156e-1612">示例</span><span class="sxs-lookup"><span data-stu-id="f156e-1612">Examples</span></span>

```js
Office.context.mailbox.item.saveAsync(
  function callback(result) {
    // Process the result.
  });
```

<span data-ttu-id="f156e-p200">下面是传递给回调函数的 `result` 参数的示例。`value` 属性包含的项目的项目 ID。</span><span class="sxs-lookup"><span data-stu-id="f156e-p200">The following is an example of the `result` parameter passed to the callback function. The `value` property contains the item ID of the item.</span></span>

```json
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

<br>

---
---

#### <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="f156e-1615">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="f156e-1615">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="f156e-1616">以异步方式将数据插入到邮件的正文或主题中。</span><span class="sxs-lookup"><span data-stu-id="f156e-1616">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="f156e-p201">`setSelectedDataAsync` 方法将指定字符串插入到项目主题或正文的光标位置，或者，如果在编辑器中已选择文本，则该方法将替换选择的文本。如果光标不在正文或主题字段中，则返回错误。插入之后，光标会位于插入内容的末尾。</span><span class="sxs-lookup"><span data-stu-id="f156e-p201">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="f156e-1620">参数</span><span class="sxs-lookup"><span data-stu-id="f156e-1620">Parameters</span></span>

|<span data-ttu-id="f156e-1621">名称</span><span class="sxs-lookup"><span data-stu-id="f156e-1621">Name</span></span>|<span data-ttu-id="f156e-1622">类型</span><span class="sxs-lookup"><span data-stu-id="f156e-1622">Type</span></span>|<span data-ttu-id="f156e-1623">属性</span><span class="sxs-lookup"><span data-stu-id="f156e-1623">Attributes</span></span>|<span data-ttu-id="f156e-1624">说明</span><span class="sxs-lookup"><span data-stu-id="f156e-1624">Description</span></span>|
|---|---|---|---|
|`data`|<span data-ttu-id="f156e-1625">字符串</span><span class="sxs-lookup"><span data-stu-id="f156e-1625">String</span></span>||<span data-ttu-id="f156e-p202">要插入的数据。数据不得超过 1,000,000 个字符。如果传入的数据超过 1,000,000 个字符，则会引发 `ArgumentOutOfRange` 异常。</span><span class="sxs-lookup"><span data-stu-id="f156e-p202">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`|<span data-ttu-id="f156e-1629">Object</span><span class="sxs-lookup"><span data-stu-id="f156e-1629">Object</span></span>|<span data-ttu-id="f156e-1630">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="f156e-1630">&lt;optional&gt;</span></span>|<span data-ttu-id="f156e-1631">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="f156e-1631">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="f156e-1632">对象</span><span class="sxs-lookup"><span data-stu-id="f156e-1632">Object</span></span>|<span data-ttu-id="f156e-1633">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="f156e-1633">&lt;optional&gt;</span></span>|<span data-ttu-id="f156e-1634">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="f156e-1634">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`|[<span data-ttu-id="f156e-1635">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="f156e-1635">Office.CoercionType</span></span>](office.md#coerciontype-string)|<span data-ttu-id="f156e-1636">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="f156e-1636">&lt;optional&gt;</span></span>|<span data-ttu-id="f156e-1637">如果为 `text`，系统在 Outlook 网页版和 Outlook 桌面版客户端中应用当前样式。</span><span class="sxs-lookup"><span data-stu-id="f156e-1637">If `text`, the current style is applied in Outlook on the web and desktop clients.</span></span> <span data-ttu-id="f156e-1638">如果字段是 HTML 编辑器，只会插入文本数据，即使数据为 HTML，也不例外。</span><span class="sxs-lookup"><span data-stu-id="f156e-1638">If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="f156e-1639">如果 `html` 和字段支持 HTML（主题不支持），系统在 Outlook 网页版中应用当前样式，而在 Outlook 桌面版客户端中则应用默认样式。</span><span class="sxs-lookup"><span data-stu-id="f156e-1639">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook on the web and the default style is applied in Outlook desktop clients.</span></span> <span data-ttu-id="f156e-1640">如果该字段是文本字段，则返回 `InvalidDataFormat` 错误。</span><span class="sxs-lookup"><span data-stu-id="f156e-1640">If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="f156e-1641">如果未设置 `coercionType`，则结果取决于该字段：如果该字段是 HTML，则使用 HTML；如果该字段是文本，则使用纯文本。</span><span class="sxs-lookup"><span data-stu-id="f156e-1641">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`|<span data-ttu-id="f156e-1642">function</span><span class="sxs-lookup"><span data-stu-id="f156e-1642">function</span></span>||<span data-ttu-id="f156e-1643">方法完成后，使用单个参数 `callback`（一个 [`asyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `AsyncResult` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="f156e-1643">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f156e-1644">Requirements</span><span class="sxs-lookup"><span data-stu-id="f156e-1644">Requirements</span></span>

|<span data-ttu-id="f156e-1645">要求</span><span class="sxs-lookup"><span data-stu-id="f156e-1645">Requirement</span></span>|<span data-ttu-id="f156e-1646">值</span><span class="sxs-lookup"><span data-stu-id="f156e-1646">Value</span></span>|
|---|---|
|[<span data-ttu-id="f156e-1647">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="f156e-1647">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="f156e-1648">1.2</span><span class="sxs-lookup"><span data-stu-id="f156e-1648">1.2</span></span>|
|[<span data-ttu-id="f156e-1649">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="f156e-1649">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="f156e-1650">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="f156e-1650">ReadWriteItem</span></span>|
|[<span data-ttu-id="f156e-1651">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="f156e-1651">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="f156e-1652">撰写</span><span class="sxs-lookup"><span data-stu-id="f156e-1652">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="f156e-1653">示例</span><span class="sxs-lookup"><span data-stu-id="f156e-1653">Example</span></span>

```js
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```
