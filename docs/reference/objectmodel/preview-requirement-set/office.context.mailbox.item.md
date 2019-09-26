---
title: "\"Context.subname\"-\"邮箱\"-预览要求集"
description: ''
ms.date: 09/23/2019
localization_priority: Normal
ms.openlocfilehash: 4a209ebde75a2857f4caa6d246c83adbd2cf7c10
ms.sourcegitcommit: 3c84fe6302341668c3f9f6dd64e636a97d03023c
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/26/2019
ms.locfileid: "37167373"
---
# <a name="item"></a><span data-ttu-id="9a31e-102">item</span><span class="sxs-lookup"><span data-stu-id="9a31e-102">item</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmditem"></a><span data-ttu-id="9a31e-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span><span class="sxs-lookup"><span data-stu-id="9a31e-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span></span>

<span data-ttu-id="9a31e-p101">`item` 命名空间用于访问当前选定的邮件、会议请求或约会。可以通过使用 [itemType](#itemtype-officemailboxenumsitemtype) 属性确定 `item` 的类型。</span><span class="sxs-lookup"><span data-stu-id="9a31e-p101">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="9a31e-106">Requirements</span><span class="sxs-lookup"><span data-stu-id="9a31e-106">Requirements</span></span>

|<span data-ttu-id="9a31e-107">要求</span><span class="sxs-lookup"><span data-stu-id="9a31e-107">Requirement</span></span>|<span data-ttu-id="9a31e-108">值</span><span class="sxs-lookup"><span data-stu-id="9a31e-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="9a31e-109">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="9a31e-109">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="9a31e-110">1.0</span><span class="sxs-lookup"><span data-stu-id="9a31e-110">1.0</span></span>|
|[<span data-ttu-id="9a31e-111">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="9a31e-111">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="9a31e-112">受限</span><span class="sxs-lookup"><span data-stu-id="9a31e-112">Restricted</span></span>|
|[<span data-ttu-id="9a31e-113">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="9a31e-113">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="9a31e-114">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="9a31e-114">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="9a31e-115">成员和方法</span><span class="sxs-lookup"><span data-stu-id="9a31e-115">Members and methods</span></span>

| <span data-ttu-id="9a31e-116">成员</span><span class="sxs-lookup"><span data-stu-id="9a31e-116">Member</span></span> | <span data-ttu-id="9a31e-117">类型</span><span class="sxs-lookup"><span data-stu-id="9a31e-117">Type</span></span> |
|--------|------|
| [<span data-ttu-id="9a31e-118">attachments</span><span class="sxs-lookup"><span data-stu-id="9a31e-118">attachments</span></span>](#attachments-arrayattachmentdetails) | <span data-ttu-id="9a31e-119">成员</span><span class="sxs-lookup"><span data-stu-id="9a31e-119">Member</span></span> |
| [<span data-ttu-id="9a31e-120">bcc</span><span class="sxs-lookup"><span data-stu-id="9a31e-120">bcc</span></span>](#bcc-recipients) | <span data-ttu-id="9a31e-121">成员</span><span class="sxs-lookup"><span data-stu-id="9a31e-121">Member</span></span> |
| [<span data-ttu-id="9a31e-122">body</span><span class="sxs-lookup"><span data-stu-id="9a31e-122">body</span></span>](#body-body) | <span data-ttu-id="9a31e-123">成员</span><span class="sxs-lookup"><span data-stu-id="9a31e-123">Member</span></span> |
| [<span data-ttu-id="9a31e-124">种类</span><span class="sxs-lookup"><span data-stu-id="9a31e-124">categories</span></span>](#categories-categories) | <span data-ttu-id="9a31e-125">成员</span><span class="sxs-lookup"><span data-stu-id="9a31e-125">Member</span></span> |
| [<span data-ttu-id="9a31e-126">cc</span><span class="sxs-lookup"><span data-stu-id="9a31e-126">cc</span></span>](#cc-arrayemailaddressdetailsrecipients) | <span data-ttu-id="9a31e-127">成员</span><span class="sxs-lookup"><span data-stu-id="9a31e-127">Member</span></span> |
| [<span data-ttu-id="9a31e-128">conversationId</span><span class="sxs-lookup"><span data-stu-id="9a31e-128">conversationId</span></span>](#nullable-conversationid-string) | <span data-ttu-id="9a31e-129">成员</span><span class="sxs-lookup"><span data-stu-id="9a31e-129">Member</span></span> |
| [<span data-ttu-id="9a31e-130">dateTimeCreated</span><span class="sxs-lookup"><span data-stu-id="9a31e-130">dateTimeCreated</span></span>](#datetimecreated-date) | <span data-ttu-id="9a31e-131">成员</span><span class="sxs-lookup"><span data-stu-id="9a31e-131">Member</span></span> |
| [<span data-ttu-id="9a31e-132">dateTimeModified</span><span class="sxs-lookup"><span data-stu-id="9a31e-132">dateTimeModified</span></span>](#datetimemodified-date) | <span data-ttu-id="9a31e-133">成员</span><span class="sxs-lookup"><span data-stu-id="9a31e-133">Member</span></span> |
| [<span data-ttu-id="9a31e-134">end</span><span class="sxs-lookup"><span data-stu-id="9a31e-134">end</span></span>](#end-datetime) | <span data-ttu-id="9a31e-135">成员</span><span class="sxs-lookup"><span data-stu-id="9a31e-135">Member</span></span> |
| [<span data-ttu-id="9a31e-136">enhancedLocation</span><span class="sxs-lookup"><span data-stu-id="9a31e-136">enhancedLocation</span></span>](#enhancedlocation-enhancedlocation) | <span data-ttu-id="9a31e-137">成员</span><span class="sxs-lookup"><span data-stu-id="9a31e-137">Member</span></span> |
| [<span data-ttu-id="9a31e-138">from</span><span class="sxs-lookup"><span data-stu-id="9a31e-138">from</span></span>](#from-emailaddressdetailsfrom) | <span data-ttu-id="9a31e-139">成员</span><span class="sxs-lookup"><span data-stu-id="9a31e-139">Member</span></span> |
| [<span data-ttu-id="9a31e-140">internetHeaders</span><span class="sxs-lookup"><span data-stu-id="9a31e-140">internetHeaders</span></span>](#internetheaders-internetheaders) | <span data-ttu-id="9a31e-141">成员</span><span class="sxs-lookup"><span data-stu-id="9a31e-141">Member</span></span> |
| [<span data-ttu-id="9a31e-142">internetMessageId</span><span class="sxs-lookup"><span data-stu-id="9a31e-142">internetMessageId</span></span>](#internetmessageid-string) | <span data-ttu-id="9a31e-143">成员</span><span class="sxs-lookup"><span data-stu-id="9a31e-143">Member</span></span> |
| [<span data-ttu-id="9a31e-144">itemClass</span><span class="sxs-lookup"><span data-stu-id="9a31e-144">itemClass</span></span>](#itemclass-string) | <span data-ttu-id="9a31e-145">成员</span><span class="sxs-lookup"><span data-stu-id="9a31e-145">Member</span></span> |
| [<span data-ttu-id="9a31e-146">itemId</span><span class="sxs-lookup"><span data-stu-id="9a31e-146">itemId</span></span>](#nullable-itemid-string) | <span data-ttu-id="9a31e-147">成员</span><span class="sxs-lookup"><span data-stu-id="9a31e-147">Member</span></span> |
| [<span data-ttu-id="9a31e-148">itemType</span><span class="sxs-lookup"><span data-stu-id="9a31e-148">itemType</span></span>](#itemtype-officemailboxenumsitemtype) | <span data-ttu-id="9a31e-149">成员</span><span class="sxs-lookup"><span data-stu-id="9a31e-149">Member</span></span> |
| [<span data-ttu-id="9a31e-150">location</span><span class="sxs-lookup"><span data-stu-id="9a31e-150">location</span></span>](#location-stringlocation) | <span data-ttu-id="9a31e-151">成员</span><span class="sxs-lookup"><span data-stu-id="9a31e-151">Member</span></span> |
| [<span data-ttu-id="9a31e-152">normalizedSubject</span><span class="sxs-lookup"><span data-stu-id="9a31e-152">normalizedSubject</span></span>](#normalizedsubject-string) | <span data-ttu-id="9a31e-153">成员</span><span class="sxs-lookup"><span data-stu-id="9a31e-153">Member</span></span> |
| [<span data-ttu-id="9a31e-154">notificationMessages</span><span class="sxs-lookup"><span data-stu-id="9a31e-154">notificationMessages</span></span>](#notificationmessages-notificationmessages) | <span data-ttu-id="9a31e-155">Member</span><span class="sxs-lookup"><span data-stu-id="9a31e-155">Member</span></span> |
| [<span data-ttu-id="9a31e-156">optionalAttendees</span><span class="sxs-lookup"><span data-stu-id="9a31e-156">optionalAttendees</span></span>](#optionalattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="9a31e-157">成员</span><span class="sxs-lookup"><span data-stu-id="9a31e-157">Member</span></span> |
| [<span data-ttu-id="9a31e-158">organizer</span><span class="sxs-lookup"><span data-stu-id="9a31e-158">organizer</span></span>](#organizer-emailaddressdetailsorganizer) | <span data-ttu-id="9a31e-159">成员</span><span class="sxs-lookup"><span data-stu-id="9a31e-159">Member</span></span> |
| [<span data-ttu-id="9a31e-160">recurrence</span><span class="sxs-lookup"><span data-stu-id="9a31e-160">recurrence</span></span>](#nullable-recurrence-recurrence) | <span data-ttu-id="9a31e-161">成员</span><span class="sxs-lookup"><span data-stu-id="9a31e-161">Member</span></span> |
| [<span data-ttu-id="9a31e-162">requiredAttendees</span><span class="sxs-lookup"><span data-stu-id="9a31e-162">requiredAttendees</span></span>](#requiredattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="9a31e-163">成员</span><span class="sxs-lookup"><span data-stu-id="9a31e-163">Member</span></span> |
| [<span data-ttu-id="9a31e-164">sender</span><span class="sxs-lookup"><span data-stu-id="9a31e-164">sender</span></span>](#sender-emailaddressdetails) | <span data-ttu-id="9a31e-165">Member</span><span class="sxs-lookup"><span data-stu-id="9a31e-165">Member</span></span> |
| [<span data-ttu-id="9a31e-166">Webcasts&seriesid</span><span class="sxs-lookup"><span data-stu-id="9a31e-166">seriesId</span></span>](#nullable-seriesid-string) | <span data-ttu-id="9a31e-167">Member</span><span class="sxs-lookup"><span data-stu-id="9a31e-167">Member</span></span> |
| [<span data-ttu-id="9a31e-168">start</span><span class="sxs-lookup"><span data-stu-id="9a31e-168">start</span></span>](#start-datetime) | <span data-ttu-id="9a31e-169">Member</span><span class="sxs-lookup"><span data-stu-id="9a31e-169">Member</span></span> |
| [<span data-ttu-id="9a31e-170">subject</span><span class="sxs-lookup"><span data-stu-id="9a31e-170">subject</span></span>](#subject-stringsubject) | <span data-ttu-id="9a31e-171">成员</span><span class="sxs-lookup"><span data-stu-id="9a31e-171">Member</span></span> |
| [<span data-ttu-id="9a31e-172">to</span><span class="sxs-lookup"><span data-stu-id="9a31e-172">to</span></span>](#to-arrayemailaddressdetailsrecipients) | <span data-ttu-id="9a31e-173">成员</span><span class="sxs-lookup"><span data-stu-id="9a31e-173">Member</span></span> |
| [<span data-ttu-id="9a31e-174">addFileAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="9a31e-174">addFileAttachmentAsync</span></span>](#addfileattachmentasyncuri-attachmentname-options-callback) | <span data-ttu-id="9a31e-175">方法</span><span class="sxs-lookup"><span data-stu-id="9a31e-175">Method</span></span> |
| [<span data-ttu-id="9a31e-176">addFileAttachmentFromBase64Async</span><span class="sxs-lookup"><span data-stu-id="9a31e-176">addFileAttachmentFromBase64Async</span></span>](#addfileattachmentfrombase64asyncbase64file-attachmentname-options-callback) | <span data-ttu-id="9a31e-177">方法</span><span class="sxs-lookup"><span data-stu-id="9a31e-177">Method</span></span> |
| [<span data-ttu-id="9a31e-178">addHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="9a31e-178">addHandlerAsync</span></span>](#addhandlerasynceventtype-handler-options-callback) | <span data-ttu-id="9a31e-179">方法</span><span class="sxs-lookup"><span data-stu-id="9a31e-179">Method</span></span> |
| [<span data-ttu-id="9a31e-180">addItemAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="9a31e-180">addItemAttachmentAsync</span></span>](#additemattachmentasyncitemid-attachmentname-options-callback) | <span data-ttu-id="9a31e-181">方法</span><span class="sxs-lookup"><span data-stu-id="9a31e-181">Method</span></span> |
| [<span data-ttu-id="9a31e-182">close</span><span class="sxs-lookup"><span data-stu-id="9a31e-182">close</span></span>](#close) | <span data-ttu-id="9a31e-183">方法</span><span class="sxs-lookup"><span data-stu-id="9a31e-183">Method</span></span> |
| [<span data-ttu-id="9a31e-184">displayReplyAllForm</span><span class="sxs-lookup"><span data-stu-id="9a31e-184">displayReplyAllForm</span></span>](#displayreplyallformformdata-callback) | <span data-ttu-id="9a31e-185">方法</span><span class="sxs-lookup"><span data-stu-id="9a31e-185">Method</span></span> |
| [<span data-ttu-id="9a31e-186">displayReplyForm</span><span class="sxs-lookup"><span data-stu-id="9a31e-186">displayReplyForm</span></span>](#displayreplyformformdata-callback) | <span data-ttu-id="9a31e-187">方法</span><span class="sxs-lookup"><span data-stu-id="9a31e-187">Method</span></span> |
| [<span data-ttu-id="9a31e-188">getAttachmentContentAsync</span><span class="sxs-lookup"><span data-stu-id="9a31e-188">getAttachmentContentAsync</span></span>](#getattachmentcontentasyncattachmentid-options-callback--attachmentcontent) | <span data-ttu-id="9a31e-189">方法</span><span class="sxs-lookup"><span data-stu-id="9a31e-189">Method</span></span> |
| [<span data-ttu-id="9a31e-190">getAttachmentsAsync</span><span class="sxs-lookup"><span data-stu-id="9a31e-190">getAttachmentsAsync</span></span>](#getattachmentsasyncoptions-callback--arrayattachmentdetails) | <span data-ttu-id="9a31e-191">方法</span><span class="sxs-lookup"><span data-stu-id="9a31e-191">Method</span></span> |
| [<span data-ttu-id="9a31e-192">getEntities</span><span class="sxs-lookup"><span data-stu-id="9a31e-192">getEntities</span></span>](#getentities--entities) | <span data-ttu-id="9a31e-193">方法</span><span class="sxs-lookup"><span data-stu-id="9a31e-193">Method</span></span> |
| [<span data-ttu-id="9a31e-194">getEntitiesByType</span><span class="sxs-lookup"><span data-stu-id="9a31e-194">getEntitiesByType</span></span>](#getentitiesbytypeentitytype--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="9a31e-195">方法</span><span class="sxs-lookup"><span data-stu-id="9a31e-195">Method</span></span> |
| [<span data-ttu-id="9a31e-196">getFilteredEntitiesByName</span><span class="sxs-lookup"><span data-stu-id="9a31e-196">getFilteredEntitiesByName</span></span>](#getfilteredentitiesbynamename--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="9a31e-197">方法</span><span class="sxs-lookup"><span data-stu-id="9a31e-197">Method</span></span> |
| [<span data-ttu-id="9a31e-198">Office.context.mailbox.item.getinitializationcontextasync</span><span class="sxs-lookup"><span data-stu-id="9a31e-198">getInitializationContextAsync</span></span>](#getinitializationcontextasyncoptions-callback) | <span data-ttu-id="9a31e-199">方法</span><span class="sxs-lookup"><span data-stu-id="9a31e-199">Method</span></span> |
| [<span data-ttu-id="9a31e-200">getItemIdAsync</span><span class="sxs-lookup"><span data-stu-id="9a31e-200">getItemIdAsync</span></span>](#getitemidasyncoptions-callback) | <span data-ttu-id="9a31e-201">方法</span><span class="sxs-lookup"><span data-stu-id="9a31e-201">Method</span></span> |
| [<span data-ttu-id="9a31e-202">getRegExMatches</span><span class="sxs-lookup"><span data-stu-id="9a31e-202">getRegExMatches</span></span>](#getregexmatches--object) | <span data-ttu-id="9a31e-203">方法</span><span class="sxs-lookup"><span data-stu-id="9a31e-203">Method</span></span> |
| [<span data-ttu-id="9a31e-204">getRegExMatchesByName</span><span class="sxs-lookup"><span data-stu-id="9a31e-204">getRegExMatchesByName</span></span>](#getregexmatchesbynamename--nullable-array-string-) | <span data-ttu-id="9a31e-205">方法</span><span class="sxs-lookup"><span data-stu-id="9a31e-205">Method</span></span> |
| [<span data-ttu-id="9a31e-206">getSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="9a31e-206">getSelectedDataAsync</span></span>](#getselecteddataasynccoerciontype-options-callback--string) | <span data-ttu-id="9a31e-207">方法</span><span class="sxs-lookup"><span data-stu-id="9a31e-207">Method</span></span> |
| [<span data-ttu-id="9a31e-208">Office.context.mailbox.item.getselectedentities</span><span class="sxs-lookup"><span data-stu-id="9a31e-208">getSelectedEntities</span></span>](#getselectedentities--entities) | <span data-ttu-id="9a31e-209">方法</span><span class="sxs-lookup"><span data-stu-id="9a31e-209">Method</span></span> |
| [<span data-ttu-id="9a31e-210">Office.context.mailbox.item.getselectedregexmatches</span><span class="sxs-lookup"><span data-stu-id="9a31e-210">getSelectedRegExMatches</span></span>](#getselectedregexmatches--object) | <span data-ttu-id="9a31e-211">方法</span><span class="sxs-lookup"><span data-stu-id="9a31e-211">Method</span></span> |
| [<span data-ttu-id="9a31e-212">getSharedPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="9a31e-212">getSharedPropertiesAsync</span></span>](#getsharedpropertiesasyncoptions-callback) | <span data-ttu-id="9a31e-213">方法</span><span class="sxs-lookup"><span data-stu-id="9a31e-213">Method</span></span> |
| [<span data-ttu-id="9a31e-214">loadCustomPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="9a31e-214">loadCustomPropertiesAsync</span></span>](#loadcustompropertiesasynccallback-usercontext) | <span data-ttu-id="9a31e-215">方法</span><span class="sxs-lookup"><span data-stu-id="9a31e-215">Method</span></span> |
| [<span data-ttu-id="9a31e-216">removeAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="9a31e-216">removeAttachmentAsync</span></span>](#removeattachmentasyncattachmentid-options-callback) | <span data-ttu-id="9a31e-217">方法</span><span class="sxs-lookup"><span data-stu-id="9a31e-217">Method</span></span> |
| [<span data-ttu-id="9a31e-218">removeHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="9a31e-218">removeHandlerAsync</span></span>](#removehandlerasynceventtype-options-callback) | <span data-ttu-id="9a31e-219">方法</span><span class="sxs-lookup"><span data-stu-id="9a31e-219">Method</span></span> |
| [<span data-ttu-id="9a31e-220">saveAsync</span><span class="sxs-lookup"><span data-stu-id="9a31e-220">saveAsync</span></span>](#saveasyncoptions-callback) | <span data-ttu-id="9a31e-221">方法</span><span class="sxs-lookup"><span data-stu-id="9a31e-221">Method</span></span> |
| [<span data-ttu-id="9a31e-222">setSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="9a31e-222">setSelectedDataAsync</span></span>](#setselecteddataasyncdata-options-callback) | <span data-ttu-id="9a31e-223">方法</span><span class="sxs-lookup"><span data-stu-id="9a31e-223">Method</span></span> |

### <a name="example"></a><span data-ttu-id="9a31e-224">示例</span><span class="sxs-lookup"><span data-stu-id="9a31e-224">Example</span></span>

<span data-ttu-id="9a31e-225">以下 JavaScript 代码示例显示了如何访问 Outlook 中当前项目的 `subject` 属性。</span><span class="sxs-lookup"><span data-stu-id="9a31e-225">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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

### <a name="members"></a><span data-ttu-id="9a31e-226">成员</span><span class="sxs-lookup"><span data-stu-id="9a31e-226">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetails"></a><span data-ttu-id="9a31e-227">附件： Array. <[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="9a31e-227">attachments: Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span></span>

<span data-ttu-id="9a31e-228">以数组的形式获取项目的附件。</span><span class="sxs-lookup"><span data-stu-id="9a31e-228">Gets the item's attachments as an array.</span></span> <span data-ttu-id="9a31e-229">仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="9a31e-229">Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="9a31e-230">由于潜在的安全问题，某些类型的文件会受到 Outlook 阻止，并且不会返回。</span><span class="sxs-lookup"><span data-stu-id="9a31e-230">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="9a31e-231">如需了解更多信息，请参阅 [Outlook 中阻止的附件](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519)。</span><span class="sxs-lookup"><span data-stu-id="9a31e-231">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="9a31e-232">类型</span><span class="sxs-lookup"><span data-stu-id="9a31e-232">Type</span></span>

*   <span data-ttu-id="9a31e-233">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="9a31e-233">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span></span>

##### <a name="requirements"></a><span data-ttu-id="9a31e-234">要求</span><span class="sxs-lookup"><span data-stu-id="9a31e-234">Requirements</span></span>

|<span data-ttu-id="9a31e-235">要求</span><span class="sxs-lookup"><span data-stu-id="9a31e-235">Requirement</span></span>|<span data-ttu-id="9a31e-236">值</span><span class="sxs-lookup"><span data-stu-id="9a31e-236">Value</span></span>|
|---|---|
|[<span data-ttu-id="9a31e-237">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="9a31e-237">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="9a31e-238">1.0</span><span class="sxs-lookup"><span data-stu-id="9a31e-238">1.0</span></span>|
|[<span data-ttu-id="9a31e-239">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="9a31e-239">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="9a31e-240">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9a31e-240">ReadItem</span></span>|
|[<span data-ttu-id="9a31e-241">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="9a31e-241">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="9a31e-242">阅读</span><span class="sxs-lookup"><span data-stu-id="9a31e-242">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="9a31e-243">示例</span><span class="sxs-lookup"><span data-stu-id="9a31e-243">Example</span></span>

<span data-ttu-id="9a31e-244">以下代码使用当前项目上所有附件的详细信息构成 HTML 字符串。</span><span class="sxs-lookup"><span data-stu-id="9a31e-244">The following code builds an HTML string with details of all attachments on the current item.</span></span>

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

#### <a name="bcc-recipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="9a31e-245">密件抄送：[收件人](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="9a31e-245">bcc: [Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="9a31e-246">获取对象，该对象提供用于获取或更新邮件的密件抄送 (Bcc) 行上的收件人的方法。</span><span class="sxs-lookup"><span data-stu-id="9a31e-246">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="9a31e-247">仅限撰写模式。</span><span class="sxs-lookup"><span data-stu-id="9a31e-247">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="9a31e-248">类型</span><span class="sxs-lookup"><span data-stu-id="9a31e-248">Type</span></span>

*   [<span data-ttu-id="9a31e-249">收件人</span><span class="sxs-lookup"><span data-stu-id="9a31e-249">Recipients</span></span>](/javascript/api/outlook/office.recipients)

##### <a name="requirements"></a><span data-ttu-id="9a31e-250">要求</span><span class="sxs-lookup"><span data-stu-id="9a31e-250">Requirements</span></span>

|<span data-ttu-id="9a31e-251">要求</span><span class="sxs-lookup"><span data-stu-id="9a31e-251">Requirement</span></span>|<span data-ttu-id="9a31e-252">值</span><span class="sxs-lookup"><span data-stu-id="9a31e-252">Value</span></span>|
|---|---|
|[<span data-ttu-id="9a31e-253">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="9a31e-253">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="9a31e-254">1.1</span><span class="sxs-lookup"><span data-stu-id="9a31e-254">1.1</span></span>|
|[<span data-ttu-id="9a31e-255">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="9a31e-255">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="9a31e-256">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9a31e-256">ReadItem</span></span>|
|[<span data-ttu-id="9a31e-257">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="9a31e-257">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="9a31e-258">撰写</span><span class="sxs-lookup"><span data-stu-id="9a31e-258">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="9a31e-259">示例</span><span class="sxs-lookup"><span data-stu-id="9a31e-259">Example</span></span>

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

#### <a name="body-bodyjavascriptapioutlookofficebody"></a><span data-ttu-id="9a31e-260">正文：[正文](/javascript/api/outlook/office.body)</span><span class="sxs-lookup"><span data-stu-id="9a31e-260">body: [Body](/javascript/api/outlook/office.body)</span></span>

<span data-ttu-id="9a31e-261">获取一个提供用于处理项目正文的方法的对象。</span><span class="sxs-lookup"><span data-stu-id="9a31e-261">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="9a31e-262">类型</span><span class="sxs-lookup"><span data-stu-id="9a31e-262">Type</span></span>

*   [<span data-ttu-id="9a31e-263">Body</span><span class="sxs-lookup"><span data-stu-id="9a31e-263">Body</span></span>](/javascript/api/outlook/office.body)

##### <a name="requirements"></a><span data-ttu-id="9a31e-264">要求</span><span class="sxs-lookup"><span data-stu-id="9a31e-264">Requirements</span></span>

|<span data-ttu-id="9a31e-265">要求</span><span class="sxs-lookup"><span data-stu-id="9a31e-265">Requirement</span></span>|<span data-ttu-id="9a31e-266">值</span><span class="sxs-lookup"><span data-stu-id="9a31e-266">Value</span></span>|
|---|---|
|[<span data-ttu-id="9a31e-267">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="9a31e-267">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="9a31e-268">1.1</span><span class="sxs-lookup"><span data-stu-id="9a31e-268">1.1</span></span>|
|[<span data-ttu-id="9a31e-269">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="9a31e-269">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="9a31e-270">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9a31e-270">ReadItem</span></span>|
|[<span data-ttu-id="9a31e-271">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="9a31e-271">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="9a31e-272">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="9a31e-272">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="9a31e-273">示例</span><span class="sxs-lookup"><span data-stu-id="9a31e-273">Example</span></span>

<span data-ttu-id="9a31e-274">本示例获取纯文本格式的邮件正文。</span><span class="sxs-lookup"><span data-stu-id="9a31e-274">This example gets the body of the message in plain text.</span></span>

```js
Office.context.mailbox.item.body.getAsync(
  "text",
  { asyncContext: "This is passed to the callback" },
  function callback(result) {
    // Do something with the result.
  });

```

<span data-ttu-id="9a31e-275">以下是传递到回调函数的结果参数的示例。</span><span class="sxs-lookup"><span data-stu-id="9a31e-275">The following is an example of the result parameter passed to the callback function.</span></span>

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

#### <a name="categories-categoriesjavascriptapioutlookofficecategories"></a><span data-ttu-id="9a31e-276">类别：[类别](/javascript/api/outlook/office.categories)</span><span class="sxs-lookup"><span data-stu-id="9a31e-276">categories: [Categories](/javascript/api/outlook/office.categories)</span></span>

<span data-ttu-id="9a31e-277">获取一个对象，该对象提供用于管理项的类别的方法。</span><span class="sxs-lookup"><span data-stu-id="9a31e-277">Gets an object that provides methods for managing the item's categories.</span></span>

> [!NOTE]
> <span data-ttu-id="9a31e-278">IOS 或 Android 上的 Outlook 不支持此成员。</span><span class="sxs-lookup"><span data-stu-id="9a31e-278">This member is not supported in Outlook on iOS or Android.</span></span>

##### <a name="type"></a><span data-ttu-id="9a31e-279">类型</span><span class="sxs-lookup"><span data-stu-id="9a31e-279">Type</span></span>

*   [<span data-ttu-id="9a31e-280">Categories</span><span class="sxs-lookup"><span data-stu-id="9a31e-280">Categories</span></span>](/javascript/api/outlook/office.categories)

##### <a name="requirements"></a><span data-ttu-id="9a31e-281">要求</span><span class="sxs-lookup"><span data-stu-id="9a31e-281">Requirements</span></span>

|<span data-ttu-id="9a31e-282">要求</span><span class="sxs-lookup"><span data-stu-id="9a31e-282">Requirement</span></span>|<span data-ttu-id="9a31e-283">值</span><span class="sxs-lookup"><span data-stu-id="9a31e-283">Value</span></span>|
|---|---|
|[<span data-ttu-id="9a31e-284">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="9a31e-284">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="9a31e-285">预览</span><span class="sxs-lookup"><span data-stu-id="9a31e-285">Preview</span></span>|
|[<span data-ttu-id="9a31e-286">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="9a31e-286">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="9a31e-287">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9a31e-287">ReadItem</span></span>|
|[<span data-ttu-id="9a31e-288">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="9a31e-288">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="9a31e-289">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="9a31e-289">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="9a31e-290">示例</span><span class="sxs-lookup"><span data-stu-id="9a31e-290">Example</span></span>

<span data-ttu-id="9a31e-291">此示例获取项的类别。</span><span class="sxs-lookup"><span data-stu-id="9a31e-291">This example gets the item's categories.</span></span>

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

#### <a name="cc-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="9a31e-292"><[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[收件人](/javascript/api/outlook/office.recipients)的抄送： Array</span><span class="sxs-lookup"><span data-stu-id="9a31e-292">cc: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="9a31e-293">提供对邮件的抄送 (Cc) 收件人的访问权限。</span><span class="sxs-lookup"><span data-stu-id="9a31e-293">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="9a31e-294">对象的类型和访问级别取决于当前项目的模式。</span><span class="sxs-lookup"><span data-stu-id="9a31e-294">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="9a31e-295">阅读模式</span><span class="sxs-lookup"><span data-stu-id="9a31e-295">Read mode</span></span>

<span data-ttu-id="9a31e-p106">`cc` 属性返回包含邮件的**抄送**行上所列的每个收件人的 `EmailAddressDetails` 对象的数组。集合上限为 100 个成员。</span><span class="sxs-lookup"><span data-stu-id="9a31e-p106">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.</span></span>

```js
console.log(JSON.stringify(Office.context.mailbox.item.cc));
```

##### <a name="compose-mode"></a><span data-ttu-id="9a31e-298">撰写模式</span><span class="sxs-lookup"><span data-stu-id="9a31e-298">Compose mode</span></span>

<span data-ttu-id="9a31e-299">`cc` 属性返回一个 `Recipients` 对象，该对象提供用于获取或更新邮件的“**抄送**”行上收件人的方法。</span><span class="sxs-lookup"><span data-stu-id="9a31e-299">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span>

```js
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="9a31e-300">类型</span><span class="sxs-lookup"><span data-stu-id="9a31e-300">Type</span></span>

*   <span data-ttu-id="9a31e-301">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="9a31e-301">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="9a31e-302">要求</span><span class="sxs-lookup"><span data-stu-id="9a31e-302">Requirements</span></span>

|<span data-ttu-id="9a31e-303">要求</span><span class="sxs-lookup"><span data-stu-id="9a31e-303">Requirement</span></span>|<span data-ttu-id="9a31e-304">值</span><span class="sxs-lookup"><span data-stu-id="9a31e-304">Value</span></span>|
|---|---|
|[<span data-ttu-id="9a31e-305">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="9a31e-305">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="9a31e-306">1.0</span><span class="sxs-lookup"><span data-stu-id="9a31e-306">1.0</span></span>|
|[<span data-ttu-id="9a31e-307">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="9a31e-307">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="9a31e-308">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9a31e-308">ReadItem</span></span>|
|[<span data-ttu-id="9a31e-309">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="9a31e-309">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="9a31e-310">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="9a31e-310">Compose or Read</span></span>|

<br>

---
---

#### <a name="nullable-conversationid-string"></a><span data-ttu-id="9a31e-311">（可以为 null） conversationId： String</span><span class="sxs-lookup"><span data-stu-id="9a31e-311">(nullable) conversationId: String</span></span>

<span data-ttu-id="9a31e-312">获取包含特定消息的电子邮件会话的标识符。</span><span class="sxs-lookup"><span data-stu-id="9a31e-312">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="9a31e-p107">如果在阅读窗体或撰写窗体的回复中激活邮件应用程序，则此属性可以获得一个整数值。如果用户随后更改了回复邮件的主题（若发送回复），则该邮件的对话 ID 将改变且之前获取的值将不适用。</span><span class="sxs-lookup"><span data-stu-id="9a31e-p107">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="9a31e-p108">对于撰写窗体的新项目，此属性获得一个 null 值。如果用户设置一个主题并保存该项目，`conversationId` 属性将返回一个值。</span><span class="sxs-lookup"><span data-stu-id="9a31e-p108">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="9a31e-317">Type</span><span class="sxs-lookup"><span data-stu-id="9a31e-317">Type</span></span>

*   <span data-ttu-id="9a31e-318">String</span><span class="sxs-lookup"><span data-stu-id="9a31e-318">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="9a31e-319">要求</span><span class="sxs-lookup"><span data-stu-id="9a31e-319">Requirements</span></span>

|<span data-ttu-id="9a31e-320">要求</span><span class="sxs-lookup"><span data-stu-id="9a31e-320">Requirement</span></span>|<span data-ttu-id="9a31e-321">值</span><span class="sxs-lookup"><span data-stu-id="9a31e-321">Value</span></span>|
|---|---|
|[<span data-ttu-id="9a31e-322">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="9a31e-322">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="9a31e-323">1.0</span><span class="sxs-lookup"><span data-stu-id="9a31e-323">1.0</span></span>|
|[<span data-ttu-id="9a31e-324">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="9a31e-324">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="9a31e-325">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9a31e-325">ReadItem</span></span>|
|[<span data-ttu-id="9a31e-326">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="9a31e-326">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="9a31e-327">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="9a31e-327">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="9a31e-328">示例</span><span class="sxs-lookup"><span data-stu-id="9a31e-328">Example</span></span>

```js
var conversationId = Office.context.mailbox.item.conversationId;
console.log("conversationId: " + conversationId);
```

<br>

---
---

#### <a name="datetimecreated-date"></a><span data-ttu-id="9a31e-329">dateTimeCreated： Date</span><span class="sxs-lookup"><span data-stu-id="9a31e-329">dateTimeCreated: Date</span></span>

<span data-ttu-id="9a31e-p109">获取项目创建的日期和时间。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="9a31e-p109">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="9a31e-332">类型</span><span class="sxs-lookup"><span data-stu-id="9a31e-332">Type</span></span>

*   <span data-ttu-id="9a31e-333">日期</span><span class="sxs-lookup"><span data-stu-id="9a31e-333">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="9a31e-334">要求</span><span class="sxs-lookup"><span data-stu-id="9a31e-334">Requirements</span></span>

|<span data-ttu-id="9a31e-335">要求</span><span class="sxs-lookup"><span data-stu-id="9a31e-335">Requirement</span></span>|<span data-ttu-id="9a31e-336">值</span><span class="sxs-lookup"><span data-stu-id="9a31e-336">Value</span></span>|
|---|---|
|[<span data-ttu-id="9a31e-337">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="9a31e-337">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="9a31e-338">1.0</span><span class="sxs-lookup"><span data-stu-id="9a31e-338">1.0</span></span>|
|[<span data-ttu-id="9a31e-339">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="9a31e-339">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="9a31e-340">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9a31e-340">ReadItem</span></span>|
|[<span data-ttu-id="9a31e-341">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="9a31e-341">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="9a31e-342">阅读</span><span class="sxs-lookup"><span data-stu-id="9a31e-342">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="9a31e-343">示例</span><span class="sxs-lookup"><span data-stu-id="9a31e-343">Example</span></span>

```js
var dateTimeCreated = Office.context.mailbox.item.dateTimeCreated;
console.log("Date and time created: " + dateTimeCreated);
```

<br>

---
---

#### <a name="datetimemodified-date"></a><span data-ttu-id="9a31e-344">dateTimeModified： Date</span><span class="sxs-lookup"><span data-stu-id="9a31e-344">dateTimeModified: Date</span></span>

<span data-ttu-id="9a31e-345">获取项目最近一次修改的日期和时间。</span><span class="sxs-lookup"><span data-stu-id="9a31e-345">Gets the date and time that an item was last modified.</span></span> <span data-ttu-id="9a31e-346">仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="9a31e-346">Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="9a31e-347">IOS 或 Android 上的 Outlook 不支持此成员。</span><span class="sxs-lookup"><span data-stu-id="9a31e-347">This member is not supported in Outlook on iOS or Android.</span></span>

##### <a name="type"></a><span data-ttu-id="9a31e-348">类型</span><span class="sxs-lookup"><span data-stu-id="9a31e-348">Type</span></span>

*   <span data-ttu-id="9a31e-349">日期</span><span class="sxs-lookup"><span data-stu-id="9a31e-349">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="9a31e-350">要求</span><span class="sxs-lookup"><span data-stu-id="9a31e-350">Requirements</span></span>

|<span data-ttu-id="9a31e-351">要求</span><span class="sxs-lookup"><span data-stu-id="9a31e-351">Requirement</span></span>|<span data-ttu-id="9a31e-352">值</span><span class="sxs-lookup"><span data-stu-id="9a31e-352">Value</span></span>|
|---|---|
|[<span data-ttu-id="9a31e-353">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="9a31e-353">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="9a31e-354">1.0</span><span class="sxs-lookup"><span data-stu-id="9a31e-354">1.0</span></span>|
|[<span data-ttu-id="9a31e-355">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="9a31e-355">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="9a31e-356">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9a31e-356">ReadItem</span></span>|
|[<span data-ttu-id="9a31e-357">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="9a31e-357">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="9a31e-358">阅读</span><span class="sxs-lookup"><span data-stu-id="9a31e-358">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="9a31e-359">示例</span><span class="sxs-lookup"><span data-stu-id="9a31e-359">Example</span></span>

```js
var dateTimeModified = Office.context.mailbox.item.dateTimeModified;
console.log("Date and time modified: " + dateTimeModified);
```

<br>

---
---

#### <a name="end-datetimejavascriptapioutlookofficetime"></a><span data-ttu-id="9a31e-360">结束：日期 |[时间](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="9a31e-360">end: Date|[Time](/javascript/api/outlook/office.time)</span></span>

<span data-ttu-id="9a31e-361">获取或设置约会结束的日期和时间。</span><span class="sxs-lookup"><span data-stu-id="9a31e-361">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="9a31e-p111">将 `end` 属性表示为协调世界时 (UTC) 的日期和时间值。可使用 [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) 方法将 end 属性值转换为客户端的本地日期和时间。</span><span class="sxs-lookup"><span data-stu-id="9a31e-p111">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="9a31e-364">阅读模式</span><span class="sxs-lookup"><span data-stu-id="9a31e-364">Read mode</span></span>

<span data-ttu-id="9a31e-365">`end` 属性返回 `Date` 对象。</span><span class="sxs-lookup"><span data-stu-id="9a31e-365">The `end` property returns a `Date` object.</span></span>

```js
var end = Office.context.mailbox.item.end;
console.log("Appointment end: " + end);
```

##### <a name="compose-mode"></a><span data-ttu-id="9a31e-366">撰写模式</span><span class="sxs-lookup"><span data-stu-id="9a31e-366">Compose mode</span></span>

<span data-ttu-id="9a31e-367">`end` 属性返回 `Time` 对象。</span><span class="sxs-lookup"><span data-stu-id="9a31e-367">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="9a31e-368">使用 [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) 方法设置结束时间时，应使用 [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) 方法将客户端的本地时间转换为服务器的 UTC。</span><span class="sxs-lookup"><span data-stu-id="9a31e-368">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="9a31e-369">以下示例使用 `Time` 对象的 [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) 方法设置约会的结束时间。</span><span class="sxs-lookup"><span data-stu-id="9a31e-369">The following example sets the end time of an appointment by using the [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="9a31e-370">类型</span><span class="sxs-lookup"><span data-stu-id="9a31e-370">Type</span></span>

*   <span data-ttu-id="9a31e-371">Date | [Time](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="9a31e-371">Date | [Time](/javascript/api/outlook/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="9a31e-372">要求</span><span class="sxs-lookup"><span data-stu-id="9a31e-372">Requirements</span></span>

|<span data-ttu-id="9a31e-373">要求</span><span class="sxs-lookup"><span data-stu-id="9a31e-373">Requirement</span></span>|<span data-ttu-id="9a31e-374">值</span><span class="sxs-lookup"><span data-stu-id="9a31e-374">Value</span></span>|
|---|---|
|[<span data-ttu-id="9a31e-375">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="9a31e-375">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="9a31e-376">1.0</span><span class="sxs-lookup"><span data-stu-id="9a31e-376">1.0</span></span>|
|[<span data-ttu-id="9a31e-377">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="9a31e-377">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="9a31e-378">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9a31e-378">ReadItem</span></span>|
|[<span data-ttu-id="9a31e-379">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="9a31e-379">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="9a31e-380">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="9a31e-380">Compose or Read</span></span>|

<br>

---
---

#### <a name="enhancedlocation-enhancedlocationjavascriptapioutlookofficeenhancedlocation"></a><span data-ttu-id="9a31e-381">enhancedLocation： [enhancedLocation](/javascript/api/outlook/office.enhancedlocation)</span><span class="sxs-lookup"><span data-stu-id="9a31e-381">enhancedLocation: [EnhancedLocation](/javascript/api/outlook/office.enhancedlocation)</span></span>

<span data-ttu-id="9a31e-382">获取或设置约会的位置。</span><span class="sxs-lookup"><span data-stu-id="9a31e-382">Gets or sets the locations of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="9a31e-383">阅读模式</span><span class="sxs-lookup"><span data-stu-id="9a31e-383">Read mode</span></span>

<span data-ttu-id="9a31e-384">该`enhancedLocation`属性返回一个[EnhancedLocation](/javascript/api/outlook/office.enhancedlocation)对象，该对象允许您获取与约会关联的一组位置（每个由[LocationDetails](/javascript/api/outlook/office.locationdetails)对象表示）。</span><span class="sxs-lookup"><span data-stu-id="9a31e-384">The `enhancedLocation` property returns an [EnhancedLocation](/javascript/api/outlook/office.enhancedlocation) object that allows you to get the set of locations (each represented by a [LocationDetails](/javascript/api/outlook/office.locationdetails) object) associated with the appointment.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="9a31e-385">撰写模式</span><span class="sxs-lookup"><span data-stu-id="9a31e-385">Compose mode</span></span>

<span data-ttu-id="9a31e-386">该`enhancedLocation`属性返回一个[EnhancedLocation](/javascript/api/outlook/office.enhancedlocation)对象，该对象提供用于获取、删除或添加约会位置的方法。</span><span class="sxs-lookup"><span data-stu-id="9a31e-386">The `enhancedLocation` property returns an [EnhancedLocation](/javascript/api/outlook/office.enhancedlocation) object that provides methods to get, remove, or add locations on an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="9a31e-387">类型</span><span class="sxs-lookup"><span data-stu-id="9a31e-387">Type</span></span>

*   [<span data-ttu-id="9a31e-388">EnhancedLocation</span><span class="sxs-lookup"><span data-stu-id="9a31e-388">EnhancedLocation</span></span>](/javascript/api/outlook/office.enhancedlocation)

##### <a name="requirements"></a><span data-ttu-id="9a31e-389">要求</span><span class="sxs-lookup"><span data-stu-id="9a31e-389">Requirements</span></span>

|<span data-ttu-id="9a31e-390">要求</span><span class="sxs-lookup"><span data-stu-id="9a31e-390">Requirement</span></span>|<span data-ttu-id="9a31e-391">值</span><span class="sxs-lookup"><span data-stu-id="9a31e-391">Value</span></span>|
|---|---|
|[<span data-ttu-id="9a31e-392">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="9a31e-392">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="9a31e-393">预览</span><span class="sxs-lookup"><span data-stu-id="9a31e-393">Preview</span></span>|
|[<span data-ttu-id="9a31e-394">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="9a31e-394">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="9a31e-395">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9a31e-395">ReadItem</span></span>|
|[<span data-ttu-id="9a31e-396">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="9a31e-396">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="9a31e-397">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="9a31e-397">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="9a31e-398">示例</span><span class="sxs-lookup"><span data-stu-id="9a31e-398">Example</span></span>

<span data-ttu-id="9a31e-399">下面的示例将获取与约会相关联的当前位置。</span><span class="sxs-lookup"><span data-stu-id="9a31e-399">The following example gets the current locations associated with the appointment.</span></span>

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

#### <a name="from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsfromjavascriptapioutlookofficefrom"></a><span data-ttu-id="9a31e-400">发件人： [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[From](/javascript/api/outlook/office.from)</span><span class="sxs-lookup"><span data-stu-id="9a31e-400">from: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[From](/javascript/api/outlook/office.from)</span></span>

<span data-ttu-id="9a31e-401">获取邮件发件人的电子邮件地址。</span><span class="sxs-lookup"><span data-stu-id="9a31e-401">Gets the email address of the sender of a message.</span></span>

<span data-ttu-id="9a31e-p112">`from` 和 [`sender`](#sender-emailaddressdetails) 属性表示同一个人，邮件由代理人发送的除外。在此情况下，`from` 属性表示代理程序，而 sender 属性表示代理人。</span><span class="sxs-lookup"><span data-stu-id="9a31e-p112">The `from` and [`sender`](#sender-emailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="9a31e-404">`from` 属性中 `EmailAddressDetails` 对象的 `recipientType` 属性为 `undefined`。</span><span class="sxs-lookup"><span data-stu-id="9a31e-404">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="9a31e-405">阅读模式</span><span class="sxs-lookup"><span data-stu-id="9a31e-405">Read mode</span></span>

<span data-ttu-id="9a31e-406">`from`属性返回一个`EmailAddressDetails`对象。</span><span class="sxs-lookup"><span data-stu-id="9a31e-406">The `from` property returns an `EmailAddressDetails` object.</span></span>

```js
var from = Office.context.mailbox.item.from;
console.log("From " + from);
```

##### <a name="compose-mode"></a><span data-ttu-id="9a31e-407">撰写模式</span><span class="sxs-lookup"><span data-stu-id="9a31e-407">Compose mode</span></span>

<span data-ttu-id="9a31e-408">`from`属性返回一个`From`对象，该对象提供用于获取 "起始" 值的方法。</span><span class="sxs-lookup"><span data-stu-id="9a31e-408">The `from` property returns a `From` object that provides a method to get the from value.</span></span>

```js
Office.context.mailbox.item.from.getAsync(callback);

function callback(asyncResult) {
  var from = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="9a31e-409">类型</span><span class="sxs-lookup"><span data-stu-id="9a31e-409">Type</span></span>

*   <span data-ttu-id="9a31e-410">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [](/javascript/api/outlook/office.from)</span><span class="sxs-lookup"><span data-stu-id="9a31e-410">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [From](/javascript/api/outlook/office.from)</span></span>

##### <a name="requirements"></a><span data-ttu-id="9a31e-411">要求</span><span class="sxs-lookup"><span data-stu-id="9a31e-411">Requirements</span></span>

|<span data-ttu-id="9a31e-412">要求</span><span class="sxs-lookup"><span data-stu-id="9a31e-412">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="9a31e-413">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="9a31e-413">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="9a31e-414">1.0</span><span class="sxs-lookup"><span data-stu-id="9a31e-414">1.0</span></span>|<span data-ttu-id="9a31e-415">1.7</span><span class="sxs-lookup"><span data-stu-id="9a31e-415">1.7</span></span>|
|[<span data-ttu-id="9a31e-416">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="9a31e-416">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="9a31e-417">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9a31e-417">ReadItem</span></span>|<span data-ttu-id="9a31e-418">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="9a31e-418">ReadWriteItem</span></span>|
|[<span data-ttu-id="9a31e-419">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="9a31e-419">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="9a31e-420">阅读</span><span class="sxs-lookup"><span data-stu-id="9a31e-420">Read</span></span>|<span data-ttu-id="9a31e-421">撰写</span><span class="sxs-lookup"><span data-stu-id="9a31e-421">Compose</span></span>|

<br>

---
---

#### <a name="internetheaders-internetheadersjavascriptapioutlookofficeinternetheaders"></a><span data-ttu-id="9a31e-422">internetHeaders： [internetHeaders](/javascript/api/outlook/office.internetheaders)</span><span class="sxs-lookup"><span data-stu-id="9a31e-422">internetHeaders: [InternetHeaders](/javascript/api/outlook/office.internetheaders)</span></span>

<span data-ttu-id="9a31e-423">获取或设置邮件的自定义 internet 邮件头。</span><span class="sxs-lookup"><span data-stu-id="9a31e-423">Gets or sets custom internet headers on a message.</span></span>

##### <a name="type"></a><span data-ttu-id="9a31e-424">类型</span><span class="sxs-lookup"><span data-stu-id="9a31e-424">Type</span></span>

*   [<span data-ttu-id="9a31e-425">InternetHeaders</span><span class="sxs-lookup"><span data-stu-id="9a31e-425">InternetHeaders</span></span>](/javascript/api/outlook/office.internetheaders)

##### <a name="requirements"></a><span data-ttu-id="9a31e-426">要求</span><span class="sxs-lookup"><span data-stu-id="9a31e-426">Requirements</span></span>

|<span data-ttu-id="9a31e-427">要求</span><span class="sxs-lookup"><span data-stu-id="9a31e-427">Requirement</span></span>|<span data-ttu-id="9a31e-428">值</span><span class="sxs-lookup"><span data-stu-id="9a31e-428">Value</span></span>|
|---|---|
|[<span data-ttu-id="9a31e-429">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="9a31e-429">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="9a31e-430">预览</span><span class="sxs-lookup"><span data-stu-id="9a31e-430">Preview</span></span>|
|[<span data-ttu-id="9a31e-431">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="9a31e-431">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="9a31e-432">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9a31e-432">ReadItem</span></span>|
|[<span data-ttu-id="9a31e-433">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="9a31e-433">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="9a31e-434">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="9a31e-434">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="9a31e-435">示例</span><span class="sxs-lookup"><span data-stu-id="9a31e-435">Example</span></span>

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

#### <a name="internetmessageid-string"></a><span data-ttu-id="9a31e-436">internetMessageId： String</span><span class="sxs-lookup"><span data-stu-id="9a31e-436">internetMessageId: String</span></span>

<span data-ttu-id="9a31e-p113">获取电子邮件的 Internet 消息标识符。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="9a31e-p113">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="9a31e-439">Type</span><span class="sxs-lookup"><span data-stu-id="9a31e-439">Type</span></span>

*   <span data-ttu-id="9a31e-440">String</span><span class="sxs-lookup"><span data-stu-id="9a31e-440">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="9a31e-441">要求</span><span class="sxs-lookup"><span data-stu-id="9a31e-441">Requirements</span></span>

|<span data-ttu-id="9a31e-442">要求</span><span class="sxs-lookup"><span data-stu-id="9a31e-442">Requirement</span></span>|<span data-ttu-id="9a31e-443">值</span><span class="sxs-lookup"><span data-stu-id="9a31e-443">Value</span></span>|
|---|---|
|[<span data-ttu-id="9a31e-444">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="9a31e-444">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="9a31e-445">1.0</span><span class="sxs-lookup"><span data-stu-id="9a31e-445">1.0</span></span>|
|[<span data-ttu-id="9a31e-446">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="9a31e-446">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="9a31e-447">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9a31e-447">ReadItem</span></span>|
|[<span data-ttu-id="9a31e-448">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="9a31e-448">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="9a31e-449">阅读</span><span class="sxs-lookup"><span data-stu-id="9a31e-449">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="9a31e-450">示例</span><span class="sxs-lookup"><span data-stu-id="9a31e-450">Example</span></span>

```js
var internetMessageId = Office.context.mailbox.item.internetMessageId;
console.log("internetMessageId: " + internetMessageId);
```

<br>

---
---

#### <a name="itemclass-string"></a><span data-ttu-id="9a31e-451">itemClass： String</span><span class="sxs-lookup"><span data-stu-id="9a31e-451">itemClass: String</span></span>

<span data-ttu-id="9a31e-p114">获取选定项目的 Exchange Web 服务项目类。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="9a31e-p114">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="9a31e-p115">`itemClass` 属性指定所选项目的邮件类别。以下是邮件或约会项目的默认邮件类别。</span><span class="sxs-lookup"><span data-stu-id="9a31e-p115">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

|<span data-ttu-id="9a31e-456">类型</span><span class="sxs-lookup"><span data-stu-id="9a31e-456">Type</span></span>|<span data-ttu-id="9a31e-457">说明</span><span class="sxs-lookup"><span data-stu-id="9a31e-457">Description</span></span>|<span data-ttu-id="9a31e-458">项目类</span><span class="sxs-lookup"><span data-stu-id="9a31e-458">item class</span></span>|
|---|---|---|
|<span data-ttu-id="9a31e-459">约会项目</span><span class="sxs-lookup"><span data-stu-id="9a31e-459">Appointment items</span></span>|<span data-ttu-id="9a31e-460">这些是项目类 `IPM.Appointment` 或 `IPM.Appointment.Occurrence` 的日历项目。</span><span class="sxs-lookup"><span data-stu-id="9a31e-460">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurrence`.</span></span>|`IPM.Appointment`<br />`IPM.Appointment.Occurrence`|
|<span data-ttu-id="9a31e-461">邮件项目</span><span class="sxs-lookup"><span data-stu-id="9a31e-461">Message items</span></span>|<span data-ttu-id="9a31e-462">这些项目包括具有默认邮件类别 `IPM.Note` 的电子邮件，以及使用 `IPM.Schedule.Meeting` 作为基础邮件类别的会议请求、响应和取消。</span><span class="sxs-lookup"><span data-stu-id="9a31e-462">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span>|`IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled`|

<span data-ttu-id="9a31e-463">你可以创建用于扩展默认邮件类别的自定义邮件类别，例如，自定义约会邮件类别 `IPM.Appointment.Contoso`。</span><span class="sxs-lookup"><span data-stu-id="9a31e-463">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="9a31e-464">类型</span><span class="sxs-lookup"><span data-stu-id="9a31e-464">Type</span></span>

*   <span data-ttu-id="9a31e-465">String</span><span class="sxs-lookup"><span data-stu-id="9a31e-465">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="9a31e-466">要求</span><span class="sxs-lookup"><span data-stu-id="9a31e-466">Requirements</span></span>

|<span data-ttu-id="9a31e-467">要求</span><span class="sxs-lookup"><span data-stu-id="9a31e-467">Requirement</span></span>|<span data-ttu-id="9a31e-468">值</span><span class="sxs-lookup"><span data-stu-id="9a31e-468">Value</span></span>|
|---|---|
|[<span data-ttu-id="9a31e-469">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="9a31e-469">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="9a31e-470">1.0</span><span class="sxs-lookup"><span data-stu-id="9a31e-470">1.0</span></span>|
|[<span data-ttu-id="9a31e-471">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="9a31e-471">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="9a31e-472">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9a31e-472">ReadItem</span></span>|
|[<span data-ttu-id="9a31e-473">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="9a31e-473">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="9a31e-474">阅读</span><span class="sxs-lookup"><span data-stu-id="9a31e-474">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="9a31e-475">示例</span><span class="sxs-lookup"><span data-stu-id="9a31e-475">Example</span></span>

```js
var itemClass = Office.context.mailbox.item.itemClass;
console.log("Item class: " + itemClass);
```

<br>

---
---

#### <a name="nullable-itemid-string"></a><span data-ttu-id="9a31e-476">（可以为 null） itemId： String</span><span class="sxs-lookup"><span data-stu-id="9a31e-476">(nullable) itemId: String</span></span>

<span data-ttu-id="9a31e-p116">获取当前项目的 Exchange Web 服务项目标识符。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="9a31e-p116">Gets the Exchange Web Services item identifier for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="9a31e-479">`itemId` 属性返回的标识符与 Exchange Web 服务项目标识符相同。</span><span class="sxs-lookup"><span data-stu-id="9a31e-479">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="9a31e-480">`itemId` 属性与 Outlook 条目 ID 或 Outlook REST API 使用的 ID 不同。</span><span class="sxs-lookup"><span data-stu-id="9a31e-480">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="9a31e-481">使用此值进行 REST API 调用前，应使用 [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string) 对它进行转换。</span><span class="sxs-lookup"><span data-stu-id="9a31e-481">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="9a31e-482">有关详细信息，请参阅[从 Outlook 加载项使用 Outlook REST API](/outlook/add-ins/use-rest-api#get-the-item-id)。</span><span class="sxs-lookup"><span data-stu-id="9a31e-482">For more details, see [Use the Outlook REST APIs from an Outlook add-in](/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

<span data-ttu-id="9a31e-p118">`itemId` 属性在撰写模式下不可用。如果需要项目标识符，[`saveAsync`](#saveasyncoptions-callback) 方法可用于将项目保存到存储，这将在回调函数的 [`AsyncResult.value`](/javascript/api/office/office.asyncresult) 参数中返回项目标识符。</span><span class="sxs-lookup"><span data-stu-id="9a31e-p118">The `itemId` property is not available in compose mode. If an item identifier is required, the [`saveAsync`](#saveasyncoptions-callback) method can be used to save the item to the store, which will return the item identifier in the [`AsyncResult.value`](/javascript/api/office/office.asyncresult) parameter in the callback function.</span></span>

##### <a name="type"></a><span data-ttu-id="9a31e-485">Type</span><span class="sxs-lookup"><span data-stu-id="9a31e-485">Type</span></span>

*   <span data-ttu-id="9a31e-486">String</span><span class="sxs-lookup"><span data-stu-id="9a31e-486">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="9a31e-487">要求</span><span class="sxs-lookup"><span data-stu-id="9a31e-487">Requirements</span></span>

|<span data-ttu-id="9a31e-488">要求</span><span class="sxs-lookup"><span data-stu-id="9a31e-488">Requirement</span></span>|<span data-ttu-id="9a31e-489">值</span><span class="sxs-lookup"><span data-stu-id="9a31e-489">Value</span></span>|
|---|---|
|[<span data-ttu-id="9a31e-490">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="9a31e-490">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="9a31e-491">1.0</span><span class="sxs-lookup"><span data-stu-id="9a31e-491">1.0</span></span>|
|[<span data-ttu-id="9a31e-492">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="9a31e-492">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="9a31e-493">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9a31e-493">ReadItem</span></span>|
|[<span data-ttu-id="9a31e-494">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="9a31e-494">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="9a31e-495">阅读</span><span class="sxs-lookup"><span data-stu-id="9a31e-495">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="9a31e-496">示例</span><span class="sxs-lookup"><span data-stu-id="9a31e-496">Example</span></span>

<span data-ttu-id="9a31e-p119">以下代码检查项目标识符是否存在。如果 `itemId` 属性返回 `null` 或 `undefined`，则将项目保存到存储，并从异步结果中获取项目标识符。</span><span class="sxs-lookup"><span data-stu-id="9a31e-p119">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

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

#### <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtype"></a><span data-ttu-id="9a31e-499">itemType： [MailboxEnums](/javascript/api/outlook/office.mailboxenums.itemtype)</span><span class="sxs-lookup"><span data-stu-id="9a31e-499">itemType: [Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype)</span></span>

<span data-ttu-id="9a31e-500">获取实例表示的项的类型。</span><span class="sxs-lookup"><span data-stu-id="9a31e-500">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="9a31e-501">`itemType` 属性返回其中一个 `ItemType` 枚举值，指示 `item` 对象实例是邮件还是约会。</span><span class="sxs-lookup"><span data-stu-id="9a31e-501">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="9a31e-502">类型</span><span class="sxs-lookup"><span data-stu-id="9a31e-502">Type</span></span>

*   [<span data-ttu-id="9a31e-503">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="9a31e-503">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook/office.mailboxenums.itemtype)

##### <a name="requirements"></a><span data-ttu-id="9a31e-504">要求</span><span class="sxs-lookup"><span data-stu-id="9a31e-504">Requirements</span></span>

|<span data-ttu-id="9a31e-505">要求</span><span class="sxs-lookup"><span data-stu-id="9a31e-505">Requirement</span></span>|<span data-ttu-id="9a31e-506">值</span><span class="sxs-lookup"><span data-stu-id="9a31e-506">Value</span></span>|
|---|---|
|[<span data-ttu-id="9a31e-507">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="9a31e-507">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="9a31e-508">1.0</span><span class="sxs-lookup"><span data-stu-id="9a31e-508">1.0</span></span>|
|[<span data-ttu-id="9a31e-509">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="9a31e-509">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="9a31e-510">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9a31e-510">ReadItem</span></span>|
|[<span data-ttu-id="9a31e-511">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="9a31e-511">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="9a31e-512">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="9a31e-512">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="9a31e-513">示例</span><span class="sxs-lookup"><span data-stu-id="9a31e-513">Example</span></span>

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

#### <a name="location-stringlocationjavascriptapioutlookofficelocation"></a><span data-ttu-id="9a31e-514">位置：字符串 |[位置](/javascript/api/outlook/office.location)</span><span class="sxs-lookup"><span data-stu-id="9a31e-514">location: String|[Location](/javascript/api/outlook/office.location)</span></span>

<span data-ttu-id="9a31e-515">获取或设置约会的位置。</span><span class="sxs-lookup"><span data-stu-id="9a31e-515">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="9a31e-516">阅读模式</span><span class="sxs-lookup"><span data-stu-id="9a31e-516">Read mode</span></span>

<span data-ttu-id="9a31e-517">`location` 属性返回一个包含约会位置的字符串。</span><span class="sxs-lookup"><span data-stu-id="9a31e-517">The `location` property returns a string that contains the location of the appointment.</span></span>

```js
var location = Office.context.mailbox.item.location;
console.log("location: " + location);
```

##### <a name="compose-mode"></a><span data-ttu-id="9a31e-518">撰写模式</span><span class="sxs-lookup"><span data-stu-id="9a31e-518">Compose mode</span></span>

<span data-ttu-id="9a31e-519">`location` 属性返回一个 `Location` 对象，该对象提供用于获取和设置约会位置的方法。</span><span class="sxs-lookup"><span data-stu-id="9a31e-519">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

```js
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="9a31e-520">类型</span><span class="sxs-lookup"><span data-stu-id="9a31e-520">Type</span></span>

*   <span data-ttu-id="9a31e-521">String | [Location](/javascript/api/outlook/office.location)</span><span class="sxs-lookup"><span data-stu-id="9a31e-521">String | [Location](/javascript/api/outlook/office.location)</span></span>

##### <a name="requirements"></a><span data-ttu-id="9a31e-522">要求</span><span class="sxs-lookup"><span data-stu-id="9a31e-522">Requirements</span></span>

|<span data-ttu-id="9a31e-523">要求</span><span class="sxs-lookup"><span data-stu-id="9a31e-523">Requirement</span></span>|<span data-ttu-id="9a31e-524">值</span><span class="sxs-lookup"><span data-stu-id="9a31e-524">Value</span></span>|
|---|---|
|[<span data-ttu-id="9a31e-525">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="9a31e-525">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="9a31e-526">1.0</span><span class="sxs-lookup"><span data-stu-id="9a31e-526">1.0</span></span>|
|[<span data-ttu-id="9a31e-527">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="9a31e-527">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="9a31e-528">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9a31e-528">ReadItem</span></span>|
|[<span data-ttu-id="9a31e-529">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="9a31e-529">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="9a31e-530">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="9a31e-530">Compose or Read</span></span>|

<br>

---
---

#### <a name="normalizedsubject-string"></a><span data-ttu-id="9a31e-531">normalizedSubject： String</span><span class="sxs-lookup"><span data-stu-id="9a31e-531">normalizedSubject: String</span></span>

<span data-ttu-id="9a31e-p120">获取删除了所有前缀（包括 `RE:` 和 `FWD:`）的项目主题。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="9a31e-p120">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="9a31e-p121">normalizedSubject 属性获取包含由电子邮件程序添加的任何标准前缀（如 `RE:` 和 `FW:`）的项目主题。若要获取包含完整前缀的项目主题，请使用 [`subject`](#subject-stringsubject) 属性。</span><span class="sxs-lookup"><span data-stu-id="9a31e-p121">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="9a31e-536">Type</span><span class="sxs-lookup"><span data-stu-id="9a31e-536">Type</span></span>

*   <span data-ttu-id="9a31e-537">String</span><span class="sxs-lookup"><span data-stu-id="9a31e-537">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="9a31e-538">要求</span><span class="sxs-lookup"><span data-stu-id="9a31e-538">Requirements</span></span>

|<span data-ttu-id="9a31e-539">要求</span><span class="sxs-lookup"><span data-stu-id="9a31e-539">Requirement</span></span>|<span data-ttu-id="9a31e-540">值</span><span class="sxs-lookup"><span data-stu-id="9a31e-540">Value</span></span>|
|---|---|
|[<span data-ttu-id="9a31e-541">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="9a31e-541">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="9a31e-542">1.0</span><span class="sxs-lookup"><span data-stu-id="9a31e-542">1.0</span></span>|
|[<span data-ttu-id="9a31e-543">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="9a31e-543">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="9a31e-544">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9a31e-544">ReadItem</span></span>|
|[<span data-ttu-id="9a31e-545">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="9a31e-545">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="9a31e-546">阅读</span><span class="sxs-lookup"><span data-stu-id="9a31e-546">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="9a31e-547">示例</span><span class="sxs-lookup"><span data-stu-id="9a31e-547">Example</span></span>

```js
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
console.log("Normalized subject: " + normalizedSubject);
```

<br>

---
---

#### <a name="notificationmessages-notificationmessagesjavascriptapioutlookofficenotificationmessages"></a><span data-ttu-id="9a31e-548">notificationMessages： [notificationMessages](/javascript/api/outlook/office.notificationmessages)</span><span class="sxs-lookup"><span data-stu-id="9a31e-548">notificationMessages: [NotificationMessages](/javascript/api/outlook/office.notificationmessages)</span></span>

<span data-ttu-id="9a31e-549">获取项目的通知邮件。</span><span class="sxs-lookup"><span data-stu-id="9a31e-549">Gets the notification messages for an item.</span></span>

##### <a name="type"></a><span data-ttu-id="9a31e-550">类型</span><span class="sxs-lookup"><span data-stu-id="9a31e-550">Type</span></span>

*   [<span data-ttu-id="9a31e-551">NotificationMessages</span><span class="sxs-lookup"><span data-stu-id="9a31e-551">NotificationMessages</span></span>](/javascript/api/outlook/office.notificationmessages)

##### <a name="requirements"></a><span data-ttu-id="9a31e-552">要求</span><span class="sxs-lookup"><span data-stu-id="9a31e-552">Requirements</span></span>

|<span data-ttu-id="9a31e-553">要求</span><span class="sxs-lookup"><span data-stu-id="9a31e-553">Requirement</span></span>|<span data-ttu-id="9a31e-554">值</span><span class="sxs-lookup"><span data-stu-id="9a31e-554">Value</span></span>|
|---|---|
|[<span data-ttu-id="9a31e-555">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="9a31e-555">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="9a31e-556">1.3</span><span class="sxs-lookup"><span data-stu-id="9a31e-556">1.3</span></span>|
|[<span data-ttu-id="9a31e-557">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="9a31e-557">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="9a31e-558">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9a31e-558">ReadItem</span></span>|
|[<span data-ttu-id="9a31e-559">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="9a31e-559">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="9a31e-560">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="9a31e-560">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="9a31e-561">示例</span><span class="sxs-lookup"><span data-stu-id="9a31e-561">Example</span></span>

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

#### <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="9a31e-562">optionalAttendees： <[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[收件人](/javascript/api/outlook/office.recipients)的数组</span><span class="sxs-lookup"><span data-stu-id="9a31e-562">optionalAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="9a31e-563">提供对事件的可选与会者的访问权限。</span><span class="sxs-lookup"><span data-stu-id="9a31e-563">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="9a31e-564">对象的类型和访问级别取决于当前项目的模式。</span><span class="sxs-lookup"><span data-stu-id="9a31e-564">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="9a31e-565">阅读模式</span><span class="sxs-lookup"><span data-stu-id="9a31e-565">Read mode</span></span>

<span data-ttu-id="9a31e-566">`optionalAttendees` 属性返回一个数组，其中包含每个可选与会者的 `EmailAddressDetails` 对象。</span><span class="sxs-lookup"><span data-stu-id="9a31e-566">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span>

```js
var optionalAttendees = Office.context.mailbox.item.optionalAttendees;
console.log("Optional attendees: " + JSON.stringify(optionalAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="9a31e-567">撰写模式</span><span class="sxs-lookup"><span data-stu-id="9a31e-567">Compose mode</span></span>

<span data-ttu-id="9a31e-568">`optionalAttendees` 属性返回一个 `Recipients` 对象，该对象提供用于获取或更新可选与会者的方法。</span><span class="sxs-lookup"><span data-stu-id="9a31e-568">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span>

```js
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="9a31e-569">类型</span><span class="sxs-lookup"><span data-stu-id="9a31e-569">Type</span></span>

*   <span data-ttu-id="9a31e-570">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="9a31e-570">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="9a31e-571">要求</span><span class="sxs-lookup"><span data-stu-id="9a31e-571">Requirements</span></span>

|<span data-ttu-id="9a31e-572">要求</span><span class="sxs-lookup"><span data-stu-id="9a31e-572">Requirement</span></span>|<span data-ttu-id="9a31e-573">值</span><span class="sxs-lookup"><span data-stu-id="9a31e-573">Value</span></span>|
|---|---|
|[<span data-ttu-id="9a31e-574">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="9a31e-574">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="9a31e-575">1.0</span><span class="sxs-lookup"><span data-stu-id="9a31e-575">1.0</span></span>|
|[<span data-ttu-id="9a31e-576">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="9a31e-576">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="9a31e-577">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9a31e-577">ReadItem</span></span>|
|[<span data-ttu-id="9a31e-578">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="9a31e-578">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="9a31e-579">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="9a31e-579">Compose or Read</span></span>|

<br>

---
---

#### <a name="organizer-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsorganizerjavascriptapioutlookofficeorganizer"></a><span data-ttu-id="9a31e-580">组织者： [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[组织者](/javascript/api/outlook/office.organizer)</span><span class="sxs-lookup"><span data-stu-id="9a31e-580">organizer: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[Organizer](/javascript/api/outlook/office.organizer)</span></span>

<span data-ttu-id="9a31e-581">获取指定会议的组织者的电子邮件地址。</span><span class="sxs-lookup"><span data-stu-id="9a31e-581">Gets the email address of the organizer for a specified meeting.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="9a31e-582">阅读模式</span><span class="sxs-lookup"><span data-stu-id="9a31e-582">Read mode</span></span>

<span data-ttu-id="9a31e-583">该`organizer`属性返回一个[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)对象，该对象代表会议组织者。</span><span class="sxs-lookup"><span data-stu-id="9a31e-583">The `organizer` property returns an [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) object that represents the meeting organizer.</span></span>

```js
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
console.log("Organizer: " + organizerName + " (" + organizerAddress + ")");
```

##### <a name="compose-mode"></a><span data-ttu-id="9a31e-584">撰写模式</span><span class="sxs-lookup"><span data-stu-id="9a31e-584">Compose mode</span></span>

<span data-ttu-id="9a31e-585">该`organizer`属性返回一个[管理](/javascript/api/outlook/office.organizer)器对象，该对象提供获取组织者值的方法。</span><span class="sxs-lookup"><span data-stu-id="9a31e-585">The `organizer` property returns an [Organizer](/javascript/api/outlook/office.organizer) object that provides a method to get the organizer value.</span></span>

```js
Office.context.mailbox.item.organizer.getAsync(
  function(asyncResult) {
    console.log(JSON.stringify(asyncResult));
  }
);
```

##### <a name="type"></a><span data-ttu-id="9a31e-586">类型</span><span class="sxs-lookup"><span data-stu-id="9a31e-586">Type</span></span>

*   <span data-ttu-id="9a31e-587">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [组织者](/javascript/api/outlook/office.organizer)</span><span class="sxs-lookup"><span data-stu-id="9a31e-587">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [Organizer](/javascript/api/outlook/office.organizer)</span></span>

##### <a name="requirements"></a><span data-ttu-id="9a31e-588">要求</span><span class="sxs-lookup"><span data-stu-id="9a31e-588">Requirements</span></span>

|<span data-ttu-id="9a31e-589">要求</span><span class="sxs-lookup"><span data-stu-id="9a31e-589">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="9a31e-590">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="9a31e-590">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="9a31e-591">1.0</span><span class="sxs-lookup"><span data-stu-id="9a31e-591">1.0</span></span>|<span data-ttu-id="9a31e-592">1.7</span><span class="sxs-lookup"><span data-stu-id="9a31e-592">1.7</span></span>|
|[<span data-ttu-id="9a31e-593">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="9a31e-593">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="9a31e-594">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9a31e-594">ReadItem</span></span>|<span data-ttu-id="9a31e-595">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="9a31e-595">ReadWriteItem</span></span>|
|[<span data-ttu-id="9a31e-596">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="9a31e-596">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="9a31e-597">阅读</span><span class="sxs-lookup"><span data-stu-id="9a31e-597">Read</span></span>|<span data-ttu-id="9a31e-598">撰写</span><span class="sxs-lookup"><span data-stu-id="9a31e-598">Compose</span></span>|

<br>

---
---

#### <a name="nullable-recurrence-recurrencejavascriptapioutlookofficerecurrence"></a><span data-ttu-id="9a31e-599">（可以为 null）定期：[定期](/javascript/api/outlook/office.recurrence)</span><span class="sxs-lookup"><span data-stu-id="9a31e-599">(nullable) recurrence: [Recurrence](/javascript/api/outlook/office.recurrence)</span></span>

<span data-ttu-id="9a31e-600">获取或设置约会的定期模式。</span><span class="sxs-lookup"><span data-stu-id="9a31e-600">Gets or sets the recurrence pattern of an appointment.</span></span> <span data-ttu-id="9a31e-601">获取会议请求的定期模式。</span><span class="sxs-lookup"><span data-stu-id="9a31e-601">Gets the recurrence pattern of a meeting request.</span></span> <span data-ttu-id="9a31e-602">约会项目的阅读和撰写模式。</span><span class="sxs-lookup"><span data-stu-id="9a31e-602">Read and compose modes for appointment items.</span></span> <span data-ttu-id="9a31e-603">会议请求项目的阅读模式。</span><span class="sxs-lookup"><span data-stu-id="9a31e-603">Read mode for meeting request items.</span></span>

<span data-ttu-id="9a31e-604">如果`recurrence`项目是系列中的一个系列或一个实例，则该属性返回定期约会或会议请求的[定期](/javascript/api/outlook/office.recurrence)对象。</span><span class="sxs-lookup"><span data-stu-id="9a31e-604">The `recurrence` property returns a [recurrence](/javascript/api/outlook/office.recurrence) object for recurring appointments or meetings requests if an item is a series or an instance in a series.</span></span> <span data-ttu-id="9a31e-605">`null`返回单个约会的单个约会和会议请求。</span><span class="sxs-lookup"><span data-stu-id="9a31e-605">`null` is returned for single appointments and meeting requests of single appointments.</span></span> <span data-ttu-id="9a31e-606">`undefined`对于不是会议请求的邮件，将返回。</span><span class="sxs-lookup"><span data-stu-id="9a31e-606">`undefined` is returned for messages that are not meeting requests.</span></span>

> <span data-ttu-id="9a31e-607">注意：会议请求的`itemClass`值为 IPM。Schedule. 会议请求。</span><span class="sxs-lookup"><span data-stu-id="9a31e-607">Note: Meeting requests have an `itemClass` value of IPM.Schedule.Meeting.Request.</span></span>

> <span data-ttu-id="9a31e-608">注意：如果定期对象为`null`，则表示该对象是单个约会的单个约会或会议请求，而不是某个系列的一部分。</span><span class="sxs-lookup"><span data-stu-id="9a31e-608">Note: If the recurrence object is `null`, this indicates that the object is a single appointment or a meeting request of a single appointment and NOT a part of a series.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="9a31e-609">阅读模式</span><span class="sxs-lookup"><span data-stu-id="9a31e-609">Read mode</span></span>

<span data-ttu-id="9a31e-610">该`recurrence`属性返回一个代表约会定期的[定期](/javascript/api/outlook/office.recurrence)对象。</span><span class="sxs-lookup"><span data-stu-id="9a31e-610">The `recurrence` property returns a [Recurrence](/javascript/api/outlook/office.recurrence) object that represents the appointment recurrence.</span></span> <span data-ttu-id="9a31e-611">此功能适用于约会和会议请求。</span><span class="sxs-lookup"><span data-stu-id="9a31e-611">This is available for appointments and meeting requests.</span></span>

```js
var recurrence = Office.context.mailbox.item.recurrence;
console.log("Recurrence: " + JSON.stringify(recurrence));
```

##### <a name="compose-mode"></a><span data-ttu-id="9a31e-612">撰写模式</span><span class="sxs-lookup"><span data-stu-id="9a31e-612">Compose mode</span></span>

<span data-ttu-id="9a31e-613">该`recurrence`属性返回一个[定期](/javascript/api/outlook/office.recurrence)对象，该对象提供用于管理约会周期的方法。</span><span class="sxs-lookup"><span data-stu-id="9a31e-613">The `recurrence` property returns a [Recurrence](/javascript/api/outlook/office.recurrence) object that provides methods to manage the appointment recurrence.</span></span> <span data-ttu-id="9a31e-614">这可用于约会。</span><span class="sxs-lookup"><span data-stu-id="9a31e-614">This is available for appointments.</span></span>

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

##### <a name="type"></a><span data-ttu-id="9a31e-615">类型</span><span class="sxs-lookup"><span data-stu-id="9a31e-615">Type</span></span>

* [<span data-ttu-id="9a31e-616">循环</span><span class="sxs-lookup"><span data-stu-id="9a31e-616">Recurrence</span></span>](/javascript/api/outlook/office.recurrence)

|<span data-ttu-id="9a31e-617">要求</span><span class="sxs-lookup"><span data-stu-id="9a31e-617">Requirement</span></span>|<span data-ttu-id="9a31e-618">值</span><span class="sxs-lookup"><span data-stu-id="9a31e-618">Value</span></span>|
|---|---|
|[<span data-ttu-id="9a31e-619">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="9a31e-619">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="9a31e-620">1.7</span><span class="sxs-lookup"><span data-stu-id="9a31e-620">1.7</span></span>|
|[<span data-ttu-id="9a31e-621">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="9a31e-621">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="9a31e-622">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9a31e-622">ReadItem</span></span>|
|[<span data-ttu-id="9a31e-623">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="9a31e-623">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="9a31e-624">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="9a31e-624">Compose or Read</span></span>|

<br>

---
---

#### <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="9a31e-625">requiredAttendees： <[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[收件人](/javascript/api/outlook/office.recipients)的数组</span><span class="sxs-lookup"><span data-stu-id="9a31e-625">requiredAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="9a31e-626">提供对事件的必需与会者的访问权限。</span><span class="sxs-lookup"><span data-stu-id="9a31e-626">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="9a31e-627">对象的类型和访问级别取决于当前项目的模式。</span><span class="sxs-lookup"><span data-stu-id="9a31e-627">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="9a31e-628">阅读模式</span><span class="sxs-lookup"><span data-stu-id="9a31e-628">Read mode</span></span>

<span data-ttu-id="9a31e-629">`requiredAttendees` 属性返回一个数组，其中包含每个必需与会者的 `EmailAddressDetails` 对象。</span><span class="sxs-lookup"><span data-stu-id="9a31e-629">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span>

```js
var requiredAttendees = Office.context.mailbox.item.requiredAttendees;
console.log("Required attendees: " + JSON.stringify(requiredAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="9a31e-630">撰写模式</span><span class="sxs-lookup"><span data-stu-id="9a31e-630">Compose mode</span></span>

<span data-ttu-id="9a31e-631">`requiredAttendees` 属性返回一个 `Recipients` 对象，该对象提供用于获取或更新必需与会者的方法。</span><span class="sxs-lookup"><span data-stu-id="9a31e-631">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span>

```js
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
  console.log(JSON.stringify(arrayOfRequiredAttendeesRecipients));
}
```

##### <a name="type"></a><span data-ttu-id="9a31e-632">类型</span><span class="sxs-lookup"><span data-stu-id="9a31e-632">Type</span></span>

*   <span data-ttu-id="9a31e-633">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="9a31e-633">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="9a31e-634">要求</span><span class="sxs-lookup"><span data-stu-id="9a31e-634">Requirements</span></span>

|<span data-ttu-id="9a31e-635">要求</span><span class="sxs-lookup"><span data-stu-id="9a31e-635">Requirement</span></span>|<span data-ttu-id="9a31e-636">值</span><span class="sxs-lookup"><span data-stu-id="9a31e-636">Value</span></span>|
|---|---|
|[<span data-ttu-id="9a31e-637">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="9a31e-637">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="9a31e-638">1.0</span><span class="sxs-lookup"><span data-stu-id="9a31e-638">1.0</span></span>|
|[<span data-ttu-id="9a31e-639">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="9a31e-639">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="9a31e-640">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9a31e-640">ReadItem</span></span>|
|[<span data-ttu-id="9a31e-641">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="9a31e-641">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="9a31e-642">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="9a31e-642">Compose or Read</span></span>|

<br>

---
---

#### <a name="sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetails"></a><span data-ttu-id="9a31e-643">发件人： [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="9a31e-643">sender: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)</span></span>

<span data-ttu-id="9a31e-p128">获取电子邮件发件人的电子邮件地址。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="9a31e-p128">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="9a31e-p129">[`from`](#from-emailaddressdetailsfrom) 和 `sender` 属性表示同一个人，邮件由代理人发送的除外。在此情况下，`from` 属性表示代理程序，而 sender 属性表示代理人。</span><span class="sxs-lookup"><span data-stu-id="9a31e-p129">The [`from`](#from-emailaddressdetailsfrom) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="9a31e-648">`sender` 属性中 `EmailAddressDetails` 对象的 `recipientType` 属性为 `undefined`。</span><span class="sxs-lookup"><span data-stu-id="9a31e-648">The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="9a31e-649">类型</span><span class="sxs-lookup"><span data-stu-id="9a31e-649">Type</span></span>

*   [<span data-ttu-id="9a31e-650">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="9a31e-650">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="9a31e-651">要求</span><span class="sxs-lookup"><span data-stu-id="9a31e-651">Requirements</span></span>

|<span data-ttu-id="9a31e-652">要求</span><span class="sxs-lookup"><span data-stu-id="9a31e-652">Requirement</span></span>|<span data-ttu-id="9a31e-653">值</span><span class="sxs-lookup"><span data-stu-id="9a31e-653">Value</span></span>|
|---|---|
|[<span data-ttu-id="9a31e-654">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="9a31e-654">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="9a31e-655">1.0</span><span class="sxs-lookup"><span data-stu-id="9a31e-655">1.0</span></span>|
|[<span data-ttu-id="9a31e-656">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="9a31e-656">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="9a31e-657">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9a31e-657">ReadItem</span></span>|
|[<span data-ttu-id="9a31e-658">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="9a31e-658">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="9a31e-659">阅读</span><span class="sxs-lookup"><span data-stu-id="9a31e-659">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="9a31e-660">示例</span><span class="sxs-lookup"><span data-stu-id="9a31e-660">Example</span></span>

```js
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
console.log("Sender: " + senderName + " (" + senderAddress + ")");
```

<br>

---
---

#### <a name="nullable-seriesid-string"></a><span data-ttu-id="9a31e-661">（可以为 null） Webcasts&seriesid： String</span><span class="sxs-lookup"><span data-stu-id="9a31e-661">(nullable) seriesId: String</span></span>

<span data-ttu-id="9a31e-662">获取实例所属的系列的 id。</span><span class="sxs-lookup"><span data-stu-id="9a31e-662">Gets the id of the series that an instance belongs to.</span></span>

<span data-ttu-id="9a31e-663">在 web 上的 Outlook 和桌面客户端中`seriesId` ，返回此项所属的父（系列）项的 Exchange web 服务（EWS） ID。</span><span class="sxs-lookup"><span data-stu-id="9a31e-663">In Outlook on the web and desktop clients, the `seriesId` returns the Exchange Web Services (EWS) ID of the parent (series) item that this item belongs to.</span></span> <span data-ttu-id="9a31e-664">但是，在 iOS 和 Android 中， `seriesId`将返回父项的 REST ID。</span><span class="sxs-lookup"><span data-stu-id="9a31e-664">However, in iOS and Android, the `seriesId` returns the REST ID of the parent item.</span></span>

> [!NOTE]
> <span data-ttu-id="9a31e-665">`seriesId` 属性返回的标识符与 Exchange Web 服务项目标识符相同。</span><span class="sxs-lookup"><span data-stu-id="9a31e-665">The identifier returned by the `seriesId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="9a31e-666">`seriesId`属性与 OUTLOOK REST API 使用的 outlook id 不相同。</span><span class="sxs-lookup"><span data-stu-id="9a31e-666">The `seriesId` property is not identical to the Outlook IDs used by the Outlook REST API.</span></span> <span data-ttu-id="9a31e-667">使用此值进行 REST API 调用前，应使用 [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string) 对它进行转换。</span><span class="sxs-lookup"><span data-stu-id="9a31e-667">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="9a31e-668">有关详细信息，请参阅[从 Outlook 加载项使用 Outlook REST API](/outlook/add-ins/use-rest-api)。</span><span class="sxs-lookup"><span data-stu-id="9a31e-668">For more details, see [Use the Outlook REST APIs from an Outlook add-in](/outlook/add-ins/use-rest-api).</span></span>

<span data-ttu-id="9a31e-669">对于`seriesId`不包含`null`父项（如单个约会、系列项或会议请求）的项，该属性将返回， `undefined`对于不是会议请求的任何其他项，该属性返回。</span><span class="sxs-lookup"><span data-stu-id="9a31e-669">The `seriesId` property returns `null` for items that do not have parent items such as single appointments, series items, or meeting requests and returns `undefined` for any other items that are not meeting requests.</span></span>

##### <a name="type"></a><span data-ttu-id="9a31e-670">Type</span><span class="sxs-lookup"><span data-stu-id="9a31e-670">Type</span></span>

* <span data-ttu-id="9a31e-671">String</span><span class="sxs-lookup"><span data-stu-id="9a31e-671">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="9a31e-672">要求</span><span class="sxs-lookup"><span data-stu-id="9a31e-672">Requirements</span></span>

|<span data-ttu-id="9a31e-673">要求</span><span class="sxs-lookup"><span data-stu-id="9a31e-673">Requirement</span></span>|<span data-ttu-id="9a31e-674">值</span><span class="sxs-lookup"><span data-stu-id="9a31e-674">Value</span></span>|
|---|---|
|[<span data-ttu-id="9a31e-675">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="9a31e-675">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="9a31e-676">1.7</span><span class="sxs-lookup"><span data-stu-id="9a31e-676">1.7</span></span>|
|[<span data-ttu-id="9a31e-677">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="9a31e-677">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="9a31e-678">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9a31e-678">ReadItem</span></span>|
|[<span data-ttu-id="9a31e-679">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="9a31e-679">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="9a31e-680">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="9a31e-680">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="9a31e-681">示例</span><span class="sxs-lookup"><span data-stu-id="9a31e-681">Example</span></span>

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

#### <a name="start-datetimejavascriptapioutlookofficetime"></a><span data-ttu-id="9a31e-682">开始日期：日期 |[时间](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="9a31e-682">start: Date|[Time](/javascript/api/outlook/office.time)</span></span>

<span data-ttu-id="9a31e-683">获取或设置约会开始的日期和时间。</span><span class="sxs-lookup"><span data-stu-id="9a31e-683">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="9a31e-p132">将 `start` 属性表示为协调世界时 (UTC) 的日期和时间值。可使用 [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) 方法将该值转换为客户端的本地日期和时间。</span><span class="sxs-lookup"><span data-stu-id="9a31e-p132">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="9a31e-686">阅读模式</span><span class="sxs-lookup"><span data-stu-id="9a31e-686">Read mode</span></span>

<span data-ttu-id="9a31e-687">`start` 属性返回 `Date` 对象。</span><span class="sxs-lookup"><span data-stu-id="9a31e-687">The `start` property returns a `Date` object.</span></span>

```js
var start = Office.context.mailbox.item.start;
console.log("Appointment start: " + JSON.stringify(start));
```

##### <a name="compose-mode"></a><span data-ttu-id="9a31e-688">撰写模式</span><span class="sxs-lookup"><span data-stu-id="9a31e-688">Compose mode</span></span>

<span data-ttu-id="9a31e-689">`start` 属性返回 `Time` 对象。</span><span class="sxs-lookup"><span data-stu-id="9a31e-689">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="9a31e-690">使用 [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) 方法设置开始时间时，应使用 [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) 方法将客户端的本地时间转换为服务器的 UTC。</span><span class="sxs-lookup"><span data-stu-id="9a31e-690">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="9a31e-691">以下示例通过使用 `Time` 对象的 [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) 方法，设置撰写模式下约会的开始时间。</span><span class="sxs-lookup"><span data-stu-id="9a31e-691">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="9a31e-692">类型</span><span class="sxs-lookup"><span data-stu-id="9a31e-692">Type</span></span>

*   <span data-ttu-id="9a31e-693">Date | [Time](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="9a31e-693">Date | [Time](/javascript/api/outlook/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="9a31e-694">要求</span><span class="sxs-lookup"><span data-stu-id="9a31e-694">Requirements</span></span>

|<span data-ttu-id="9a31e-695">要求</span><span class="sxs-lookup"><span data-stu-id="9a31e-695">Requirement</span></span>|<span data-ttu-id="9a31e-696">值</span><span class="sxs-lookup"><span data-stu-id="9a31e-696">Value</span></span>|
|---|---|
|[<span data-ttu-id="9a31e-697">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="9a31e-697">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="9a31e-698">1.0</span><span class="sxs-lookup"><span data-stu-id="9a31e-698">1.0</span></span>|
|[<span data-ttu-id="9a31e-699">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="9a31e-699">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="9a31e-700">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9a31e-700">ReadItem</span></span>|
|[<span data-ttu-id="9a31e-701">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="9a31e-701">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="9a31e-702">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="9a31e-702">Compose or Read</span></span>|

<br>

---
---

#### <a name="subject-stringsubjectjavascriptapioutlookofficesubject"></a><span data-ttu-id="9a31e-703">subject： String |[主题](/javascript/api/outlook/office.subject)</span><span class="sxs-lookup"><span data-stu-id="9a31e-703">subject: String|[Subject](/javascript/api/outlook/office.subject)</span></span>

<span data-ttu-id="9a31e-704">获取或设置显示在项目的主题字段中的说明。</span><span class="sxs-lookup"><span data-stu-id="9a31e-704">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="9a31e-705">`subject` 属性获取或设置由电子邮件服务器发送项目时的整个主题。</span><span class="sxs-lookup"><span data-stu-id="9a31e-705">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="9a31e-706">阅读模式</span><span class="sxs-lookup"><span data-stu-id="9a31e-706">Read mode</span></span>

<span data-ttu-id="9a31e-p133">`subject` 属性返回一个字符串。使用 [`normalizedSubject`](#normalizedsubject-string) 属性获取不带任何前导前缀（如 `RE:` 和 `FW:`）的主题。</span><span class="sxs-lookup"><span data-stu-id="9a31e-p133">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

<span data-ttu-id="9a31e-709">以下 JavaScript 代码示例显示了如何访问 Outlook 中当前项目的 `subject` 属性。</span><span class="sxs-lookup"><span data-stu-id="9a31e-709">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

```js
var subject = Office.context.mailbox.item.subject;
console.log(subject);
```

##### <a name="compose-mode"></a><span data-ttu-id="9a31e-710">撰写模式</span><span class="sxs-lookup"><span data-stu-id="9a31e-710">Compose mode</span></span>
<span data-ttu-id="9a31e-711">`subject` 属性返回一个 `Subject` 对象，该对象提供用于获取和设置主题的方法。</span><span class="sxs-lookup"><span data-stu-id="9a31e-711">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```js
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
  console.log(subject);
}
```

##### <a name="type"></a><span data-ttu-id="9a31e-712">类型</span><span class="sxs-lookup"><span data-stu-id="9a31e-712">Type</span></span>

*   <span data-ttu-id="9a31e-713">String | [Subject](/javascript/api/outlook/office.subject)</span><span class="sxs-lookup"><span data-stu-id="9a31e-713">String | [Subject](/javascript/api/outlook/office.subject)</span></span>

##### <a name="requirements"></a><span data-ttu-id="9a31e-714">要求</span><span class="sxs-lookup"><span data-stu-id="9a31e-714">Requirements</span></span>

|<span data-ttu-id="9a31e-715">要求</span><span class="sxs-lookup"><span data-stu-id="9a31e-715">Requirement</span></span>|<span data-ttu-id="9a31e-716">值</span><span class="sxs-lookup"><span data-stu-id="9a31e-716">Value</span></span>|
|---|---|
|[<span data-ttu-id="9a31e-717">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="9a31e-717">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="9a31e-718">1.0</span><span class="sxs-lookup"><span data-stu-id="9a31e-718">1.0</span></span>|
|[<span data-ttu-id="9a31e-719">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="9a31e-719">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="9a31e-720">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9a31e-720">ReadItem</span></span>|
|[<span data-ttu-id="9a31e-721">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="9a31e-721">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="9a31e-722">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="9a31e-722">Compose or Read</span></span>|

<br>

---
---

#### <a name="to-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="9a31e-723">to： <[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[收件人](/javascript/api/outlook/office.recipients)的数组</span><span class="sxs-lookup"><span data-stu-id="9a31e-723">to: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="9a31e-724">提供对邮件的“**收件人**”行上的收件人的访问权限。</span><span class="sxs-lookup"><span data-stu-id="9a31e-724">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="9a31e-725">对象的类型和访问级别取决于当前项目的模式。</span><span class="sxs-lookup"><span data-stu-id="9a31e-725">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="9a31e-726">阅读模式</span><span class="sxs-lookup"><span data-stu-id="9a31e-726">Read mode</span></span>

<span data-ttu-id="9a31e-p135">`to` 属性返回包含邮件的**收件人**行上所列的每个收件人的 `EmailAddressDetails` 对象的数组。集合上限为 100 个成员。</span><span class="sxs-lookup"><span data-stu-id="9a31e-p135">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.</span></span>

```js
console.log(JSON.stringify(Office.context.mailbox.item.to));
```

##### <a name="compose-mode"></a><span data-ttu-id="9a31e-729">撰写模式</span><span class="sxs-lookup"><span data-stu-id="9a31e-729">Compose mode</span></span>

<span data-ttu-id="9a31e-730">`to` 属性返回一个 `Recipients` 对象，该对象提供用于获取或更新邮件的“**收件人**”行上收件人的方法。</span><span class="sxs-lookup"><span data-stu-id="9a31e-730">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span>

```js
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="9a31e-731">类型</span><span class="sxs-lookup"><span data-stu-id="9a31e-731">Type</span></span>

*   <span data-ttu-id="9a31e-732">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="9a31e-732">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="9a31e-733">要求</span><span class="sxs-lookup"><span data-stu-id="9a31e-733">Requirements</span></span>

|<span data-ttu-id="9a31e-734">要求</span><span class="sxs-lookup"><span data-stu-id="9a31e-734">Requirement</span></span>|<span data-ttu-id="9a31e-735">值</span><span class="sxs-lookup"><span data-stu-id="9a31e-735">Value</span></span>|
|---|---|
|[<span data-ttu-id="9a31e-736">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="9a31e-736">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="9a31e-737">1.0</span><span class="sxs-lookup"><span data-stu-id="9a31e-737">1.0</span></span>|
|[<span data-ttu-id="9a31e-738">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="9a31e-738">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="9a31e-739">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9a31e-739">ReadItem</span></span>|
|[<span data-ttu-id="9a31e-740">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="9a31e-740">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="9a31e-741">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="9a31e-741">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="9a31e-742">方法</span><span class="sxs-lookup"><span data-stu-id="9a31e-742">Methods</span></span>

#### <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="9a31e-743">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="9a31e-743">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="9a31e-744">将文件作为附件添加到邮件或约会。</span><span class="sxs-lookup"><span data-stu-id="9a31e-744">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="9a31e-745">`addFileAttachmentAsync` 方法在指定的 URI 上载文件并将其附加到撰写窗体中的项目。</span><span class="sxs-lookup"><span data-stu-id="9a31e-745">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="9a31e-746">随后可以将该标识符与 [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) 方法一同使用，以删除同一个会话中的附件。</span><span class="sxs-lookup"><span data-stu-id="9a31e-746">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="9a31e-747">参数</span><span class="sxs-lookup"><span data-stu-id="9a31e-747">Parameters</span></span>
|<span data-ttu-id="9a31e-748">名称</span><span class="sxs-lookup"><span data-stu-id="9a31e-748">Name</span></span>|<span data-ttu-id="9a31e-749">类型</span><span class="sxs-lookup"><span data-stu-id="9a31e-749">Type</span></span>|<span data-ttu-id="9a31e-750">属性</span><span class="sxs-lookup"><span data-stu-id="9a31e-750">Attributes</span></span>|<span data-ttu-id="9a31e-751">说明</span><span class="sxs-lookup"><span data-stu-id="9a31e-751">Description</span></span>|
|---|---|---|---|
|`uri`|<span data-ttu-id="9a31e-752">String</span><span class="sxs-lookup"><span data-stu-id="9a31e-752">String</span></span>||<span data-ttu-id="9a31e-p136">提供附加到邮件或约会的文件的位置的 URI。最大长度为 2048 个字符。</span><span class="sxs-lookup"><span data-stu-id="9a31e-p136">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`|<span data-ttu-id="9a31e-755">字符串</span><span class="sxs-lookup"><span data-stu-id="9a31e-755">String</span></span>||<span data-ttu-id="9a31e-p137">在附件上载过程中显示的附件名称。最大长度为 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="9a31e-p137">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="9a31e-758">Object</span><span class="sxs-lookup"><span data-stu-id="9a31e-758">Object</span></span>|<span data-ttu-id="9a31e-759">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="9a31e-759">&lt;optional&gt;</span></span>|<span data-ttu-id="9a31e-760">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="9a31e-760">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="9a31e-761">对象</span><span class="sxs-lookup"><span data-stu-id="9a31e-761">Object</span></span>|<span data-ttu-id="9a31e-762">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="9a31e-762">&lt;optional&gt;</span></span>|<span data-ttu-id="9a31e-763">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="9a31e-763">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.isInline`|<span data-ttu-id="9a31e-764">布尔值</span><span class="sxs-lookup"><span data-stu-id="9a31e-764">Boolean</span></span>|<span data-ttu-id="9a31e-765">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="9a31e-765">&lt;optional&gt;</span></span>|<span data-ttu-id="9a31e-766">如果为 `true`，则表示附件将在邮件正文中内联显示，并且不应显示在附件列表中。</span><span class="sxs-lookup"><span data-stu-id="9a31e-766">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`callback`|<span data-ttu-id="9a31e-767">函数</span><span class="sxs-lookup"><span data-stu-id="9a31e-767">function</span></span>|<span data-ttu-id="9a31e-768">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="9a31e-768">&lt;optional&gt;</span></span>|<span data-ttu-id="9a31e-769">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="9a31e-769">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="9a31e-770">如果成功，附件标识符将在 `asyncResult.value` 属性中提供。</span><span class="sxs-lookup"><span data-stu-id="9a31e-770">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="9a31e-771">如果上传附件失败，`asyncResult` 对象将包含一个提供错误说明的 `Error` 对象。</span><span class="sxs-lookup"><span data-stu-id="9a31e-771">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="9a31e-772">错误</span><span class="sxs-lookup"><span data-stu-id="9a31e-772">Errors</span></span>

|<span data-ttu-id="9a31e-773">错误代码</span><span class="sxs-lookup"><span data-stu-id="9a31e-773">Error code</span></span>|<span data-ttu-id="9a31e-774">说明</span><span class="sxs-lookup"><span data-stu-id="9a31e-774">Description</span></span>|
|------------|-------------|
|`AttachmentSizeExceeded`|<span data-ttu-id="9a31e-775">附件大小超过了允许的大小。</span><span class="sxs-lookup"><span data-stu-id="9a31e-775">The attachment is larger than allowed.</span></span>|
|`FileTypeNotSupported`|<span data-ttu-id="9a31e-776">该附件的扩展名不是允许的扩展名。</span><span class="sxs-lookup"><span data-stu-id="9a31e-776">The attachment has an extension that is not allowed.</span></span>|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="9a31e-777">邮件或约会具有的附件过多。</span><span class="sxs-lookup"><span data-stu-id="9a31e-777">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="9a31e-778">要求</span><span class="sxs-lookup"><span data-stu-id="9a31e-778">Requirements</span></span>

|<span data-ttu-id="9a31e-779">要求</span><span class="sxs-lookup"><span data-stu-id="9a31e-779">Requirement</span></span>|<span data-ttu-id="9a31e-780">值</span><span class="sxs-lookup"><span data-stu-id="9a31e-780">Value</span></span>|
|---|---|
|[<span data-ttu-id="9a31e-781">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="9a31e-781">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="9a31e-782">1.1</span><span class="sxs-lookup"><span data-stu-id="9a31e-782">1.1</span></span>|
|[<span data-ttu-id="9a31e-783">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="9a31e-783">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="9a31e-784">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="9a31e-784">ReadWriteItem</span></span>|
|[<span data-ttu-id="9a31e-785">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="9a31e-785">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="9a31e-786">撰写</span><span class="sxs-lookup"><span data-stu-id="9a31e-786">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="9a31e-787">示例</span><span class="sxs-lookup"><span data-stu-id="9a31e-787">Examples</span></span>

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

<span data-ttu-id="9a31e-788">以下示例将图像文件添加为内联附件，并在邮件正文中引用该附件。</span><span class="sxs-lookup"><span data-stu-id="9a31e-788">The following example adds an image file as an inline attachment and references the attachment in the message body.</span></span>

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

#### <a name="addfileattachmentfrombase64asyncbase64file-attachmentname-options-callback"></a><span data-ttu-id="9a31e-789">addFileAttachmentFromBase64Async （base64File，attachmentName，[options]，[callback]）</span><span class="sxs-lookup"><span data-stu-id="9a31e-789">addFileAttachmentFromBase64Async(base64File, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="9a31e-790">将 base64 编码中的文件作为附件添加到邮件或约会中。</span><span class="sxs-lookup"><span data-stu-id="9a31e-790">Adds a file from the base64 encoding to a message or appointment as an attachment.</span></span>

<span data-ttu-id="9a31e-791">该`addFileAttachmentFromBase64Async`方法从 base64 编码中上载文件，并将其附加到撰写窗体中的项目。</span><span class="sxs-lookup"><span data-stu-id="9a31e-791">The `addFileAttachmentFromBase64Async` method uploads the file from the base64 encoding and attaches it to the item in the compose form.</span></span> <span data-ttu-id="9a31e-792">此方法返回 AsyncResult 对象中的附件标识符。</span><span class="sxs-lookup"><span data-stu-id="9a31e-792">This method returns the attachment identifier in the AsyncResult.value object.</span></span>

<span data-ttu-id="9a31e-793">随后可以将该标识符与 [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) 方法一同使用，以删除同一个会话中的附件。</span><span class="sxs-lookup"><span data-stu-id="9a31e-793">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="9a31e-794">参数</span><span class="sxs-lookup"><span data-stu-id="9a31e-794">Parameters</span></span>

|<span data-ttu-id="9a31e-795">名称</span><span class="sxs-lookup"><span data-stu-id="9a31e-795">Name</span></span>|<span data-ttu-id="9a31e-796">类型</span><span class="sxs-lookup"><span data-stu-id="9a31e-796">Type</span></span>|<span data-ttu-id="9a31e-797">属性</span><span class="sxs-lookup"><span data-stu-id="9a31e-797">Attributes</span></span>|<span data-ttu-id="9a31e-798">说明</span><span class="sxs-lookup"><span data-stu-id="9a31e-798">Description</span></span>|
|---|---|---|---|
|`base64File`|<span data-ttu-id="9a31e-799">String</span><span class="sxs-lookup"><span data-stu-id="9a31e-799">String</span></span>||<span data-ttu-id="9a31e-800">要添加到电子邮件或事件的图像或文件的 base64 编码内容。</span><span class="sxs-lookup"><span data-stu-id="9a31e-800">The base64 encoded content of an image or file to be added to an email or event.</span></span>|
|`attachmentName`|<span data-ttu-id="9a31e-801">String</span><span class="sxs-lookup"><span data-stu-id="9a31e-801">String</span></span>||<span data-ttu-id="9a31e-p139">在附件上载过程中显示的附件名称。最大长度为 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="9a31e-p139">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="9a31e-804">对象</span><span class="sxs-lookup"><span data-stu-id="9a31e-804">Object</span></span>|<span data-ttu-id="9a31e-805">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="9a31e-805">&lt;optional&gt;</span></span>|<span data-ttu-id="9a31e-806">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="9a31e-806">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="9a31e-807">对象</span><span class="sxs-lookup"><span data-stu-id="9a31e-807">Object</span></span>|<span data-ttu-id="9a31e-808">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="9a31e-808">&lt;optional&gt;</span></span>|<span data-ttu-id="9a31e-809">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="9a31e-809">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.isInline`|<span data-ttu-id="9a31e-810">布尔值</span><span class="sxs-lookup"><span data-stu-id="9a31e-810">Boolean</span></span>|<span data-ttu-id="9a31e-811">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="9a31e-811">&lt;optional&gt;</span></span>|<span data-ttu-id="9a31e-812">如果为 `true`，则表示附件将在邮件正文中内联显示，并且不应显示在附件列表中。</span><span class="sxs-lookup"><span data-stu-id="9a31e-812">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`callback`|<span data-ttu-id="9a31e-813">函数</span><span class="sxs-lookup"><span data-stu-id="9a31e-813">function</span></span>|<span data-ttu-id="9a31e-814">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="9a31e-814">&lt;optional&gt;</span></span>|<span data-ttu-id="9a31e-815">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="9a31e-815">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="9a31e-816">如果成功，附件标识符将在 `asyncResult.value` 属性中提供。</span><span class="sxs-lookup"><span data-stu-id="9a31e-816">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="9a31e-817">如果上传附件失败，`asyncResult` 对象将包含一个提供错误说明的 `Error` 对象。</span><span class="sxs-lookup"><span data-stu-id="9a31e-817">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="9a31e-818">错误</span><span class="sxs-lookup"><span data-stu-id="9a31e-818">Errors</span></span>

|<span data-ttu-id="9a31e-819">错误代码</span><span class="sxs-lookup"><span data-stu-id="9a31e-819">Error code</span></span>|<span data-ttu-id="9a31e-820">说明</span><span class="sxs-lookup"><span data-stu-id="9a31e-820">Description</span></span>|
|------------|-------------|
|`AttachmentSizeExceeded`|<span data-ttu-id="9a31e-821">附件大小超过了允许的大小。</span><span class="sxs-lookup"><span data-stu-id="9a31e-821">The attachment is larger than allowed.</span></span>|
|`FileTypeNotSupported`|<span data-ttu-id="9a31e-822">该附件的扩展名不是允许的扩展名。</span><span class="sxs-lookup"><span data-stu-id="9a31e-822">The attachment has an extension that is not allowed.</span></span>|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="9a31e-823">邮件或约会具有的附件过多。</span><span class="sxs-lookup"><span data-stu-id="9a31e-823">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="9a31e-824">要求</span><span class="sxs-lookup"><span data-stu-id="9a31e-824">Requirements</span></span>

|<span data-ttu-id="9a31e-825">要求</span><span class="sxs-lookup"><span data-stu-id="9a31e-825">Requirement</span></span>|<span data-ttu-id="9a31e-826">值</span><span class="sxs-lookup"><span data-stu-id="9a31e-826">Value</span></span>|
|---|---|
|[<span data-ttu-id="9a31e-827">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="9a31e-827">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="9a31e-828">预览</span><span class="sxs-lookup"><span data-stu-id="9a31e-828">Preview</span></span>|
|[<span data-ttu-id="9a31e-829">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="9a31e-829">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="9a31e-830">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="9a31e-830">ReadWriteItem</span></span>|
|[<span data-ttu-id="9a31e-831">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="9a31e-831">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="9a31e-832">撰写</span><span class="sxs-lookup"><span data-stu-id="9a31e-832">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="9a31e-833">示例</span><span class="sxs-lookup"><span data-stu-id="9a31e-833">Examples</span></span>

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

#### <a name="addhandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="9a31e-834">addHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="9a31e-834">addHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="9a31e-835">添加支持事件的事件处理程序。</span><span class="sxs-lookup"><span data-stu-id="9a31e-835">Adds an event handler for a supported event.</span></span>

<span data-ttu-id="9a31e-836">目前，受支持的事件`Office.EventType.AttachmentsChanged`类型`Office.EventType.AppointmentTimeChanged`是`Office.EventType.EnhancedLocationsChanged`、 `Office.EventType.RecipientsChanged`、、 `Office.EventType.RecurrenceChanged`和。</span><span class="sxs-lookup"><span data-stu-id="9a31e-836">Currently the supported event types are `Office.EventType.AttachmentsChanged`, `Office.EventType.AppointmentTimeChanged`, `Office.EventType.EnhancedLocationsChanged`, `Office.EventType.RecipientsChanged`, and `Office.EventType.RecurrenceChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="9a31e-837">参数</span><span class="sxs-lookup"><span data-stu-id="9a31e-837">Parameters</span></span>

| <span data-ttu-id="9a31e-838">名称</span><span class="sxs-lookup"><span data-stu-id="9a31e-838">Name</span></span> | <span data-ttu-id="9a31e-839">类型</span><span class="sxs-lookup"><span data-stu-id="9a31e-839">Type</span></span> | <span data-ttu-id="9a31e-840">属性</span><span class="sxs-lookup"><span data-stu-id="9a31e-840">Attributes</span></span> | <span data-ttu-id="9a31e-841">说明</span><span class="sxs-lookup"><span data-stu-id="9a31e-841">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="9a31e-842">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="9a31e-842">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="9a31e-843">应调用处理程序的事件。</span><span class="sxs-lookup"><span data-stu-id="9a31e-843">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="9a31e-844">函数</span><span class="sxs-lookup"><span data-stu-id="9a31e-844">Function</span></span> || <span data-ttu-id="9a31e-p140">用于处理事件的函数。此函数必须接受一个参数，即对象文本。参数上的 `type` 属性将匹配传递给 `addHandlerAsync` 的 `eventType` 参数。</span><span class="sxs-lookup"><span data-stu-id="9a31e-p140">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `addHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="9a31e-848">Object</span><span class="sxs-lookup"><span data-stu-id="9a31e-848">Object</span></span> | <span data-ttu-id="9a31e-849">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="9a31e-849">&lt;optional&gt;</span></span> | <span data-ttu-id="9a31e-850">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="9a31e-850">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="9a31e-851">对象</span><span class="sxs-lookup"><span data-stu-id="9a31e-851">Object</span></span> | <span data-ttu-id="9a31e-852">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="9a31e-852">&lt;optional&gt;</span></span> | <span data-ttu-id="9a31e-853">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="9a31e-853">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="9a31e-854">函数</span><span class="sxs-lookup"><span data-stu-id="9a31e-854">function</span></span>| <span data-ttu-id="9a31e-855">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="9a31e-855">&lt;optional&gt;</span></span>|<span data-ttu-id="9a31e-856">方法完成后，使用单个参数 `callback`（一个 [`asyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `AsyncResult` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="9a31e-856">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="9a31e-857">Requirements</span><span class="sxs-lookup"><span data-stu-id="9a31e-857">Requirements</span></span>

|<span data-ttu-id="9a31e-858">要求</span><span class="sxs-lookup"><span data-stu-id="9a31e-858">Requirement</span></span>| <span data-ttu-id="9a31e-859">值</span><span class="sxs-lookup"><span data-stu-id="9a31e-859">Value</span></span>|
|---|---|
|[<span data-ttu-id="9a31e-860">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="9a31e-860">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9a31e-861">1.7</span><span class="sxs-lookup"><span data-stu-id="9a31e-861">1.7</span></span> |
|[<span data-ttu-id="9a31e-862">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="9a31e-862">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9a31e-863">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9a31e-863">ReadItem</span></span> |
|[<span data-ttu-id="9a31e-864">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="9a31e-864">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="9a31e-865">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="9a31e-865">Compose or Read</span></span> |

##### <a name="example"></a><span data-ttu-id="9a31e-866">示例</span><span class="sxs-lookup"><span data-stu-id="9a31e-866">Example</span></span>

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

#### <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="9a31e-867">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="9a31e-867">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="9a31e-868">将 Exchange 项目（如邮件）作为附件添加到邮件或约会。</span><span class="sxs-lookup"><span data-stu-id="9a31e-868">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="9a31e-p141">`addItemAttachmentAsync` 方法将包含指定 Exchange 标识符的项目附加到撰写窗体中的项目。如果指定一个回调方法，此方法使用 `asyncResult` 参数调用，该参数包含一个附件标识符或代码，指示附加项目过程中出现的任何错误。可以使用 `options` 参数将状态信息传递给回调方法（如果需要）。</span><span class="sxs-lookup"><span data-stu-id="9a31e-p141">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="9a31e-872">随后可以将该标识符与 [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) 方法一同使用，以删除同一个会话中的附件。</span><span class="sxs-lookup"><span data-stu-id="9a31e-872">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="9a31e-873">如果 Office 外接程序在 web 上的 Outlook 中运行，则该`addItemAttachmentAsync`方法可以将项目附加到您正在编辑的项目之外的项目中;但是，不支持这种情况，建议不要这样做。</span><span class="sxs-lookup"><span data-stu-id="9a31e-873">If your Office Add-in is running in Outlook on the web, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="9a31e-874">参数</span><span class="sxs-lookup"><span data-stu-id="9a31e-874">Parameters</span></span>

|<span data-ttu-id="9a31e-875">名称</span><span class="sxs-lookup"><span data-stu-id="9a31e-875">Name</span></span>|<span data-ttu-id="9a31e-876">类型</span><span class="sxs-lookup"><span data-stu-id="9a31e-876">Type</span></span>|<span data-ttu-id="9a31e-877">属性</span><span class="sxs-lookup"><span data-stu-id="9a31e-877">Attributes</span></span>|<span data-ttu-id="9a31e-878">说明</span><span class="sxs-lookup"><span data-stu-id="9a31e-878">Description</span></span>|
|---|---|---|---|
|`itemId`|<span data-ttu-id="9a31e-879">字符串</span><span class="sxs-lookup"><span data-stu-id="9a31e-879">String</span></span>||<span data-ttu-id="9a31e-p142">要附加的项目的 Exchange 标识符。最大长度为 100 个字符。</span><span class="sxs-lookup"><span data-stu-id="9a31e-p142">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`|<span data-ttu-id="9a31e-882">String</span><span class="sxs-lookup"><span data-stu-id="9a31e-882">String</span></span>||<span data-ttu-id="9a31e-883">要附加的项目的主题。</span><span class="sxs-lookup"><span data-stu-id="9a31e-883">The subject of the item to be attached.</span></span> <span data-ttu-id="9a31e-884">最大长度为 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="9a31e-884">The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="9a31e-885">对象</span><span class="sxs-lookup"><span data-stu-id="9a31e-885">Object</span></span>|<span data-ttu-id="9a31e-886">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="9a31e-886">&lt;optional&gt;</span></span>|<span data-ttu-id="9a31e-887">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="9a31e-887">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="9a31e-888">对象</span><span class="sxs-lookup"><span data-stu-id="9a31e-888">Object</span></span>|<span data-ttu-id="9a31e-889">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="9a31e-889">&lt;optional&gt;</span></span>|<span data-ttu-id="9a31e-890">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="9a31e-890">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="9a31e-891">函数</span><span class="sxs-lookup"><span data-stu-id="9a31e-891">function</span></span>|<span data-ttu-id="9a31e-892">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="9a31e-892">&lt;optional&gt;</span></span>|<span data-ttu-id="9a31e-893">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="9a31e-893">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="9a31e-894">如果成功，附件标识符将在 `asyncResult.value` 属性中提供。</span><span class="sxs-lookup"><span data-stu-id="9a31e-894">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="9a31e-895">如果添加附件失败，`asyncResult` 对象将包含一个提供错误说明的 `Error` 对象。</span><span class="sxs-lookup"><span data-stu-id="9a31e-895">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="9a31e-896">错误</span><span class="sxs-lookup"><span data-stu-id="9a31e-896">Errors</span></span>

|<span data-ttu-id="9a31e-897">错误代码</span><span class="sxs-lookup"><span data-stu-id="9a31e-897">Error code</span></span>|<span data-ttu-id="9a31e-898">说明</span><span class="sxs-lookup"><span data-stu-id="9a31e-898">Description</span></span>|
|------------|-------------|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="9a31e-899">邮件或约会具有的附件过多。</span><span class="sxs-lookup"><span data-stu-id="9a31e-899">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="9a31e-900">要求</span><span class="sxs-lookup"><span data-stu-id="9a31e-900">Requirements</span></span>

|<span data-ttu-id="9a31e-901">要求</span><span class="sxs-lookup"><span data-stu-id="9a31e-901">Requirement</span></span>|<span data-ttu-id="9a31e-902">值</span><span class="sxs-lookup"><span data-stu-id="9a31e-902">Value</span></span>|
|---|---|
|[<span data-ttu-id="9a31e-903">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="9a31e-903">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="9a31e-904">1.1</span><span class="sxs-lookup"><span data-stu-id="9a31e-904">1.1</span></span>|
|[<span data-ttu-id="9a31e-905">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="9a31e-905">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="9a31e-906">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="9a31e-906">ReadWriteItem</span></span>|
|[<span data-ttu-id="9a31e-907">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="9a31e-907">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="9a31e-908">撰写</span><span class="sxs-lookup"><span data-stu-id="9a31e-908">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="9a31e-909">示例</span><span class="sxs-lookup"><span data-stu-id="9a31e-909">Example</span></span>

<span data-ttu-id="9a31e-910">以下示例将现有的 Outlook 项目添加为名为 `My Attachment` 的附件。</span><span class="sxs-lookup"><span data-stu-id="9a31e-910">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

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

#### <a name="close"></a><span data-ttu-id="9a31e-911">close()</span><span class="sxs-lookup"><span data-stu-id="9a31e-911">close()</span></span>

<span data-ttu-id="9a31e-912">关闭当前正在撰写的项目。</span><span class="sxs-lookup"><span data-stu-id="9a31e-912">Closes the current item that is being composed.</span></span>

<span data-ttu-id="9a31e-p144">
            \`close\` 方法的行为取决于要撰写的项目的当前状态。如果项目具有未保存的更改，客户端将提示用户保存、放弃或取消关闭操作。</span><span class="sxs-lookup"><span data-stu-id="9a31e-p144">The behavior of the `close` method depends on the current state of the item being composed. If the item has unsaved changes, the client prompts the user to save, discard, or cancel the close action.</span></span>

> [!NOTE]
> <span data-ttu-id="9a31e-915">在 Outlook 网页版中，如果该项目是约会并且之前已使用 `saveAsync` 保存，则即使自上次保存项目后未发生任何更改，也会提示用户保存、放弃或取消。</span><span class="sxs-lookup"><span data-stu-id="9a31e-915">In Outlook on the web, if the item is an appointment and it has previously been saved using `saveAsync`, the user is prompted to save, discard, or cancel even if no changes have occurred since the item was last saved.</span></span>

<span data-ttu-id="9a31e-916">在 Outlook 桌面客户端中，如果邮件是内联答复，`close` 方法不起作用。</span><span class="sxs-lookup"><span data-stu-id="9a31e-916">In the Outlook desktop client, if the message is an inline reply, the `close` method has no effect.</span></span>

##### <a name="requirements"></a><span data-ttu-id="9a31e-917">要求</span><span class="sxs-lookup"><span data-stu-id="9a31e-917">Requirements</span></span>

|<span data-ttu-id="9a31e-918">要求</span><span class="sxs-lookup"><span data-stu-id="9a31e-918">Requirement</span></span>|<span data-ttu-id="9a31e-919">值</span><span class="sxs-lookup"><span data-stu-id="9a31e-919">Value</span></span>|
|---|---|
|[<span data-ttu-id="9a31e-920">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="9a31e-920">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="9a31e-921">1.3</span><span class="sxs-lookup"><span data-stu-id="9a31e-921">1.3</span></span>|
|[<span data-ttu-id="9a31e-922">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="9a31e-922">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="9a31e-923">受限</span><span class="sxs-lookup"><span data-stu-id="9a31e-923">Restricted</span></span>|
|[<span data-ttu-id="9a31e-924">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="9a31e-924">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="9a31e-925">撰写</span><span class="sxs-lookup"><span data-stu-id="9a31e-925">Compose</span></span>|

<br>

---
---

#### <a name="displayreplyallformformdata-callback"></a><span data-ttu-id="9a31e-926">displayReplyAllForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="9a31e-926">displayReplyAllForm(formData, [callback])</span></span>

<span data-ttu-id="9a31e-927">显示答复窗体，其中包括所选邮件的发件人和所有收件人或所选约会的组织者和所有与会者。</span><span class="sxs-lookup"><span data-stu-id="9a31e-927">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="9a31e-928">IOS 或 Android 上的 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="9a31e-928">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="9a31e-929">在 web 上的 Outlook 中，答复窗体显示为3列视图中的弹出窗体和2列或1列视图中的弹出式窗体。</span><span class="sxs-lookup"><span data-stu-id="9a31e-929">In Outlook on the web, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="9a31e-930">如果任意字符串参数超出其限制，`displayReplyAllForm` 将引发异常。</span><span class="sxs-lookup"><span data-stu-id="9a31e-930">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="9a31e-931">如果在`formData.attachments`参数中指定了附件，则 web 上的 Outlook 和桌面客户端将尝试下载所有附件并将其附加到答复窗体。</span><span class="sxs-lookup"><span data-stu-id="9a31e-931">When attachments are specified in the `formData.attachments` parameter, Outlook on the web and desktop clients attempt to download all attachments and attach them to the reply form.</span></span> <span data-ttu-id="9a31e-932">如果无法添加任何附件，则在窗体 UI 中显示错误。</span><span class="sxs-lookup"><span data-stu-id="9a31e-932">If any attachments fail to be added, an error is shown in the form UI.</span></span> <span data-ttu-id="9a31e-933">如果这不可能，则不引发错误消息。</span><span class="sxs-lookup"><span data-stu-id="9a31e-933">If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="9a31e-934">参数</span><span class="sxs-lookup"><span data-stu-id="9a31e-934">Parameters</span></span>

|<span data-ttu-id="9a31e-935">名称</span><span class="sxs-lookup"><span data-stu-id="9a31e-935">Name</span></span>|<span data-ttu-id="9a31e-936">类型</span><span class="sxs-lookup"><span data-stu-id="9a31e-936">Type</span></span>|<span data-ttu-id="9a31e-937">属性</span><span class="sxs-lookup"><span data-stu-id="9a31e-937">Attributes</span></span>|<span data-ttu-id="9a31e-938">说明</span><span class="sxs-lookup"><span data-stu-id="9a31e-938">Description</span></span>|
|---|---|---|---|
|`formData`|<span data-ttu-id="9a31e-939">字符串 &#124; 对象</span><span class="sxs-lookup"><span data-stu-id="9a31e-939">String &#124; Object</span></span>||<span data-ttu-id="9a31e-p146">一个包含文本和 HTML 且表示答复窗体的正文的字符串。字符串限制为 32 KB。</span><span class="sxs-lookup"><span data-stu-id="9a31e-p146">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="9a31e-942">**或**</span><span class="sxs-lookup"><span data-stu-id="9a31e-942">**OR**</span></span><br/><span data-ttu-id="9a31e-p147">包含正文或附件数据和回调函数的对象。对象定义如下。</span><span class="sxs-lookup"><span data-stu-id="9a31e-p147">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span>|
|`formData.htmlBody`|<span data-ttu-id="9a31e-945">String</span><span class="sxs-lookup"><span data-stu-id="9a31e-945">String</span></span>|<span data-ttu-id="9a31e-946">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="9a31e-946">&lt;optional&gt;</span></span>|<span data-ttu-id="9a31e-p148">一个包含文本和 HTML 且表示答复窗体的正文的字符串。字符串限制为 32 KB。</span><span class="sxs-lookup"><span data-stu-id="9a31e-p148">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
|`formData.attachments`|<span data-ttu-id="9a31e-949">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="9a31e-949">Array.&lt;Object&gt;</span></span>|<span data-ttu-id="9a31e-950">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="9a31e-950">&lt;optional&gt;</span></span>|<span data-ttu-id="9a31e-951">JSON 对象（是文件或项目附件）的数组。</span><span class="sxs-lookup"><span data-stu-id="9a31e-951">An array of JSON objects that are either file or item attachments.</span></span>|
|`formData.attachments.type`|<span data-ttu-id="9a31e-952">String</span><span class="sxs-lookup"><span data-stu-id="9a31e-952">String</span></span>||<span data-ttu-id="9a31e-p149">指示附件的类型。必须是文件附件的 `file` 或项目附件的 `item`。</span><span class="sxs-lookup"><span data-stu-id="9a31e-p149">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span>|
|`formData.attachments.name`|<span data-ttu-id="9a31e-955">字符串</span><span class="sxs-lookup"><span data-stu-id="9a31e-955">String</span></span>||<span data-ttu-id="9a31e-956">一个包含附件的名称的字符串，最多包含 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="9a31e-956">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
|`formData.attachments.url`|<span data-ttu-id="9a31e-957">字符串</span><span class="sxs-lookup"><span data-stu-id="9a31e-957">String</span></span>||<span data-ttu-id="9a31e-p150">仅在将 `type` 设置为 `file` 时使用。文件的位置的 URI。</span><span class="sxs-lookup"><span data-stu-id="9a31e-p150">Only used if `type` is set to `file`. The URI of the location for the file.</span></span>|
|`formData.attachments.isInline`|<span data-ttu-id="9a31e-960">布尔</span><span class="sxs-lookup"><span data-stu-id="9a31e-960">Boolean</span></span>||<span data-ttu-id="9a31e-p151">仅在将 `type` 设置为 `file` 时使用。如果为 `true`，则表示附件将在邮件正文中内联显示，并且不应显示在附件列表中。</span><span class="sxs-lookup"><span data-stu-id="9a31e-p151">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`formData.attachments.itemId`|<span data-ttu-id="9a31e-963">String</span><span class="sxs-lookup"><span data-stu-id="9a31e-963">String</span></span>||<span data-ttu-id="9a31e-p152">仅在将 `type` 设置为 `item` 时使用。附件的 EWS 项目 ID。字符串最长为 100 个字符。</span><span class="sxs-lookup"><span data-stu-id="9a31e-p152">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span>|
|`callback`|<span data-ttu-id="9a31e-967">函数</span><span class="sxs-lookup"><span data-stu-id="9a31e-967">function</span></span>|<span data-ttu-id="9a31e-968">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="9a31e-968">&lt;optional&gt;</span></span>|<span data-ttu-id="9a31e-969">方法完成后，使用单个参数 `asyncResult`（一个 [AsyncResult](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="9a31e-969">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="9a31e-970">要求</span><span class="sxs-lookup"><span data-stu-id="9a31e-970">Requirements</span></span>

|<span data-ttu-id="9a31e-971">要求</span><span class="sxs-lookup"><span data-stu-id="9a31e-971">Requirement</span></span>|<span data-ttu-id="9a31e-972">值</span><span class="sxs-lookup"><span data-stu-id="9a31e-972">Value</span></span>|
|---|---|
|[<span data-ttu-id="9a31e-973">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="9a31e-973">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="9a31e-974">1.0</span><span class="sxs-lookup"><span data-stu-id="9a31e-974">1.0</span></span>|
|[<span data-ttu-id="9a31e-975">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="9a31e-975">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="9a31e-976">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9a31e-976">ReadItem</span></span>|
|[<span data-ttu-id="9a31e-977">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="9a31e-977">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="9a31e-978">阅读</span><span class="sxs-lookup"><span data-stu-id="9a31e-978">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="9a31e-979">示例</span><span class="sxs-lookup"><span data-stu-id="9a31e-979">Examples</span></span>

<span data-ttu-id="9a31e-980">以下代码将一个字符串传递到 `displayReplyAllForm` 函数。</span><span class="sxs-lookup"><span data-stu-id="9a31e-980">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="9a31e-981">使用空白正文答复。</span><span class="sxs-lookup"><span data-stu-id="9a31e-981">Reply with an empty body.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="9a31e-982">仅使用正文答复。</span><span class="sxs-lookup"><span data-stu-id="9a31e-982">Reply with just a body.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="9a31e-983">使用正文和文件附件答复。</span><span class="sxs-lookup"><span data-stu-id="9a31e-983">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="9a31e-984">使用正文和项目附件答复。</span><span class="sxs-lookup"><span data-stu-id="9a31e-984">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="9a31e-985">使用正文、文件附件、项目附件和回调答复。</span><span class="sxs-lookup"><span data-stu-id="9a31e-985">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="displayreplyformformdata-callback"></a><span data-ttu-id="9a31e-986">displayReplyForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="9a31e-986">displayReplyForm(formData, [callback])</span></span>

<span data-ttu-id="9a31e-987">显示答复窗体，其中仅包括所选邮件的发件人或所选约会的组织者。</span><span class="sxs-lookup"><span data-stu-id="9a31e-987">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="9a31e-988">IOS 或 Android 上的 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="9a31e-988">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="9a31e-989">在 web 上的 Outlook 中，答复窗体显示为3列视图中的弹出窗体和2列或1列视图中的弹出式窗体。</span><span class="sxs-lookup"><span data-stu-id="9a31e-989">In Outlook on the web, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="9a31e-990">如果任意字符串参数超出其限制，`displayReplyForm` 将引发异常。</span><span class="sxs-lookup"><span data-stu-id="9a31e-990">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="9a31e-991">如果在`formData.attachments`参数中指定了附件，则 web 上的 Outlook 和桌面客户端将尝试下载所有附件并将其附加到答复窗体。</span><span class="sxs-lookup"><span data-stu-id="9a31e-991">When attachments are specified in the `formData.attachments` parameter, Outlook on the web and desktop clients attempt to download all attachments and attach them to the reply form.</span></span> <span data-ttu-id="9a31e-992">如果无法添加任何附件，则在窗体 UI 中显示错误。</span><span class="sxs-lookup"><span data-stu-id="9a31e-992">If any attachments fail to be added, an error is shown in the form UI.</span></span> <span data-ttu-id="9a31e-993">如果这不可能，则不引发错误消息。</span><span class="sxs-lookup"><span data-stu-id="9a31e-993">If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="9a31e-994">参数</span><span class="sxs-lookup"><span data-stu-id="9a31e-994">Parameters</span></span>

|<span data-ttu-id="9a31e-995">名称</span><span class="sxs-lookup"><span data-stu-id="9a31e-995">Name</span></span>|<span data-ttu-id="9a31e-996">类型</span><span class="sxs-lookup"><span data-stu-id="9a31e-996">Type</span></span>|<span data-ttu-id="9a31e-997">属性</span><span class="sxs-lookup"><span data-stu-id="9a31e-997">Attributes</span></span>|<span data-ttu-id="9a31e-998">说明</span><span class="sxs-lookup"><span data-stu-id="9a31e-998">Description</span></span>|
|---|---|---|---|
|`formData`|<span data-ttu-id="9a31e-999">字符串 &#124; 对象</span><span class="sxs-lookup"><span data-stu-id="9a31e-999">String &#124; Object</span></span>||<span data-ttu-id="9a31e-p154">一个包含文本和 HTML 且表示答复窗体的正文的字符串。字符串限制为 32 KB。</span><span class="sxs-lookup"><span data-stu-id="9a31e-p154">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="9a31e-1002">**或**</span><span class="sxs-lookup"><span data-stu-id="9a31e-1002">**OR**</span></span><br/><span data-ttu-id="9a31e-p155">包含正文或附件数据和回调函数的对象。对象定义如下。</span><span class="sxs-lookup"><span data-stu-id="9a31e-p155">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span>|
|`formData.htmlBody`|<span data-ttu-id="9a31e-1005">String</span><span class="sxs-lookup"><span data-stu-id="9a31e-1005">String</span></span>|<span data-ttu-id="9a31e-1006">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="9a31e-1006">&lt;optional&gt;</span></span>|<span data-ttu-id="9a31e-p156">一个包含文本和 HTML 且表示答复窗体的正文的字符串。字符串限制为 32 KB。</span><span class="sxs-lookup"><span data-stu-id="9a31e-p156">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
|`formData.attachments`|<span data-ttu-id="9a31e-1009">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="9a31e-1009">Array.&lt;Object&gt;</span></span>|<span data-ttu-id="9a31e-1010">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="9a31e-1010">&lt;optional&gt;</span></span>|<span data-ttu-id="9a31e-1011">JSON 对象（是文件或项目附件）的数组。</span><span class="sxs-lookup"><span data-stu-id="9a31e-1011">An array of JSON objects that are either file or item attachments.</span></span>|
|`formData.attachments.type`|<span data-ttu-id="9a31e-1012">String</span><span class="sxs-lookup"><span data-stu-id="9a31e-1012">String</span></span>||<span data-ttu-id="9a31e-p157">指示附件的类型。必须是文件附件的 `file` 或项目附件的 `item`。</span><span class="sxs-lookup"><span data-stu-id="9a31e-p157">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span>|
|`formData.attachments.name`|<span data-ttu-id="9a31e-1015">字符串</span><span class="sxs-lookup"><span data-stu-id="9a31e-1015">String</span></span>||<span data-ttu-id="9a31e-1016">一个包含附件的名称的字符串，最多包含 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="9a31e-1016">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
|`formData.attachments.url`|<span data-ttu-id="9a31e-1017">String</span><span class="sxs-lookup"><span data-stu-id="9a31e-1017">String</span></span>||<span data-ttu-id="9a31e-p158">仅在将 `type` 设置为 `file` 时使用。文件的位置的 URI。</span><span class="sxs-lookup"><span data-stu-id="9a31e-p158">Only used if `type` is set to `file`. The URI of the location for the file.</span></span>|
|`formData.attachments.isInline`|<span data-ttu-id="9a31e-1020">布尔</span><span class="sxs-lookup"><span data-stu-id="9a31e-1020">Boolean</span></span>||<span data-ttu-id="9a31e-p159">仅在将 `type` 设置为 `file` 时使用。如果为 `true`，则表示附件将在邮件正文中内联显示，并且不应显示在附件列表中。</span><span class="sxs-lookup"><span data-stu-id="9a31e-p159">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`formData.attachments.itemId`|<span data-ttu-id="9a31e-1023">String</span><span class="sxs-lookup"><span data-stu-id="9a31e-1023">String</span></span>||<span data-ttu-id="9a31e-p160">仅在将 `type` 设置为 `item` 时使用。附件的 EWS 项目 ID。字符串最长为 100 个字符。</span><span class="sxs-lookup"><span data-stu-id="9a31e-p160">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span>|
|`callback`|<span data-ttu-id="9a31e-1027">函数</span><span class="sxs-lookup"><span data-stu-id="9a31e-1027">function</span></span>|<span data-ttu-id="9a31e-1028">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="9a31e-1028">&lt;optional&gt;</span></span>|<span data-ttu-id="9a31e-1029">方法完成后，使用单个参数 `asyncResult`（一个 [AsyncResult](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="9a31e-1029">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="9a31e-1030">要求</span><span class="sxs-lookup"><span data-stu-id="9a31e-1030">Requirements</span></span>

|<span data-ttu-id="9a31e-1031">要求</span><span class="sxs-lookup"><span data-stu-id="9a31e-1031">Requirement</span></span>|<span data-ttu-id="9a31e-1032">值</span><span class="sxs-lookup"><span data-stu-id="9a31e-1032">Value</span></span>|
|---|---|
|[<span data-ttu-id="9a31e-1033">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="9a31e-1033">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="9a31e-1034">1.0</span><span class="sxs-lookup"><span data-stu-id="9a31e-1034">1.0</span></span>|
|[<span data-ttu-id="9a31e-1035">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="9a31e-1035">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="9a31e-1036">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9a31e-1036">ReadItem</span></span>|
|[<span data-ttu-id="9a31e-1037">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="9a31e-1037">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="9a31e-1038">阅读</span><span class="sxs-lookup"><span data-stu-id="9a31e-1038">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="9a31e-1039">示例</span><span class="sxs-lookup"><span data-stu-id="9a31e-1039">Examples</span></span>

<span data-ttu-id="9a31e-1040">以下代码将一个字符串传递到 `displayReplyForm` 函数。</span><span class="sxs-lookup"><span data-stu-id="9a31e-1040">The following code passes a string to the `displayReplyForm` function.</span></span>

```js
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="9a31e-1041">使用空白正文答复。</span><span class="sxs-lookup"><span data-stu-id="9a31e-1041">Reply with an empty body.</span></span>

```js
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="9a31e-1042">仅使用正文答复。</span><span class="sxs-lookup"><span data-stu-id="9a31e-1042">Reply with just a body.</span></span>

```js
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="9a31e-1043">使用正文和文件附件答复。</span><span class="sxs-lookup"><span data-stu-id="9a31e-1043">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="9a31e-1044">使用正文和项目附件答复。</span><span class="sxs-lookup"><span data-stu-id="9a31e-1044">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="9a31e-1045">使用正文、文件附件、项目附件和回调答复。</span><span class="sxs-lookup"><span data-stu-id="9a31e-1045">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="getattachmentcontentasyncattachmentid-options-callback--attachmentcontentjavascriptapioutlookofficeattachmentcontent"></a><span data-ttu-id="9a31e-1046">getAttachmentContentAsync （attachmentId，[options]，[callback]）→ [AttachmentContent](/javascript/api/outlook/office.attachmentcontent)</span><span class="sxs-lookup"><span data-stu-id="9a31e-1046">getAttachmentContentAsync(attachmentId, [options], [callback]) → [AttachmentContent](/javascript/api/outlook/office.attachmentcontent)</span></span>

<span data-ttu-id="9a31e-1047">从邮件或约会中获取指定附件并将其作为`AttachmentContent`对象返回。</span><span class="sxs-lookup"><span data-stu-id="9a31e-1047">Gets the specified attachment from a message or appointment and returns it as an `AttachmentContent` object.</span></span>

<span data-ttu-id="9a31e-1048">该`getAttachmentContentAsync`方法从项目中获取具有指定标识符的附件。</span><span class="sxs-lookup"><span data-stu-id="9a31e-1048">The `getAttachmentContentAsync` method gets the attachment with the specified identifier from the item.</span></span> <span data-ttu-id="9a31e-1049">作为一种最佳做法，您应使用标识符在与`getAttachmentsAsync` or `item.attachments`调用一起检索到会话的同一会话中检索附件。</span><span class="sxs-lookup"><span data-stu-id="9a31e-1049">As a best practice, you should use the identifier to retrieve an attachment in the same session that the attachmentIds were retrieved with the `getAttachmentsAsync` or `item.attachments` call.</span></span> <span data-ttu-id="9a31e-1050">在 web 和移动设备上的 Outlook 中，附件标识符仅在同一个会话中有效。</span><span class="sxs-lookup"><span data-stu-id="9a31e-1050">In Outlook on the web and mobile devices, the attachment identifier is valid only within the same session.</span></span> <span data-ttu-id="9a31e-1051">当用户关闭应用程序时，或者如果用户开始撰写内嵌窗体，随后弹出窗体以继续在单独的窗口中。</span><span class="sxs-lookup"><span data-stu-id="9a31e-1051">A session is over when the user closes the app, or if the user starts composing an inline form then subsequently pops out the form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="9a31e-1052">参数</span><span class="sxs-lookup"><span data-stu-id="9a31e-1052">Parameters</span></span>

|<span data-ttu-id="9a31e-1053">名称</span><span class="sxs-lookup"><span data-stu-id="9a31e-1053">Name</span></span>|<span data-ttu-id="9a31e-1054">类型</span><span class="sxs-lookup"><span data-stu-id="9a31e-1054">Type</span></span>|<span data-ttu-id="9a31e-1055">属性</span><span class="sxs-lookup"><span data-stu-id="9a31e-1055">Attributes</span></span>|<span data-ttu-id="9a31e-1056">说明</span><span class="sxs-lookup"><span data-stu-id="9a31e-1056">Description</span></span>|
|---|---|---|---|
|`attachmentId`|<span data-ttu-id="9a31e-1057">String</span><span class="sxs-lookup"><span data-stu-id="9a31e-1057">String</span></span>||<span data-ttu-id="9a31e-1058">要获取的附件的标识符。</span><span class="sxs-lookup"><span data-stu-id="9a31e-1058">The identifier of the attachment you want to get.</span></span>|
|`options`|<span data-ttu-id="9a31e-1059">对象</span><span class="sxs-lookup"><span data-stu-id="9a31e-1059">Object</span></span>|<span data-ttu-id="9a31e-1060">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="9a31e-1060">&lt;optional&gt;</span></span>|<span data-ttu-id="9a31e-1061">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="9a31e-1061">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="9a31e-1062">对象</span><span class="sxs-lookup"><span data-stu-id="9a31e-1062">Object</span></span>|<span data-ttu-id="9a31e-1063">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="9a31e-1063">&lt;optional&gt;</span></span>|<span data-ttu-id="9a31e-1064">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="9a31e-1064">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="9a31e-1065">函数</span><span class="sxs-lookup"><span data-stu-id="9a31e-1065">function</span></span>|<span data-ttu-id="9a31e-1066">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="9a31e-1066">&lt;optional&gt;</span></span>|<span data-ttu-id="9a31e-1067">方法完成后，使用单个参数 `asyncResult`（一个 [AsyncResult](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="9a31e-1067">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="9a31e-1068">要求</span><span class="sxs-lookup"><span data-stu-id="9a31e-1068">Requirements</span></span>

|<span data-ttu-id="9a31e-1069">要求</span><span class="sxs-lookup"><span data-stu-id="9a31e-1069">Requirement</span></span>|<span data-ttu-id="9a31e-1070">值</span><span class="sxs-lookup"><span data-stu-id="9a31e-1070">Value</span></span>|
|---|---|
|[<span data-ttu-id="9a31e-1071">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="9a31e-1071">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="9a31e-1072">预览</span><span class="sxs-lookup"><span data-stu-id="9a31e-1072">Preview</span></span>|
|[<span data-ttu-id="9a31e-1073">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="9a31e-1073">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="9a31e-1074">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9a31e-1074">ReadItem</span></span>|
|[<span data-ttu-id="9a31e-1075">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="9a31e-1075">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="9a31e-1076">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="9a31e-1076">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="9a31e-1077">返回：</span><span class="sxs-lookup"><span data-stu-id="9a31e-1077">Returns:</span></span>

<span data-ttu-id="9a31e-1078">类型： [AttachmentContent](/javascript/api/outlook/office.attachmentcontent)</span><span class="sxs-lookup"><span data-stu-id="9a31e-1078">Type: [AttachmentContent](/javascript/api/outlook/office.attachmentcontent)</span></span>

##### <a name="example"></a><span data-ttu-id="9a31e-1079">示例</span><span class="sxs-lookup"><span data-stu-id="9a31e-1079">Example</span></span>

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

#### <a name="getattachmentsasyncoptions-callback--arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetails"></a><span data-ttu-id="9a31e-1080">getAttachmentsAsync （[options]，[callback]）→ Array. <[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="9a31e-1080">getAttachmentsAsync([options], [callback]) → Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span></span>

<span data-ttu-id="9a31e-1081">以数组的形式获取项目的附件。</span><span class="sxs-lookup"><span data-stu-id="9a31e-1081">Gets the item's attachments as an array.</span></span> <span data-ttu-id="9a31e-1082">仅限撰写模式。</span><span class="sxs-lookup"><span data-stu-id="9a31e-1082">Compose mode only.</span></span>

##### <a name="parameters"></a><span data-ttu-id="9a31e-1083">参数</span><span class="sxs-lookup"><span data-stu-id="9a31e-1083">Parameters</span></span>

|<span data-ttu-id="9a31e-1084">名称</span><span class="sxs-lookup"><span data-stu-id="9a31e-1084">Name</span></span>|<span data-ttu-id="9a31e-1085">类型</span><span class="sxs-lookup"><span data-stu-id="9a31e-1085">Type</span></span>|<span data-ttu-id="9a31e-1086">属性</span><span class="sxs-lookup"><span data-stu-id="9a31e-1086">Attributes</span></span>|<span data-ttu-id="9a31e-1087">说明</span><span class="sxs-lookup"><span data-stu-id="9a31e-1087">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="9a31e-1088">对象</span><span class="sxs-lookup"><span data-stu-id="9a31e-1088">Object</span></span>|<span data-ttu-id="9a31e-1089">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="9a31e-1089">&lt;optional&gt;</span></span>|<span data-ttu-id="9a31e-1090">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="9a31e-1090">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="9a31e-1091">对象</span><span class="sxs-lookup"><span data-stu-id="9a31e-1091">Object</span></span>|<span data-ttu-id="9a31e-1092">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="9a31e-1092">&lt;optional&gt;</span></span>|<span data-ttu-id="9a31e-1093">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="9a31e-1093">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="9a31e-1094">函数</span><span class="sxs-lookup"><span data-stu-id="9a31e-1094">function</span></span>|<span data-ttu-id="9a31e-1095">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="9a31e-1095">&lt;optional&gt;</span></span>|<span data-ttu-id="9a31e-1096">方法完成后，使用单个参数 `asyncResult`（一个 [AsyncResult](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="9a31e-1096">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="9a31e-1097">要求</span><span class="sxs-lookup"><span data-stu-id="9a31e-1097">Requirements</span></span>

|<span data-ttu-id="9a31e-1098">要求</span><span class="sxs-lookup"><span data-stu-id="9a31e-1098">Requirement</span></span>|<span data-ttu-id="9a31e-1099">值</span><span class="sxs-lookup"><span data-stu-id="9a31e-1099">Value</span></span>|
|---|---|
|[<span data-ttu-id="9a31e-1100">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="9a31e-1100">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="9a31e-1101">预览</span><span class="sxs-lookup"><span data-stu-id="9a31e-1101">Preview</span></span>|
|[<span data-ttu-id="9a31e-1102">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="9a31e-1102">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="9a31e-1103">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9a31e-1103">ReadItem</span></span>|
|[<span data-ttu-id="9a31e-1104">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="9a31e-1104">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="9a31e-1105">撰写</span><span class="sxs-lookup"><span data-stu-id="9a31e-1105">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="9a31e-1106">返回：</span><span class="sxs-lookup"><span data-stu-id="9a31e-1106">Returns:</span></span>

<span data-ttu-id="9a31e-1107">类型： Array. <[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="9a31e-1107">Type: Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span></span>

##### <a name="example"></a><span data-ttu-id="9a31e-1108">示例</span><span class="sxs-lookup"><span data-stu-id="9a31e-1108">Example</span></span>

<span data-ttu-id="9a31e-1109">下面的示例将生成一个 HTML 字符串，其中包含当前项目上所有附件的详细信息。</span><span class="sxs-lookup"><span data-stu-id="9a31e-1109">The following example builds an HTML string with details of all attachments on the current item.</span></span>

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

#### <a name="getentities--entitiesjavascriptapioutlookofficeentities"></a><span data-ttu-id="9a31e-1110">getEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="9a31e-1110">getEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span></span>

<span data-ttu-id="9a31e-1111">获取在所选项目的正文中找到的实体。</span><span class="sxs-lookup"><span data-stu-id="9a31e-1111">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="9a31e-1112">IOS 或 Android 上的 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="9a31e-1112">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="9a31e-1113">要求</span><span class="sxs-lookup"><span data-stu-id="9a31e-1113">Requirements</span></span>

|<span data-ttu-id="9a31e-1114">要求</span><span class="sxs-lookup"><span data-stu-id="9a31e-1114">Requirement</span></span>|<span data-ttu-id="9a31e-1115">值</span><span class="sxs-lookup"><span data-stu-id="9a31e-1115">Value</span></span>|
|---|---|
|[<span data-ttu-id="9a31e-1116">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="9a31e-1116">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="9a31e-1117">1.0</span><span class="sxs-lookup"><span data-stu-id="9a31e-1117">1.0</span></span>|
|[<span data-ttu-id="9a31e-1118">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="9a31e-1118">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="9a31e-1119">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9a31e-1119">ReadItem</span></span>|
|[<span data-ttu-id="9a31e-1120">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="9a31e-1120">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="9a31e-1121">阅读</span><span class="sxs-lookup"><span data-stu-id="9a31e-1121">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="9a31e-1122">返回：</span><span class="sxs-lookup"><span data-stu-id="9a31e-1122">Returns:</span></span>

<span data-ttu-id="9a31e-1123">类型：[Entities](/javascript/api/outlook/office.entities)</span><span class="sxs-lookup"><span data-stu-id="9a31e-1123">Type: [Entities](/javascript/api/outlook/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="9a31e-1124">示例</span><span class="sxs-lookup"><span data-stu-id="9a31e-1124">Example</span></span>

<span data-ttu-id="9a31e-1125">以下示例访问当前项目的正文中的联系人实体。</span><span class="sxs-lookup"><span data-stu-id="9a31e-1125">The following example accesses the contacts entities in the current item's body.</span></span>

```js
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

<br>

---
---

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlookofficecontactmeetingsuggestionjavascriptapioutlookofficemeetingsuggestionphonenumberjavascriptapioutlookofficephonenumbertasksuggestionjavascriptapioutlookofficetasksuggestion"></a><span data-ttu-id="9a31e-1126">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="9a31e-1126">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span></span>

<span data-ttu-id="9a31e-1127">获取所选项目的正文中找到的指定实体类型的所有实体的数组。</span><span class="sxs-lookup"><span data-stu-id="9a31e-1127">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="9a31e-1128">IOS 或 Android 上的 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="9a31e-1128">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="9a31e-1129">参数</span><span class="sxs-lookup"><span data-stu-id="9a31e-1129">Parameters</span></span>

|<span data-ttu-id="9a31e-1130">名称</span><span class="sxs-lookup"><span data-stu-id="9a31e-1130">Name</span></span>|<span data-ttu-id="9a31e-1131">类型</span><span class="sxs-lookup"><span data-stu-id="9a31e-1131">Type</span></span>|<span data-ttu-id="9a31e-1132">说明</span><span class="sxs-lookup"><span data-stu-id="9a31e-1132">Description</span></span>|
|---|---|---|
|`entityType`|[<span data-ttu-id="9a31e-1133">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="9a31e-1133">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook/office.mailboxenums.entitytype)|<span data-ttu-id="9a31e-1134">EntityType 枚举值之一。</span><span class="sxs-lookup"><span data-stu-id="9a31e-1134">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="9a31e-1135">要求</span><span class="sxs-lookup"><span data-stu-id="9a31e-1135">Requirements</span></span>

|<span data-ttu-id="9a31e-1136">要求</span><span class="sxs-lookup"><span data-stu-id="9a31e-1136">Requirement</span></span>|<span data-ttu-id="9a31e-1137">值</span><span class="sxs-lookup"><span data-stu-id="9a31e-1137">Value</span></span>|
|---|---|
|[<span data-ttu-id="9a31e-1138">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="9a31e-1138">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="9a31e-1139">1.0</span><span class="sxs-lookup"><span data-stu-id="9a31e-1139">1.0</span></span>|
|[<span data-ttu-id="9a31e-1140">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="9a31e-1140">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="9a31e-1141">受限</span><span class="sxs-lookup"><span data-stu-id="9a31e-1141">Restricted</span></span>|
|[<span data-ttu-id="9a31e-1142">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="9a31e-1142">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="9a31e-1143">阅读</span><span class="sxs-lookup"><span data-stu-id="9a31e-1143">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="9a31e-1144">返回：</span><span class="sxs-lookup"><span data-stu-id="9a31e-1144">Returns:</span></span>

<span data-ttu-id="9a31e-1145">如果在 `entityType` 中传递的值不是 `EntityType` 枚举的有效成员，该方法返回 null。</span><span class="sxs-lookup"><span data-stu-id="9a31e-1145">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="9a31e-1146">如果指定类型的任何实体都不存在于该项目的正文中，该方法将返回空数组。</span><span class="sxs-lookup"><span data-stu-id="9a31e-1146">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="9a31e-1147">否则，返回的数组中对象的类型取决于 `entityType` 参数中请求实体的类型。</span><span class="sxs-lookup"><span data-stu-id="9a31e-1147">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="9a31e-1148">当使用此方法的最低权限级别**受限**时，某些实体类型需要 **ReadItem** 才能进行访问，如下表中所指定。</span><span class="sxs-lookup"><span data-stu-id="9a31e-1148">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

|<span data-ttu-id="9a31e-1149">`entityType` 的值</span><span class="sxs-lookup"><span data-stu-id="9a31e-1149">Value of `entityType`</span></span>|<span data-ttu-id="9a31e-1150">返回的数组中对象的类型</span><span class="sxs-lookup"><span data-stu-id="9a31e-1150">Type of objects in returned array</span></span>|<span data-ttu-id="9a31e-1151">所需权限级别</span><span class="sxs-lookup"><span data-stu-id="9a31e-1151">Required Permission Level</span></span>|
|---|---|---|
|`Address`|<span data-ttu-id="9a31e-1152">字符串</span><span class="sxs-lookup"><span data-stu-id="9a31e-1152">String</span></span>|<span data-ttu-id="9a31e-1153">**受限**</span><span class="sxs-lookup"><span data-stu-id="9a31e-1153">**Restricted**</span></span>|
|`Contact`|<span data-ttu-id="9a31e-1154">Contact</span><span class="sxs-lookup"><span data-stu-id="9a31e-1154">Contact</span></span>|<span data-ttu-id="9a31e-1155">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="9a31e-1155">**ReadItem**</span></span>|
|`EmailAddress`|<span data-ttu-id="9a31e-1156">String</span><span class="sxs-lookup"><span data-stu-id="9a31e-1156">String</span></span>|<span data-ttu-id="9a31e-1157">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="9a31e-1157">**ReadItem**</span></span>|
|`MeetingSuggestion`|<span data-ttu-id="9a31e-1158">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="9a31e-1158">MeetingSuggestion</span></span>|<span data-ttu-id="9a31e-1159">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="9a31e-1159">**ReadItem**</span></span>|
|`PhoneNumber`|<span data-ttu-id="9a31e-1160">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="9a31e-1160">PhoneNumber</span></span>|<span data-ttu-id="9a31e-1161">**受限**</span><span class="sxs-lookup"><span data-stu-id="9a31e-1161">**Restricted**</span></span>|
|`TaskSuggestion`|<span data-ttu-id="9a31e-1162">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="9a31e-1162">TaskSuggestion</span></span>|<span data-ttu-id="9a31e-1163">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="9a31e-1163">**ReadItem**</span></span>|
|`URL`|<span data-ttu-id="9a31e-1164">字符串</span><span class="sxs-lookup"><span data-stu-id="9a31e-1164">String</span></span>|<span data-ttu-id="9a31e-1165">**受限**</span><span class="sxs-lookup"><span data-stu-id="9a31e-1165">**Restricted**</span></span>|

<span data-ttu-id="9a31e-1166">类型：Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="9a31e-1166">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span></span>

##### <a name="example"></a><span data-ttu-id="9a31e-1167">示例</span><span class="sxs-lookup"><span data-stu-id="9a31e-1167">Example</span></span>

<span data-ttu-id="9a31e-1168">以下示例显示了如何访问表示当前项目的正文中的邮政地址的字符串数组。</span><span class="sxs-lookup"><span data-stu-id="9a31e-1168">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlookofficecontactmeetingsuggestionjavascriptapioutlookofficemeetingsuggestionphonenumberjavascriptapioutlookofficephonenumbertasksuggestionjavascriptapioutlookofficetasksuggestion"></a><span data-ttu-id="9a31e-1169">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="9a31e-1169">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span></span>

<span data-ttu-id="9a31e-1170">返回传递清单 XML 文件中定义的命名筛选器的所选项目中的已知实体。</span><span class="sxs-lookup"><span data-stu-id="9a31e-1170">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="9a31e-1171">IOS 或 Android 上的 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="9a31e-1171">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="9a31e-1172">`getFilteredEntitiesByName` 方法返回匹配在具有指定 `FilterName` 元素值的清单 XML 文件中的 [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) 规则元素中定义的正则表达式的实体。</span><span class="sxs-lookup"><span data-stu-id="9a31e-1172">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="9a31e-1173">参数</span><span class="sxs-lookup"><span data-stu-id="9a31e-1173">Parameters</span></span>

|<span data-ttu-id="9a31e-1174">名称</span><span class="sxs-lookup"><span data-stu-id="9a31e-1174">Name</span></span>|<span data-ttu-id="9a31e-1175">类型</span><span class="sxs-lookup"><span data-stu-id="9a31e-1175">Type</span></span>|<span data-ttu-id="9a31e-1176">说明</span><span class="sxs-lookup"><span data-stu-id="9a31e-1176">Description</span></span>|
|---|---|---|
|`name`|<span data-ttu-id="9a31e-1177">字符串</span><span class="sxs-lookup"><span data-stu-id="9a31e-1177">String</span></span>|<span data-ttu-id="9a31e-1178">定义筛选器匹配的 `ItemHasKnownEntity` 规则元素的名称。</span><span class="sxs-lookup"><span data-stu-id="9a31e-1178">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="9a31e-1179">要求</span><span class="sxs-lookup"><span data-stu-id="9a31e-1179">Requirements</span></span>

|<span data-ttu-id="9a31e-1180">要求</span><span class="sxs-lookup"><span data-stu-id="9a31e-1180">Requirement</span></span>|<span data-ttu-id="9a31e-1181">值</span><span class="sxs-lookup"><span data-stu-id="9a31e-1181">Value</span></span>|
|---|---|
|[<span data-ttu-id="9a31e-1182">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="9a31e-1182">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="9a31e-1183">1.0</span><span class="sxs-lookup"><span data-stu-id="9a31e-1183">1.0</span></span>|
|[<span data-ttu-id="9a31e-1184">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="9a31e-1184">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="9a31e-1185">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9a31e-1185">ReadItem</span></span>|
|[<span data-ttu-id="9a31e-1186">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="9a31e-1186">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="9a31e-1187">阅读</span><span class="sxs-lookup"><span data-stu-id="9a31e-1187">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="9a31e-1188">返回：</span><span class="sxs-lookup"><span data-stu-id="9a31e-1188">Returns:</span></span>

<span data-ttu-id="9a31e-p164">如果具有匹配 `name` 参数的 `FilterName` 元素值的清单中没有任何 `ItemHasKnownEntity` 元素，则该方法返回 `null`。如果 `name` 参数匹配清单中的 `ItemHasKnownEntity` 元素，但在匹配的当前项目中没有实体，则该方法返回一个空数组。</span><span class="sxs-lookup"><span data-stu-id="9a31e-p164">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="9a31e-1191">类型：Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="9a31e-1191">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span></span>

<br>

---
---

#### <a name="getinitializationcontextasyncoptions-callback"></a><span data-ttu-id="9a31e-1192">Office.context.mailbox.item.getinitializationcontextasync （[options]，[callback]）</span><span class="sxs-lookup"><span data-stu-id="9a31e-1192">getInitializationContextAsync([options], [callback])</span></span>

<span data-ttu-id="9a31e-1193">获取[通过可操作邮件激活](/outlook/actionable-messages/invoke-add-in-from-actionable-message)加载项时传递的初始化数据。</span><span class="sxs-lookup"><span data-stu-id="9a31e-1193">Gets initialization data passed when the add-in is [activated by an actionable message](/outlook/actionable-messages/invoke-add-in-from-actionable-message).</span></span>

> [!NOTE]
> <span data-ttu-id="9a31e-1194">仅 Outlook 2016 或更高版本（高于16.0.8413.1000 的即点即用版本）和适用于 Office 365 的 Outlook 网页版支持此方法。</span><span class="sxs-lookup"><span data-stu-id="9a31e-1194">This method is only supported by Outlook 2016 or later on Windows (Click-to-Run versions later than 16.0.8413.1000) and Outlook on the web for Office 365.</span></span>

##### <a name="parameters"></a><span data-ttu-id="9a31e-1195">参数</span><span class="sxs-lookup"><span data-stu-id="9a31e-1195">Parameters</span></span>

|<span data-ttu-id="9a31e-1196">名称</span><span class="sxs-lookup"><span data-stu-id="9a31e-1196">Name</span></span>|<span data-ttu-id="9a31e-1197">类型</span><span class="sxs-lookup"><span data-stu-id="9a31e-1197">Type</span></span>|<span data-ttu-id="9a31e-1198">属性</span><span class="sxs-lookup"><span data-stu-id="9a31e-1198">Attributes</span></span>|<span data-ttu-id="9a31e-1199">说明</span><span class="sxs-lookup"><span data-stu-id="9a31e-1199">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="9a31e-1200">对象</span><span class="sxs-lookup"><span data-stu-id="9a31e-1200">Object</span></span>|<span data-ttu-id="9a31e-1201">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="9a31e-1201">&lt;optional&gt;</span></span>|<span data-ttu-id="9a31e-1202">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="9a31e-1202">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="9a31e-1203">对象</span><span class="sxs-lookup"><span data-stu-id="9a31e-1203">Object</span></span>|<span data-ttu-id="9a31e-1204">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="9a31e-1204">&lt;optional&gt;</span></span>|<span data-ttu-id="9a31e-1205">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="9a31e-1205">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="9a31e-1206">函数</span><span class="sxs-lookup"><span data-stu-id="9a31e-1206">function</span></span>|<span data-ttu-id="9a31e-1207">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="9a31e-1207">&lt;optional&gt;</span></span>|<span data-ttu-id="9a31e-1208">方法完成后，使用单个参数 `callback`（一个 [`asyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `AsyncResult` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="9a31e-1208">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="9a31e-1209">如果成功，初始化数据在`asyncResult.value`属性中提供为字符串。</span><span class="sxs-lookup"><span data-stu-id="9a31e-1209">On success, the initialization data is provided in the `asyncResult.value` property as a string.</span></span><br/><span data-ttu-id="9a31e-1210">如果没有初始化上下文，该`asyncResult`对象将包含其`Error` `code`属性设置为`9020`的对象及其`name`属性设置为。 `GenericResponseError`</span><span class="sxs-lookup"><span data-stu-id="9a31e-1210">If there is no initialization context, the `asyncResult` object will contain an `Error` object with its `code` property set to `9020` and its `name` property set to `GenericResponseError`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="9a31e-1211">要求</span><span class="sxs-lookup"><span data-stu-id="9a31e-1211">Requirements</span></span>

|<span data-ttu-id="9a31e-1212">要求</span><span class="sxs-lookup"><span data-stu-id="9a31e-1212">Requirement</span></span>|<span data-ttu-id="9a31e-1213">值</span><span class="sxs-lookup"><span data-stu-id="9a31e-1213">Value</span></span>|
|---|---|
|[<span data-ttu-id="9a31e-1214">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="9a31e-1214">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="9a31e-1215">预览</span><span class="sxs-lookup"><span data-stu-id="9a31e-1215">Preview</span></span>|
|[<span data-ttu-id="9a31e-1216">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="9a31e-1216">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="9a31e-1217">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9a31e-1217">ReadItem</span></span>|
|[<span data-ttu-id="9a31e-1218">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="9a31e-1218">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="9a31e-1219">阅读</span><span class="sxs-lookup"><span data-stu-id="9a31e-1219">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="9a31e-1220">示例</span><span class="sxs-lookup"><span data-stu-id="9a31e-1220">Example</span></span>

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

#### <a name="getitemidasyncoptions-callback"></a><span data-ttu-id="9a31e-1221">getItemIdAsync （[options]，回拨）</span><span class="sxs-lookup"><span data-stu-id="9a31e-1221">getItemIdAsync([options], callback)</span></span>

<span data-ttu-id="9a31e-1222">异步获取已保存项的 ID。</span><span class="sxs-lookup"><span data-stu-id="9a31e-1222">Asynchronously gets the ID of a saved item.</span></span> <span data-ttu-id="9a31e-1223">仅限撰写模式。</span><span class="sxs-lookup"><span data-stu-id="9a31e-1223">Compose mode only.</span></span>

<span data-ttu-id="9a31e-1224">调用此方法时，此方法通过回调方法返回项 ID。</span><span class="sxs-lookup"><span data-stu-id="9a31e-1224">When invoked, this method returns the item ID via the callback method.</span></span>

> [!NOTE]
> <span data-ttu-id="9a31e-1225">如果你的外接程序`getItemIdAsync`对撰写模式中的项（例如，要获取`itemId`使用 EWS 或 REST API 的使用）调用，请注意，当 Outlook 处于缓存模式下时，可能需要一段时间才能将项目同步到服务器。</span><span class="sxs-lookup"><span data-stu-id="9a31e-1225">If your add-in calls `getItemIdAsync` on an item in compose mode (e.g., to get an `itemId` to use with EWS or the REST API), be aware that when Outlook is in cached mode, it may take some time before the item is synced to the server.</span></span> <span data-ttu-id="9a31e-1226">在同步项目之前，无法识别`itemId`该项目并使用它将返回错误。</span><span class="sxs-lookup"><span data-stu-id="9a31e-1226">Until the item is synced, the `itemId` is not recognized and using it returns an error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="9a31e-1227">参数</span><span class="sxs-lookup"><span data-stu-id="9a31e-1227">Parameters</span></span>

|<span data-ttu-id="9a31e-1228">名称</span><span class="sxs-lookup"><span data-stu-id="9a31e-1228">Name</span></span>|<span data-ttu-id="9a31e-1229">类型</span><span class="sxs-lookup"><span data-stu-id="9a31e-1229">Type</span></span>|<span data-ttu-id="9a31e-1230">属性</span><span class="sxs-lookup"><span data-stu-id="9a31e-1230">Attributes</span></span>|<span data-ttu-id="9a31e-1231">说明</span><span class="sxs-lookup"><span data-stu-id="9a31e-1231">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="9a31e-1232">对象</span><span class="sxs-lookup"><span data-stu-id="9a31e-1232">Object</span></span>|<span data-ttu-id="9a31e-1233">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="9a31e-1233">&lt;optional&gt;</span></span>|<span data-ttu-id="9a31e-1234">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="9a31e-1234">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="9a31e-1235">对象</span><span class="sxs-lookup"><span data-stu-id="9a31e-1235">Object</span></span>|<span data-ttu-id="9a31e-1236">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="9a31e-1236">&lt;optional&gt;</span></span>|<span data-ttu-id="9a31e-1237">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="9a31e-1237">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="9a31e-1238">函数</span><span class="sxs-lookup"><span data-stu-id="9a31e-1238">function</span></span>||<span data-ttu-id="9a31e-1239">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="9a31e-1239">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="9a31e-1240">如果成功，该项目标识符将在 `asyncResult.value` 属性中提供。</span><span class="sxs-lookup"><span data-stu-id="9a31e-1240">On success, the item identifier is provided in the `asyncResult.value` property.</span></span>|

##### <a name="errors"></a><span data-ttu-id="9a31e-1241">错误</span><span class="sxs-lookup"><span data-stu-id="9a31e-1241">Errors</span></span>

|<span data-ttu-id="9a31e-1242">错误代码</span><span class="sxs-lookup"><span data-stu-id="9a31e-1242">Error code</span></span>|<span data-ttu-id="9a31e-1243">说明</span><span class="sxs-lookup"><span data-stu-id="9a31e-1243">Description</span></span>|
|------------|-------------|
|`ItemNotSaved`|<span data-ttu-id="9a31e-1244">在保存项目之前，无法检索此 id。</span><span class="sxs-lookup"><span data-stu-id="9a31e-1244">The id can't be retrieved until the item is saved.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="9a31e-1245">要求</span><span class="sxs-lookup"><span data-stu-id="9a31e-1245">Requirements</span></span>

|<span data-ttu-id="9a31e-1246">要求</span><span class="sxs-lookup"><span data-stu-id="9a31e-1246">Requirement</span></span>|<span data-ttu-id="9a31e-1247">值</span><span class="sxs-lookup"><span data-stu-id="9a31e-1247">Value</span></span>|
|---|---|
|[<span data-ttu-id="9a31e-1248">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="9a31e-1248">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="9a31e-1249">预览</span><span class="sxs-lookup"><span data-stu-id="9a31e-1249">Preview</span></span>|
|[<span data-ttu-id="9a31e-1250">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="9a31e-1250">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="9a31e-1251">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9a31e-1251">ReadItem</span></span>|
|[<span data-ttu-id="9a31e-1252">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="9a31e-1252">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="9a31e-1253">撰写</span><span class="sxs-lookup"><span data-stu-id="9a31e-1253">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="9a31e-1254">示例</span><span class="sxs-lookup"><span data-stu-id="9a31e-1254">Examples</span></span>

```js
Office.context.mailbox.item.getItemIdAsync(
  function callback(result) {
    // Process the result.
  });
```

<span data-ttu-id="9a31e-1255">下面的示例演示传递给回调函数`result`的参数的结构。</span><span class="sxs-lookup"><span data-stu-id="9a31e-1255">The following example shows the structure of the `result` parameter that's passed to the callback function.</span></span> <span data-ttu-id="9a31e-1256">`value`属性包含项 ID。</span><span class="sxs-lookup"><span data-stu-id="9a31e-1256">The `value` property contains the item ID.</span></span>

```json
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

<br>

---
---

#### <a name="getregexmatches--object"></a><span data-ttu-id="9a31e-1257">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="9a31e-1257">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="9a31e-1258">返回所选项目中匹配在清单 XML 文件中定义的正则表达式的字符串值。</span><span class="sxs-lookup"><span data-stu-id="9a31e-1258">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="9a31e-1259">IOS 或 Android 上的 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="9a31e-1259">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="9a31e-p168">`getRegExMatches` 方法返回匹配在清单 XML 文件中的每个 `ItemHasRegularExpressionMatch` 或 `ItemHasKnownEntity` 规则元素中定义的正则表达式的字符串。对于 `ItemHasRegularExpressionMatch` 规则，匹配字符串必须发生在该规则指定的项目的属性中。`PropertyName` 简单类型定义支持的属性。</span><span class="sxs-lookup"><span data-stu-id="9a31e-p168">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="9a31e-1263">例如，考虑一个外接程序清单具有以下 `Rule` 元素：</span><span class="sxs-lookup"><span data-stu-id="9a31e-1263">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="9a31e-1264">从 `getRegExMatches` 返回的对象应有两个属性：`fruits` 和 `veggies`。</span><span class="sxs-lookup"><span data-stu-id="9a31e-1264">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="9a31e-p169">如果在项目的正文属性上指定 `ItemHasRegularExpressionMatch` 规则，则正则表达式应进一步筛选正文，不应尝试返回该项目的整个正文。使用正则表达式（如 `.*`）获取项目的整个正文并不总是返回预期的结果。而是使用 [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) 方法检索整个正文。</span><span class="sxs-lookup"><span data-stu-id="9a31e-p169">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="9a31e-1268">Requirements</span><span class="sxs-lookup"><span data-stu-id="9a31e-1268">Requirements</span></span>

|<span data-ttu-id="9a31e-1269">要求</span><span class="sxs-lookup"><span data-stu-id="9a31e-1269">Requirement</span></span>|<span data-ttu-id="9a31e-1270">值</span><span class="sxs-lookup"><span data-stu-id="9a31e-1270">Value</span></span>|
|---|---|
|[<span data-ttu-id="9a31e-1271">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="9a31e-1271">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="9a31e-1272">1.0</span><span class="sxs-lookup"><span data-stu-id="9a31e-1272">1.0</span></span>|
|[<span data-ttu-id="9a31e-1273">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="9a31e-1273">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="9a31e-1274">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9a31e-1274">ReadItem</span></span>|
|[<span data-ttu-id="9a31e-1275">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="9a31e-1275">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="9a31e-1276">阅读</span><span class="sxs-lookup"><span data-stu-id="9a31e-1276">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="9a31e-1277">返回：</span><span class="sxs-lookup"><span data-stu-id="9a31e-1277">Returns:</span></span>

<span data-ttu-id="9a31e-p170">一个包含匹配在清单 XML 文件中定义的正则表达式的字符串数组的对象。每个数组的名称等于匹配 `ItemHasRegularExpressionMatch` 规则的 `RegExName` 属性或匹配 `ItemHasKnownEntity` 规则的 `FilterName` 属性的相应值。</span><span class="sxs-lookup"><span data-stu-id="9a31e-p170">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<dl class="param-type"><span data-ttu-id="9a31e-1280">

<dt>
类型</dt>


</span><span class="sxs-lookup"><span data-stu-id="9a31e-1280">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="9a31e-1281">对象</span><span class="sxs-lookup"><span data-stu-id="9a31e-1281">Object</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="9a31e-1282">示例</span><span class="sxs-lookup"><span data-stu-id="9a31e-1282">Example</span></span>

<span data-ttu-id="9a31e-1283">以下示例显示了如何访问正则表达式规则元素 `fruits` 和 `veggies` 的匹配项的数组，这些元素在清单中指定。</span><span class="sxs-lookup"><span data-stu-id="9a31e-1283">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```js
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veggies = allMatches.veggies;
```

<br>

---
---

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="9a31e-1284">getRegExMatchesByName(name) → (nullable) {Array.<String>}</span><span class="sxs-lookup"><span data-stu-id="9a31e-1284">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="9a31e-1285">返回匹配在清单 XML 文件中定义的命名正则表达式的所选项目中的字符串值。</span><span class="sxs-lookup"><span data-stu-id="9a31e-1285">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="9a31e-1286">IOS 或 Android 上的 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="9a31e-1286">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="9a31e-1287">`getRegExMatchesByName` 方法返回匹配在具有指定 `RegExName` 元素值的清单 XML 文件中的 `ItemHasRegularExpressionMatch` 规则元素中定义的正则表达式的字符串。</span><span class="sxs-lookup"><span data-stu-id="9a31e-1287">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="9a31e-p171">如果在项目的正文属性上指定 `ItemHasRegularExpressionMatch` 规则，则正则表达式应进一步筛选正文，不应尝试返回该项目的整个正文。使用正则表达式（如 `.*`）获取项目的整个正文并不总是返回预期的结果。</span><span class="sxs-lookup"><span data-stu-id="9a31e-p171">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="9a31e-1290">参数</span><span class="sxs-lookup"><span data-stu-id="9a31e-1290">Parameters</span></span>

|<span data-ttu-id="9a31e-1291">名称</span><span class="sxs-lookup"><span data-stu-id="9a31e-1291">Name</span></span>|<span data-ttu-id="9a31e-1292">类型</span><span class="sxs-lookup"><span data-stu-id="9a31e-1292">Type</span></span>|<span data-ttu-id="9a31e-1293">说明</span><span class="sxs-lookup"><span data-stu-id="9a31e-1293">Description</span></span>|
|---|---|---|
|`name`|<span data-ttu-id="9a31e-1294">字符串</span><span class="sxs-lookup"><span data-stu-id="9a31e-1294">String</span></span>|<span data-ttu-id="9a31e-1295">定义筛选器匹配的 `ItemHasRegularExpressionMatch` 规则元素的名称。</span><span class="sxs-lookup"><span data-stu-id="9a31e-1295">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="9a31e-1296">要求</span><span class="sxs-lookup"><span data-stu-id="9a31e-1296">Requirements</span></span>

|<span data-ttu-id="9a31e-1297">要求</span><span class="sxs-lookup"><span data-stu-id="9a31e-1297">Requirement</span></span>|<span data-ttu-id="9a31e-1298">值</span><span class="sxs-lookup"><span data-stu-id="9a31e-1298">Value</span></span>|
|---|---|
|[<span data-ttu-id="9a31e-1299">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="9a31e-1299">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="9a31e-1300">1.0</span><span class="sxs-lookup"><span data-stu-id="9a31e-1300">1.0</span></span>|
|[<span data-ttu-id="9a31e-1301">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="9a31e-1301">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="9a31e-1302">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9a31e-1302">ReadItem</span></span>|
|[<span data-ttu-id="9a31e-1303">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="9a31e-1303">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="9a31e-1304">阅读</span><span class="sxs-lookup"><span data-stu-id="9a31e-1304">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="9a31e-1305">返回：</span><span class="sxs-lookup"><span data-stu-id="9a31e-1305">Returns:</span></span>

<span data-ttu-id="9a31e-1306">一个包含匹配在清单 XML 文件中定义的正则表达式的字符串的数组。</span><span class="sxs-lookup"><span data-stu-id="9a31e-1306">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<span data-ttu-id="9a31e-1307">类型： Array. < 字符串 ></span><span class="sxs-lookup"><span data-stu-id="9a31e-1307">Type: Array.< String ></span></span>

##### <a name="example"></a><span data-ttu-id="9a31e-1308">示例</span><span class="sxs-lookup"><span data-stu-id="9a31e-1308">Example</span></span>

```js
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

<br>

---
---

#### <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="9a31e-1309">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="9a31e-1309">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="9a31e-1310">以异步方式返回邮件的主题或正文中选定的数据。</span><span class="sxs-lookup"><span data-stu-id="9a31e-1310">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="9a31e-p172">如果没有选定内容，但光标位于正文或主题中，此方法将会为所选数据返回 null。如果选定的是字段，而不是正文或主题，则此方法返回 `InvalidSelection` 错误。</span><span class="sxs-lookup"><span data-stu-id="9a31e-p172">If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="9a31e-1313">参数</span><span class="sxs-lookup"><span data-stu-id="9a31e-1313">Parameters</span></span>

|<span data-ttu-id="9a31e-1314">名称</span><span class="sxs-lookup"><span data-stu-id="9a31e-1314">Name</span></span>|<span data-ttu-id="9a31e-1315">类型</span><span class="sxs-lookup"><span data-stu-id="9a31e-1315">Type</span></span>|<span data-ttu-id="9a31e-1316">属性</span><span class="sxs-lookup"><span data-stu-id="9a31e-1316">Attributes</span></span>|<span data-ttu-id="9a31e-1317">说明</span><span class="sxs-lookup"><span data-stu-id="9a31e-1317">Description</span></span>|
|---|---|---|---|
|`coercionType`|[<span data-ttu-id="9a31e-1318">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="9a31e-1318">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="9a31e-p173">请求数据的格式。如果为文本，则此方法返回纯文本作为字符串，删除任何显示的 HTML 标记。如果为 HTML，则此方法返回所选文本，不论是纯文本还是 HTML。</span><span class="sxs-lookup"><span data-stu-id="9a31e-p173">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`|<span data-ttu-id="9a31e-1322">对象</span><span class="sxs-lookup"><span data-stu-id="9a31e-1322">Object</span></span>|<span data-ttu-id="9a31e-1323">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="9a31e-1323">&lt;optional&gt;</span></span>|<span data-ttu-id="9a31e-1324">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="9a31e-1324">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="9a31e-1325">Object</span><span class="sxs-lookup"><span data-stu-id="9a31e-1325">Object</span></span>|<span data-ttu-id="9a31e-1326">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="9a31e-1326">&lt;optional&gt;</span></span>|<span data-ttu-id="9a31e-1327">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="9a31e-1327">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="9a31e-1328">函数</span><span class="sxs-lookup"><span data-stu-id="9a31e-1328">function</span></span>||<span data-ttu-id="9a31e-1329">方法完成后，使用单个参数 `callback`（一个 [`asyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `AsyncResult` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="9a31e-1329">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="9a31e-1330">若要从回调方法访问所选数据，请调用 `asyncResult.value.data`。</span><span class="sxs-lookup"><span data-stu-id="9a31e-1330">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="9a31e-1331">若要访问选定内容的源属性，请调用 `asyncResult.value.sourceProperty`，这将为 `body` 或 `subject`。</span><span class="sxs-lookup"><span data-stu-id="9a31e-1331">To access the source property that the selection comes from, call `asyncResult.value.sourceProperty`, which will be either `body` or `subject`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="9a31e-1332">要求</span><span class="sxs-lookup"><span data-stu-id="9a31e-1332">Requirements</span></span>

|<span data-ttu-id="9a31e-1333">要求</span><span class="sxs-lookup"><span data-stu-id="9a31e-1333">Requirement</span></span>|<span data-ttu-id="9a31e-1334">值</span><span class="sxs-lookup"><span data-stu-id="9a31e-1334">Value</span></span>|
|---|---|
|[<span data-ttu-id="9a31e-1335">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="9a31e-1335">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="9a31e-1336">1.2</span><span class="sxs-lookup"><span data-stu-id="9a31e-1336">1.2</span></span>|
|[<span data-ttu-id="9a31e-1337">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="9a31e-1337">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="9a31e-1338">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9a31e-1338">ReadItem</span></span>|
|[<span data-ttu-id="9a31e-1339">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="9a31e-1339">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="9a31e-1340">撰写</span><span class="sxs-lookup"><span data-stu-id="9a31e-1340">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="9a31e-1341">返回：</span><span class="sxs-lookup"><span data-stu-id="9a31e-1341">Returns:</span></span>

<span data-ttu-id="9a31e-1342">作为字符串的所选数据的格式由 `coercionType` 确定。</span><span class="sxs-lookup"><span data-stu-id="9a31e-1342">The selected data as a string with format determined by `coercionType`.</span></span>

<span data-ttu-id="9a31e-1343">类型：字符串</span><span class="sxs-lookup"><span data-stu-id="9a31e-1343">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="9a31e-1344">示例</span><span class="sxs-lookup"><span data-stu-id="9a31e-1344">Example</span></span>

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

#### <a name="getselectedentities--entitiesjavascriptapioutlookofficeentities"></a><span data-ttu-id="9a31e-1345">getSelectedEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="9a31e-1345">getSelectedEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span></span>

<span data-ttu-id="9a31e-1346">获取在用户已选择的突出显示匹配项中找到的实体。</span><span class="sxs-lookup"><span data-stu-id="9a31e-1346">Gets the entities found in a highlighted match a user has selected.</span></span> <span data-ttu-id="9a31e-1347">突出显示匹配项适用于[上下文外接程序](/outlook/add-ins/contextual-outlook-add-ins)。</span><span class="sxs-lookup"><span data-stu-id="9a31e-1347">Highlighted matches apply to [contextual add-ins](/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="9a31e-1348">IOS 或 Android 上的 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="9a31e-1348">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="9a31e-1349">要求</span><span class="sxs-lookup"><span data-stu-id="9a31e-1349">Requirements</span></span>

|<span data-ttu-id="9a31e-1350">要求</span><span class="sxs-lookup"><span data-stu-id="9a31e-1350">Requirement</span></span>|<span data-ttu-id="9a31e-1351">值</span><span class="sxs-lookup"><span data-stu-id="9a31e-1351">Value</span></span>|
|---|---|
|[<span data-ttu-id="9a31e-1352">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="9a31e-1352">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="9a31e-1353">1.6</span><span class="sxs-lookup"><span data-stu-id="9a31e-1353">1.6</span></span>|
|[<span data-ttu-id="9a31e-1354">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="9a31e-1354">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="9a31e-1355">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9a31e-1355">ReadItem</span></span>|
|[<span data-ttu-id="9a31e-1356">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="9a31e-1356">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="9a31e-1357">阅读</span><span class="sxs-lookup"><span data-stu-id="9a31e-1357">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="9a31e-1358">返回：</span><span class="sxs-lookup"><span data-stu-id="9a31e-1358">Returns:</span></span>

<span data-ttu-id="9a31e-1359">类型：[Entities](/javascript/api/outlook/office.entities)</span><span class="sxs-lookup"><span data-stu-id="9a31e-1359">Type: [Entities](/javascript/api/outlook/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="9a31e-1360">示例</span><span class="sxs-lookup"><span data-stu-id="9a31e-1360">Example</span></span>

<span data-ttu-id="9a31e-1361">以下示例访问用户选择的突出显示匹配项中的地址实体。</span><span class="sxs-lookup"><span data-stu-id="9a31e-1361">The following example accesses the addresses entities in the highlighted match selected by the user.</span></span>

```js
var contacts = Office.context.mailbox.item.getSelectedEntities().addresses;
```

<br>

---
---

#### <a name="getselectedregexmatches--object"></a><span data-ttu-id="9a31e-1362">getSelectedRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="9a31e-1362">getSelectedRegExMatches() → {Object}</span></span>

<span data-ttu-id="9a31e-p176">返回突出显示匹配项中匹配在清单 XML 文件中定义的正则表达式的字符串值。突出显示匹配项适用于[上下文外接程序](/outlook/add-ins/contextual-outlook-add-ins)。</span><span class="sxs-lookup"><span data-stu-id="9a31e-p176">Returns string values in a highlighted match that match the regular expressions defined in the manifest XML file. Highlighted matches apply to [contextual add-ins](/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="9a31e-1365">IOS 或 Android 上的 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="9a31e-1365">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="9a31e-p177">`getSelectedRegExMatches` 方法返回匹配在清单 XML 文件中的每个 `ItemHasRegularExpressionMatch` 或 `ItemHasKnownEntity` 规则元素中定义的正则表达式的字符串。对于 `ItemHasRegularExpressionMatch` 规则，匹配字符串必须发生在该规则指定的项目的属性中。`PropertyName` 简单类型定义支持的属性。</span><span class="sxs-lookup"><span data-stu-id="9a31e-p177">The `getSelectedRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="9a31e-1369">例如，考虑一个外接程序清单具有以下 `Rule` 元素：</span><span class="sxs-lookup"><span data-stu-id="9a31e-1369">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="9a31e-1370">从 `getRegExMatches` 返回的对象应有两个属性：`fruits` 和 `veggies`。</span><span class="sxs-lookup"><span data-stu-id="9a31e-1370">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="9a31e-p178">如果在项目的正文属性上指定 `ItemHasRegularExpressionMatch` 规则，则正则表达式应进一步筛选正文，不应尝试返回该项目的整个正文。使用正则表达式（如 `.*`）获取项目的整个正文并不总是返回预期的结果。而是使用 [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) 方法检索整个正文。</span><span class="sxs-lookup"><span data-stu-id="9a31e-p178">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="9a31e-1374">Requirements</span><span class="sxs-lookup"><span data-stu-id="9a31e-1374">Requirements</span></span>

|<span data-ttu-id="9a31e-1375">要求</span><span class="sxs-lookup"><span data-stu-id="9a31e-1375">Requirement</span></span>|<span data-ttu-id="9a31e-1376">值</span><span class="sxs-lookup"><span data-stu-id="9a31e-1376">Value</span></span>|
|---|---|
|[<span data-ttu-id="9a31e-1377">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="9a31e-1377">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="9a31e-1378">1.6</span><span class="sxs-lookup"><span data-stu-id="9a31e-1378">1.6</span></span>|
|[<span data-ttu-id="9a31e-1379">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="9a31e-1379">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="9a31e-1380">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9a31e-1380">ReadItem</span></span>|
|[<span data-ttu-id="9a31e-1381">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="9a31e-1381">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="9a31e-1382">阅读</span><span class="sxs-lookup"><span data-stu-id="9a31e-1382">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="9a31e-1383">返回：</span><span class="sxs-lookup"><span data-stu-id="9a31e-1383">Returns:</span></span>

<span data-ttu-id="9a31e-p179">一个包含匹配在清单 XML 文件中定义的正则表达式的字符串数组的对象。每个数组的名称等于匹配 `ItemHasRegularExpressionMatch` 规则的 `RegExName` 属性或匹配 `ItemHasKnownEntity` 规则的 `FilterName` 属性的相应值。</span><span class="sxs-lookup"><span data-stu-id="9a31e-p179">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

##### <a name="example"></a><span data-ttu-id="9a31e-1386">示例</span><span class="sxs-lookup"><span data-stu-id="9a31e-1386">Example</span></span>

<span data-ttu-id="9a31e-1387">以下示例显示了如何访问正则表达式规则元素 `fruits` 和 `veggies` 的匹配项的数组，这些元素在清单中指定。</span><span class="sxs-lookup"><span data-stu-id="9a31e-1387">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```js
var selectedMatches = Office.context.mailbox.item.getSelectedRegExMatches();
var fruits = selectedMatches.fruits;
var veggies = selectedMatches.veggies;
```

<br>

---
---

#### <a name="getsharedpropertiesasyncoptions-callback"></a><span data-ttu-id="9a31e-1388">getSharedPropertiesAsync （[options]，回拨）</span><span class="sxs-lookup"><span data-stu-id="9a31e-1388">getSharedPropertiesAsync([options], callback)</span></span>

<span data-ttu-id="9a31e-1389">获取共享文件夹、日历或邮箱中的所选约会或邮件的属性。</span><span class="sxs-lookup"><span data-stu-id="9a31e-1389">Gets the properties of the selected appointment or message in a shared folder, calendar, or mailbox.</span></span>

##### <a name="parameters"></a><span data-ttu-id="9a31e-1390">参数</span><span class="sxs-lookup"><span data-stu-id="9a31e-1390">Parameters</span></span>

|<span data-ttu-id="9a31e-1391">名称</span><span class="sxs-lookup"><span data-stu-id="9a31e-1391">Name</span></span>|<span data-ttu-id="9a31e-1392">类型</span><span class="sxs-lookup"><span data-stu-id="9a31e-1392">Type</span></span>|<span data-ttu-id="9a31e-1393">属性</span><span class="sxs-lookup"><span data-stu-id="9a31e-1393">Attributes</span></span>|<span data-ttu-id="9a31e-1394">说明</span><span class="sxs-lookup"><span data-stu-id="9a31e-1394">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="9a31e-1395">对象</span><span class="sxs-lookup"><span data-stu-id="9a31e-1395">Object</span></span>|<span data-ttu-id="9a31e-1396">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="9a31e-1396">&lt;optional&gt;</span></span>|<span data-ttu-id="9a31e-1397">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="9a31e-1397">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="9a31e-1398">对象</span><span class="sxs-lookup"><span data-stu-id="9a31e-1398">Object</span></span>|<span data-ttu-id="9a31e-1399">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="9a31e-1399">&lt;optional&gt;</span></span>|<span data-ttu-id="9a31e-1400">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="9a31e-1400">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="9a31e-1401">function</span><span class="sxs-lookup"><span data-stu-id="9a31e-1401">function</span></span>||<span data-ttu-id="9a31e-1402">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="9a31e-1402">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="9a31e-1403">共享属性作为[`SharedProperties`](/javascript/api/outlook/office.sharedproperties) `asyncResult.value`属性中的对象提供。</span><span class="sxs-lookup"><span data-stu-id="9a31e-1403">The shared properties are provided as a [`SharedProperties`](/javascript/api/outlook/office.sharedproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="9a31e-1404">此对象可用于获取项目的共享属性。</span><span class="sxs-lookup"><span data-stu-id="9a31e-1404">This object can be used to get the item's shared properties.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="9a31e-1405">要求</span><span class="sxs-lookup"><span data-stu-id="9a31e-1405">Requirements</span></span>

|<span data-ttu-id="9a31e-1406">要求</span><span class="sxs-lookup"><span data-stu-id="9a31e-1406">Requirement</span></span>|<span data-ttu-id="9a31e-1407">值</span><span class="sxs-lookup"><span data-stu-id="9a31e-1407">Value</span></span>|
|---|---|
|[<span data-ttu-id="9a31e-1408">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="9a31e-1408">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="9a31e-1409">预览</span><span class="sxs-lookup"><span data-stu-id="9a31e-1409">Preview</span></span>|
|[<span data-ttu-id="9a31e-1410">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="9a31e-1410">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="9a31e-1411">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9a31e-1411">ReadItem</span></span>|
|[<span data-ttu-id="9a31e-1412">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="9a31e-1412">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="9a31e-1413">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="9a31e-1413">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="9a31e-1414">示例</span><span class="sxs-lookup"><span data-stu-id="9a31e-1414">Example</span></span>

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

#### <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="9a31e-1415">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="9a31e-1415">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="9a31e-1416">异步加载所选项目上此外接程序的自定义属性。</span><span class="sxs-lookup"><span data-stu-id="9a31e-1416">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="9a31e-p181">自定义属性基于每个应用、每个项目存储为键/值对。此方法在回调中返回 `CustomProperties` 对象，该回调提供访问特定于当前项目和当前外接程序的自定义属性的方法。自定义属性未在项目上加密，因此这不应用作安全存储。</span><span class="sxs-lookup"><span data-stu-id="9a31e-p181">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="9a31e-1420">参数</span><span class="sxs-lookup"><span data-stu-id="9a31e-1420">Parameters</span></span>

|<span data-ttu-id="9a31e-1421">名称</span><span class="sxs-lookup"><span data-stu-id="9a31e-1421">Name</span></span>|<span data-ttu-id="9a31e-1422">类型</span><span class="sxs-lookup"><span data-stu-id="9a31e-1422">Type</span></span>|<span data-ttu-id="9a31e-1423">属性</span><span class="sxs-lookup"><span data-stu-id="9a31e-1423">Attributes</span></span>|<span data-ttu-id="9a31e-1424">说明</span><span class="sxs-lookup"><span data-stu-id="9a31e-1424">Description</span></span>|
|---|---|---|---|
|`callback`|<span data-ttu-id="9a31e-1425">函数</span><span class="sxs-lookup"><span data-stu-id="9a31e-1425">function</span></span>||<span data-ttu-id="9a31e-1426">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="9a31e-1426">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="9a31e-1427">自定义属性作为 `asyncResult.value` 属性中的 [`CustomProperties`](/javascript/api/outlook/office.customproperties) 对象提供。</span><span class="sxs-lookup"><span data-stu-id="9a31e-1427">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook/office.customproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="9a31e-1428">此对象可用于获取、设置以及从项目中删除自定义属性，并将自定义属性集的更改重新保存到服务器。</span><span class="sxs-lookup"><span data-stu-id="9a31e-1428">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`|<span data-ttu-id="9a31e-1429">对象</span><span class="sxs-lookup"><span data-stu-id="9a31e-1429">Object</span></span>|<span data-ttu-id="9a31e-1430">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="9a31e-1430">&lt;optional&gt;</span></span>|<span data-ttu-id="9a31e-1431">开发人员可以提供他们想要在回调函数中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="9a31e-1431">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="9a31e-1432">此对象可以通过回调函数中的 `asyncResult.asyncContext` 属性进行访问。</span><span class="sxs-lookup"><span data-stu-id="9a31e-1432">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="9a31e-1433">要求</span><span class="sxs-lookup"><span data-stu-id="9a31e-1433">Requirements</span></span>

|<span data-ttu-id="9a31e-1434">要求</span><span class="sxs-lookup"><span data-stu-id="9a31e-1434">Requirement</span></span>|<span data-ttu-id="9a31e-1435">值</span><span class="sxs-lookup"><span data-stu-id="9a31e-1435">Value</span></span>|
|---|---|
|[<span data-ttu-id="9a31e-1436">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="9a31e-1436">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="9a31e-1437">1.0</span><span class="sxs-lookup"><span data-stu-id="9a31e-1437">1.0</span></span>|
|[<span data-ttu-id="9a31e-1438">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="9a31e-1438">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="9a31e-1439">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9a31e-1439">ReadItem</span></span>|
|[<span data-ttu-id="9a31e-1440">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="9a31e-1440">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="9a31e-1441">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="9a31e-1441">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="9a31e-1442">示例</span><span class="sxs-lookup"><span data-stu-id="9a31e-1442">Example</span></span>

<span data-ttu-id="9a31e-p184">以下代码示例显示了如何使用 `loadCustomPropertiesAsync` 方法异步加载特定于当前项目的自定义属性。该示例还显示了如何使用 `CustomProperties.saveAsync` 方法将这些属性重新保存到服务器。加载自定义属性后，该代码示例将使用 `CustomProperties.get` 方法读取自定义属性 `myProp`，使用 `CustomProperties.set` 方法写入自定义属性 `otherProp`，最后调用 `saveAsync` 方法保存这些自定义属性。</span><span class="sxs-lookup"><span data-stu-id="9a31e-p184">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

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

#### <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="9a31e-1446">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="9a31e-1446">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="9a31e-1447">将附件从邮件或约会中删除。</span><span class="sxs-lookup"><span data-stu-id="9a31e-1447">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="9a31e-1448">`removeAttachmentAsync` 方法删除项目中带指定标识符的附件。</span><span class="sxs-lookup"><span data-stu-id="9a31e-1448">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item.</span></span> <span data-ttu-id="9a31e-1449">最佳做法是，仅当同一个邮件应用程序在同一会话中添加了一个附件时，你才应使用该附件标识符来删除该附件。</span><span class="sxs-lookup"><span data-stu-id="9a31e-1449">As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session.</span></span> <span data-ttu-id="9a31e-1450">在 web 和移动设备上的 Outlook 中，附件标识符仅在同一个会话中有效。</span><span class="sxs-lookup"><span data-stu-id="9a31e-1450">In Outlook on the web and mobile devices, the attachment identifier is valid only within the same session.</span></span> <span data-ttu-id="9a31e-1451">当用户关闭应用程序时，或者如果用户开始撰写内嵌窗体，随后弹出窗体以继续在单独的窗口中。</span><span class="sxs-lookup"><span data-stu-id="9a31e-1451">A session is over when the user closes the app, or if the user starts composing an inline form then subsequently pops out the form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="9a31e-1452">参数</span><span class="sxs-lookup"><span data-stu-id="9a31e-1452">Parameters</span></span>

|<span data-ttu-id="9a31e-1453">名称</span><span class="sxs-lookup"><span data-stu-id="9a31e-1453">Name</span></span>|<span data-ttu-id="9a31e-1454">类型</span><span class="sxs-lookup"><span data-stu-id="9a31e-1454">Type</span></span>|<span data-ttu-id="9a31e-1455">属性</span><span class="sxs-lookup"><span data-stu-id="9a31e-1455">Attributes</span></span>|<span data-ttu-id="9a31e-1456">说明</span><span class="sxs-lookup"><span data-stu-id="9a31e-1456">Description</span></span>|
|---|---|---|---|
|`attachmentId`|<span data-ttu-id="9a31e-1457">字符串</span><span class="sxs-lookup"><span data-stu-id="9a31e-1457">String</span></span>||<span data-ttu-id="9a31e-1458">要删除的附件的标识符。</span><span class="sxs-lookup"><span data-stu-id="9a31e-1458">The identifier of the attachment to remove.</span></span>|
|`options`|<span data-ttu-id="9a31e-1459">对象</span><span class="sxs-lookup"><span data-stu-id="9a31e-1459">Object</span></span>|<span data-ttu-id="9a31e-1460">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="9a31e-1460">&lt;optional&gt;</span></span>|<span data-ttu-id="9a31e-1461">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="9a31e-1461">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="9a31e-1462">对象</span><span class="sxs-lookup"><span data-stu-id="9a31e-1462">Object</span></span>|<span data-ttu-id="9a31e-1463">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="9a31e-1463">&lt;optional&gt;</span></span>|<span data-ttu-id="9a31e-1464">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="9a31e-1464">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="9a31e-1465">函数</span><span class="sxs-lookup"><span data-stu-id="9a31e-1465">function</span></span>|<span data-ttu-id="9a31e-1466">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="9a31e-1466">&lt;optional&gt;</span></span>|<span data-ttu-id="9a31e-1467">方法完成后，使用单个参数 `callback`（一个 [`asyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `AsyncResult` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="9a31e-1467">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="9a31e-1468">如果删除附件失败，`asyncResult.error` 属性将包含一个说明失败原因的错误代码。</span><span class="sxs-lookup"><span data-stu-id="9a31e-1468">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="9a31e-1469">错误</span><span class="sxs-lookup"><span data-stu-id="9a31e-1469">Errors</span></span>

|<span data-ttu-id="9a31e-1470">错误代码</span><span class="sxs-lookup"><span data-stu-id="9a31e-1470">Error code</span></span>|<span data-ttu-id="9a31e-1471">说明</span><span class="sxs-lookup"><span data-stu-id="9a31e-1471">Description</span></span>|
|------------|-------------|
|`InvalidAttachmentId`|<span data-ttu-id="9a31e-1472">附件标识符不存在。</span><span class="sxs-lookup"><span data-stu-id="9a31e-1472">The attachment identifier does not exist.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="9a31e-1473">要求</span><span class="sxs-lookup"><span data-stu-id="9a31e-1473">Requirements</span></span>

|<span data-ttu-id="9a31e-1474">要求</span><span class="sxs-lookup"><span data-stu-id="9a31e-1474">Requirement</span></span>|<span data-ttu-id="9a31e-1475">值</span><span class="sxs-lookup"><span data-stu-id="9a31e-1475">Value</span></span>|
|---|---|
|[<span data-ttu-id="9a31e-1476">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="9a31e-1476">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="9a31e-1477">1.1</span><span class="sxs-lookup"><span data-stu-id="9a31e-1477">1.1</span></span>|
|[<span data-ttu-id="9a31e-1478">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="9a31e-1478">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="9a31e-1479">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="9a31e-1479">ReadWriteItem</span></span>|
|[<span data-ttu-id="9a31e-1480">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="9a31e-1480">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="9a31e-1481">撰写</span><span class="sxs-lookup"><span data-stu-id="9a31e-1481">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="9a31e-1482">示例</span><span class="sxs-lookup"><span data-stu-id="9a31e-1482">Example</span></span>

<span data-ttu-id="9a31e-1483">以下代码删除包含标识符 '0' 的附件。</span><span class="sxs-lookup"><span data-stu-id="9a31e-1483">The following code removes an attachment with an identifier of '0'.</span></span>

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

#### <a name="removehandlerasynceventtype-options-callback"></a><span data-ttu-id="9a31e-1484">removeHandlerAsync(eventType, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="9a31e-1484">removeHandlerAsync(eventType, [options], [callback])</span></span>

<span data-ttu-id="9a31e-1485">删除受支持事件类型的事件处理程序。</span><span class="sxs-lookup"><span data-stu-id="9a31e-1485">Removes the event handlers for a supported event type.</span></span>

<span data-ttu-id="9a31e-1486">目前，受支持的事件`Office.EventType.AttachmentsChanged`类型`Office.EventType.AppointmentTimeChanged`是`Office.EventType.EnhancedLocationsChanged`、 `Office.EventType.RecipientsChanged`、、 `Office.EventType.RecurrenceChanged`和。</span><span class="sxs-lookup"><span data-stu-id="9a31e-1486">Currently the supported event types are `Office.EventType.AttachmentsChanged`, `Office.EventType.AppointmentTimeChanged`, `Office.EventType.EnhancedLocationsChanged`, `Office.EventType.RecipientsChanged`, and `Office.EventType.RecurrenceChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="9a31e-1487">参数</span><span class="sxs-lookup"><span data-stu-id="9a31e-1487">Parameters</span></span>

| <span data-ttu-id="9a31e-1488">名称</span><span class="sxs-lookup"><span data-stu-id="9a31e-1488">Name</span></span> | <span data-ttu-id="9a31e-1489">类型</span><span class="sxs-lookup"><span data-stu-id="9a31e-1489">Type</span></span> | <span data-ttu-id="9a31e-1490">属性</span><span class="sxs-lookup"><span data-stu-id="9a31e-1490">Attributes</span></span> | <span data-ttu-id="9a31e-1491">说明</span><span class="sxs-lookup"><span data-stu-id="9a31e-1491">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="9a31e-1492">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="9a31e-1492">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="9a31e-1493">应撤销处理程序的事件。</span><span class="sxs-lookup"><span data-stu-id="9a31e-1493">The event that should revoke the handler.</span></span> |
| `options` | <span data-ttu-id="9a31e-1494">对象</span><span class="sxs-lookup"><span data-stu-id="9a31e-1494">Object</span></span> | <span data-ttu-id="9a31e-1495">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="9a31e-1495">&lt;optional&gt;</span></span> | <span data-ttu-id="9a31e-1496">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="9a31e-1496">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="9a31e-1497">对象</span><span class="sxs-lookup"><span data-stu-id="9a31e-1497">Object</span></span> | <span data-ttu-id="9a31e-1498">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="9a31e-1498">&lt;optional&gt;</span></span> | <span data-ttu-id="9a31e-1499">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="9a31e-1499">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="9a31e-1500">函数</span><span class="sxs-lookup"><span data-stu-id="9a31e-1500">function</span></span>| <span data-ttu-id="9a31e-1501">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="9a31e-1501">&lt;optional&gt;</span></span>|<span data-ttu-id="9a31e-1502">方法完成后，使用单个参数 `callback`（一个 [`asyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `AsyncResult` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="9a31e-1502">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="9a31e-1503">Requirements</span><span class="sxs-lookup"><span data-stu-id="9a31e-1503">Requirements</span></span>

|<span data-ttu-id="9a31e-1504">要求</span><span class="sxs-lookup"><span data-stu-id="9a31e-1504">Requirement</span></span>| <span data-ttu-id="9a31e-1505">值</span><span class="sxs-lookup"><span data-stu-id="9a31e-1505">Value</span></span>|
|---|---|
|[<span data-ttu-id="9a31e-1506">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="9a31e-1506">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9a31e-1507">1.7</span><span class="sxs-lookup"><span data-stu-id="9a31e-1507">1.7</span></span> |
|[<span data-ttu-id="9a31e-1508">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="9a31e-1508">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9a31e-1509">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9a31e-1509">ReadItem</span></span> |
|[<span data-ttu-id="9a31e-1510">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="9a31e-1510">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="9a31e-1511">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="9a31e-1511">Compose or Read</span></span> |

<br>

---
---

#### <a name="saveasyncoptions-callback"></a><span data-ttu-id="9a31e-1512">saveAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="9a31e-1512">saveAsync([options], callback)</span></span>

<span data-ttu-id="9a31e-1513">异步保存项目。</span><span class="sxs-lookup"><span data-stu-id="9a31e-1513">Asynchronously saves an item.</span></span>

<span data-ttu-id="9a31e-1514">调用时，此方法将当前邮件保存为草稿，并通过回调方法返回项目 ID。</span><span class="sxs-lookup"><span data-stu-id="9a31e-1514">When invoked, this method saves the current message as a draft and returns the item id via the callback method.</span></span> <span data-ttu-id="9a31e-1515">在 Outlook 网页或 Outlook 的联机模式中，将项目保存到服务器。</span><span class="sxs-lookup"><span data-stu-id="9a31e-1515">In Outlook on the web or Outlook in online mode, the item is saved to the server.</span></span> <span data-ttu-id="9a31e-1516">在 Outlook 缓存模式下，该项目被保存到本地缓存中。</span><span class="sxs-lookup"><span data-stu-id="9a31e-1516">In Outlook in cached mode, the item is saved to the local cache.</span></span>

> [!NOTE]
> <span data-ttu-id="9a31e-1517">如果加载项在撰写模式下对某个项目调用 `saveAsync` 来获得 `itemId`，以便与 EWS 或 REST API 一同使用，请注意，当 Outlook 处于高速缓存模式时，可能需要一段时间项目才能真正同步到服务器。</span><span class="sxs-lookup"><span data-stu-id="9a31e-1517">If your add-in calls `saveAsync` on an item in compose mode in order to get an `itemId` to use with EWS or the REST API, be aware that when Outlook is in cached mode, it may take some time before the item is actually synced to the server.</span></span> <span data-ttu-id="9a31e-1518">在项目同步前，使用 `itemId` 将返回一个错误。</span><span class="sxs-lookup"><span data-stu-id="9a31e-1518">Until the item is synced, using the `itemId` will return an error.</span></span>

<span data-ttu-id="9a31e-p188">由于约会没有草稿状态，如果以撰写模式在约会中调用 `saveAsync`，则该项将被保存为用户日历中的正常约会。对于之前未保存过的新约会，则不会发送邀请。保存现有约会将向添加或删除的与会者发送更新。</span><span class="sxs-lookup"><span data-stu-id="9a31e-p188">Since appointments have no draft state, if `saveAsync` is called on an appointment in compose mode, the item will be saved as a normal appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent. Saving an existing appointment will send an update to added or removed attendees.</span></span>

> [!NOTE]
> <span data-ttu-id="9a31e-1522">以下客户端在撰写模式下对约会上的 `saveAsync` 具有不同的行为：</span><span class="sxs-lookup"><span data-stu-id="9a31e-1522">The following clients have different behavior for `saveAsync` on appointments in compose mode:</span></span>
>
> - <span data-ttu-id="9a31e-1523">Mac 上的 Outlook 不支持保存会议。</span><span class="sxs-lookup"><span data-stu-id="9a31e-1523">Outlook on Mac does not support saving a meeting.</span></span> <span data-ttu-id="9a31e-1524">在`saveAsync`撰写模式下从会议中调用时，此方法将失败。</span><span class="sxs-lookup"><span data-stu-id="9a31e-1524">The `saveAsync` method fails when called from a meeting in compose mode.</span></span> <span data-ttu-id="9a31e-1525">若要解决此问题，请参阅[使用 OFFICE JS API 将会议保存为 Outlook For Mac 中的草稿](https://support.microsoft.com/help/4505745)。</span><span class="sxs-lookup"><span data-stu-id="9a31e-1525">See [Cannot save a meeting as a draft in Outlook for Mac by using Office JS API](https://support.microsoft.com/help/4505745) for a workaround.</span></span>
> - <span data-ttu-id="9a31e-1526">在撰写模式下的约会上调用 `saveAsync` 时，Outlook 网页版始终发送邀请或更新。</span><span class="sxs-lookup"><span data-stu-id="9a31e-1526">Outlook on the web always sends an invitation or update when `saveAsync` is called on an appointment in compose mode.</span></span>

##### <a name="parameters"></a><span data-ttu-id="9a31e-1527">参数</span><span class="sxs-lookup"><span data-stu-id="9a31e-1527">Parameters</span></span>

|<span data-ttu-id="9a31e-1528">名称</span><span class="sxs-lookup"><span data-stu-id="9a31e-1528">Name</span></span>|<span data-ttu-id="9a31e-1529">类型</span><span class="sxs-lookup"><span data-stu-id="9a31e-1529">Type</span></span>|<span data-ttu-id="9a31e-1530">属性</span><span class="sxs-lookup"><span data-stu-id="9a31e-1530">Attributes</span></span>|<span data-ttu-id="9a31e-1531">说明</span><span class="sxs-lookup"><span data-stu-id="9a31e-1531">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="9a31e-1532">对象</span><span class="sxs-lookup"><span data-stu-id="9a31e-1532">Object</span></span>|<span data-ttu-id="9a31e-1533">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="9a31e-1533">&lt;optional&gt;</span></span>|<span data-ttu-id="9a31e-1534">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="9a31e-1534">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="9a31e-1535">对象</span><span class="sxs-lookup"><span data-stu-id="9a31e-1535">Object</span></span>|<span data-ttu-id="9a31e-1536">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="9a31e-1536">&lt;optional&gt;</span></span>|<span data-ttu-id="9a31e-1537">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="9a31e-1537">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="9a31e-1538">函数</span><span class="sxs-lookup"><span data-stu-id="9a31e-1538">function</span></span>||<span data-ttu-id="9a31e-1539">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="9a31e-1539">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="9a31e-1540">如果成功，该项目标识符将在 `asyncResult.value` 属性中提供。</span><span class="sxs-lookup"><span data-stu-id="9a31e-1540">On success, the item identifier is provided in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="9a31e-1541">要求</span><span class="sxs-lookup"><span data-stu-id="9a31e-1541">Requirements</span></span>

|<span data-ttu-id="9a31e-1542">要求</span><span class="sxs-lookup"><span data-stu-id="9a31e-1542">Requirement</span></span>|<span data-ttu-id="9a31e-1543">值</span><span class="sxs-lookup"><span data-stu-id="9a31e-1543">Value</span></span>|
|---|---|
|[<span data-ttu-id="9a31e-1544">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="9a31e-1544">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="9a31e-1545">1.3</span><span class="sxs-lookup"><span data-stu-id="9a31e-1545">1.3</span></span>|
|[<span data-ttu-id="9a31e-1546">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="9a31e-1546">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="9a31e-1547">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="9a31e-1547">ReadWriteItem</span></span>|
|[<span data-ttu-id="9a31e-1548">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="9a31e-1548">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="9a31e-1549">撰写</span><span class="sxs-lookup"><span data-stu-id="9a31e-1549">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="9a31e-1550">示例</span><span class="sxs-lookup"><span data-stu-id="9a31e-1550">Examples</span></span>

```js
Office.context.mailbox.item.saveAsync(
  function callback(result) {
    // Process the result.
  });
```

<span data-ttu-id="9a31e-p190">下面是传递给回调函数的 `result` 参数的示例。`value` 属性包含的项目的项目 ID。</span><span class="sxs-lookup"><span data-stu-id="9a31e-p190">The following is an example of the `result` parameter passed to the callback function. The `value` property contains the item ID of the item.</span></span>

```json
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

<br>

---
---

#### <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="9a31e-1553">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="9a31e-1553">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="9a31e-1554">以异步方式将数据插入到邮件的正文或主题中。</span><span class="sxs-lookup"><span data-stu-id="9a31e-1554">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="9a31e-p191">`setSelectedDataAsync` 方法将指定字符串插入到项目主题或正文的光标位置，或者，如果在编辑器中已选择文本，则该方法将替换选择的文本。如果光标不在正文或主题字段中，则返回错误。插入之后，光标会位于插入内容的末尾。</span><span class="sxs-lookup"><span data-stu-id="9a31e-p191">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="9a31e-1558">参数</span><span class="sxs-lookup"><span data-stu-id="9a31e-1558">Parameters</span></span>

|<span data-ttu-id="9a31e-1559">名称</span><span class="sxs-lookup"><span data-stu-id="9a31e-1559">Name</span></span>|<span data-ttu-id="9a31e-1560">类型</span><span class="sxs-lookup"><span data-stu-id="9a31e-1560">Type</span></span>|<span data-ttu-id="9a31e-1561">属性</span><span class="sxs-lookup"><span data-stu-id="9a31e-1561">Attributes</span></span>|<span data-ttu-id="9a31e-1562">说明</span><span class="sxs-lookup"><span data-stu-id="9a31e-1562">Description</span></span>|
|---|---|---|---|
|`data`|<span data-ttu-id="9a31e-1563">字符串</span><span class="sxs-lookup"><span data-stu-id="9a31e-1563">String</span></span>||<span data-ttu-id="9a31e-p192">要插入的数据。数据不得超过 1,000,000 个字符。如果传入的数据超过 1,000,000 个字符，则会引发 `ArgumentOutOfRange` 异常。</span><span class="sxs-lookup"><span data-stu-id="9a31e-p192">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`|<span data-ttu-id="9a31e-1567">Object</span><span class="sxs-lookup"><span data-stu-id="9a31e-1567">Object</span></span>|<span data-ttu-id="9a31e-1568">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="9a31e-1568">&lt;optional&gt;</span></span>|<span data-ttu-id="9a31e-1569">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="9a31e-1569">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="9a31e-1570">对象</span><span class="sxs-lookup"><span data-stu-id="9a31e-1570">Object</span></span>|<span data-ttu-id="9a31e-1571">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="9a31e-1571">&lt;optional&gt;</span></span>|<span data-ttu-id="9a31e-1572">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="9a31e-1572">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`|[<span data-ttu-id="9a31e-1573">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="9a31e-1573">Office.CoercionType</span></span>](office.md#coerciontype-string)|<span data-ttu-id="9a31e-1574">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="9a31e-1574">&lt;optional&gt;</span></span>|<span data-ttu-id="9a31e-1575">如果`text`为，则当前样式应用于 web 上的 Outlook 和桌面客户端。</span><span class="sxs-lookup"><span data-stu-id="9a31e-1575">If `text`, the current style is applied in Outlook on the web and desktop clients.</span></span> <span data-ttu-id="9a31e-1576">如果该字段是 HTML 编辑器，则仅插入文本数据，即使数据为 HTML。</span><span class="sxs-lookup"><span data-stu-id="9a31e-1576">If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="9a31e-1577">如果`html`和字段支持 HTML （主题不），则当前样式应用于 web 上的 outlook，并且在 outlook 桌面客户端中应用了默认样式。</span><span class="sxs-lookup"><span data-stu-id="9a31e-1577">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook on the web and the default style is applied in Outlook desktop clients.</span></span> <span data-ttu-id="9a31e-1578">如果该字段是文本字段，则返回 `InvalidDataFormat` 错误。</span><span class="sxs-lookup"><span data-stu-id="9a31e-1578">If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="9a31e-1579">如果未设置 `coercionType`，则结果取决于该字段：如果该字段是 HTML，则使用 HTML；如果该字段是文本，则使用纯文本。</span><span class="sxs-lookup"><span data-stu-id="9a31e-1579">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`|<span data-ttu-id="9a31e-1580">function</span><span class="sxs-lookup"><span data-stu-id="9a31e-1580">function</span></span>||<span data-ttu-id="9a31e-1581">方法完成后，使用单个参数 `callback`（一个 [`asyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `AsyncResult` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="9a31e-1581">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="9a31e-1582">Requirements</span><span class="sxs-lookup"><span data-stu-id="9a31e-1582">Requirements</span></span>

|<span data-ttu-id="9a31e-1583">要求</span><span class="sxs-lookup"><span data-stu-id="9a31e-1583">Requirement</span></span>|<span data-ttu-id="9a31e-1584">值</span><span class="sxs-lookup"><span data-stu-id="9a31e-1584">Value</span></span>|
|---|---|
|[<span data-ttu-id="9a31e-1585">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="9a31e-1585">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="9a31e-1586">1.2</span><span class="sxs-lookup"><span data-stu-id="9a31e-1586">1.2</span></span>|
|[<span data-ttu-id="9a31e-1587">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="9a31e-1587">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="9a31e-1588">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="9a31e-1588">ReadWriteItem</span></span>|
|[<span data-ttu-id="9a31e-1589">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="9a31e-1589">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="9a31e-1590">撰写</span><span class="sxs-lookup"><span data-stu-id="9a31e-1590">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="9a31e-1591">示例</span><span class="sxs-lookup"><span data-stu-id="9a31e-1591">Example</span></span>

```js
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```
