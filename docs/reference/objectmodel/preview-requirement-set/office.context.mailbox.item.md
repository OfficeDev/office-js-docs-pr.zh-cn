---
title: "\"context.subname\"-\"邮箱\"-预览要求集"
description: ''
ms.date: 04/17/2019
localization_priority: Normal
ms.openlocfilehash: cb9c298302bf0df9d7842fde4706d9d0c9710ae4
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 04/24/2019
ms.locfileid: "32450392"
---
# <a name="item"></a><span data-ttu-id="cc78d-102">item</span><span class="sxs-lookup"><span data-stu-id="cc78d-102">item</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmditem"></a><span data-ttu-id="cc78d-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span><span class="sxs-lookup"><span data-stu-id="cc78d-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span></span>

<span data-ttu-id="cc78d-p101">`item` 命名空间用于访问当前选定的邮件、会议请求或约会。可以通过使用 [itemType](#itemtype-officemailboxenumsitemtype) 属性确定 `item` 的类型。</span><span class="sxs-lookup"><span data-stu-id="cc78d-p101">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="cc78d-106">Requirements</span><span class="sxs-lookup"><span data-stu-id="cc78d-106">Requirements</span></span>

|<span data-ttu-id="cc78d-107">要求</span><span class="sxs-lookup"><span data-stu-id="cc78d-107">Requirement</span></span>|<span data-ttu-id="cc78d-108">值</span><span class="sxs-lookup"><span data-stu-id="cc78d-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="cc78d-109">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="cc78d-109">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="cc78d-110">1.0</span><span class="sxs-lookup"><span data-stu-id="cc78d-110">1.0</span></span>|
|[<span data-ttu-id="cc78d-111">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="cc78d-111">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="cc78d-112">受限</span><span class="sxs-lookup"><span data-stu-id="cc78d-112">Restricted</span></span>|
|[<span data-ttu-id="cc78d-113">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="cc78d-113">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="cc78d-114">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="cc78d-114">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="cc78d-115">成员和方法</span><span class="sxs-lookup"><span data-stu-id="cc78d-115">Members and methods</span></span>

| <span data-ttu-id="cc78d-116">成员</span><span class="sxs-lookup"><span data-stu-id="cc78d-116">Member</span></span> | <span data-ttu-id="cc78d-117">类型</span><span class="sxs-lookup"><span data-stu-id="cc78d-117">Type</span></span> |
|--------|------|
| [<span data-ttu-id="cc78d-118">attachments</span><span class="sxs-lookup"><span data-stu-id="cc78d-118">attachments</span></span>](#attachments-arrayattachmentdetails) | <span data-ttu-id="cc78d-119">成员</span><span class="sxs-lookup"><span data-stu-id="cc78d-119">Member</span></span> |
| [<span data-ttu-id="cc78d-120">bcc</span><span class="sxs-lookup"><span data-stu-id="cc78d-120">bcc</span></span>](#bcc-recipients) | <span data-ttu-id="cc78d-121">成员</span><span class="sxs-lookup"><span data-stu-id="cc78d-121">Member</span></span> |
| [<span data-ttu-id="cc78d-122">body</span><span class="sxs-lookup"><span data-stu-id="cc78d-122">body</span></span>](#body-body) | <span data-ttu-id="cc78d-123">成员</span><span class="sxs-lookup"><span data-stu-id="cc78d-123">Member</span></span> |
| [<span data-ttu-id="cc78d-124">种类</span><span class="sxs-lookup"><span data-stu-id="cc78d-124">categories</span></span>](#categories-categories) | <span data-ttu-id="cc78d-125">成员</span><span class="sxs-lookup"><span data-stu-id="cc78d-125">Member</span></span> |
| [<span data-ttu-id="cc78d-126">cc</span><span class="sxs-lookup"><span data-stu-id="cc78d-126">cc</span></span>](#cc-arrayemailaddressdetailsrecipients) | <span data-ttu-id="cc78d-127">成员</span><span class="sxs-lookup"><span data-stu-id="cc78d-127">Member</span></span> |
| [<span data-ttu-id="cc78d-128">conversationId</span><span class="sxs-lookup"><span data-stu-id="cc78d-128">conversationId</span></span>](#nullable-conversationid-string) | <span data-ttu-id="cc78d-129">成员</span><span class="sxs-lookup"><span data-stu-id="cc78d-129">Member</span></span> |
| [<span data-ttu-id="cc78d-130">dateTimeCreated</span><span class="sxs-lookup"><span data-stu-id="cc78d-130">dateTimeCreated</span></span>](#datetimecreated-date) | <span data-ttu-id="cc78d-131">成员</span><span class="sxs-lookup"><span data-stu-id="cc78d-131">Member</span></span> |
| [<span data-ttu-id="cc78d-132">dateTimeModified</span><span class="sxs-lookup"><span data-stu-id="cc78d-132">dateTimeModified</span></span>](#datetimemodified-date) | <span data-ttu-id="cc78d-133">成员</span><span class="sxs-lookup"><span data-stu-id="cc78d-133">Member</span></span> |
| [<span data-ttu-id="cc78d-134">end</span><span class="sxs-lookup"><span data-stu-id="cc78d-134">end</span></span>](#end-datetime) | <span data-ttu-id="cc78d-135">成员</span><span class="sxs-lookup"><span data-stu-id="cc78d-135">Member</span></span> |
| [<span data-ttu-id="cc78d-136">enhancedLocation</span><span class="sxs-lookup"><span data-stu-id="cc78d-136">enhancedLocation</span></span>](#enhancedlocation-enhancedlocation) | <span data-ttu-id="cc78d-137">成员</span><span class="sxs-lookup"><span data-stu-id="cc78d-137">Member</span></span> |
| [<span data-ttu-id="cc78d-138">from</span><span class="sxs-lookup"><span data-stu-id="cc78d-138">from</span></span>](#from-emailaddressdetailsfrom) | <span data-ttu-id="cc78d-139">成员</span><span class="sxs-lookup"><span data-stu-id="cc78d-139">Member</span></span> |
| [<span data-ttu-id="cc78d-140">internetHeaders</span><span class="sxs-lookup"><span data-stu-id="cc78d-140">internetHeaders</span></span>](#internetheaders-internetheaders) | <span data-ttu-id="cc78d-141">成员</span><span class="sxs-lookup"><span data-stu-id="cc78d-141">Member</span></span> |
| [<span data-ttu-id="cc78d-142">internetMessageId</span><span class="sxs-lookup"><span data-stu-id="cc78d-142">internetMessageId</span></span>](#internetmessageid-string) | <span data-ttu-id="cc78d-143">成员</span><span class="sxs-lookup"><span data-stu-id="cc78d-143">Member</span></span> |
| [<span data-ttu-id="cc78d-144">itemClass</span><span class="sxs-lookup"><span data-stu-id="cc78d-144">itemClass</span></span>](#itemclass-string) | <span data-ttu-id="cc78d-145">成员</span><span class="sxs-lookup"><span data-stu-id="cc78d-145">Member</span></span> |
| [<span data-ttu-id="cc78d-146">itemId</span><span class="sxs-lookup"><span data-stu-id="cc78d-146">itemId</span></span>](#nullable-itemid-string) | <span data-ttu-id="cc78d-147">成员</span><span class="sxs-lookup"><span data-stu-id="cc78d-147">Member</span></span> |
| [<span data-ttu-id="cc78d-148">itemType</span><span class="sxs-lookup"><span data-stu-id="cc78d-148">itemType</span></span>](#itemtype-officemailboxenumsitemtype) | <span data-ttu-id="cc78d-149">成员</span><span class="sxs-lookup"><span data-stu-id="cc78d-149">Member</span></span> |
| [<span data-ttu-id="cc78d-150">location</span><span class="sxs-lookup"><span data-stu-id="cc78d-150">location</span></span>](#location-stringlocation) | <span data-ttu-id="cc78d-151">成员</span><span class="sxs-lookup"><span data-stu-id="cc78d-151">Member</span></span> |
| [<span data-ttu-id="cc78d-152">normalizedSubject</span><span class="sxs-lookup"><span data-stu-id="cc78d-152">normalizedSubject</span></span>](#normalizedsubject-string) | <span data-ttu-id="cc78d-153">成员</span><span class="sxs-lookup"><span data-stu-id="cc78d-153">Member</span></span> |
| [<span data-ttu-id="cc78d-154">notificationMessages</span><span class="sxs-lookup"><span data-stu-id="cc78d-154">notificationMessages</span></span>](#notificationmessages-notificationmessages) | <span data-ttu-id="cc78d-155">Member</span><span class="sxs-lookup"><span data-stu-id="cc78d-155">Member</span></span> |
| [<span data-ttu-id="cc78d-156">optionalAttendees</span><span class="sxs-lookup"><span data-stu-id="cc78d-156">optionalAttendees</span></span>](#optionalattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="cc78d-157">成员</span><span class="sxs-lookup"><span data-stu-id="cc78d-157">Member</span></span> |
| [<span data-ttu-id="cc78d-158">organizer</span><span class="sxs-lookup"><span data-stu-id="cc78d-158">organizer</span></span>](#organizer-emailaddressdetailsorganizer) | <span data-ttu-id="cc78d-159">成员</span><span class="sxs-lookup"><span data-stu-id="cc78d-159">Member</span></span> |
| [<span data-ttu-id="cc78d-160">定期</span><span class="sxs-lookup"><span data-stu-id="cc78d-160">recurrence</span></span>](#nullable-recurrence-recurrence) | <span data-ttu-id="cc78d-161">成员</span><span class="sxs-lookup"><span data-stu-id="cc78d-161">Member</span></span> |
| [<span data-ttu-id="cc78d-162">requiredAttendees</span><span class="sxs-lookup"><span data-stu-id="cc78d-162">requiredAttendees</span></span>](#requiredattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="cc78d-163">成员</span><span class="sxs-lookup"><span data-stu-id="cc78d-163">Member</span></span> |
| [<span data-ttu-id="cc78d-164">sender</span><span class="sxs-lookup"><span data-stu-id="cc78d-164">sender</span></span>](#sender-emailaddressdetails) | <span data-ttu-id="cc78d-165">Member</span><span class="sxs-lookup"><span data-stu-id="cc78d-165">Member</span></span> |
| [<span data-ttu-id="cc78d-166">webcasts&seriesid</span><span class="sxs-lookup"><span data-stu-id="cc78d-166">seriesId</span></span>](#nullable-seriesid-string) | <span data-ttu-id="cc78d-167">Member</span><span class="sxs-lookup"><span data-stu-id="cc78d-167">Member</span></span> |
| [<span data-ttu-id="cc78d-168">start</span><span class="sxs-lookup"><span data-stu-id="cc78d-168">start</span></span>](#start-datetime) | <span data-ttu-id="cc78d-169">Member</span><span class="sxs-lookup"><span data-stu-id="cc78d-169">Member</span></span> |
| [<span data-ttu-id="cc78d-170">subject</span><span class="sxs-lookup"><span data-stu-id="cc78d-170">subject</span></span>](#subject-stringsubject) | <span data-ttu-id="cc78d-171">成员</span><span class="sxs-lookup"><span data-stu-id="cc78d-171">Member</span></span> |
| [<span data-ttu-id="cc78d-172">to</span><span class="sxs-lookup"><span data-stu-id="cc78d-172">to</span></span>](#to-arrayemailaddressdetailsrecipients) | <span data-ttu-id="cc78d-173">成员</span><span class="sxs-lookup"><span data-stu-id="cc78d-173">Member</span></span> |
| [<span data-ttu-id="cc78d-174">addFileAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="cc78d-174">addFileAttachmentAsync</span></span>](#addfileattachmentasyncuri-attachmentname-options-callback) | <span data-ttu-id="cc78d-175">方法</span><span class="sxs-lookup"><span data-stu-id="cc78d-175">Method</span></span> |
| [<span data-ttu-id="cc78d-176">addFileAttachmentFromBase64Async</span><span class="sxs-lookup"><span data-stu-id="cc78d-176">addFileAttachmentFromBase64Async</span></span>](#addfileattachmentfrombase64asyncbase64file-attachmentname-options-callback) | <span data-ttu-id="cc78d-177">方法</span><span class="sxs-lookup"><span data-stu-id="cc78d-177">Method</span></span> |
| [<span data-ttu-id="cc78d-178">addHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="cc78d-178">addHandlerAsync</span></span>](#addhandlerasynceventtype-handler-options-callback) | <span data-ttu-id="cc78d-179">方法</span><span class="sxs-lookup"><span data-stu-id="cc78d-179">Method</span></span> |
| [<span data-ttu-id="cc78d-180">addItemAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="cc78d-180">addItemAttachmentAsync</span></span>](#additemattachmentasyncitemid-attachmentname-options-callback) | <span data-ttu-id="cc78d-181">方法</span><span class="sxs-lookup"><span data-stu-id="cc78d-181">Method</span></span> |
| [<span data-ttu-id="cc78d-182">close</span><span class="sxs-lookup"><span data-stu-id="cc78d-182">close</span></span>](#close) | <span data-ttu-id="cc78d-183">方法</span><span class="sxs-lookup"><span data-stu-id="cc78d-183">Method</span></span> |
| [<span data-ttu-id="cc78d-184">displayReplyAllForm</span><span class="sxs-lookup"><span data-stu-id="cc78d-184">displayReplyAllForm</span></span>](#displayreplyallformformdata-callback) | <span data-ttu-id="cc78d-185">方法</span><span class="sxs-lookup"><span data-stu-id="cc78d-185">Method</span></span> |
| [<span data-ttu-id="cc78d-186">displayReplyForm</span><span class="sxs-lookup"><span data-stu-id="cc78d-186">displayReplyForm</span></span>](#displayreplyformformdata-callback) | <span data-ttu-id="cc78d-187">方法</span><span class="sxs-lookup"><span data-stu-id="cc78d-187">Method</span></span> |
| [<span data-ttu-id="cc78d-188">getAttachmentContentAsync</span><span class="sxs-lookup"><span data-stu-id="cc78d-188">getAttachmentContentAsync</span></span>](#getattachmentcontentasyncattachmentid-options-callback--attachmentcontent) | <span data-ttu-id="cc78d-189">方法</span><span class="sxs-lookup"><span data-stu-id="cc78d-189">Method</span></span> |
| [<span data-ttu-id="cc78d-190">getAttachmentsAsync</span><span class="sxs-lookup"><span data-stu-id="cc78d-190">getAttachmentsAsync</span></span>](#getattachmentsasyncoptions-callback--arrayattachmentdetails) | <span data-ttu-id="cc78d-191">方法</span><span class="sxs-lookup"><span data-stu-id="cc78d-191">Method</span></span> |
| [<span data-ttu-id="cc78d-192">getEntities</span><span class="sxs-lookup"><span data-stu-id="cc78d-192">getEntities</span></span>](#getentities--entities) | <span data-ttu-id="cc78d-193">方法</span><span class="sxs-lookup"><span data-stu-id="cc78d-193">Method</span></span> |
| [<span data-ttu-id="cc78d-194">getEntitiesByType</span><span class="sxs-lookup"><span data-stu-id="cc78d-194">getEntitiesByType</span></span>](#getentitiesbytypeentitytype--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="cc78d-195">方法</span><span class="sxs-lookup"><span data-stu-id="cc78d-195">Method</span></span> |
| [<span data-ttu-id="cc78d-196">getFilteredEntitiesByName</span><span class="sxs-lookup"><span data-stu-id="cc78d-196">getFilteredEntitiesByName</span></span>](#getfilteredentitiesbynamename--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="cc78d-197">方法</span><span class="sxs-lookup"><span data-stu-id="cc78d-197">Method</span></span> |
| [<span data-ttu-id="cc78d-198">office.context.mailbox.item.getinitializationcontextasync</span><span class="sxs-lookup"><span data-stu-id="cc78d-198">getInitializationContextAsync</span></span>](#getinitializationcontextasyncoptions-callback) | <span data-ttu-id="cc78d-199">方法</span><span class="sxs-lookup"><span data-stu-id="cc78d-199">Method</span></span> |
| [<span data-ttu-id="cc78d-200">getRegExMatches</span><span class="sxs-lookup"><span data-stu-id="cc78d-200">getRegExMatches</span></span>](#getregexmatches--object) | <span data-ttu-id="cc78d-201">方法</span><span class="sxs-lookup"><span data-stu-id="cc78d-201">Method</span></span> |
| [<span data-ttu-id="cc78d-202">getRegExMatchesByName</span><span class="sxs-lookup"><span data-stu-id="cc78d-202">getRegExMatchesByName</span></span>](#getregexmatchesbynamename--nullable-array-string-) | <span data-ttu-id="cc78d-203">方法</span><span class="sxs-lookup"><span data-stu-id="cc78d-203">Method</span></span> |
| [<span data-ttu-id="cc78d-204">getSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="cc78d-204">getSelectedDataAsync</span></span>](#getselecteddataasynccoerciontype-options-callback--string) | <span data-ttu-id="cc78d-205">方法</span><span class="sxs-lookup"><span data-stu-id="cc78d-205">Method</span></span> |
| [<span data-ttu-id="cc78d-206">office.context.mailbox.item.getselectedentities</span><span class="sxs-lookup"><span data-stu-id="cc78d-206">getSelectedEntities</span></span>](#getselectedentities--entities) | <span data-ttu-id="cc78d-207">方法</span><span class="sxs-lookup"><span data-stu-id="cc78d-207">Method</span></span> |
| [<span data-ttu-id="cc78d-208">office.context.mailbox.item.getselectedregexmatches</span><span class="sxs-lookup"><span data-stu-id="cc78d-208">getSelectedRegExMatches</span></span>](#getselectedregexmatches--object) | <span data-ttu-id="cc78d-209">方法</span><span class="sxs-lookup"><span data-stu-id="cc78d-209">Method</span></span> |
| [<span data-ttu-id="cc78d-210">getSharedPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="cc78d-210">getSharedPropertiesAsync</span></span>](#getsharedpropertiesasyncoptions-callback) | <span data-ttu-id="cc78d-211">方法</span><span class="sxs-lookup"><span data-stu-id="cc78d-211">Method</span></span> |
| [<span data-ttu-id="cc78d-212">loadCustomPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="cc78d-212">loadCustomPropertiesAsync</span></span>](#loadcustompropertiesasynccallback-usercontext) | <span data-ttu-id="cc78d-213">方法</span><span class="sxs-lookup"><span data-stu-id="cc78d-213">Method</span></span> |
| [<span data-ttu-id="cc78d-214">removeAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="cc78d-214">removeAttachmentAsync</span></span>](#removeattachmentasyncattachmentid-options-callback) | <span data-ttu-id="cc78d-215">方法</span><span class="sxs-lookup"><span data-stu-id="cc78d-215">Method</span></span> |
| [<span data-ttu-id="cc78d-216">removeHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="cc78d-216">removeHandlerAsync</span></span>](#removehandlerasynceventtype-options-callback) | <span data-ttu-id="cc78d-217">方法</span><span class="sxs-lookup"><span data-stu-id="cc78d-217">Method</span></span> |
| [<span data-ttu-id="cc78d-218">saveAsync</span><span class="sxs-lookup"><span data-stu-id="cc78d-218">saveAsync</span></span>](#saveasyncoptions-callback) | <span data-ttu-id="cc78d-219">方法</span><span class="sxs-lookup"><span data-stu-id="cc78d-219">Method</span></span> |
| [<span data-ttu-id="cc78d-220">setSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="cc78d-220">setSelectedDataAsync</span></span>](#setselecteddataasyncdata-options-callback) | <span data-ttu-id="cc78d-221">方法</span><span class="sxs-lookup"><span data-stu-id="cc78d-221">Method</span></span> |

### <a name="example"></a><span data-ttu-id="cc78d-222">示例</span><span class="sxs-lookup"><span data-stu-id="cc78d-222">Example</span></span>

<span data-ttu-id="cc78d-223">以下 JavaScript 代码示例显示了如何访问 Outlook 中当前项目的 `subject` 属性。</span><span class="sxs-lookup"><span data-stu-id="cc78d-223">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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

### <a name="members"></a><span data-ttu-id="cc78d-224">成员</span><span class="sxs-lookup"><span data-stu-id="cc78d-224">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetails"></a><span data-ttu-id="cc78d-225">attachments :Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="cc78d-225">attachments :Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span></span>

<span data-ttu-id="cc78d-226">以数组的形式获取项目的附件。</span><span class="sxs-lookup"><span data-stu-id="cc78d-226">Gets the item's attachments as an array.</span></span> <span data-ttu-id="cc78d-227">仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="cc78d-227">Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="cc78d-228">由于潜在的安全问题，某些类型的文件会受到 Outlook 阻止，并且不会返回。</span><span class="sxs-lookup"><span data-stu-id="cc78d-228">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="cc78d-229">如需了解更多信息，请参阅 [Outlook 中阻止的附件](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519)。</span><span class="sxs-lookup"><span data-stu-id="cc78d-229">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="cc78d-230">Type</span><span class="sxs-lookup"><span data-stu-id="cc78d-230">Type</span></span>

*   <span data-ttu-id="cc78d-231">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="cc78d-231">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span></span>

##### <a name="requirements"></a><span data-ttu-id="cc78d-232">Requirements</span><span class="sxs-lookup"><span data-stu-id="cc78d-232">Requirements</span></span>

|<span data-ttu-id="cc78d-233">要求</span><span class="sxs-lookup"><span data-stu-id="cc78d-233">Requirement</span></span>|<span data-ttu-id="cc78d-234">值</span><span class="sxs-lookup"><span data-stu-id="cc78d-234">Value</span></span>|
|---|---|
|[<span data-ttu-id="cc78d-235">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="cc78d-235">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="cc78d-236">1.0</span><span class="sxs-lookup"><span data-stu-id="cc78d-236">1.0</span></span>|
|[<span data-ttu-id="cc78d-237">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="cc78d-237">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="cc78d-238">ReadItem</span><span class="sxs-lookup"><span data-stu-id="cc78d-238">ReadItem</span></span>|
|[<span data-ttu-id="cc78d-239">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="cc78d-239">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="cc78d-240">阅读</span><span class="sxs-lookup"><span data-stu-id="cc78d-240">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="cc78d-241">示例</span><span class="sxs-lookup"><span data-stu-id="cc78d-241">Example</span></span>

<span data-ttu-id="cc78d-242">以下代码使用当前项目上所有附件的详细信息构成 HTML 字符串。</span><span class="sxs-lookup"><span data-stu-id="cc78d-242">The following code builds an HTML string with details of all attachments on the current item.</span></span>

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

---
---

####  <a name="bcc-recipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="cc78d-243">bcc :[Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="cc78d-243">bcc :[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="cc78d-244">获取对象，该对象提供用于获取或更新邮件的密件抄送 (Bcc) 行上的收件人的方法。</span><span class="sxs-lookup"><span data-stu-id="cc78d-244">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="cc78d-245">仅限撰写模式。</span><span class="sxs-lookup"><span data-stu-id="cc78d-245">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="cc78d-246">类型</span><span class="sxs-lookup"><span data-stu-id="cc78d-246">Type</span></span>

*   [<span data-ttu-id="cc78d-247">收件人</span><span class="sxs-lookup"><span data-stu-id="cc78d-247">Recipients</span></span>](/javascript/api/outlook/office.recipients)

##### <a name="requirements"></a><span data-ttu-id="cc78d-248">要求</span><span class="sxs-lookup"><span data-stu-id="cc78d-248">Requirements</span></span>

|<span data-ttu-id="cc78d-249">要求</span><span class="sxs-lookup"><span data-stu-id="cc78d-249">Requirement</span></span>|<span data-ttu-id="cc78d-250">值</span><span class="sxs-lookup"><span data-stu-id="cc78d-250">Value</span></span>|
|---|---|
|[<span data-ttu-id="cc78d-251">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="cc78d-251">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="cc78d-252">1.1</span><span class="sxs-lookup"><span data-stu-id="cc78d-252">1.1</span></span>|
|[<span data-ttu-id="cc78d-253">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="cc78d-253">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="cc78d-254">ReadItem</span><span class="sxs-lookup"><span data-stu-id="cc78d-254">ReadItem</span></span>|
|[<span data-ttu-id="cc78d-255">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="cc78d-255">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="cc78d-256">撰写</span><span class="sxs-lookup"><span data-stu-id="cc78d-256">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="cc78d-257">示例</span><span class="sxs-lookup"><span data-stu-id="cc78d-257">Example</span></span>

```javascript
Office.context.mailbox.item.bcc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.bcc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.bcc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfBccRecipients = asyncResult.value;
}
```

---
---

####  <a name="body-bodyjavascriptapioutlookofficebody"></a><span data-ttu-id="cc78d-258">body :[Body](/javascript/api/outlook/office.body)</span><span class="sxs-lookup"><span data-stu-id="cc78d-258">body :[Body](/javascript/api/outlook/office.body)</span></span>

<span data-ttu-id="cc78d-259">获取一个提供用于处理项目正文的方法的对象。</span><span class="sxs-lookup"><span data-stu-id="cc78d-259">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="cc78d-260">Type</span><span class="sxs-lookup"><span data-stu-id="cc78d-260">Type</span></span>

*   [<span data-ttu-id="cc78d-261">Body</span><span class="sxs-lookup"><span data-stu-id="cc78d-261">Body</span></span>](/javascript/api/outlook/office.body)

##### <a name="requirements"></a><span data-ttu-id="cc78d-262">Requirements</span><span class="sxs-lookup"><span data-stu-id="cc78d-262">Requirements</span></span>

|<span data-ttu-id="cc78d-263">要求</span><span class="sxs-lookup"><span data-stu-id="cc78d-263">Requirement</span></span>|<span data-ttu-id="cc78d-264">值</span><span class="sxs-lookup"><span data-stu-id="cc78d-264">Value</span></span>|
|---|---|
|[<span data-ttu-id="cc78d-265">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="cc78d-265">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="cc78d-266">1.1</span><span class="sxs-lookup"><span data-stu-id="cc78d-266">1.1</span></span>|
|[<span data-ttu-id="cc78d-267">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="cc78d-267">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="cc78d-268">ReadItem</span><span class="sxs-lookup"><span data-stu-id="cc78d-268">ReadItem</span></span>|
|[<span data-ttu-id="cc78d-269">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="cc78d-269">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="cc78d-270">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="cc78d-270">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="cc78d-271">示例</span><span class="sxs-lookup"><span data-stu-id="cc78d-271">Example</span></span>

<span data-ttu-id="cc78d-272">本示例获取纯文本格式的邮件正文。</span><span class="sxs-lookup"><span data-stu-id="cc78d-272">This example gets the body of the message in plain text.</span></span>

```javascript
Office.context.mailbox.item.body.getAsync(
  "text",
  { asyncContext: "This is passed to the callback" },
  function callback(result) {
    // Do something with the result.
  });

```

<span data-ttu-id="cc78d-273">以下是传递到回调函数的结果参数的示例。</span><span class="sxs-lookup"><span data-stu-id="cc78d-273">The following is an example of the result parameter passed to the callback function.</span></span>

```json
{
  "value": "TEXT of whole body (including threads below)",
  "status": "succeeded",
  "asyncContext": "This is passed to the callback"
}
```

---
---

####  <a name="categories-categoriesjavascriptapioutlookofficecategories"></a><span data-ttu-id="cc78d-274">类别:[类别](/javascript/api/outlook/office.categories)</span><span class="sxs-lookup"><span data-stu-id="cc78d-274">categories :[Categories](/javascript/api/outlook/office.categories)</span></span>

<span data-ttu-id="cc78d-275">获取一个对象, 该对象提供用于管理项的类别的方法。</span><span class="sxs-lookup"><span data-stu-id="cc78d-275">Gets an object that provides methods for managing the item's categories.</span></span>

> [!NOTE]
> <span data-ttu-id="cc78d-276">在 Outlook for iOS 或 Outlook for Android 中不支持此成员。</span><span class="sxs-lookup"><span data-stu-id="cc78d-276">This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="type"></a><span data-ttu-id="cc78d-277">类型</span><span class="sxs-lookup"><span data-stu-id="cc78d-277">Type</span></span>

*   [<span data-ttu-id="cc78d-278">Categories</span><span class="sxs-lookup"><span data-stu-id="cc78d-278">Categories</span></span>](/javascript/api/outlook/office.categories)

##### <a name="requirements"></a><span data-ttu-id="cc78d-279">Requirements</span><span class="sxs-lookup"><span data-stu-id="cc78d-279">Requirements</span></span>

|<span data-ttu-id="cc78d-280">要求</span><span class="sxs-lookup"><span data-stu-id="cc78d-280">Requirement</span></span>|<span data-ttu-id="cc78d-281">值</span><span class="sxs-lookup"><span data-stu-id="cc78d-281">Value</span></span>|
|---|---|
|[<span data-ttu-id="cc78d-282">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="cc78d-282">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="cc78d-283">预览</span><span class="sxs-lookup"><span data-stu-id="cc78d-283">Preview</span></span>|
|[<span data-ttu-id="cc78d-284">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="cc78d-284">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="cc78d-285">ReadItem</span><span class="sxs-lookup"><span data-stu-id="cc78d-285">ReadItem</span></span>|
|[<span data-ttu-id="cc78d-286">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="cc78d-286">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="cc78d-287">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="cc78d-287">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="cc78d-288">示例</span><span class="sxs-lookup"><span data-stu-id="cc78d-288">Example</span></span>

<span data-ttu-id="cc78d-289">此示例获取项的类别。</span><span class="sxs-lookup"><span data-stu-id="cc78d-289">This example gets the item's categories.</span></span>

```javascript
Office.context.mailbox.item.categories.getAsync(function (asyncResult) {
  if (asyncResult.status === Office.AsyncResultStatus.Failed) {
    console.log("Action failed with error: " + asyncResult.error.message);
  } else {
    console.log("Categories: " + JSON.stringify(asyncResult.value));
  }
});
```

---
---

####  <a name="cc-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="cc78d-290">cc :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="cc78d-290">cc :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="cc78d-291">提供对邮件的抄送 (Cc) 收件人的访问权限。</span><span class="sxs-lookup"><span data-stu-id="cc78d-291">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="cc78d-292">对象的类型和访问级别取决于当前项目的模式。</span><span class="sxs-lookup"><span data-stu-id="cc78d-292">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="cc78d-293">阅读模式</span><span class="sxs-lookup"><span data-stu-id="cc78d-293">Read mode</span></span>

<span data-ttu-id="cc78d-p106">`cc` 属性返回包含邮件的**抄送**行上所列的每个收件人的 `EmailAddressDetails` 对象的数组。集合上限为 100 个成员。</span><span class="sxs-lookup"><span data-stu-id="cc78d-p106">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.</span></span>

```javascript
console.log(JSON.stringify(Office.context.mailbox.item.cc));
```

##### <a name="compose-mode"></a><span data-ttu-id="cc78d-296">撰写模式</span><span class="sxs-lookup"><span data-stu-id="cc78d-296">Compose mode</span></span>

<span data-ttu-id="cc78d-297">`cc` 属性返回一个 `Recipients` 对象，该对象提供用于获取或更新邮件的“**抄送**”行上收件人的方法。</span><span class="sxs-lookup"><span data-stu-id="cc78d-297">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span>

```javascript
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="cc78d-298">Type</span><span class="sxs-lookup"><span data-stu-id="cc78d-298">Type</span></span>

*   <span data-ttu-id="cc78d-299">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="cc78d-299">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="cc78d-300">Requirements</span><span class="sxs-lookup"><span data-stu-id="cc78d-300">Requirements</span></span>

|<span data-ttu-id="cc78d-301">要求</span><span class="sxs-lookup"><span data-stu-id="cc78d-301">Requirement</span></span>|<span data-ttu-id="cc78d-302">值</span><span class="sxs-lookup"><span data-stu-id="cc78d-302">Value</span></span>|
|---|---|
|[<span data-ttu-id="cc78d-303">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="cc78d-303">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="cc78d-304">1.0</span><span class="sxs-lookup"><span data-stu-id="cc78d-304">1.0</span></span>|
|[<span data-ttu-id="cc78d-305">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="cc78d-305">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="cc78d-306">ReadItem</span><span class="sxs-lookup"><span data-stu-id="cc78d-306">ReadItem</span></span>|
|[<span data-ttu-id="cc78d-307">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="cc78d-307">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="cc78d-308">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="cc78d-308">Compose or Read</span></span>|

---
---

####  <a name="nullable-conversationid-string"></a><span data-ttu-id="cc78d-309">(nullable) conversationId :String</span><span class="sxs-lookup"><span data-stu-id="cc78d-309">(nullable) conversationId :String</span></span>

<span data-ttu-id="cc78d-310">获取包含特定消息的电子邮件会话的标识符。</span><span class="sxs-lookup"><span data-stu-id="cc78d-310">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="cc78d-p107">如果在阅读窗体或撰写窗体的回复中激活邮件应用程序，则此属性可以获得一个整数值。如果用户随后更改了回复邮件的主题（若发送回复），则该邮件的对话 ID 将改变且之前获取的值将不适用。</span><span class="sxs-lookup"><span data-stu-id="cc78d-p107">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="cc78d-p108">对于撰写窗体的新项目，此属性获得一个 null 值。如果用户设置一个主题并保存该项目，`conversationId` 属性将返回一个值。</span><span class="sxs-lookup"><span data-stu-id="cc78d-p108">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="cc78d-315">Type</span><span class="sxs-lookup"><span data-stu-id="cc78d-315">Type</span></span>

*   <span data-ttu-id="cc78d-316">String</span><span class="sxs-lookup"><span data-stu-id="cc78d-316">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="cc78d-317">Requirements</span><span class="sxs-lookup"><span data-stu-id="cc78d-317">Requirements</span></span>

|<span data-ttu-id="cc78d-318">要求</span><span class="sxs-lookup"><span data-stu-id="cc78d-318">Requirement</span></span>|<span data-ttu-id="cc78d-319">值</span><span class="sxs-lookup"><span data-stu-id="cc78d-319">Value</span></span>|
|---|---|
|[<span data-ttu-id="cc78d-320">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="cc78d-320">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="cc78d-321">1.0</span><span class="sxs-lookup"><span data-stu-id="cc78d-321">1.0</span></span>|
|[<span data-ttu-id="cc78d-322">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="cc78d-322">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="cc78d-323">ReadItem</span><span class="sxs-lookup"><span data-stu-id="cc78d-323">ReadItem</span></span>|
|[<span data-ttu-id="cc78d-324">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="cc78d-324">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="cc78d-325">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="cc78d-325">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="cc78d-326">示例</span><span class="sxs-lookup"><span data-stu-id="cc78d-326">Example</span></span>

```javascript
var conversationId = Office.context.mailbox.item.conversationId;
console.log("conversationId: " + conversationId);
```

---
---

#### <a name="datetimecreated-date"></a><span data-ttu-id="cc78d-327">dateTimeCreated :Date</span><span class="sxs-lookup"><span data-stu-id="cc78d-327">dateTimeCreated :Date</span></span>

<span data-ttu-id="cc78d-p109">获取项目创建的日期和时间。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="cc78d-p109">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="cc78d-330">类型</span><span class="sxs-lookup"><span data-stu-id="cc78d-330">Type</span></span>

*   <span data-ttu-id="cc78d-331">日期</span><span class="sxs-lookup"><span data-stu-id="cc78d-331">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="cc78d-332">要求</span><span class="sxs-lookup"><span data-stu-id="cc78d-332">Requirements</span></span>

|<span data-ttu-id="cc78d-333">要求</span><span class="sxs-lookup"><span data-stu-id="cc78d-333">Requirement</span></span>|<span data-ttu-id="cc78d-334">值</span><span class="sxs-lookup"><span data-stu-id="cc78d-334">Value</span></span>|
|---|---|
|[<span data-ttu-id="cc78d-335">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="cc78d-335">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="cc78d-336">1.0</span><span class="sxs-lookup"><span data-stu-id="cc78d-336">1.0</span></span>|
|[<span data-ttu-id="cc78d-337">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="cc78d-337">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="cc78d-338">ReadItem</span><span class="sxs-lookup"><span data-stu-id="cc78d-338">ReadItem</span></span>|
|[<span data-ttu-id="cc78d-339">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="cc78d-339">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="cc78d-340">阅读</span><span class="sxs-lookup"><span data-stu-id="cc78d-340">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="cc78d-341">示例</span><span class="sxs-lookup"><span data-stu-id="cc78d-341">Example</span></span>

```javascript
var dateTimeCreated = Office.context.mailbox.item.dateTimeCreated;
console.log("Date and time created: " + dateTimeCreated);
```

---
---

#### <a name="datetimemodified-date"></a><span data-ttu-id="cc78d-342">dateTimeModified :Date</span><span class="sxs-lookup"><span data-stu-id="cc78d-342">dateTimeModified :Date</span></span>

<span data-ttu-id="cc78d-p110">获取项目最近一次修改的日期和时间。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="cc78d-p110">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="cc78d-345">在 Outlook for iOS 或 Outlook for Android 中不支持此成员。</span><span class="sxs-lookup"><span data-stu-id="cc78d-345">This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="type"></a><span data-ttu-id="cc78d-346">类型</span><span class="sxs-lookup"><span data-stu-id="cc78d-346">Type</span></span>

*   <span data-ttu-id="cc78d-347">日期</span><span class="sxs-lookup"><span data-stu-id="cc78d-347">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="cc78d-348">要求</span><span class="sxs-lookup"><span data-stu-id="cc78d-348">Requirements</span></span>

|<span data-ttu-id="cc78d-349">要求</span><span class="sxs-lookup"><span data-stu-id="cc78d-349">Requirement</span></span>|<span data-ttu-id="cc78d-350">值</span><span class="sxs-lookup"><span data-stu-id="cc78d-350">Value</span></span>|
|---|---|
|[<span data-ttu-id="cc78d-351">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="cc78d-351">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="cc78d-352">1.0</span><span class="sxs-lookup"><span data-stu-id="cc78d-352">1.0</span></span>|
|[<span data-ttu-id="cc78d-353">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="cc78d-353">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="cc78d-354">ReadItem</span><span class="sxs-lookup"><span data-stu-id="cc78d-354">ReadItem</span></span>|
|[<span data-ttu-id="cc78d-355">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="cc78d-355">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="cc78d-356">阅读</span><span class="sxs-lookup"><span data-stu-id="cc78d-356">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="cc78d-357">示例</span><span class="sxs-lookup"><span data-stu-id="cc78d-357">Example</span></span>

```javascript
var dateTimeModified = Office.context.mailbox.item.dateTimeModified;
console.log("Date and time modified: " + dateTimeModified);
```

---
---

####  <a name="end-datetimejavascriptapioutlookofficetime"></a><span data-ttu-id="cc78d-358">end :Date|[Time](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="cc78d-358">end :Date|[Time](/javascript/api/outlook/office.time)</span></span>

<span data-ttu-id="cc78d-359">获取或设置约会结束的日期和时间。</span><span class="sxs-lookup"><span data-stu-id="cc78d-359">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="cc78d-p111">将 `end` 属性表示为协调世界时 (UTC) 的日期和时间值。可使用 [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) 方法将 end 属性值转换为客户端的本地日期和时间。</span><span class="sxs-lookup"><span data-stu-id="cc78d-p111">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="cc78d-362">阅读模式</span><span class="sxs-lookup"><span data-stu-id="cc78d-362">Read mode</span></span>

<span data-ttu-id="cc78d-363">`end` 属性返回 `Date` 对象。</span><span class="sxs-lookup"><span data-stu-id="cc78d-363">The `end` property returns a `Date` object.</span></span>

```javascript
var end = Office.context.mailbox.item.end;
console.log("Appointment end: " + end);
```

##### <a name="compose-mode"></a><span data-ttu-id="cc78d-364">撰写模式</span><span class="sxs-lookup"><span data-stu-id="cc78d-364">Compose mode</span></span>

<span data-ttu-id="cc78d-365">`end` 属性返回 `Time` 对象。</span><span class="sxs-lookup"><span data-stu-id="cc78d-365">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="cc78d-366">使用 [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) 方法设置结束时间时，应使用 [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) 方法将客户端的本地时间转换为服务器的 UTC。</span><span class="sxs-lookup"><span data-stu-id="cc78d-366">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="cc78d-367">以下示例使用 `Time` 对象的 [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) 方法设置约会的结束时间。</span><span class="sxs-lookup"><span data-stu-id="cc78d-367">The following example sets the end time of an appointment by using the [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="cc78d-368">类型</span><span class="sxs-lookup"><span data-stu-id="cc78d-368">Type</span></span>

*   <span data-ttu-id="cc78d-369">Date | [Time](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="cc78d-369">Date | [Time](/javascript/api/outlook/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="cc78d-370">要求</span><span class="sxs-lookup"><span data-stu-id="cc78d-370">Requirements</span></span>

|<span data-ttu-id="cc78d-371">要求</span><span class="sxs-lookup"><span data-stu-id="cc78d-371">Requirement</span></span>|<span data-ttu-id="cc78d-372">值</span><span class="sxs-lookup"><span data-stu-id="cc78d-372">Value</span></span>|
|---|---|
|[<span data-ttu-id="cc78d-373">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="cc78d-373">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="cc78d-374">1.0</span><span class="sxs-lookup"><span data-stu-id="cc78d-374">1.0</span></span>|
|[<span data-ttu-id="cc78d-375">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="cc78d-375">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="cc78d-376">ReadItem</span><span class="sxs-lookup"><span data-stu-id="cc78d-376">ReadItem</span></span>|
|[<span data-ttu-id="cc78d-377">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="cc78d-377">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="cc78d-378">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="cc78d-378">Compose or Read</span></span>|

---
---

#### <a name="enhancedlocation-enhancedlocationjavascriptapioutlookofficeenhancedlocation"></a><span data-ttu-id="cc78d-379">enhancedLocation:[enhancedLocation](/javascript/api/outlook/office.enhancedlocation)</span><span class="sxs-lookup"><span data-stu-id="cc78d-379">enhancedLocation :[EnhancedLocation](/javascript/api/outlook/office.enhancedlocation)</span></span>

<span data-ttu-id="cc78d-380">获取或设置约会的位置。</span><span class="sxs-lookup"><span data-stu-id="cc78d-380">Gets or sets the locations of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="cc78d-381">阅读模式</span><span class="sxs-lookup"><span data-stu-id="cc78d-381">Read mode</span></span>

<span data-ttu-id="cc78d-382">该`enhancedLocation`属性返回一个[EnhancedLocation](/javascript/api/outlook/office.enhancedlocation)对象, 该对象允许您获取与约会关联的一组位置 (每个由[LocationDetails](/javascript/api/outlook/office.locationdetails)对象表示)。</span><span class="sxs-lookup"><span data-stu-id="cc78d-382">The `enhancedLocation` property returns an [EnhancedLocation](/javascript/api/outlook/office.enhancedlocation) object that allows you to get the set of locations (each represented by a [LocationDetails](/javascript/api/outlook/office.locationdetails) object) associated with the appointment.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="cc78d-383">撰写模式</span><span class="sxs-lookup"><span data-stu-id="cc78d-383">Compose mode</span></span>

<span data-ttu-id="cc78d-384">该`enhancedLocation`属性返回一个[EnhancedLocation](/javascript/api/outlook/office.enhancedlocation)对象, 该对象提供用于获取、删除或添加约会位置的方法。</span><span class="sxs-lookup"><span data-stu-id="cc78d-384">The `enhancedLocation` property returns an [EnhancedLocation](/javascript/api/outlook/office.enhancedlocation) object that provides methods to get, remove, or add locations on an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="cc78d-385">Type</span><span class="sxs-lookup"><span data-stu-id="cc78d-385">Type</span></span>

*   [<span data-ttu-id="cc78d-386">EnhancedLocation</span><span class="sxs-lookup"><span data-stu-id="cc78d-386">EnhancedLocation</span></span>](/javascript/api/outlook/office.enhancedlocation)

##### <a name="requirements"></a><span data-ttu-id="cc78d-387">Requirements</span><span class="sxs-lookup"><span data-stu-id="cc78d-387">Requirements</span></span>

|<span data-ttu-id="cc78d-388">要求</span><span class="sxs-lookup"><span data-stu-id="cc78d-388">Requirement</span></span>|<span data-ttu-id="cc78d-389">值</span><span class="sxs-lookup"><span data-stu-id="cc78d-389">Value</span></span>|
|---|---|
|[<span data-ttu-id="cc78d-390">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="cc78d-390">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="cc78d-391">预览</span><span class="sxs-lookup"><span data-stu-id="cc78d-391">Preview</span></span>|
|[<span data-ttu-id="cc78d-392">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="cc78d-392">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="cc78d-393">ReadItem</span><span class="sxs-lookup"><span data-stu-id="cc78d-393">ReadItem</span></span>|
|[<span data-ttu-id="cc78d-394">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="cc78d-394">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="cc78d-395">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="cc78d-395">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="cc78d-396">示例</span><span class="sxs-lookup"><span data-stu-id="cc78d-396">Example</span></span>

<span data-ttu-id="cc78d-397">下面的示例将获取与约会相关联的当前位置。</span><span class="sxs-lookup"><span data-stu-id="cc78d-397">The following example gets the current locations associated with the appointment.</span></span>

```javascript
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

---
---

#### <a name="from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsfromjavascriptapioutlookofficefrom"></a><span data-ttu-id="cc78d-398">发件人:[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[From](/javascript/api/outlook/office.from)</span><span class="sxs-lookup"><span data-stu-id="cc78d-398">from :[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[From](/javascript/api/outlook/office.from)</span></span>

<span data-ttu-id="cc78d-399">获取邮件发件人的电子邮件地址。</span><span class="sxs-lookup"><span data-stu-id="cc78d-399">Gets the email address of the sender of a message.</span></span>

<span data-ttu-id="cc78d-p112">`from` 和 [`sender`](#sender-emailaddressdetails) 属性表示同一个人，邮件由代理人发送的除外。在此情况下，`from` 属性表示代理程序，而 sender 属性表示代理人。</span><span class="sxs-lookup"><span data-stu-id="cc78d-p112">The `from` and [`sender`](#sender-emailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="cc78d-402">`from` 属性中 `EmailAddressDetails` 对象的 `recipientType` 属性为 `undefined`。</span><span class="sxs-lookup"><span data-stu-id="cc78d-402">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="cc78d-403">阅读模式</span><span class="sxs-lookup"><span data-stu-id="cc78d-403">Read mode</span></span>

<span data-ttu-id="cc78d-404">`from`属性返回一个`EmailAddressDetails`对象。</span><span class="sxs-lookup"><span data-stu-id="cc78d-404">The `from` property returns an `EmailAddressDetails` object.</span></span>

```javascript
var from = Office.context.mailbox.item.from;
console.log("From " + from);
```

##### <a name="compose-mode"></a><span data-ttu-id="cc78d-405">撰写模式</span><span class="sxs-lookup"><span data-stu-id="cc78d-405">Compose mode</span></span>

<span data-ttu-id="cc78d-406">`from`属性返回一个`From`对象, 该对象提供用于获取 "起始" 值的方法。</span><span class="sxs-lookup"><span data-stu-id="cc78d-406">The `from` property returns a `From` object that provides a method to get the from value.</span></span>

```javascript
Office.context.mailbox.item.from.getAsync(callback);

function callback(asyncResult) {
  var from = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="cc78d-407">Type</span><span class="sxs-lookup"><span data-stu-id="cc78d-407">Type</span></span>

*   <span data-ttu-id="cc78d-408">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [](/javascript/api/outlook/office.from)</span><span class="sxs-lookup"><span data-stu-id="cc78d-408">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [From](/javascript/api/outlook/office.from)</span></span>

##### <a name="requirements"></a><span data-ttu-id="cc78d-409">Requirements</span><span class="sxs-lookup"><span data-stu-id="cc78d-409">Requirements</span></span>

|<span data-ttu-id="cc78d-410">要求</span><span class="sxs-lookup"><span data-stu-id="cc78d-410">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="cc78d-411">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="cc78d-411">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="cc78d-412">1.0</span><span class="sxs-lookup"><span data-stu-id="cc78d-412">1.0</span></span>|<span data-ttu-id="cc78d-413">1.7</span><span class="sxs-lookup"><span data-stu-id="cc78d-413">1.7</span></span>|
|[<span data-ttu-id="cc78d-414">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="cc78d-414">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="cc78d-415">ReadItem</span><span class="sxs-lookup"><span data-stu-id="cc78d-415">ReadItem</span></span>|<span data-ttu-id="cc78d-416">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="cc78d-416">ReadWriteItem</span></span>|
|[<span data-ttu-id="cc78d-417">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="cc78d-417">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="cc78d-418">阅读</span><span class="sxs-lookup"><span data-stu-id="cc78d-418">Read</span></span>|<span data-ttu-id="cc78d-419">撰写</span><span class="sxs-lookup"><span data-stu-id="cc78d-419">Compose</span></span>|

---
---

#### <a name="internetheaders-internetheadersjavascriptapioutlookofficeinternetheaders"></a><span data-ttu-id="cc78d-420">internetHeaders:[internetHeaders](/javascript/api/outlook/office.internetheaders)</span><span class="sxs-lookup"><span data-stu-id="cc78d-420">internetHeaders :[InternetHeaders](/javascript/api/outlook/office.internetheaders)</span></span>

<span data-ttu-id="cc78d-421">获取或设置邮件的 internet 邮件头。</span><span class="sxs-lookup"><span data-stu-id="cc78d-421">Gets or sets the internet headers of a message.</span></span>

##### <a name="type"></a><span data-ttu-id="cc78d-422">Type</span><span class="sxs-lookup"><span data-stu-id="cc78d-422">Type</span></span>

*   [<span data-ttu-id="cc78d-423">InternetHeaders</span><span class="sxs-lookup"><span data-stu-id="cc78d-423">InternetHeaders</span></span>](/javascript/api/outlook/office.internetheaders)

##### <a name="requirements"></a><span data-ttu-id="cc78d-424">Requirements</span><span class="sxs-lookup"><span data-stu-id="cc78d-424">Requirements</span></span>

|<span data-ttu-id="cc78d-425">要求</span><span class="sxs-lookup"><span data-stu-id="cc78d-425">Requirement</span></span>|<span data-ttu-id="cc78d-426">值</span><span class="sxs-lookup"><span data-stu-id="cc78d-426">Value</span></span>|
|---|---|
|[<span data-ttu-id="cc78d-427">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="cc78d-427">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="cc78d-428">预览</span><span class="sxs-lookup"><span data-stu-id="cc78d-428">Preview</span></span>|
|[<span data-ttu-id="cc78d-429">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="cc78d-429">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="cc78d-430">ReadItem</span><span class="sxs-lookup"><span data-stu-id="cc78d-430">ReadItem</span></span>|
|[<span data-ttu-id="cc78d-431">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="cc78d-431">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="cc78d-432">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="cc78d-432">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="cc78d-433">示例</span><span class="sxs-lookup"><span data-stu-id="cc78d-433">Example</span></span>

```javascript
Office.context.mailbox.item.internetHeaders.getAsync(["header1", "header2"], callback);

function callback(asyncResult) {
  var dictionary = asyncResult.value;
  var header1_value = dictionary["header1"];
}
```

---
---

#### <a name="internetmessageid-string"></a><span data-ttu-id="cc78d-434">internetMessageId :String</span><span class="sxs-lookup"><span data-stu-id="cc78d-434">internetMessageId :String</span></span>

<span data-ttu-id="cc78d-p113">获取电子邮件的 Internet 消息标识符。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="cc78d-p113">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="cc78d-437">Type</span><span class="sxs-lookup"><span data-stu-id="cc78d-437">Type</span></span>

*   <span data-ttu-id="cc78d-438">String</span><span class="sxs-lookup"><span data-stu-id="cc78d-438">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="cc78d-439">Requirements</span><span class="sxs-lookup"><span data-stu-id="cc78d-439">Requirements</span></span>

|<span data-ttu-id="cc78d-440">要求</span><span class="sxs-lookup"><span data-stu-id="cc78d-440">Requirement</span></span>|<span data-ttu-id="cc78d-441">值</span><span class="sxs-lookup"><span data-stu-id="cc78d-441">Value</span></span>|
|---|---|
|[<span data-ttu-id="cc78d-442">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="cc78d-442">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="cc78d-443">1.0</span><span class="sxs-lookup"><span data-stu-id="cc78d-443">1.0</span></span>|
|[<span data-ttu-id="cc78d-444">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="cc78d-444">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="cc78d-445">ReadItem</span><span class="sxs-lookup"><span data-stu-id="cc78d-445">ReadItem</span></span>|
|[<span data-ttu-id="cc78d-446">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="cc78d-446">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="cc78d-447">阅读</span><span class="sxs-lookup"><span data-stu-id="cc78d-447">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="cc78d-448">示例</span><span class="sxs-lookup"><span data-stu-id="cc78d-448">Example</span></span>

```javascript
var internetMessageId = Office.context.mailbox.item.internetMessageId;
console.log("internetMessageId: " + internetMessageId);
```

---
---

#### <a name="itemclass-string"></a><span data-ttu-id="cc78d-449">itemClass :String</span><span class="sxs-lookup"><span data-stu-id="cc78d-449">itemClass :String</span></span>

<span data-ttu-id="cc78d-p114">获取选定项目的 Exchange Web 服务项目类。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="cc78d-p114">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="cc78d-p115">`itemClass` 属性指定所选项目的邮件类别。以下是邮件或约会项目的默认邮件类别。</span><span class="sxs-lookup"><span data-stu-id="cc78d-p115">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

|<span data-ttu-id="cc78d-454">类型</span><span class="sxs-lookup"><span data-stu-id="cc78d-454">Type</span></span>|<span data-ttu-id="cc78d-455">说明</span><span class="sxs-lookup"><span data-stu-id="cc78d-455">Description</span></span>|<span data-ttu-id="cc78d-456">项目类</span><span class="sxs-lookup"><span data-stu-id="cc78d-456">item class</span></span>|
|---|---|---|
|<span data-ttu-id="cc78d-457">约会项目</span><span class="sxs-lookup"><span data-stu-id="cc78d-457">Appointment items</span></span>|<span data-ttu-id="cc78d-458">这些是项目类 `IPM.Appointment` 或 `IPM.Appointment.Occurrence` 的日历项目。</span><span class="sxs-lookup"><span data-stu-id="cc78d-458">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurrence`.</span></span>|`IPM.Appointment`<br />`IPM.Appointment.Occurrence`|
|<span data-ttu-id="cc78d-459">邮件项目</span><span class="sxs-lookup"><span data-stu-id="cc78d-459">Message items</span></span>|<span data-ttu-id="cc78d-460">这些项目包括具有默认邮件类别 `IPM.Note` 的电子邮件，以及使用 `IPM.Schedule.Meeting` 作为基础邮件类别的会议请求、响应和取消。</span><span class="sxs-lookup"><span data-stu-id="cc78d-460">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span>|`IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled`|

<span data-ttu-id="cc78d-461">你可以创建用于扩展默认邮件类别的自定义邮件类别，例如，自定义约会邮件类别 `IPM.Appointment.Contoso`。</span><span class="sxs-lookup"><span data-stu-id="cc78d-461">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="cc78d-462">类型</span><span class="sxs-lookup"><span data-stu-id="cc78d-462">Type</span></span>

*   <span data-ttu-id="cc78d-463">String</span><span class="sxs-lookup"><span data-stu-id="cc78d-463">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="cc78d-464">Requirements</span><span class="sxs-lookup"><span data-stu-id="cc78d-464">Requirements</span></span>

|<span data-ttu-id="cc78d-465">要求</span><span class="sxs-lookup"><span data-stu-id="cc78d-465">Requirement</span></span>|<span data-ttu-id="cc78d-466">值</span><span class="sxs-lookup"><span data-stu-id="cc78d-466">Value</span></span>|
|---|---|
|[<span data-ttu-id="cc78d-467">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="cc78d-467">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="cc78d-468">1.0</span><span class="sxs-lookup"><span data-stu-id="cc78d-468">1.0</span></span>|
|[<span data-ttu-id="cc78d-469">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="cc78d-469">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="cc78d-470">ReadItem</span><span class="sxs-lookup"><span data-stu-id="cc78d-470">ReadItem</span></span>|
|[<span data-ttu-id="cc78d-471">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="cc78d-471">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="cc78d-472">阅读</span><span class="sxs-lookup"><span data-stu-id="cc78d-472">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="cc78d-473">示例</span><span class="sxs-lookup"><span data-stu-id="cc78d-473">Example</span></span>

```javascript
var itemClass = Office.context.mailbox.item.itemClass;
console.log("Item class: " + itemClass);
```

---
---

#### <a name="nullable-itemid-string"></a><span data-ttu-id="cc78d-474">(nullable) itemId :String</span><span class="sxs-lookup"><span data-stu-id="cc78d-474">(nullable) itemId :String</span></span>

<span data-ttu-id="cc78d-p116">获取当前项目的 Exchange Web 服务项目标识符。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="cc78d-p116">Gets the Exchange Web Services item identifier for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="cc78d-477">`itemId` 属性返回的标识符与 Exchange Web 服务项目标识符相同。</span><span class="sxs-lookup"><span data-stu-id="cc78d-477">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="cc78d-478">`itemId` 属性与 Outlook 条目 ID 或 Outlook REST API 使用的 ID 不同。</span><span class="sxs-lookup"><span data-stu-id="cc78d-478">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="cc78d-479">使用此值进行 REST API 调用前，应使用 [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string) 对它进行转换。</span><span class="sxs-lookup"><span data-stu-id="cc78d-479">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="cc78d-480">有关详细信息，请参阅[从 Outlook 加载项使用 Outlook REST API](/outlook/add-ins/use-rest-api#get-the-item-id)。</span><span class="sxs-lookup"><span data-stu-id="cc78d-480">For more details, see [Use the Outlook REST APIs from an Outlook add-in](/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

<span data-ttu-id="cc78d-p118">`itemId` 属性在撰写模式下不可用。如果需要项目标识符，[`saveAsync`](#saveasyncoptions-callback) 方法可用于将项目保存到存储，这将在回调函数的 [`AsyncResult.value`](/javascript/api/office/office.asyncresult) 参数中返回项目标识符。</span><span class="sxs-lookup"><span data-stu-id="cc78d-p118">The `itemId` property is not available in compose mode. If an item identifier is required, the [`saveAsync`](#saveasyncoptions-callback) method can be used to save the item to the store, which will return the item identifier in the [`AsyncResult.value`](/javascript/api/office/office.asyncresult) parameter in the callback function.</span></span>

##### <a name="type"></a><span data-ttu-id="cc78d-483">Type</span><span class="sxs-lookup"><span data-stu-id="cc78d-483">Type</span></span>

*   <span data-ttu-id="cc78d-484">String</span><span class="sxs-lookup"><span data-stu-id="cc78d-484">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="cc78d-485">Requirements</span><span class="sxs-lookup"><span data-stu-id="cc78d-485">Requirements</span></span>

|<span data-ttu-id="cc78d-486">要求</span><span class="sxs-lookup"><span data-stu-id="cc78d-486">Requirement</span></span>|<span data-ttu-id="cc78d-487">值</span><span class="sxs-lookup"><span data-stu-id="cc78d-487">Value</span></span>|
|---|---|
|[<span data-ttu-id="cc78d-488">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="cc78d-488">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="cc78d-489">1.0</span><span class="sxs-lookup"><span data-stu-id="cc78d-489">1.0</span></span>|
|[<span data-ttu-id="cc78d-490">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="cc78d-490">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="cc78d-491">ReadItem</span><span class="sxs-lookup"><span data-stu-id="cc78d-491">ReadItem</span></span>|
|[<span data-ttu-id="cc78d-492">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="cc78d-492">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="cc78d-493">阅读</span><span class="sxs-lookup"><span data-stu-id="cc78d-493">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="cc78d-494">示例</span><span class="sxs-lookup"><span data-stu-id="cc78d-494">Example</span></span>

<span data-ttu-id="cc78d-p119">以下代码检查项目标识符是否存在。如果 `itemId` 属性返回 `null` 或 `undefined`，则将项目保存到存储，并从异步结果中获取项目标识符。</span><span class="sxs-lookup"><span data-stu-id="cc78d-p119">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

```javascript
var itemId = Office.context.mailbox.item.itemId;
if (itemId === null || itemId == undefined) {
  Office.context.mailbox.item.saveAsync(function(result) {
    itemId = result.value;
  });
}
```

---
---

####  <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtype"></a><span data-ttu-id="cc78d-497">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype)</span><span class="sxs-lookup"><span data-stu-id="cc78d-497">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype)</span></span>

<span data-ttu-id="cc78d-498">获取实例表示的项的类型。</span><span class="sxs-lookup"><span data-stu-id="cc78d-498">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="cc78d-499">`itemType` 属性返回其中一个 `ItemType` 枚举值，指示 `item` 对象实例是邮件还是约会。</span><span class="sxs-lookup"><span data-stu-id="cc78d-499">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="cc78d-500">类型</span><span class="sxs-lookup"><span data-stu-id="cc78d-500">Type</span></span>

*   [<span data-ttu-id="cc78d-501">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="cc78d-501">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook/office.mailboxenums.itemtype)

##### <a name="requirements"></a><span data-ttu-id="cc78d-502">Requirements</span><span class="sxs-lookup"><span data-stu-id="cc78d-502">Requirements</span></span>

|<span data-ttu-id="cc78d-503">要求</span><span class="sxs-lookup"><span data-stu-id="cc78d-503">Requirement</span></span>|<span data-ttu-id="cc78d-504">值</span><span class="sxs-lookup"><span data-stu-id="cc78d-504">Value</span></span>|
|---|---|
|[<span data-ttu-id="cc78d-505">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="cc78d-505">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="cc78d-506">1.0</span><span class="sxs-lookup"><span data-stu-id="cc78d-506">1.0</span></span>|
|[<span data-ttu-id="cc78d-507">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="cc78d-507">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="cc78d-508">ReadItem</span><span class="sxs-lookup"><span data-stu-id="cc78d-508">ReadItem</span></span>|
|[<span data-ttu-id="cc78d-509">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="cc78d-509">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="cc78d-510">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="cc78d-510">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="cc78d-511">示例</span><span class="sxs-lookup"><span data-stu-id="cc78d-511">Example</span></span>

```javascript
if (Office.context.mailbox.item.itemType === Office.MailboxEnums.ItemType.Message) {
  // Do something.
} else {
  // Do something else.
}
```

---
---

####  <a name="location-stringlocationjavascriptapioutlookofficelocation"></a><span data-ttu-id="cc78d-512">location :String|[Location](/javascript/api/outlook/office.location)</span><span class="sxs-lookup"><span data-stu-id="cc78d-512">location :String|[Location](/javascript/api/outlook/office.location)</span></span>

<span data-ttu-id="cc78d-513">获取或设置约会的位置。</span><span class="sxs-lookup"><span data-stu-id="cc78d-513">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="cc78d-514">阅读模式</span><span class="sxs-lookup"><span data-stu-id="cc78d-514">Read mode</span></span>

<span data-ttu-id="cc78d-515">`location` 属性返回一个包含约会位置的字符串。</span><span class="sxs-lookup"><span data-stu-id="cc78d-515">The `location` property returns a string that contains the location of the appointment.</span></span>

```javascript
var location = Office.context.mailbox.item.location;
console.log("location: " + location);
```

##### <a name="compose-mode"></a><span data-ttu-id="cc78d-516">撰写模式</span><span class="sxs-lookup"><span data-stu-id="cc78d-516">Compose mode</span></span>

<span data-ttu-id="cc78d-517">`location` 属性返回一个 `Location` 对象，该对象提供用于获取和设置约会位置的方法。</span><span class="sxs-lookup"><span data-stu-id="cc78d-517">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

```javascript
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="cc78d-518">类型</span><span class="sxs-lookup"><span data-stu-id="cc78d-518">Type</span></span>

*   <span data-ttu-id="cc78d-519">String | [Location](/javascript/api/outlook/office.location)</span><span class="sxs-lookup"><span data-stu-id="cc78d-519">String | [Location](/javascript/api/outlook/office.location)</span></span>

##### <a name="requirements"></a><span data-ttu-id="cc78d-520">Requirements</span><span class="sxs-lookup"><span data-stu-id="cc78d-520">Requirements</span></span>

|<span data-ttu-id="cc78d-521">要求</span><span class="sxs-lookup"><span data-stu-id="cc78d-521">Requirement</span></span>|<span data-ttu-id="cc78d-522">值</span><span class="sxs-lookup"><span data-stu-id="cc78d-522">Value</span></span>|
|---|---|
|[<span data-ttu-id="cc78d-523">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="cc78d-523">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="cc78d-524">1.0</span><span class="sxs-lookup"><span data-stu-id="cc78d-524">1.0</span></span>|
|[<span data-ttu-id="cc78d-525">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="cc78d-525">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="cc78d-526">ReadItem</span><span class="sxs-lookup"><span data-stu-id="cc78d-526">ReadItem</span></span>|
|[<span data-ttu-id="cc78d-527">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="cc78d-527">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="cc78d-528">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="cc78d-528">Compose or Read</span></span>|

---
---

#### <a name="normalizedsubject-string"></a><span data-ttu-id="cc78d-529">normalizedSubject :String</span><span class="sxs-lookup"><span data-stu-id="cc78d-529">normalizedSubject :String</span></span>

<span data-ttu-id="cc78d-p120">获取删除了所有前缀（包括 `RE:` 和 `FWD:`）的项目主题。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="cc78d-p120">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="cc78d-p121">normalizedSubject 属性获取包含由电子邮件程序添加的任何标准前缀（如 `RE:` 和 `FW:`）的项目主题。若要获取包含完整前缀的项目主题，请使用 [`subject`](#subject-stringsubject) 属性。</span><span class="sxs-lookup"><span data-stu-id="cc78d-p121">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="cc78d-534">Type</span><span class="sxs-lookup"><span data-stu-id="cc78d-534">Type</span></span>

*   <span data-ttu-id="cc78d-535">String</span><span class="sxs-lookup"><span data-stu-id="cc78d-535">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="cc78d-536">Requirements</span><span class="sxs-lookup"><span data-stu-id="cc78d-536">Requirements</span></span>

|<span data-ttu-id="cc78d-537">要求</span><span class="sxs-lookup"><span data-stu-id="cc78d-537">Requirement</span></span>|<span data-ttu-id="cc78d-538">值</span><span class="sxs-lookup"><span data-stu-id="cc78d-538">Value</span></span>|
|---|---|
|[<span data-ttu-id="cc78d-539">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="cc78d-539">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="cc78d-540">1.0</span><span class="sxs-lookup"><span data-stu-id="cc78d-540">1.0</span></span>|
|[<span data-ttu-id="cc78d-541">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="cc78d-541">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="cc78d-542">ReadItem</span><span class="sxs-lookup"><span data-stu-id="cc78d-542">ReadItem</span></span>|
|[<span data-ttu-id="cc78d-543">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="cc78d-543">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="cc78d-544">阅读</span><span class="sxs-lookup"><span data-stu-id="cc78d-544">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="cc78d-545">示例</span><span class="sxs-lookup"><span data-stu-id="cc78d-545">Example</span></span>

```javascript
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
console.log("Normalized subject: " + normalizedSubject);
```

---
---

####  <a name="notificationmessages-notificationmessagesjavascriptapioutlookofficenotificationmessages"></a><span data-ttu-id="cc78d-546">notificationMessages :[NotificationMessages](/javascript/api/outlook/office.notificationmessages)</span><span class="sxs-lookup"><span data-stu-id="cc78d-546">notificationMessages :[NotificationMessages](/javascript/api/outlook/office.notificationmessages)</span></span>

<span data-ttu-id="cc78d-547">获取项目的通知邮件。</span><span class="sxs-lookup"><span data-stu-id="cc78d-547">Gets the notification messages for an item.</span></span>

##### <a name="type"></a><span data-ttu-id="cc78d-548">类型</span><span class="sxs-lookup"><span data-stu-id="cc78d-548">Type</span></span>

*   [<span data-ttu-id="cc78d-549">NotificationMessages</span><span class="sxs-lookup"><span data-stu-id="cc78d-549">NotificationMessages</span></span>](/javascript/api/outlook/office.notificationmessages)

##### <a name="requirements"></a><span data-ttu-id="cc78d-550">要求</span><span class="sxs-lookup"><span data-stu-id="cc78d-550">Requirements</span></span>

|<span data-ttu-id="cc78d-551">要求</span><span class="sxs-lookup"><span data-stu-id="cc78d-551">Requirement</span></span>|<span data-ttu-id="cc78d-552">值</span><span class="sxs-lookup"><span data-stu-id="cc78d-552">Value</span></span>|
|---|---|
|[<span data-ttu-id="cc78d-553">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="cc78d-553">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="cc78d-554">1.3</span><span class="sxs-lookup"><span data-stu-id="cc78d-554">1.3</span></span>|
|[<span data-ttu-id="cc78d-555">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="cc78d-555">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="cc78d-556">ReadItem</span><span class="sxs-lookup"><span data-stu-id="cc78d-556">ReadItem</span></span>|
|[<span data-ttu-id="cc78d-557">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="cc78d-557">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="cc78d-558">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="cc78d-558">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="cc78d-559">示例</span><span class="sxs-lookup"><span data-stu-id="cc78d-559">Example</span></span>

```javascript
// Get all notifications.
Office.context.mailbox.item.notificationMessages.getAllAsync(
  function (asyncResult) {
    console.log(JSON.stringify(asyncResult));
  }
);
```

---
---

####  <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="cc78d-560">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="cc78d-560">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="cc78d-561">提供对事件的可选与会者的访问权限。</span><span class="sxs-lookup"><span data-stu-id="cc78d-561">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="cc78d-562">对象的类型和访问级别取决于当前项目的模式。</span><span class="sxs-lookup"><span data-stu-id="cc78d-562">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="cc78d-563">阅读模式</span><span class="sxs-lookup"><span data-stu-id="cc78d-563">Read mode</span></span>

<span data-ttu-id="cc78d-564">`optionalAttendees` 属性返回一个数组，其中包含每个可选与会者的 `EmailAddressDetails` 对象。</span><span class="sxs-lookup"><span data-stu-id="cc78d-564">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span>

```javascript
var optionalAttendees = Office.context.mailbox.item.optionalAttendees;
console.log("Optional attendees: " + JSON.stringify(optionalAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="cc78d-565">撰写模式</span><span class="sxs-lookup"><span data-stu-id="cc78d-565">Compose mode</span></span>

<span data-ttu-id="cc78d-566">`optionalAttendees` 属性返回一个 `Recipients` 对象，该对象提供用于获取或更新可选与会者的方法。</span><span class="sxs-lookup"><span data-stu-id="cc78d-566">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span>

```javascript
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="cc78d-567">类型</span><span class="sxs-lookup"><span data-stu-id="cc78d-567">Type</span></span>

*   <span data-ttu-id="cc78d-568">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="cc78d-568">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="cc78d-569">Requirements</span><span class="sxs-lookup"><span data-stu-id="cc78d-569">Requirements</span></span>

|<span data-ttu-id="cc78d-570">要求</span><span class="sxs-lookup"><span data-stu-id="cc78d-570">Requirement</span></span>|<span data-ttu-id="cc78d-571">值</span><span class="sxs-lookup"><span data-stu-id="cc78d-571">Value</span></span>|
|---|---|
|[<span data-ttu-id="cc78d-572">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="cc78d-572">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="cc78d-573">1.0</span><span class="sxs-lookup"><span data-stu-id="cc78d-573">1.0</span></span>|
|[<span data-ttu-id="cc78d-574">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="cc78d-574">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="cc78d-575">ReadItem</span><span class="sxs-lookup"><span data-stu-id="cc78d-575">ReadItem</span></span>|
|[<span data-ttu-id="cc78d-576">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="cc78d-576">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="cc78d-577">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="cc78d-577">Compose or Read</span></span>|

---
---

#### <a name="organizer-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsorganizerjavascriptapioutlookofficeorganizer"></a><span data-ttu-id="cc78d-578">组织者:[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[组织者](/javascript/api/outlook/office.organizer)</span><span class="sxs-lookup"><span data-stu-id="cc78d-578">organizer :[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[Organizer](/javascript/api/outlook/office.organizer)</span></span>

<span data-ttu-id="cc78d-579">获取指定会议的组织者的电子邮件地址。</span><span class="sxs-lookup"><span data-stu-id="cc78d-579">Gets the email address of the organizer for a specified meeting.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="cc78d-580">阅读模式</span><span class="sxs-lookup"><span data-stu-id="cc78d-580">Read mode</span></span>

<span data-ttu-id="cc78d-581">该`organizer`属性返回一个[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)对象, 该对象代表会议组织者。</span><span class="sxs-lookup"><span data-stu-id="cc78d-581">The `organizer` property returns an [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) object that represents the meeting organizer.</span></span>

```javascript
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
console.log("Organizer: " + organizerName + " (" + organizerAddress + ")");
```

##### <a name="compose-mode"></a><span data-ttu-id="cc78d-582">撰写模式</span><span class="sxs-lookup"><span data-stu-id="cc78d-582">Compose mode</span></span>

<span data-ttu-id="cc78d-583">该`organizer`属性返回一个[管理](/javascript/api/outlook/office.organizer)器对象, 该对象提供获取组织者值的方法。</span><span class="sxs-lookup"><span data-stu-id="cc78d-583">The `organizer` property returns an [Organizer](/javascript/api/outlook/office.organizer) object that provides a method to get the organizer value.</span></span>

```javascript
Office.context.mailbox.item.organizer.getAsync(
  function(asyncResult) {
    console.log(JSON.stringify(asyncResult));
  }
);
```

##### <a name="type"></a><span data-ttu-id="cc78d-584">Type</span><span class="sxs-lookup"><span data-stu-id="cc78d-584">Type</span></span>

*   <span data-ttu-id="cc78d-585">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [组织者](/javascript/api/outlook/office.organizer)</span><span class="sxs-lookup"><span data-stu-id="cc78d-585">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [Organizer](/javascript/api/outlook/office.organizer)</span></span>

##### <a name="requirements"></a><span data-ttu-id="cc78d-586">Requirements</span><span class="sxs-lookup"><span data-stu-id="cc78d-586">Requirements</span></span>

|<span data-ttu-id="cc78d-587">要求</span><span class="sxs-lookup"><span data-stu-id="cc78d-587">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="cc78d-588">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="cc78d-588">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="cc78d-589">1.0</span><span class="sxs-lookup"><span data-stu-id="cc78d-589">1.0</span></span>|<span data-ttu-id="cc78d-590">1.7</span><span class="sxs-lookup"><span data-stu-id="cc78d-590">1.7</span></span>|
|[<span data-ttu-id="cc78d-591">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="cc78d-591">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="cc78d-592">ReadItem</span><span class="sxs-lookup"><span data-stu-id="cc78d-592">ReadItem</span></span>|<span data-ttu-id="cc78d-593">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="cc78d-593">ReadWriteItem</span></span>|
|[<span data-ttu-id="cc78d-594">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="cc78d-594">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="cc78d-595">阅读</span><span class="sxs-lookup"><span data-stu-id="cc78d-595">Read</span></span>|<span data-ttu-id="cc78d-596">撰写</span><span class="sxs-lookup"><span data-stu-id="cc78d-596">Compose</span></span>|

---
---

#### <a name="nullable-recurrence-recurrencejavascriptapioutlookofficerecurrence"></a><span data-ttu-id="cc78d-597">(可以为 null) 定期:[定期](/javascript/api/outlook/office.recurrence)</span><span class="sxs-lookup"><span data-stu-id="cc78d-597">(nullable) recurrence :[Recurrence](/javascript/api/outlook/office.recurrence)</span></span>

<span data-ttu-id="cc78d-598">获取或设置约会的定期模式。</span><span class="sxs-lookup"><span data-stu-id="cc78d-598">Gets or sets the recurrence pattern of an appointment.</span></span> <span data-ttu-id="cc78d-599">获取会议请求的定期模式。</span><span class="sxs-lookup"><span data-stu-id="cc78d-599">Gets the recurrence pattern of a meeting request.</span></span> <span data-ttu-id="cc78d-600">约会项目的阅读和撰写模式。</span><span class="sxs-lookup"><span data-stu-id="cc78d-600">Read and compose modes for appointment items.</span></span> <span data-ttu-id="cc78d-601">会议请求项目的阅读模式。</span><span class="sxs-lookup"><span data-stu-id="cc78d-601">Read mode for meeting request items.</span></span>

<span data-ttu-id="cc78d-602">如果`recurrence`项目是系列中的一个系列或一个实例, 则该属性返回定期约会或会议请求的[定期](/javascript/api/outlook/office.recurrence)对象。</span><span class="sxs-lookup"><span data-stu-id="cc78d-602">The `recurrence` property returns a [recurrence](/javascript/api/outlook/office.recurrence) object for recurring appointments or meetings requests if an item is a series or an instance in a series.</span></span> <span data-ttu-id="cc78d-603">`null`返回单个约会的单个约会和会议请求。</span><span class="sxs-lookup"><span data-stu-id="cc78d-603">`null` is returned for single appointments and meeting requests of single appointments.</span></span> <span data-ttu-id="cc78d-604">`undefined`对于不是会议请求的邮件, 将返回。</span><span class="sxs-lookup"><span data-stu-id="cc78d-604">`undefined` is returned for messages that are not meeting requests.</span></span>

> <span data-ttu-id="cc78d-605">注意: 会议请求的`itemClass`值为 IPM。Schedule. 会议请求。</span><span class="sxs-lookup"><span data-stu-id="cc78d-605">Note: Meeting requests have an `itemClass` value of IPM.Schedule.Meeting.Request.</span></span>

> <span data-ttu-id="cc78d-606">注意: 如果定期对象为`null`, 则表示该对象是单个约会的单个约会或会议请求, 而不是某个系列的一部分。</span><span class="sxs-lookup"><span data-stu-id="cc78d-606">Note: If the recurrence object is `null`, this indicates that the object is a single appointment or a meeting request of a single appointment and NOT a part of a series.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="cc78d-607">阅读模式</span><span class="sxs-lookup"><span data-stu-id="cc78d-607">Read mode</span></span>

<span data-ttu-id="cc78d-608">该`recurrence`属性返回一个代表约会定期的[定期](/javascript/api/outlook/office.recurrence)对象。</span><span class="sxs-lookup"><span data-stu-id="cc78d-608">The `recurrence` property returns a [Recurrence](/javascript/api/outlook/office.recurrence) object that represents the appointment recurrence.</span></span> <span data-ttu-id="cc78d-609">此功能适用于约会和会议请求。</span><span class="sxs-lookup"><span data-stu-id="cc78d-609">This is available for appointments and meeting requests.</span></span>

```javascript
var recurrence = Office.context.mailbox.item.recurrence;
console.log("Recurrence: " + JSON.stringify(recurrence));
```

##### <a name="compose-mode"></a><span data-ttu-id="cc78d-610">撰写模式</span><span class="sxs-lookup"><span data-stu-id="cc78d-610">Compose mode</span></span>

<span data-ttu-id="cc78d-611">该`recurrence`属性返回一个[定期](/javascript/api/outlook/office.recurrence)对象, 该对象提供用于管理约会周期的方法。</span><span class="sxs-lookup"><span data-stu-id="cc78d-611">The `recurrence` property returns a [Recurrence](/javascript/api/outlook/office.recurrence) object that provides methods to manage the appointment recurrence.</span></span> <span data-ttu-id="cc78d-612">这可用于约会。</span><span class="sxs-lookup"><span data-stu-id="cc78d-612">This is available for appointments.</span></span>

```javascript
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

##### <a name="type"></a><span data-ttu-id="cc78d-613">Type</span><span class="sxs-lookup"><span data-stu-id="cc78d-613">Type</span></span>

* [<span data-ttu-id="cc78d-614">循环</span><span class="sxs-lookup"><span data-stu-id="cc78d-614">Recurrence</span></span>](/javascript/api/outlook/office.recurrence)

|<span data-ttu-id="cc78d-615">要求</span><span class="sxs-lookup"><span data-stu-id="cc78d-615">Requirement</span></span>|<span data-ttu-id="cc78d-616">值</span><span class="sxs-lookup"><span data-stu-id="cc78d-616">Value</span></span>|
|---|---|
|[<span data-ttu-id="cc78d-617">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="cc78d-617">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="cc78d-618">1.7</span><span class="sxs-lookup"><span data-stu-id="cc78d-618">1.7</span></span>|
|[<span data-ttu-id="cc78d-619">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="cc78d-619">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="cc78d-620">ReadItem</span><span class="sxs-lookup"><span data-stu-id="cc78d-620">ReadItem</span></span>|
|[<span data-ttu-id="cc78d-621">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="cc78d-621">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="cc78d-622">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="cc78d-622">Compose or Read</span></span>|

---
---

####  <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="cc78d-623">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="cc78d-623">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="cc78d-624">提供对事件的必需与会者的访问权限。</span><span class="sxs-lookup"><span data-stu-id="cc78d-624">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="cc78d-625">对象的类型和访问级别取决于当前项目的模式。</span><span class="sxs-lookup"><span data-stu-id="cc78d-625">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="cc78d-626">阅读模式</span><span class="sxs-lookup"><span data-stu-id="cc78d-626">Read mode</span></span>

<span data-ttu-id="cc78d-627">`requiredAttendees` 属性返回一个数组，其中包含每个必需与会者的 `EmailAddressDetails` 对象。</span><span class="sxs-lookup"><span data-stu-id="cc78d-627">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span>

```javascript
var requiredAttendees = Office.context.mailbox.item.requiredAttendees;
console.log("Required attendees: " + JSON.stringify(requiredAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="cc78d-628">撰写模式</span><span class="sxs-lookup"><span data-stu-id="cc78d-628">Compose mode</span></span>

<span data-ttu-id="cc78d-629">`requiredAttendees` 属性返回一个 `Recipients` 对象，该对象提供用于获取或更新必需与会者的方法。</span><span class="sxs-lookup"><span data-stu-id="cc78d-629">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span>

```javascript
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
  console.log(JSON.stringify(arrayOfRequiredAttendeesRecipients));
}
```

##### <a name="type"></a><span data-ttu-id="cc78d-630">类型</span><span class="sxs-lookup"><span data-stu-id="cc78d-630">Type</span></span>

*   <span data-ttu-id="cc78d-631">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="cc78d-631">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="cc78d-632">Requirements</span><span class="sxs-lookup"><span data-stu-id="cc78d-632">Requirements</span></span>

|<span data-ttu-id="cc78d-633">要求</span><span class="sxs-lookup"><span data-stu-id="cc78d-633">Requirement</span></span>|<span data-ttu-id="cc78d-634">值</span><span class="sxs-lookup"><span data-stu-id="cc78d-634">Value</span></span>|
|---|---|
|[<span data-ttu-id="cc78d-635">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="cc78d-635">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="cc78d-636">1.0</span><span class="sxs-lookup"><span data-stu-id="cc78d-636">1.0</span></span>|
|[<span data-ttu-id="cc78d-637">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="cc78d-637">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="cc78d-638">ReadItem</span><span class="sxs-lookup"><span data-stu-id="cc78d-638">ReadItem</span></span>|
|[<span data-ttu-id="cc78d-639">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="cc78d-639">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="cc78d-640">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="cc78d-640">Compose or Read</span></span>|

---
---

#### <a name="sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetails"></a><span data-ttu-id="cc78d-641">sender :[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="cc78d-641">sender :[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)</span></span>

<span data-ttu-id="cc78d-p128">获取电子邮件发件人的电子邮件地址。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="cc78d-p128">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="cc78d-p129">[`from`](#from-emailaddressdetailsfrom) 和 `sender` 属性表示同一个人，邮件由代理人发送的除外。在此情况下，`from` 属性表示代理程序，而 sender 属性表示代理人。</span><span class="sxs-lookup"><span data-stu-id="cc78d-p129">The [`from`](#from-emailaddressdetailsfrom) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="cc78d-646">`sender` 属性中 `EmailAddressDetails` 对象的 `recipientType` 属性为 `undefined`。</span><span class="sxs-lookup"><span data-stu-id="cc78d-646">The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="cc78d-647">类型</span><span class="sxs-lookup"><span data-stu-id="cc78d-647">Type</span></span>

*   [<span data-ttu-id="cc78d-648">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="cc78d-648">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="cc78d-649">要求</span><span class="sxs-lookup"><span data-stu-id="cc78d-649">Requirements</span></span>

|<span data-ttu-id="cc78d-650">要求</span><span class="sxs-lookup"><span data-stu-id="cc78d-650">Requirement</span></span>|<span data-ttu-id="cc78d-651">值</span><span class="sxs-lookup"><span data-stu-id="cc78d-651">Value</span></span>|
|---|---|
|[<span data-ttu-id="cc78d-652">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="cc78d-652">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="cc78d-653">1.0</span><span class="sxs-lookup"><span data-stu-id="cc78d-653">1.0</span></span>|
|[<span data-ttu-id="cc78d-654">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="cc78d-654">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="cc78d-655">ReadItem</span><span class="sxs-lookup"><span data-stu-id="cc78d-655">ReadItem</span></span>|
|[<span data-ttu-id="cc78d-656">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="cc78d-656">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="cc78d-657">阅读</span><span class="sxs-lookup"><span data-stu-id="cc78d-657">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="cc78d-658">示例</span><span class="sxs-lookup"><span data-stu-id="cc78d-658">Example</span></span>

```javascript
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
console.log("Sender: " + senderName + " (" + senderAddress + ")");
```

---
---

#### <a name="nullable-seriesid-string"></a><span data-ttu-id="cc78d-659">(可以为 null) webcasts&seriesid: String</span><span class="sxs-lookup"><span data-stu-id="cc78d-659">(nullable) seriesId :String</span></span>

<span data-ttu-id="cc78d-660">获取实例所属的系列的 id。</span><span class="sxs-lookup"><span data-stu-id="cc78d-660">Gets the id of the series that an instance belongs to.</span></span>

<span data-ttu-id="cc78d-661">在 OWA 和 Outlook 中, `seriesId`返回此项所属的父 (系列) 项的 Exchange Web 服务 (EWS) ID。</span><span class="sxs-lookup"><span data-stu-id="cc78d-661">In OWA and Outlook, the `seriesId` returns the Exchange Web Services (EWS) ID of the parent (series) item that this item belongs to.</span></span> <span data-ttu-id="cc78d-662">但是, 在 iOS 和 Android 中, `seriesId`将返回父项的 REST ID。</span><span class="sxs-lookup"><span data-stu-id="cc78d-662">However, in iOS and Android, the `seriesId` returns the REST ID of the parent item.</span></span>

> [!NOTE]
> <span data-ttu-id="cc78d-663">`seriesId` 属性返回的标识符与 Exchange Web 服务项目标识符相同。</span><span class="sxs-lookup"><span data-stu-id="cc78d-663">The identifier returned by the `seriesId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="cc78d-664">`seriesId`属性与 outlook REST API 使用的 outlook id 不相同。</span><span class="sxs-lookup"><span data-stu-id="cc78d-664">The `seriesId` property is not identical to the Outlook IDs used by the Outlook REST API.</span></span> <span data-ttu-id="cc78d-665">使用此值进行 REST API 调用前，应使用 [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string) 对它进行转换。</span><span class="sxs-lookup"><span data-stu-id="cc78d-665">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="cc78d-666">有关详细信息，请参阅[从 Outlook 加载项使用 Outlook REST API](/outlook/add-ins/use-rest-api)。</span><span class="sxs-lookup"><span data-stu-id="cc78d-666">For more details, see [Use the Outlook REST APIs from an Outlook add-in](/outlook/add-ins/use-rest-api).</span></span>

<span data-ttu-id="cc78d-667">对于`seriesId`不包含`null`父项 (如单个约会、系列项或会议请求) 的项, 该属性将返回, `undefined`对于不是会议请求的任何其他项, 该属性返回。</span><span class="sxs-lookup"><span data-stu-id="cc78d-667">The `seriesId` property returns `null` for items that do not have parent items such as single appointments, series items, or meeting requests and returns `undefined` for any other items that are not meeting requests.</span></span>

##### <a name="type"></a><span data-ttu-id="cc78d-668">Type</span><span class="sxs-lookup"><span data-stu-id="cc78d-668">Type</span></span>

* <span data-ttu-id="cc78d-669">String</span><span class="sxs-lookup"><span data-stu-id="cc78d-669">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="cc78d-670">Requirements</span><span class="sxs-lookup"><span data-stu-id="cc78d-670">Requirements</span></span>

|<span data-ttu-id="cc78d-671">要求</span><span class="sxs-lookup"><span data-stu-id="cc78d-671">Requirement</span></span>|<span data-ttu-id="cc78d-672">值</span><span class="sxs-lookup"><span data-stu-id="cc78d-672">Value</span></span>|
|---|---|
|[<span data-ttu-id="cc78d-673">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="cc78d-673">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="cc78d-674">1.7</span><span class="sxs-lookup"><span data-stu-id="cc78d-674">1.7</span></span>|
|[<span data-ttu-id="cc78d-675">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="cc78d-675">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="cc78d-676">ReadItem</span><span class="sxs-lookup"><span data-stu-id="cc78d-676">ReadItem</span></span>|
|[<span data-ttu-id="cc78d-677">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="cc78d-677">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="cc78d-678">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="cc78d-678">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="cc78d-679">示例</span><span class="sxs-lookup"><span data-stu-id="cc78d-679">Example</span></span>

```javascript
var seriesId = Office.context.mailbox.item.seriesId;

// The seriesId property returns null for items that do
// not have parent items (such as single appointments,
// series items, or meeting requests) and returns
// undefined for messages that are not meeting requests.
var isSeriesInstance = (seriesId != null);
console.log("SeriesId is " + seriesId + " and isSeriesInstance is " + isSeriesInstance);
```

---
---

####  <a name="start-datetimejavascriptapioutlookofficetime"></a><span data-ttu-id="cc78d-680">start :Date|[Time](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="cc78d-680">start :Date|[Time](/javascript/api/outlook/office.time)</span></span>

<span data-ttu-id="cc78d-681">获取或设置约会开始的日期和时间。</span><span class="sxs-lookup"><span data-stu-id="cc78d-681">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="cc78d-p132">将 `start` 属性表示为协调世界时 (UTC) 的日期和时间值。可使用 [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) 方法将该值转换为客户端的本地日期和时间。</span><span class="sxs-lookup"><span data-stu-id="cc78d-p132">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="cc78d-684">阅读模式</span><span class="sxs-lookup"><span data-stu-id="cc78d-684">Read mode</span></span>

<span data-ttu-id="cc78d-685">`start` 属性返回 `Date` 对象。</span><span class="sxs-lookup"><span data-stu-id="cc78d-685">The `start` property returns a `Date` object.</span></span>

```javascript
var start = Office.context.mailbox.item.start;
console.log("Appointment start: " + JSON.stringify(start));
```

##### <a name="compose-mode"></a><span data-ttu-id="cc78d-686">撰写模式</span><span class="sxs-lookup"><span data-stu-id="cc78d-686">Compose mode</span></span>

<span data-ttu-id="cc78d-687">`start` 属性返回 `Time` 对象。</span><span class="sxs-lookup"><span data-stu-id="cc78d-687">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="cc78d-688">使用 [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) 方法设置开始时间时，应使用 [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) 方法将客户端的本地时间转换为服务器的 UTC。</span><span class="sxs-lookup"><span data-stu-id="cc78d-688">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="cc78d-689">以下示例通过使用 `Time` 对象的 [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) 方法，设置撰写模式下约会的开始时间。</span><span class="sxs-lookup"><span data-stu-id="cc78d-689">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="cc78d-690">类型</span><span class="sxs-lookup"><span data-stu-id="cc78d-690">Type</span></span>

*   <span data-ttu-id="cc78d-691">Date | [Time](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="cc78d-691">Date | [Time](/javascript/api/outlook/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="cc78d-692">Requirements</span><span class="sxs-lookup"><span data-stu-id="cc78d-692">Requirements</span></span>

|<span data-ttu-id="cc78d-693">要求</span><span class="sxs-lookup"><span data-stu-id="cc78d-693">Requirement</span></span>|<span data-ttu-id="cc78d-694">值</span><span class="sxs-lookup"><span data-stu-id="cc78d-694">Value</span></span>|
|---|---|
|[<span data-ttu-id="cc78d-695">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="cc78d-695">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="cc78d-696">1.0</span><span class="sxs-lookup"><span data-stu-id="cc78d-696">1.0</span></span>|
|[<span data-ttu-id="cc78d-697">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="cc78d-697">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="cc78d-698">ReadItem</span><span class="sxs-lookup"><span data-stu-id="cc78d-698">ReadItem</span></span>|
|[<span data-ttu-id="cc78d-699">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="cc78d-699">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="cc78d-700">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="cc78d-700">Compose or Read</span></span>|

---
---

####  <a name="subject-stringsubjectjavascriptapioutlookofficesubject"></a><span data-ttu-id="cc78d-701">subject :String|[Subject](/javascript/api/outlook/office.subject)</span><span class="sxs-lookup"><span data-stu-id="cc78d-701">subject :String|[Subject](/javascript/api/outlook/office.subject)</span></span>

<span data-ttu-id="cc78d-702">获取或设置显示在项目的主题字段中的说明。</span><span class="sxs-lookup"><span data-stu-id="cc78d-702">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="cc78d-703">`subject` 属性获取或设置由电子邮件服务器发送项目时的整个主题。</span><span class="sxs-lookup"><span data-stu-id="cc78d-703">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="cc78d-704">阅读模式</span><span class="sxs-lookup"><span data-stu-id="cc78d-704">Read mode</span></span>

<span data-ttu-id="cc78d-p133">`subject` 属性返回一个字符串。使用 [`normalizedSubject`](#normalizedsubject-string) 属性获取不带任何前导前缀（如 `RE:` 和 `FW:`）的主题。</span><span class="sxs-lookup"><span data-stu-id="cc78d-p133">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

<span data-ttu-id="cc78d-707">以下 JavaScript 代码示例显示了如何访问 Outlook 中当前项目的 `subject` 属性。</span><span class="sxs-lookup"><span data-stu-id="cc78d-707">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

```javascript
var subject = Office.context.mailbox.item.subject;
console.log(subject);
```

##### <a name="compose-mode"></a><span data-ttu-id="cc78d-708">撰写模式</span><span class="sxs-lookup"><span data-stu-id="cc78d-708">Compose mode</span></span>
<span data-ttu-id="cc78d-709">`subject` 属性返回一个 `Subject` 对象，该对象提供用于获取和设置主题的方法。</span><span class="sxs-lookup"><span data-stu-id="cc78d-709">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```javascript
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
  console.log(subject);
}
```

##### <a name="type"></a><span data-ttu-id="cc78d-710">类型</span><span class="sxs-lookup"><span data-stu-id="cc78d-710">Type</span></span>

*   <span data-ttu-id="cc78d-711">String | [Subject](/javascript/api/outlook/office.subject)</span><span class="sxs-lookup"><span data-stu-id="cc78d-711">String | [Subject](/javascript/api/outlook/office.subject)</span></span>

##### <a name="requirements"></a><span data-ttu-id="cc78d-712">Requirements</span><span class="sxs-lookup"><span data-stu-id="cc78d-712">Requirements</span></span>

|<span data-ttu-id="cc78d-713">要求</span><span class="sxs-lookup"><span data-stu-id="cc78d-713">Requirement</span></span>|<span data-ttu-id="cc78d-714">值</span><span class="sxs-lookup"><span data-stu-id="cc78d-714">Value</span></span>|
|---|---|
|[<span data-ttu-id="cc78d-715">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="cc78d-715">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="cc78d-716">1.0</span><span class="sxs-lookup"><span data-stu-id="cc78d-716">1.0</span></span>|
|[<span data-ttu-id="cc78d-717">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="cc78d-717">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="cc78d-718">ReadItem</span><span class="sxs-lookup"><span data-stu-id="cc78d-718">ReadItem</span></span>|
|[<span data-ttu-id="cc78d-719">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="cc78d-719">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="cc78d-720">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="cc78d-720">Compose or Read</span></span>|

---
---

####  <a name="to-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="cc78d-721">to :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="cc78d-721">to :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="cc78d-722">提供对邮件的“**收件人**”行上的收件人的访问权限。</span><span class="sxs-lookup"><span data-stu-id="cc78d-722">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="cc78d-723">对象的类型和访问级别取决于当前项目的模式。</span><span class="sxs-lookup"><span data-stu-id="cc78d-723">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="cc78d-724">阅读模式</span><span class="sxs-lookup"><span data-stu-id="cc78d-724">Read mode</span></span>

<span data-ttu-id="cc78d-p135">`to` 属性返回包含邮件的**收件人**行上所列的每个收件人的 `EmailAddressDetails` 对象的数组。集合上限为 100 个成员。</span><span class="sxs-lookup"><span data-stu-id="cc78d-p135">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.</span></span>

```javascript
console.log(JSON.stringify(Office.context.mailbox.item.to));
```

##### <a name="compose-mode"></a><span data-ttu-id="cc78d-727">撰写模式</span><span class="sxs-lookup"><span data-stu-id="cc78d-727">Compose mode</span></span>

<span data-ttu-id="cc78d-728">`to` 属性返回一个 `Recipients` 对象，该对象提供用于获取或更新邮件的“**收件人**”行上收件人的方法。</span><span class="sxs-lookup"><span data-stu-id="cc78d-728">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span>

```javascript
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="cc78d-729">类型</span><span class="sxs-lookup"><span data-stu-id="cc78d-729">Type</span></span>

*   <span data-ttu-id="cc78d-730">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="cc78d-730">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="cc78d-731">Requirements</span><span class="sxs-lookup"><span data-stu-id="cc78d-731">Requirements</span></span>

|<span data-ttu-id="cc78d-732">要求</span><span class="sxs-lookup"><span data-stu-id="cc78d-732">Requirement</span></span>|<span data-ttu-id="cc78d-733">值</span><span class="sxs-lookup"><span data-stu-id="cc78d-733">Value</span></span>|
|---|---|
|[<span data-ttu-id="cc78d-734">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="cc78d-734">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="cc78d-735">1.0</span><span class="sxs-lookup"><span data-stu-id="cc78d-735">1.0</span></span>|
|[<span data-ttu-id="cc78d-736">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="cc78d-736">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="cc78d-737">ReadItem</span><span class="sxs-lookup"><span data-stu-id="cc78d-737">ReadItem</span></span>|
|[<span data-ttu-id="cc78d-738">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="cc78d-738">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="cc78d-739">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="cc78d-739">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="cc78d-740">方法</span><span class="sxs-lookup"><span data-stu-id="cc78d-740">Methods</span></span>

####  <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="cc78d-741">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="cc78d-741">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="cc78d-742">将文件作为附件添加到邮件或约会。</span><span class="sxs-lookup"><span data-stu-id="cc78d-742">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="cc78d-743">`addFileAttachmentAsync` 方法在指定的 URI 上载文件并将其附加到撰写窗体中的项目。</span><span class="sxs-lookup"><span data-stu-id="cc78d-743">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="cc78d-744">随后可以将该标识符与 [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) 方法一同使用，以删除同一个会话中的附件。</span><span class="sxs-lookup"><span data-stu-id="cc78d-744">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="cc78d-745">参数</span><span class="sxs-lookup"><span data-stu-id="cc78d-745">Parameters</span></span>
|<span data-ttu-id="cc78d-746">名称</span><span class="sxs-lookup"><span data-stu-id="cc78d-746">Name</span></span>|<span data-ttu-id="cc78d-747">类型</span><span class="sxs-lookup"><span data-stu-id="cc78d-747">Type</span></span>|<span data-ttu-id="cc78d-748">属性</span><span class="sxs-lookup"><span data-stu-id="cc78d-748">Attributes</span></span>|<span data-ttu-id="cc78d-749">说明</span><span class="sxs-lookup"><span data-stu-id="cc78d-749">Description</span></span>|
|---|---|---|---|
|`uri`|<span data-ttu-id="cc78d-750">字符串</span><span class="sxs-lookup"><span data-stu-id="cc78d-750">String</span></span>||<span data-ttu-id="cc78d-p136">提供附加到邮件或约会的文件的位置的 URI。最大长度为 2048 个字符。</span><span class="sxs-lookup"><span data-stu-id="cc78d-p136">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`|<span data-ttu-id="cc78d-753">字符串</span><span class="sxs-lookup"><span data-stu-id="cc78d-753">String</span></span>||<span data-ttu-id="cc78d-p137">在附件上载过程中显示的附件名称。最大长度为 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="cc78d-p137">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="cc78d-756">Object</span><span class="sxs-lookup"><span data-stu-id="cc78d-756">Object</span></span>|<span data-ttu-id="cc78d-757">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="cc78d-757">&lt;optional&gt;</span></span>|<span data-ttu-id="cc78d-758">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="cc78d-758">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="cc78d-759">对象</span><span class="sxs-lookup"><span data-stu-id="cc78d-759">Object</span></span>|<span data-ttu-id="cc78d-760">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="cc78d-760">&lt;optional&gt;</span></span>|<span data-ttu-id="cc78d-761">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="cc78d-761">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.isInline`|<span data-ttu-id="cc78d-762">布尔值</span><span class="sxs-lookup"><span data-stu-id="cc78d-762">Boolean</span></span>|<span data-ttu-id="cc78d-763">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="cc78d-763">&lt;optional&gt;</span></span>|<span data-ttu-id="cc78d-764">如果为 `true`，则表示附件将在邮件正文中内联显示，并且不应显示在附件列表中。</span><span class="sxs-lookup"><span data-stu-id="cc78d-764">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`callback`|<span data-ttu-id="cc78d-765">函数</span><span class="sxs-lookup"><span data-stu-id="cc78d-765">function</span></span>|<span data-ttu-id="cc78d-766">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="cc78d-766">&lt;optional&gt;</span></span>|<span data-ttu-id="cc78d-767">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="cc78d-767">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="cc78d-768">如果成功，附件标识符将在 `asyncResult.value` 属性中提供。</span><span class="sxs-lookup"><span data-stu-id="cc78d-768">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="cc78d-769">如果上传附件失败，`asyncResult` 对象将包含一个提供错误说明的 `Error` 对象。</span><span class="sxs-lookup"><span data-stu-id="cc78d-769">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="cc78d-770">错误</span><span class="sxs-lookup"><span data-stu-id="cc78d-770">Errors</span></span>

|<span data-ttu-id="cc78d-771">错误代码</span><span class="sxs-lookup"><span data-stu-id="cc78d-771">Error code</span></span>|<span data-ttu-id="cc78d-772">说明</span><span class="sxs-lookup"><span data-stu-id="cc78d-772">Description</span></span>|
|------------|-------------|
|`AttachmentSizeExceeded`|<span data-ttu-id="cc78d-773">附件大小超过了允许的大小。</span><span class="sxs-lookup"><span data-stu-id="cc78d-773">The attachment is larger than allowed.</span></span>|
|`FileTypeNotSupported`|<span data-ttu-id="cc78d-774">该附件的扩展名不是允许的扩展名。</span><span class="sxs-lookup"><span data-stu-id="cc78d-774">The attachment has an extension that is not allowed.</span></span>|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="cc78d-775">邮件或约会具有的附件过多。</span><span class="sxs-lookup"><span data-stu-id="cc78d-775">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="cc78d-776">要求</span><span class="sxs-lookup"><span data-stu-id="cc78d-776">Requirements</span></span>

|<span data-ttu-id="cc78d-777">要求</span><span class="sxs-lookup"><span data-stu-id="cc78d-777">Requirement</span></span>|<span data-ttu-id="cc78d-778">值</span><span class="sxs-lookup"><span data-stu-id="cc78d-778">Value</span></span>|
|---|---|
|[<span data-ttu-id="cc78d-779">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="cc78d-779">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="cc78d-780">1.1</span><span class="sxs-lookup"><span data-stu-id="cc78d-780">1.1</span></span>|
|[<span data-ttu-id="cc78d-781">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="cc78d-781">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="cc78d-782">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="cc78d-782">ReadWriteItem</span></span>|
|[<span data-ttu-id="cc78d-783">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="cc78d-783">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="cc78d-784">撰写</span><span class="sxs-lookup"><span data-stu-id="cc78d-784">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="cc78d-785">示例</span><span class="sxs-lookup"><span data-stu-id="cc78d-785">Examples</span></span>

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

<span data-ttu-id="cc78d-786">以下示例将图像文件添加为内联附件，并在邮件正文中引用该附件。</span><span class="sxs-lookup"><span data-stu-id="cc78d-786">The following example adds an image file as an inline attachment and references the attachment in the message body.</span></span>

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

---
---

#### <a name="addfileattachmentfrombase64asyncbase64file-attachmentname-options-callback"></a><span data-ttu-id="cc78d-787">addFileAttachmentFromBase64Async (base64File, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="cc78d-787">addFileAttachmentFromBase64Async(base64File, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="cc78d-788">将 base64 编码中的文件作为附件添加到邮件或约会中。</span><span class="sxs-lookup"><span data-stu-id="cc78d-788">Adds a file from the base64 encoding to a message or appointment as an attachment.</span></span>

<span data-ttu-id="cc78d-789">该`addFileAttachmentFromBase64Async`方法从 base64 编码中上载文件, 并将其附加到撰写窗体中的项目。</span><span class="sxs-lookup"><span data-stu-id="cc78d-789">The `addFileAttachmentFromBase64Async` method uploads the file from the base64 encoding and attaches it to the item in the compose form.</span></span> <span data-ttu-id="cc78d-790">此方法返回 AsyncResult 对象中的附件标识符。</span><span class="sxs-lookup"><span data-stu-id="cc78d-790">This method returns the attachment identifier in the AsyncResult.value object.</span></span>

<span data-ttu-id="cc78d-791">随后可以将该标识符与 [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) 方法一同使用，以删除同一个会话中的附件。</span><span class="sxs-lookup"><span data-stu-id="cc78d-791">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="cc78d-792">参数</span><span class="sxs-lookup"><span data-stu-id="cc78d-792">Parameters</span></span>

|<span data-ttu-id="cc78d-793">名称</span><span class="sxs-lookup"><span data-stu-id="cc78d-793">Name</span></span>|<span data-ttu-id="cc78d-794">类型</span><span class="sxs-lookup"><span data-stu-id="cc78d-794">Type</span></span>|<span data-ttu-id="cc78d-795">属性</span><span class="sxs-lookup"><span data-stu-id="cc78d-795">Attributes</span></span>|<span data-ttu-id="cc78d-796">说明</span><span class="sxs-lookup"><span data-stu-id="cc78d-796">Description</span></span>|
|---|---|---|---|
|`base64File`|<span data-ttu-id="cc78d-797">字符串</span><span class="sxs-lookup"><span data-stu-id="cc78d-797">String</span></span>||<span data-ttu-id="cc78d-798">要添加到电子邮件或事件的图像或文件的 base64 编码内容。</span><span class="sxs-lookup"><span data-stu-id="cc78d-798">The base64 encoded content of an image or file to be added to an email or event.</span></span>|
|`attachmentName`|<span data-ttu-id="cc78d-799">字符串</span><span class="sxs-lookup"><span data-stu-id="cc78d-799">String</span></span>||<span data-ttu-id="cc78d-p139">在附件上载过程中显示的附件名称。最大长度为 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="cc78d-p139">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="cc78d-802">Object</span><span class="sxs-lookup"><span data-stu-id="cc78d-802">Object</span></span>|<span data-ttu-id="cc78d-803">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="cc78d-803">&lt;optional&gt;</span></span>|<span data-ttu-id="cc78d-804">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="cc78d-804">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="cc78d-805">Object</span><span class="sxs-lookup"><span data-stu-id="cc78d-805">Object</span></span>|<span data-ttu-id="cc78d-806">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="cc78d-806">&lt;optional&gt;</span></span>|<span data-ttu-id="cc78d-807">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="cc78d-807">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.isInline`|<span data-ttu-id="cc78d-808">布尔值</span><span class="sxs-lookup"><span data-stu-id="cc78d-808">Boolean</span></span>|<span data-ttu-id="cc78d-809">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="cc78d-809">&lt;optional&gt;</span></span>|<span data-ttu-id="cc78d-810">如果为 `true`，则表示附件将在邮件正文中内联显示，并且不应显示在附件列表中。</span><span class="sxs-lookup"><span data-stu-id="cc78d-810">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`callback`|<span data-ttu-id="cc78d-811">函数</span><span class="sxs-lookup"><span data-stu-id="cc78d-811">function</span></span>|<span data-ttu-id="cc78d-812">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="cc78d-812">&lt;optional&gt;</span></span>|<span data-ttu-id="cc78d-813">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="cc78d-813">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="cc78d-814">如果成功，附件标识符将在 `asyncResult.value` 属性中提供。</span><span class="sxs-lookup"><span data-stu-id="cc78d-814">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="cc78d-815">如果上传附件失败，`asyncResult` 对象将包含一个提供错误说明的 `Error` 对象。</span><span class="sxs-lookup"><span data-stu-id="cc78d-815">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="cc78d-816">错误</span><span class="sxs-lookup"><span data-stu-id="cc78d-816">Errors</span></span>

|<span data-ttu-id="cc78d-817">错误代码</span><span class="sxs-lookup"><span data-stu-id="cc78d-817">Error code</span></span>|<span data-ttu-id="cc78d-818">说明</span><span class="sxs-lookup"><span data-stu-id="cc78d-818">Description</span></span>|
|------------|-------------|
|`AttachmentSizeExceeded`|<span data-ttu-id="cc78d-819">附件大小超过了允许的大小。</span><span class="sxs-lookup"><span data-stu-id="cc78d-819">The attachment is larger than allowed.</span></span>|
|`FileTypeNotSupported`|<span data-ttu-id="cc78d-820">该附件的扩展名不是允许的扩展名。</span><span class="sxs-lookup"><span data-stu-id="cc78d-820">The attachment has an extension that is not allowed.</span></span>|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="cc78d-821">邮件或约会具有的附件过多。</span><span class="sxs-lookup"><span data-stu-id="cc78d-821">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="cc78d-822">要求</span><span class="sxs-lookup"><span data-stu-id="cc78d-822">Requirements</span></span>

|<span data-ttu-id="cc78d-823">要求</span><span class="sxs-lookup"><span data-stu-id="cc78d-823">Requirement</span></span>|<span data-ttu-id="cc78d-824">值</span><span class="sxs-lookup"><span data-stu-id="cc78d-824">Value</span></span>|
|---|---|
|[<span data-ttu-id="cc78d-825">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="cc78d-825">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="cc78d-826">预览</span><span class="sxs-lookup"><span data-stu-id="cc78d-826">Preview</span></span>|
|[<span data-ttu-id="cc78d-827">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="cc78d-827">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="cc78d-828">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="cc78d-828">ReadWriteItem</span></span>|
|[<span data-ttu-id="cc78d-829">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="cc78d-829">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="cc78d-830">撰写</span><span class="sxs-lookup"><span data-stu-id="cc78d-830">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="cc78d-831">示例</span><span class="sxs-lookup"><span data-stu-id="cc78d-831">Examples</span></span>

```javascript
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

---
---

####  <a name="addhandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="cc78d-832">addHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="cc78d-832">addHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="cc78d-833">添加支持事件的事件处理程序。</span><span class="sxs-lookup"><span data-stu-id="cc78d-833">Adds an event handler for a supported event.</span></span>

<span data-ttu-id="cc78d-834">目前, 受支持的事件`Office.EventType.AttachmentsChanged`类型`Office.EventType.AppointmentTimeChanged`是`Office.EventType.EnhancedLocationsChanged`、 `Office.EventType.RecipientsChanged`、、 `Office.EventType.RecurrenceChanged`和。</span><span class="sxs-lookup"><span data-stu-id="cc78d-834">Currently the supported event types are `Office.EventType.AttachmentsChanged`, `Office.EventType.AppointmentTimeChanged`, `Office.EventType.EnhancedLocationsChanged`, `Office.EventType.RecipientsChanged`, and `Office.EventType.RecurrenceChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="cc78d-835">参数</span><span class="sxs-lookup"><span data-stu-id="cc78d-835">Parameters</span></span>

| <span data-ttu-id="cc78d-836">名称</span><span class="sxs-lookup"><span data-stu-id="cc78d-836">Name</span></span> | <span data-ttu-id="cc78d-837">类型</span><span class="sxs-lookup"><span data-stu-id="cc78d-837">Type</span></span> | <span data-ttu-id="cc78d-838">属性</span><span class="sxs-lookup"><span data-stu-id="cc78d-838">Attributes</span></span> | <span data-ttu-id="cc78d-839">说明</span><span class="sxs-lookup"><span data-stu-id="cc78d-839">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="cc78d-840">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="cc78d-840">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="cc78d-841">应调用处理程序的事件。</span><span class="sxs-lookup"><span data-stu-id="cc78d-841">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="cc78d-842">函数</span><span class="sxs-lookup"><span data-stu-id="cc78d-842">Function</span></span> || <span data-ttu-id="cc78d-p140">用于处理事件的函数。此函数必须接受一个参数，即对象文本。参数上的 `type` 属性将匹配传递给 `addHandlerAsync` 的 `eventType` 参数。</span><span class="sxs-lookup"><span data-stu-id="cc78d-p140">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `addHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="cc78d-846">Object</span><span class="sxs-lookup"><span data-stu-id="cc78d-846">Object</span></span> | <span data-ttu-id="cc78d-847">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="cc78d-847">&lt;optional&gt;</span></span> | <span data-ttu-id="cc78d-848">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="cc78d-848">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="cc78d-849">Object</span><span class="sxs-lookup"><span data-stu-id="cc78d-849">Object</span></span> | <span data-ttu-id="cc78d-850">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="cc78d-850">&lt;optional&gt;</span></span> | <span data-ttu-id="cc78d-851">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="cc78d-851">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="cc78d-852">函数</span><span class="sxs-lookup"><span data-stu-id="cc78d-852">function</span></span>| <span data-ttu-id="cc78d-853">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="cc78d-853">&lt;optional&gt;</span></span>|<span data-ttu-id="cc78d-854">方法完成后，使用单个参数 `callback`（一个 [`asyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `AsyncResult` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="cc78d-854">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="cc78d-855">Requirements</span><span class="sxs-lookup"><span data-stu-id="cc78d-855">Requirements</span></span>

|<span data-ttu-id="cc78d-856">要求</span><span class="sxs-lookup"><span data-stu-id="cc78d-856">Requirement</span></span>| <span data-ttu-id="cc78d-857">值</span><span class="sxs-lookup"><span data-stu-id="cc78d-857">Value</span></span>|
|---|---|
|[<span data-ttu-id="cc78d-858">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="cc78d-858">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="cc78d-859">1.7</span><span class="sxs-lookup"><span data-stu-id="cc78d-859">1.7</span></span> |
|[<span data-ttu-id="cc78d-860">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="cc78d-860">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="cc78d-861">ReadItem</span><span class="sxs-lookup"><span data-stu-id="cc78d-861">ReadItem</span></span> |
|[<span data-ttu-id="cc78d-862">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="cc78d-862">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="cc78d-863">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="cc78d-863">Compose or Read</span></span> |

##### <a name="example"></a><span data-ttu-id="cc78d-864">示例</span><span class="sxs-lookup"><span data-stu-id="cc78d-864">Example</span></span>

```javascript
function myHandlerFunction(eventarg) {
  if (eventarg.attachmentStatus === Office.MailboxEnums.AttachmentStatus.Added) {
    var attachment = eventarg.attachmentDetails;
    console.log("Event Fired and Attachment Added!");
    getAttachmentContentAsync(attachment.id, options, callback);
  }
}

Office.context.mailbox.item.addHandlerAsync(Office.EventType.AttachmentsChanged, myHandlerFunction, myCallback);
```

---
---

####  <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="cc78d-865">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="cc78d-865">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="cc78d-866">将 Exchange 项目（如邮件）作为附件添加到邮件或约会。</span><span class="sxs-lookup"><span data-stu-id="cc78d-866">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="cc78d-p141">`addItemAttachmentAsync` 方法将包含指定 Exchange 标识符的项目附加到撰写窗体中的项目。如果指定一个回调方法，此方法使用 `asyncResult` 参数调用，该参数包含一个附件标识符或代码，指示附加项目过程中出现的任何错误。可以使用 `options` 参数将状态信息传递给回调方法（如果需要）。</span><span class="sxs-lookup"><span data-stu-id="cc78d-p141">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="cc78d-870">随后可以将该标识符与 [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) 方法一同使用，以删除同一个会话中的附件。</span><span class="sxs-lookup"><span data-stu-id="cc78d-870">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="cc78d-871">如果 Office 加载项在 Outlook Web App 中运行，则 `addItemAttachmentAsync` 方法可以将项目附加到项目（正在编辑的项目除外）中；然而，不支持也不建议这样做。</span><span class="sxs-lookup"><span data-stu-id="cc78d-871">If your Office Add-in is running in Outlook Web App, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="cc78d-872">参数</span><span class="sxs-lookup"><span data-stu-id="cc78d-872">Parameters</span></span>

|<span data-ttu-id="cc78d-873">名称</span><span class="sxs-lookup"><span data-stu-id="cc78d-873">Name</span></span>|<span data-ttu-id="cc78d-874">类型</span><span class="sxs-lookup"><span data-stu-id="cc78d-874">Type</span></span>|<span data-ttu-id="cc78d-875">属性</span><span class="sxs-lookup"><span data-stu-id="cc78d-875">Attributes</span></span>|<span data-ttu-id="cc78d-876">描述</span><span class="sxs-lookup"><span data-stu-id="cc78d-876">Description</span></span>|
|---|---|---|---|
|`itemId`|<span data-ttu-id="cc78d-877">字符串</span><span class="sxs-lookup"><span data-stu-id="cc78d-877">String</span></span>||<span data-ttu-id="cc78d-p142">要附加的项目的 Exchange 标识符。最大长度为 100 个字符。</span><span class="sxs-lookup"><span data-stu-id="cc78d-p142">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`|<span data-ttu-id="cc78d-880">String</span><span class="sxs-lookup"><span data-stu-id="cc78d-880">String</span></span>||<span data-ttu-id="cc78d-881">要附加的项目的主题。</span><span class="sxs-lookup"><span data-stu-id="cc78d-881">The subject of the item to be attached.</span></span> <span data-ttu-id="cc78d-882">最大长度为 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="cc78d-882">The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="cc78d-883">对象</span><span class="sxs-lookup"><span data-stu-id="cc78d-883">Object</span></span>|<span data-ttu-id="cc78d-884">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="cc78d-884">&lt;optional&gt;</span></span>|<span data-ttu-id="cc78d-885">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="cc78d-885">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="cc78d-886">Object</span><span class="sxs-lookup"><span data-stu-id="cc78d-886">Object</span></span>|<span data-ttu-id="cc78d-887">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="cc78d-887">&lt;optional&gt;</span></span>|<span data-ttu-id="cc78d-888">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="cc78d-888">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="cc78d-889">函数</span><span class="sxs-lookup"><span data-stu-id="cc78d-889">function</span></span>|<span data-ttu-id="cc78d-890">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="cc78d-890">&lt;optional&gt;</span></span>|<span data-ttu-id="cc78d-891">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="cc78d-891">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="cc78d-892">如果成功，附件标识符将在 `asyncResult.value` 属性中提供。</span><span class="sxs-lookup"><span data-stu-id="cc78d-892">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="cc78d-893">如果添加附件失败，`asyncResult` 对象将包含一个提供错误说明的 `Error` 对象。</span><span class="sxs-lookup"><span data-stu-id="cc78d-893">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="cc78d-894">错误</span><span class="sxs-lookup"><span data-stu-id="cc78d-894">Errors</span></span>

|<span data-ttu-id="cc78d-895">错误代码</span><span class="sxs-lookup"><span data-stu-id="cc78d-895">Error code</span></span>|<span data-ttu-id="cc78d-896">说明</span><span class="sxs-lookup"><span data-stu-id="cc78d-896">Description</span></span>|
|------------|-------------|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="cc78d-897">邮件或约会具有的附件过多。</span><span class="sxs-lookup"><span data-stu-id="cc78d-897">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="cc78d-898">要求</span><span class="sxs-lookup"><span data-stu-id="cc78d-898">Requirements</span></span>

|<span data-ttu-id="cc78d-899">要求</span><span class="sxs-lookup"><span data-stu-id="cc78d-899">Requirement</span></span>|<span data-ttu-id="cc78d-900">值</span><span class="sxs-lookup"><span data-stu-id="cc78d-900">Value</span></span>|
|---|---|
|[<span data-ttu-id="cc78d-901">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="cc78d-901">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="cc78d-902">1.1</span><span class="sxs-lookup"><span data-stu-id="cc78d-902">1.1</span></span>|
|[<span data-ttu-id="cc78d-903">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="cc78d-903">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="cc78d-904">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="cc78d-904">ReadWriteItem</span></span>|
|[<span data-ttu-id="cc78d-905">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="cc78d-905">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="cc78d-906">撰写</span><span class="sxs-lookup"><span data-stu-id="cc78d-906">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="cc78d-907">示例</span><span class="sxs-lookup"><span data-stu-id="cc78d-907">Example</span></span>

<span data-ttu-id="cc78d-908">以下示例将现有的 Outlook 项目添加为名为 `My Attachment` 的附件。</span><span class="sxs-lookup"><span data-stu-id="cc78d-908">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

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

---
---

####  <a name="close"></a><span data-ttu-id="cc78d-909">close()</span><span class="sxs-lookup"><span data-stu-id="cc78d-909">close()</span></span>

<span data-ttu-id="cc78d-910">关闭当前正在撰写的项目。</span><span class="sxs-lookup"><span data-stu-id="cc78d-910">Closes the current item that is being composed.</span></span>

<span data-ttu-id="cc78d-p144">
            \`close\` 方法的行为取决于要撰写的项目的当前状态。如果项目具有未保存的更改，客户端将提示用户保存、放弃或取消关闭操作。</span><span class="sxs-lookup"><span data-stu-id="cc78d-p144">The behavior of the `close` method depends on the current state of the item being composed. If the item has unsaved changes, the client prompts the user to save, discard, or cancel the close action.</span></span>

> [!NOTE]
> <span data-ttu-id="cc78d-913">在 Outlook 网页版中，如果该项目是约会并且之前已使用 `saveAsync` 保存，则即使自上次保存项目后未发生任何更改，也会提示用户保存、放弃或取消。</span><span class="sxs-lookup"><span data-stu-id="cc78d-913">In Outlook on the web, if the item is an appointment and it has previously been saved using `saveAsync`, the user is prompted to save, discard, or cancel even if no changes have occurred since the item was last saved.</span></span>

<span data-ttu-id="cc78d-914">在 Outlook 桌面客户端中，如果邮件是内联答复，`close` 方法不起作用。</span><span class="sxs-lookup"><span data-stu-id="cc78d-914">In the Outlook desktop client, if the message is an inline reply, the `close` method has no effect.</span></span>

##### <a name="requirements"></a><span data-ttu-id="cc78d-915">Requirements</span><span class="sxs-lookup"><span data-stu-id="cc78d-915">Requirements</span></span>

|<span data-ttu-id="cc78d-916">要求</span><span class="sxs-lookup"><span data-stu-id="cc78d-916">Requirement</span></span>|<span data-ttu-id="cc78d-917">值</span><span class="sxs-lookup"><span data-stu-id="cc78d-917">Value</span></span>|
|---|---|
|[<span data-ttu-id="cc78d-918">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="cc78d-918">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="cc78d-919">1.3</span><span class="sxs-lookup"><span data-stu-id="cc78d-919">1.3</span></span>|
|[<span data-ttu-id="cc78d-920">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="cc78d-920">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="cc78d-921">受限</span><span class="sxs-lookup"><span data-stu-id="cc78d-921">Restricted</span></span>|
|[<span data-ttu-id="cc78d-922">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="cc78d-922">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="cc78d-923">撰写</span><span class="sxs-lookup"><span data-stu-id="cc78d-923">Compose</span></span>|

---
---

#### <a name="displayreplyallformformdata-callback"></a><span data-ttu-id="cc78d-924">displayReplyAllForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="cc78d-924">displayReplyAllForm(formData, [callback])</span></span>

<span data-ttu-id="cc78d-925">显示答复窗体，其中包括所选邮件的发件人和所有收件人或所选约会的组织者和所有与会者。</span><span class="sxs-lookup"><span data-stu-id="cc78d-925">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="cc78d-926">在 Outlook for iOS 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="cc78d-926">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="cc78d-927">在 Outlook Web App 中，答复窗体显示为包含 3 列视图的弹出式窗体以及包含 2 列或 1 列视图的弹出式窗体。</span><span class="sxs-lookup"><span data-stu-id="cc78d-927">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="cc78d-928">如果任意字符串参数超出其限制，`displayReplyAllForm` 将引发异常。</span><span class="sxs-lookup"><span data-stu-id="cc78d-928">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="cc78d-p145">当在 `formData.attachments` 参数中指定附件时，Outlook 和 Outlook Web App 尝试下载所有附件并将其附加到答复窗体。如果无法添加任何附件，则在窗体 UI 中显示错误。如果这不可能，则不引发错误消息。</span><span class="sxs-lookup"><span data-stu-id="cc78d-p145">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="cc78d-932">参数</span><span class="sxs-lookup"><span data-stu-id="cc78d-932">Parameters</span></span>

|<span data-ttu-id="cc78d-933">名称</span><span class="sxs-lookup"><span data-stu-id="cc78d-933">Name</span></span>|<span data-ttu-id="cc78d-934">类型</span><span class="sxs-lookup"><span data-stu-id="cc78d-934">Type</span></span>|<span data-ttu-id="cc78d-935">属性</span><span class="sxs-lookup"><span data-stu-id="cc78d-935">Attributes</span></span>|<span data-ttu-id="cc78d-936">说明</span><span class="sxs-lookup"><span data-stu-id="cc78d-936">Description</span></span>|
|---|---|---|---|
|`formData`|<span data-ttu-id="cc78d-937">字符串 &#124; 对象</span><span class="sxs-lookup"><span data-stu-id="cc78d-937">String &#124; Object</span></span>||<span data-ttu-id="cc78d-p146">一个包含文本和 HTML 且表示答复窗体的正文的字符串。字符串限制为 32 KB。</span><span class="sxs-lookup"><span data-stu-id="cc78d-p146">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="cc78d-940">**或**</span><span class="sxs-lookup"><span data-stu-id="cc78d-940">**OR**</span></span><br/><span data-ttu-id="cc78d-p147">包含正文或附件数据和回调函数的对象。对象定义如下。</span><span class="sxs-lookup"><span data-stu-id="cc78d-p147">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span>|
|`formData.htmlBody`|<span data-ttu-id="cc78d-943">String</span><span class="sxs-lookup"><span data-stu-id="cc78d-943">String</span></span>|<span data-ttu-id="cc78d-944">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="cc78d-944">&lt;optional&gt;</span></span>|<span data-ttu-id="cc78d-p148">一个包含文本和 HTML 且表示答复窗体的正文的字符串。字符串限制为 32 KB。</span><span class="sxs-lookup"><span data-stu-id="cc78d-p148">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
|`formData.attachments`|<span data-ttu-id="cc78d-947">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="cc78d-947">Array.&lt;Object&gt;</span></span>|<span data-ttu-id="cc78d-948">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="cc78d-948">&lt;optional&gt;</span></span>|<span data-ttu-id="cc78d-949">JSON 对象（是文件或项目附件）的数组。</span><span class="sxs-lookup"><span data-stu-id="cc78d-949">An array of JSON objects that are either file or item attachments.</span></span>|
|`formData.attachments.type`|<span data-ttu-id="cc78d-950">String</span><span class="sxs-lookup"><span data-stu-id="cc78d-950">String</span></span>||<span data-ttu-id="cc78d-p149">指示附件的类型。必须是文件附件的 `file` 或项目附件的 `item`。</span><span class="sxs-lookup"><span data-stu-id="cc78d-p149">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span>|
|`formData.attachments.name`|<span data-ttu-id="cc78d-953">字符串</span><span class="sxs-lookup"><span data-stu-id="cc78d-953">String</span></span>||<span data-ttu-id="cc78d-954">一个包含附件的名称的字符串，最多包含 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="cc78d-954">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
|`formData.attachments.url`|<span data-ttu-id="cc78d-955">字符串</span><span class="sxs-lookup"><span data-stu-id="cc78d-955">String</span></span>||<span data-ttu-id="cc78d-p150">仅在将 `type` 设置为 `file` 时使用。文件的位置的 URI。</span><span class="sxs-lookup"><span data-stu-id="cc78d-p150">Only used if `type` is set to `file`. The URI of the location for the file.</span></span>|
|`formData.attachments.isInline`|<span data-ttu-id="cc78d-958">布尔</span><span class="sxs-lookup"><span data-stu-id="cc78d-958">Boolean</span></span>||<span data-ttu-id="cc78d-p151">仅在将 `type` 设置为 `file` 时使用。如果为 `true`，则表示附件将在邮件正文中内联显示，并且不应显示在附件列表中。</span><span class="sxs-lookup"><span data-stu-id="cc78d-p151">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`formData.attachments.itemId`|<span data-ttu-id="cc78d-961">String</span><span class="sxs-lookup"><span data-stu-id="cc78d-961">String</span></span>||<span data-ttu-id="cc78d-p152">仅在将 `type` 设置为 `item` 时使用。附件的 EWS 项目 ID。字符串最长为 100 个字符。</span><span class="sxs-lookup"><span data-stu-id="cc78d-p152">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span>|
|`callback`|<span data-ttu-id="cc78d-965">函数</span><span class="sxs-lookup"><span data-stu-id="cc78d-965">function</span></span>|<span data-ttu-id="cc78d-966">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="cc78d-966">&lt;optional&gt;</span></span>|<span data-ttu-id="cc78d-967">方法完成后，使用单个参数 `asyncResult`（一个 [AsyncResult](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="cc78d-967">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="cc78d-968">要求</span><span class="sxs-lookup"><span data-stu-id="cc78d-968">Requirements</span></span>

|<span data-ttu-id="cc78d-969">要求</span><span class="sxs-lookup"><span data-stu-id="cc78d-969">Requirement</span></span>|<span data-ttu-id="cc78d-970">值</span><span class="sxs-lookup"><span data-stu-id="cc78d-970">Value</span></span>|
|---|---|
|[<span data-ttu-id="cc78d-971">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="cc78d-971">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="cc78d-972">1.0</span><span class="sxs-lookup"><span data-stu-id="cc78d-972">1.0</span></span>|
|[<span data-ttu-id="cc78d-973">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="cc78d-973">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="cc78d-974">ReadItem</span><span class="sxs-lookup"><span data-stu-id="cc78d-974">ReadItem</span></span>|
|[<span data-ttu-id="cc78d-975">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="cc78d-975">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="cc78d-976">阅读</span><span class="sxs-lookup"><span data-stu-id="cc78d-976">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="cc78d-977">示例</span><span class="sxs-lookup"><span data-stu-id="cc78d-977">Examples</span></span>

<span data-ttu-id="cc78d-978">以下代码将一个字符串传递到 `displayReplyAllForm` 函数。</span><span class="sxs-lookup"><span data-stu-id="cc78d-978">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="cc78d-979">使用空白正文答复。</span><span class="sxs-lookup"><span data-stu-id="cc78d-979">Reply with an empty body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="cc78d-980">仅使用正文答复。</span><span class="sxs-lookup"><span data-stu-id="cc78d-980">Reply with just a body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="cc78d-981">使用正文和文件附件答复。</span><span class="sxs-lookup"><span data-stu-id="cc78d-981">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="cc78d-982">使用正文和项目附件答复。</span><span class="sxs-lookup"><span data-stu-id="cc78d-982">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="cc78d-983">使用正文、文件附件、项目附件和回调答复。</span><span class="sxs-lookup"><span data-stu-id="cc78d-983">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

---
---

#### <a name="displayreplyformformdata-callback"></a><span data-ttu-id="cc78d-984">displayReplyForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="cc78d-984">displayReplyForm(formData, [callback])</span></span>

<span data-ttu-id="cc78d-985">显示答复窗体，其中仅包括所选邮件的发件人或所选约会的组织者。</span><span class="sxs-lookup"><span data-stu-id="cc78d-985">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="cc78d-986">在 Outlook for iOS 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="cc78d-986">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="cc78d-987">在 Outlook Web App 中，答复窗体显示为包含 3 列视图的弹出式窗体以及包含 2 列或 1 列视图的弹出式窗体。</span><span class="sxs-lookup"><span data-stu-id="cc78d-987">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="cc78d-988">如果任意字符串参数超出其限制，`displayReplyForm` 将引发异常。</span><span class="sxs-lookup"><span data-stu-id="cc78d-988">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="cc78d-p153">当在 `formData.attachments` 参数中指定附件时，Outlook 和 Outlook Web App 尝试下载所有附件并将其附加到答复窗体。如果无法添加任何附件，则在窗体 UI 中显示错误。如果这不可能，则不引发错误消息。</span><span class="sxs-lookup"><span data-stu-id="cc78d-p153">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="cc78d-992">参数</span><span class="sxs-lookup"><span data-stu-id="cc78d-992">Parameters</span></span>

|<span data-ttu-id="cc78d-993">名称</span><span class="sxs-lookup"><span data-stu-id="cc78d-993">Name</span></span>|<span data-ttu-id="cc78d-994">类型</span><span class="sxs-lookup"><span data-stu-id="cc78d-994">Type</span></span>|<span data-ttu-id="cc78d-995">属性</span><span class="sxs-lookup"><span data-stu-id="cc78d-995">Attributes</span></span>|<span data-ttu-id="cc78d-996">说明</span><span class="sxs-lookup"><span data-stu-id="cc78d-996">Description</span></span>|
|---|---|---|---|
|`formData`|<span data-ttu-id="cc78d-997">字符串 &#124; 对象</span><span class="sxs-lookup"><span data-stu-id="cc78d-997">String &#124; Object</span></span>||<span data-ttu-id="cc78d-p154">一个包含文本和 HTML 且表示答复窗体的正文的字符串。字符串限制为 32 KB。</span><span class="sxs-lookup"><span data-stu-id="cc78d-p154">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="cc78d-1000">**或**</span><span class="sxs-lookup"><span data-stu-id="cc78d-1000">**OR**</span></span><br/><span data-ttu-id="cc78d-p155">包含正文或附件数据和回调函数的对象。对象定义如下。</span><span class="sxs-lookup"><span data-stu-id="cc78d-p155">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span>|
|`formData.htmlBody`|<span data-ttu-id="cc78d-1003">String</span><span class="sxs-lookup"><span data-stu-id="cc78d-1003">String</span></span>|<span data-ttu-id="cc78d-1004">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="cc78d-1004">&lt;optional&gt;</span></span>|<span data-ttu-id="cc78d-p156">一个包含文本和 HTML 且表示答复窗体的正文的字符串。字符串限制为 32 KB。</span><span class="sxs-lookup"><span data-stu-id="cc78d-p156">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
|`formData.attachments`|<span data-ttu-id="cc78d-1007">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="cc78d-1007">Array.&lt;Object&gt;</span></span>|<span data-ttu-id="cc78d-1008">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="cc78d-1008">&lt;optional&gt;</span></span>|<span data-ttu-id="cc78d-1009">JSON 对象（是文件或项目附件）的数组。</span><span class="sxs-lookup"><span data-stu-id="cc78d-1009">An array of JSON objects that are either file or item attachments.</span></span>|
|`formData.attachments.type`|<span data-ttu-id="cc78d-1010">String</span><span class="sxs-lookup"><span data-stu-id="cc78d-1010">String</span></span>||<span data-ttu-id="cc78d-p157">指示附件的类型。必须是文件附件的 `file` 或项目附件的 `item`。</span><span class="sxs-lookup"><span data-stu-id="cc78d-p157">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span>|
|`formData.attachments.name`|<span data-ttu-id="cc78d-1013">字符串</span><span class="sxs-lookup"><span data-stu-id="cc78d-1013">String</span></span>||<span data-ttu-id="cc78d-1014">一个包含附件的名称的字符串，最多包含 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="cc78d-1014">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
|`formData.attachments.url`|<span data-ttu-id="cc78d-1015">字符串</span><span class="sxs-lookup"><span data-stu-id="cc78d-1015">String</span></span>||<span data-ttu-id="cc78d-p158">仅在将 `type` 设置为 `file` 时使用。文件的位置的 URI。</span><span class="sxs-lookup"><span data-stu-id="cc78d-p158">Only used if `type` is set to `file`. The URI of the location for the file.</span></span>|
|`formData.attachments.isInline`|<span data-ttu-id="cc78d-1018">布尔</span><span class="sxs-lookup"><span data-stu-id="cc78d-1018">Boolean</span></span>||<span data-ttu-id="cc78d-p159">仅在将 `type` 设置为 `file` 时使用。如果为 `true`，则表示附件将在邮件正文中内联显示，并且不应显示在附件列表中。</span><span class="sxs-lookup"><span data-stu-id="cc78d-p159">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`formData.attachments.itemId`|<span data-ttu-id="cc78d-1021">String</span><span class="sxs-lookup"><span data-stu-id="cc78d-1021">String</span></span>||<span data-ttu-id="cc78d-p160">仅在将 `type` 设置为 `item` 时使用。附件的 EWS 项目 ID。字符串最长为 100 个字符。</span><span class="sxs-lookup"><span data-stu-id="cc78d-p160">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span>|
|`callback`|<span data-ttu-id="cc78d-1025">函数</span><span class="sxs-lookup"><span data-stu-id="cc78d-1025">function</span></span>|<span data-ttu-id="cc78d-1026">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="cc78d-1026">&lt;optional&gt;</span></span>|<span data-ttu-id="cc78d-1027">方法完成后，使用单个参数 `asyncResult`（一个 [AsyncResult](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="cc78d-1027">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="cc78d-1028">要求</span><span class="sxs-lookup"><span data-stu-id="cc78d-1028">Requirements</span></span>

|<span data-ttu-id="cc78d-1029">要求</span><span class="sxs-lookup"><span data-stu-id="cc78d-1029">Requirement</span></span>|<span data-ttu-id="cc78d-1030">值</span><span class="sxs-lookup"><span data-stu-id="cc78d-1030">Value</span></span>|
|---|---|
|[<span data-ttu-id="cc78d-1031">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="cc78d-1031">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="cc78d-1032">1.0</span><span class="sxs-lookup"><span data-stu-id="cc78d-1032">1.0</span></span>|
|[<span data-ttu-id="cc78d-1033">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="cc78d-1033">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="cc78d-1034">ReadItem</span><span class="sxs-lookup"><span data-stu-id="cc78d-1034">ReadItem</span></span>|
|[<span data-ttu-id="cc78d-1035">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="cc78d-1035">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="cc78d-1036">阅读</span><span class="sxs-lookup"><span data-stu-id="cc78d-1036">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="cc78d-1037">示例</span><span class="sxs-lookup"><span data-stu-id="cc78d-1037">Examples</span></span>

<span data-ttu-id="cc78d-1038">以下代码将一个字符串传递到 `displayReplyForm` 函数。</span><span class="sxs-lookup"><span data-stu-id="cc78d-1038">The following code passes a string to the `displayReplyForm` function.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="cc78d-1039">使用空白正文答复。</span><span class="sxs-lookup"><span data-stu-id="cc78d-1039">Reply with an empty body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="cc78d-1040">仅使用正文答复。</span><span class="sxs-lookup"><span data-stu-id="cc78d-1040">Reply with just a body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="cc78d-1041">使用正文和文件附件答复。</span><span class="sxs-lookup"><span data-stu-id="cc78d-1041">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="cc78d-1042">使用正文和项目附件答复。</span><span class="sxs-lookup"><span data-stu-id="cc78d-1042">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="cc78d-1043">使用正文、文件附件、项目附件和回调答复。</span><span class="sxs-lookup"><span data-stu-id="cc78d-1043">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

---
---

#### <a name="getattachmentcontentasyncattachmentid-options-callback--attachmentcontentjavascriptapioutlookofficeattachmentcontent"></a><span data-ttu-id="cc78d-1044">getAttachmentContentAsync (attachmentId, [options], [callback]) → [AttachmentContent](/javascript/api/outlook/office.attachmentcontent)</span><span class="sxs-lookup"><span data-stu-id="cc78d-1044">getAttachmentContentAsync(attachmentId, [options], [callback]) → [AttachmentContent](/javascript/api/outlook/office.attachmentcontent)</span></span>

<span data-ttu-id="cc78d-1045">从邮件或约会中获取指定附件并将其作为`AttachmentContent`对象返回。</span><span class="sxs-lookup"><span data-stu-id="cc78d-1045">Gets the specified attachment from a message or appointment and returns it as an `AttachmentContent` object.</span></span>

<span data-ttu-id="cc78d-1046">该`getAttachmentContentAsync`方法从项目中获取具有指定标识符的附件。</span><span class="sxs-lookup"><span data-stu-id="cc78d-1046">The `getAttachmentContentAsync` method gets the attachment with the specified identifier from the item.</span></span> <span data-ttu-id="cc78d-1047">作为一种最佳做法, 您应使用标识符在与`getAttachmentsAsync` or `item.attachments`调用一起检索到会话的同一会话中检索附件。</span><span class="sxs-lookup"><span data-stu-id="cc78d-1047">As a best practice, you should use the identifier to retrieve an attachment in the same session that the attachmentIds were retrieved with the `getAttachmentsAsync` or `item.attachments` call.</span></span> <span data-ttu-id="cc78d-1048">在 Outlook Web App 和适用于设备的 OWA 中，附件标识符只在同一个会话中才有效。</span><span class="sxs-lookup"><span data-stu-id="cc78d-1048">In Outlook Web App and OWA for Devices, the attachment identifier is valid only within the same session.</span></span> <span data-ttu-id="cc78d-1049">当用户关闭应用程序时, 或者如果用户开始撰写内嵌窗体, 随后弹出窗体以继续在单独的窗口中。</span><span class="sxs-lookup"><span data-stu-id="cc78d-1049">A session is over when the user closes the app, or if the user starts composing an inline form then subsequently pops out the form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="cc78d-1050">参数</span><span class="sxs-lookup"><span data-stu-id="cc78d-1050">Parameters</span></span>

|<span data-ttu-id="cc78d-1051">名称</span><span class="sxs-lookup"><span data-stu-id="cc78d-1051">Name</span></span>|<span data-ttu-id="cc78d-1052">类型</span><span class="sxs-lookup"><span data-stu-id="cc78d-1052">Type</span></span>|<span data-ttu-id="cc78d-1053">属性</span><span class="sxs-lookup"><span data-stu-id="cc78d-1053">Attributes</span></span>|<span data-ttu-id="cc78d-1054">描述</span><span class="sxs-lookup"><span data-stu-id="cc78d-1054">Description</span></span>|
|---|---|---|---|
|`attachmentId`|<span data-ttu-id="cc78d-1055">字符串</span><span class="sxs-lookup"><span data-stu-id="cc78d-1055">String</span></span>||<span data-ttu-id="cc78d-1056">要获取的附件的标识符。</span><span class="sxs-lookup"><span data-stu-id="cc78d-1056">The identifier of the attachment you want to get.</span></span>|
|`options`|<span data-ttu-id="cc78d-1057">Object</span><span class="sxs-lookup"><span data-stu-id="cc78d-1057">Object</span></span>|<span data-ttu-id="cc78d-1058">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="cc78d-1058">&lt;optional&gt;</span></span>|<span data-ttu-id="cc78d-1059">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="cc78d-1059">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="cc78d-1060">Object</span><span class="sxs-lookup"><span data-stu-id="cc78d-1060">Object</span></span>|<span data-ttu-id="cc78d-1061">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="cc78d-1061">&lt;optional&gt;</span></span>|<span data-ttu-id="cc78d-1062">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="cc78d-1062">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="cc78d-1063">函数</span><span class="sxs-lookup"><span data-stu-id="cc78d-1063">function</span></span>|<span data-ttu-id="cc78d-1064">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="cc78d-1064">&lt;optional&gt;</span></span>|<span data-ttu-id="cc78d-1065">方法完成后，使用单个参数 `asyncResult`（一个 [AsyncResult](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="cc78d-1065">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="cc78d-1066">Requirements</span><span class="sxs-lookup"><span data-stu-id="cc78d-1066">Requirements</span></span>

|<span data-ttu-id="cc78d-1067">要求</span><span class="sxs-lookup"><span data-stu-id="cc78d-1067">Requirement</span></span>|<span data-ttu-id="cc78d-1068">值</span><span class="sxs-lookup"><span data-stu-id="cc78d-1068">Value</span></span>|
|---|---|
|[<span data-ttu-id="cc78d-1069">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="cc78d-1069">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="cc78d-1070">预览</span><span class="sxs-lookup"><span data-stu-id="cc78d-1070">Preview</span></span>|
|[<span data-ttu-id="cc78d-1071">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="cc78d-1071">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="cc78d-1072">ReadItem</span><span class="sxs-lookup"><span data-stu-id="cc78d-1072">ReadItem</span></span>|
|[<span data-ttu-id="cc78d-1073">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="cc78d-1073">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="cc78d-1074">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="cc78d-1074">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="cc78d-1075">返回：</span><span class="sxs-lookup"><span data-stu-id="cc78d-1075">Returns:</span></span>

<span data-ttu-id="cc78d-1076">类型: [AttachmentContent](/javascript/api/outlook/office.attachmentcontent)</span><span class="sxs-lookup"><span data-stu-id="cc78d-1076">Type: [AttachmentContent](/javascript/api/outlook/office.attachmentcontent)</span></span>

##### <a name="example"></a><span data-ttu-id="cc78d-1077">示例</span><span class="sxs-lookup"><span data-stu-id="cc78d-1077">Example</span></span>

```javascript
var item = Office.context.mailbox.item;
var listOfAttachments = [];
item.getAttachmentsAsync(callback);

function callback(result) {
  if (result.value.length > 0) {
    for (i = 0 ; i < result.value.length ; i++) {
      var options = {asyncContext: {type: result.value[i].attachmentType}};
      getAttachmentContentAsync(result.value[i].id, options, handleAttachmentsCallback);
    }
  }
}

function handleAttachmentsCallback(result) {
  // Parse string to be a url, an .eml file, a base64-encoded string, or an .icalendar file.
  if (result.format === Office.MailboxEnums.AttachmentContentFormat.Base64) {
    // Handle file attachment.
  } else if (result.format === Office.MailboxEnums.AttachmentContentFormat.Eml) {
    // Handle email item attachment.
  } else if (result.format === Office.MailboxEnums.AttachmentContentFormat.ICalendar) {
    // Handle .icalender attachment.
  } else if (result.format === Office.MailboxEnums.AttachmentContentFormat.Url) {
    // Handle cloud attachment.
  } else {
    // Handle attachment formats that are not supported.
  }
}
```

---
---

#### <a name="getattachmentsasyncoptions-callback--arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetails"></a><span data-ttu-id="cc78d-1078">getAttachmentsAsync ([options], [callback]) → Array. <[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="cc78d-1078">getAttachmentsAsync([options], [callback]) → Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span></span>

<span data-ttu-id="cc78d-1079">以数组的形式获取项目的附件。</span><span class="sxs-lookup"><span data-stu-id="cc78d-1079">Gets the item's attachments as an array.</span></span> <span data-ttu-id="cc78d-1080">仅限撰写模式。</span><span class="sxs-lookup"><span data-stu-id="cc78d-1080">Compose mode only.</span></span>

##### <a name="parameters"></a><span data-ttu-id="cc78d-1081">参数</span><span class="sxs-lookup"><span data-stu-id="cc78d-1081">Parameters</span></span>

|<span data-ttu-id="cc78d-1082">名称</span><span class="sxs-lookup"><span data-stu-id="cc78d-1082">Name</span></span>|<span data-ttu-id="cc78d-1083">类型</span><span class="sxs-lookup"><span data-stu-id="cc78d-1083">Type</span></span>|<span data-ttu-id="cc78d-1084">属性</span><span class="sxs-lookup"><span data-stu-id="cc78d-1084">Attributes</span></span>|<span data-ttu-id="cc78d-1085">说明</span><span class="sxs-lookup"><span data-stu-id="cc78d-1085">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="cc78d-1086">Object</span><span class="sxs-lookup"><span data-stu-id="cc78d-1086">Object</span></span>|<span data-ttu-id="cc78d-1087">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="cc78d-1087">&lt;optional&gt;</span></span>|<span data-ttu-id="cc78d-1088">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="cc78d-1088">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="cc78d-1089">Object</span><span class="sxs-lookup"><span data-stu-id="cc78d-1089">Object</span></span>|<span data-ttu-id="cc78d-1090">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="cc78d-1090">&lt;optional&gt;</span></span>|<span data-ttu-id="cc78d-1091">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="cc78d-1091">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="cc78d-1092">函数</span><span class="sxs-lookup"><span data-stu-id="cc78d-1092">function</span></span>|<span data-ttu-id="cc78d-1093">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="cc78d-1093">&lt;optional&gt;</span></span>|<span data-ttu-id="cc78d-1094">方法完成后，使用单个参数 `asyncResult`（一个 [AsyncResult](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="cc78d-1094">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="cc78d-1095">Requirements</span><span class="sxs-lookup"><span data-stu-id="cc78d-1095">Requirements</span></span>

|<span data-ttu-id="cc78d-1096">要求</span><span class="sxs-lookup"><span data-stu-id="cc78d-1096">Requirement</span></span>|<span data-ttu-id="cc78d-1097">值</span><span class="sxs-lookup"><span data-stu-id="cc78d-1097">Value</span></span>|
|---|---|
|[<span data-ttu-id="cc78d-1098">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="cc78d-1098">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="cc78d-1099">预览</span><span class="sxs-lookup"><span data-stu-id="cc78d-1099">Preview</span></span>|
|[<span data-ttu-id="cc78d-1100">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="cc78d-1100">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="cc78d-1101">ReadItem</span><span class="sxs-lookup"><span data-stu-id="cc78d-1101">ReadItem</span></span>|
|[<span data-ttu-id="cc78d-1102">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="cc78d-1102">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="cc78d-1103">撰写</span><span class="sxs-lookup"><span data-stu-id="cc78d-1103">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="cc78d-1104">返回：</span><span class="sxs-lookup"><span data-stu-id="cc78d-1104">Returns:</span></span>

<span data-ttu-id="cc78d-1105">类型: <[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="cc78d-1105">Type: Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span></span>

##### <a name="example"></a><span data-ttu-id="cc78d-1106">示例</span><span class="sxs-lookup"><span data-stu-id="cc78d-1106">Example</span></span>

<span data-ttu-id="cc78d-1107">下面的示例将生成一个 HTML 字符串, 其中包含当前项目上所有附件的详细信息。</span><span class="sxs-lookup"><span data-stu-id="cc78d-1107">The following example builds an HTML string with details of all attachments on the current item.</span></span>

```javascript
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

---
---

#### <a name="getentities--entitiesjavascriptapioutlookofficeentities"></a><span data-ttu-id="cc78d-1108">getEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="cc78d-1108">getEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span></span>

<span data-ttu-id="cc78d-1109">获取在所选项目的正文中找到的实体。</span><span class="sxs-lookup"><span data-stu-id="cc78d-1109">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="cc78d-1110">在 Outlook for iOS 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="cc78d-1110">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="cc78d-1111">Requirements</span><span class="sxs-lookup"><span data-stu-id="cc78d-1111">Requirements</span></span>

|<span data-ttu-id="cc78d-1112">要求</span><span class="sxs-lookup"><span data-stu-id="cc78d-1112">Requirement</span></span>|<span data-ttu-id="cc78d-1113">值</span><span class="sxs-lookup"><span data-stu-id="cc78d-1113">Value</span></span>|
|---|---|
|[<span data-ttu-id="cc78d-1114">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="cc78d-1114">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="cc78d-1115">1.0</span><span class="sxs-lookup"><span data-stu-id="cc78d-1115">1.0</span></span>|
|[<span data-ttu-id="cc78d-1116">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="cc78d-1116">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="cc78d-1117">ReadItem</span><span class="sxs-lookup"><span data-stu-id="cc78d-1117">ReadItem</span></span>|
|[<span data-ttu-id="cc78d-1118">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="cc78d-1118">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="cc78d-1119">阅读</span><span class="sxs-lookup"><span data-stu-id="cc78d-1119">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="cc78d-1120">返回：</span><span class="sxs-lookup"><span data-stu-id="cc78d-1120">Returns:</span></span>

<span data-ttu-id="cc78d-1121">类型：[Entities](/javascript/api/outlook/office.entities)</span><span class="sxs-lookup"><span data-stu-id="cc78d-1121">Type: [Entities](/javascript/api/outlook/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="cc78d-1122">示例</span><span class="sxs-lookup"><span data-stu-id="cc78d-1122">Example</span></span>

<span data-ttu-id="cc78d-1123">以下示例访问当前项目的正文中的联系人实体。</span><span class="sxs-lookup"><span data-stu-id="cc78d-1123">The following example accesses the contacts entities in the current item's body.</span></span>

```javascript
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

---
---

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlookofficecontactmeetingsuggestionjavascriptapioutlookofficemeetingsuggestionphonenumberjavascriptapioutlookofficephonenumbertasksuggestionjavascriptapioutlookofficetasksuggestion"></a><span data-ttu-id="cc78d-1124">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="cc78d-1124">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span></span>

<span data-ttu-id="cc78d-1125">获取所选项目的正文中找到的指定实体类型的所有实体的数组。</span><span class="sxs-lookup"><span data-stu-id="cc78d-1125">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="cc78d-1126">在 Outlook for iOS 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="cc78d-1126">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="cc78d-1127">参数</span><span class="sxs-lookup"><span data-stu-id="cc78d-1127">Parameters</span></span>

|<span data-ttu-id="cc78d-1128">名称</span><span class="sxs-lookup"><span data-stu-id="cc78d-1128">Name</span></span>|<span data-ttu-id="cc78d-1129">类型</span><span class="sxs-lookup"><span data-stu-id="cc78d-1129">Type</span></span>|<span data-ttu-id="cc78d-1130">描述</span><span class="sxs-lookup"><span data-stu-id="cc78d-1130">Description</span></span>|
|---|---|---|
|`entityType`|[<span data-ttu-id="cc78d-1131">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="cc78d-1131">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook/office.mailboxenums.entitytype)|<span data-ttu-id="cc78d-1132">EntityType 枚举值之一。</span><span class="sxs-lookup"><span data-stu-id="cc78d-1132">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="cc78d-1133">要求</span><span class="sxs-lookup"><span data-stu-id="cc78d-1133">Requirements</span></span>

|<span data-ttu-id="cc78d-1134">要求</span><span class="sxs-lookup"><span data-stu-id="cc78d-1134">Requirement</span></span>|<span data-ttu-id="cc78d-1135">值</span><span class="sxs-lookup"><span data-stu-id="cc78d-1135">Value</span></span>|
|---|---|
|[<span data-ttu-id="cc78d-1136">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="cc78d-1136">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="cc78d-1137">1.0</span><span class="sxs-lookup"><span data-stu-id="cc78d-1137">1.0</span></span>|
|[<span data-ttu-id="cc78d-1138">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="cc78d-1138">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="cc78d-1139">受限</span><span class="sxs-lookup"><span data-stu-id="cc78d-1139">Restricted</span></span>|
|[<span data-ttu-id="cc78d-1140">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="cc78d-1140">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="cc78d-1141">阅读</span><span class="sxs-lookup"><span data-stu-id="cc78d-1141">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="cc78d-1142">返回：</span><span class="sxs-lookup"><span data-stu-id="cc78d-1142">Returns:</span></span>

<span data-ttu-id="cc78d-1143">如果在 `entityType` 中传递的值不是 `EntityType` 枚举的有效成员，该方法返回 null。</span><span class="sxs-lookup"><span data-stu-id="cc78d-1143">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="cc78d-1144">如果指定类型的任何实体都不存在于该项目的正文中，该方法将返回空数组。</span><span class="sxs-lookup"><span data-stu-id="cc78d-1144">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="cc78d-1145">否则，返回的数组中对象的类型取决于 `entityType` 参数中请求实体的类型。</span><span class="sxs-lookup"><span data-stu-id="cc78d-1145">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="cc78d-1146">当使用此方法的最低权限级别**受限**时，某些实体类型需要 **ReadItem** 才能进行访问，如下表中所指定。</span><span class="sxs-lookup"><span data-stu-id="cc78d-1146">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

|<span data-ttu-id="cc78d-1147">`entityType` 的值</span><span class="sxs-lookup"><span data-stu-id="cc78d-1147">Value of `entityType`</span></span>|<span data-ttu-id="cc78d-1148">返回的数组中对象的类型</span><span class="sxs-lookup"><span data-stu-id="cc78d-1148">Type of objects in returned array</span></span>|<span data-ttu-id="cc78d-1149">所需权限级别</span><span class="sxs-lookup"><span data-stu-id="cc78d-1149">Required Permission Level</span></span>|
|---|---|---|
|`Address`|<span data-ttu-id="cc78d-1150">字符串</span><span class="sxs-lookup"><span data-stu-id="cc78d-1150">String</span></span>|<span data-ttu-id="cc78d-1151">**受限**</span><span class="sxs-lookup"><span data-stu-id="cc78d-1151">**Restricted**</span></span>|
|`Contact`|<span data-ttu-id="cc78d-1152">Contact</span><span class="sxs-lookup"><span data-stu-id="cc78d-1152">Contact</span></span>|<span data-ttu-id="cc78d-1153">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="cc78d-1153">**ReadItem**</span></span>|
|`EmailAddress`|<span data-ttu-id="cc78d-1154">String</span><span class="sxs-lookup"><span data-stu-id="cc78d-1154">String</span></span>|<span data-ttu-id="cc78d-1155">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="cc78d-1155">**ReadItem**</span></span>|
|`MeetingSuggestion`|<span data-ttu-id="cc78d-1156">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="cc78d-1156">MeetingSuggestion</span></span>|<span data-ttu-id="cc78d-1157">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="cc78d-1157">**ReadItem**</span></span>|
|`PhoneNumber`|<span data-ttu-id="cc78d-1158">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="cc78d-1158">PhoneNumber</span></span>|<span data-ttu-id="cc78d-1159">**受限**</span><span class="sxs-lookup"><span data-stu-id="cc78d-1159">**Restricted**</span></span>|
|`TaskSuggestion`|<span data-ttu-id="cc78d-1160">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="cc78d-1160">TaskSuggestion</span></span>|<span data-ttu-id="cc78d-1161">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="cc78d-1161">**ReadItem**</span></span>|
|`URL`|<span data-ttu-id="cc78d-1162">String</span><span class="sxs-lookup"><span data-stu-id="cc78d-1162">String</span></span>|<span data-ttu-id="cc78d-1163">**受限**</span><span class="sxs-lookup"><span data-stu-id="cc78d-1163">**Restricted**</span></span>|

<span data-ttu-id="cc78d-1164">类型：Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="cc78d-1164">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span></span>

##### <a name="example"></a><span data-ttu-id="cc78d-1165">示例</span><span class="sxs-lookup"><span data-stu-id="cc78d-1165">Example</span></span>

<span data-ttu-id="cc78d-1166">以下示例显示了如何访问表示当前项目的正文中的邮政地址的字符串数组。</span><span class="sxs-lookup"><span data-stu-id="cc78d-1166">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

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
};
```

---
---

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlookofficecontactmeetingsuggestionjavascriptapioutlookofficemeetingsuggestionphonenumberjavascriptapioutlookofficephonenumbertasksuggestionjavascriptapioutlookofficetasksuggestion"></a><span data-ttu-id="cc78d-1167">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="cc78d-1167">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span></span>

<span data-ttu-id="cc78d-1168">返回传递清单 XML 文件中定义的命名筛选器的所选项目中的已知实体。</span><span class="sxs-lookup"><span data-stu-id="cc78d-1168">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="cc78d-1169">在 Outlook for iOS 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="cc78d-1169">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="cc78d-1170">`getFilteredEntitiesByName` 方法返回匹配在具有指定 `FilterName` 元素值的清单 XML 文件中的 [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) 规则元素中定义的正则表达式的实体。</span><span class="sxs-lookup"><span data-stu-id="cc78d-1170">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="cc78d-1171">参数</span><span class="sxs-lookup"><span data-stu-id="cc78d-1171">Parameters</span></span>

|<span data-ttu-id="cc78d-1172">名称</span><span class="sxs-lookup"><span data-stu-id="cc78d-1172">Name</span></span>|<span data-ttu-id="cc78d-1173">类型</span><span class="sxs-lookup"><span data-stu-id="cc78d-1173">Type</span></span>|<span data-ttu-id="cc78d-1174">描述</span><span class="sxs-lookup"><span data-stu-id="cc78d-1174">Description</span></span>|
|---|---|---|
|`name`|<span data-ttu-id="cc78d-1175">字符串</span><span class="sxs-lookup"><span data-stu-id="cc78d-1175">String</span></span>|<span data-ttu-id="cc78d-1176">定义筛选器匹配的 `ItemHasKnownEntity` 规则元素的名称。</span><span class="sxs-lookup"><span data-stu-id="cc78d-1176">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="cc78d-1177">要求</span><span class="sxs-lookup"><span data-stu-id="cc78d-1177">Requirements</span></span>

|<span data-ttu-id="cc78d-1178">要求</span><span class="sxs-lookup"><span data-stu-id="cc78d-1178">Requirement</span></span>|<span data-ttu-id="cc78d-1179">值</span><span class="sxs-lookup"><span data-stu-id="cc78d-1179">Value</span></span>|
|---|---|
|[<span data-ttu-id="cc78d-1180">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="cc78d-1180">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="cc78d-1181">1.0</span><span class="sxs-lookup"><span data-stu-id="cc78d-1181">1.0</span></span>|
|[<span data-ttu-id="cc78d-1182">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="cc78d-1182">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="cc78d-1183">ReadItem</span><span class="sxs-lookup"><span data-stu-id="cc78d-1183">ReadItem</span></span>|
|[<span data-ttu-id="cc78d-1184">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="cc78d-1184">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="cc78d-1185">阅读</span><span class="sxs-lookup"><span data-stu-id="cc78d-1185">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="cc78d-1186">返回：</span><span class="sxs-lookup"><span data-stu-id="cc78d-1186">Returns:</span></span>

<span data-ttu-id="cc78d-p164">如果具有匹配 `name` 参数的 `FilterName` 元素值的清单中没有任何 `ItemHasKnownEntity` 元素，则该方法返回 `null`。如果 `name` 参数匹配清单中的 `ItemHasKnownEntity` 元素，但在匹配的当前项目中没有实体，则该方法返回一个空数组。</span><span class="sxs-lookup"><span data-stu-id="cc78d-p164">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="cc78d-1189">类型：Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="cc78d-1189">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span></span>

---
---

#### <a name="getinitializationcontextasyncoptions-callback"></a><span data-ttu-id="cc78d-1190">office.context.mailbox.item.getinitializationcontextasync ([options], [callback])</span><span class="sxs-lookup"><span data-stu-id="cc78d-1190">getInitializationContextAsync([options], [callback])</span></span>

<span data-ttu-id="cc78d-1191">获取[通过可操作邮件激活](/outlook/actionable-messages/invoke-add-in-from-actionable-message)加载项时传递的初始化数据。</span><span class="sxs-lookup"><span data-stu-id="cc78d-1191">Gets initialization data passed when the add-in is [activated by an actionable message](/outlook/actionable-messages/invoke-add-in-from-actionable-message).</span></span>

> [!NOTE]
> <span data-ttu-id="cc78d-1192">仅 outlook 2016 或更高版本 (早于16.0.8413.1000 的即点即用版本) 和适用于 Office 365 的 outlook 网页版支持此方法。</span><span class="sxs-lookup"><span data-stu-id="cc78d-1192">This method is only supported by Outlook 2016 or later for Windows (Click-to-Run versions later than 16.0.8413.1000) and Outlook on the web for Office 365.</span></span>

##### <a name="parameters"></a><span data-ttu-id="cc78d-1193">参数</span><span class="sxs-lookup"><span data-stu-id="cc78d-1193">Parameters</span></span>

|<span data-ttu-id="cc78d-1194">名称</span><span class="sxs-lookup"><span data-stu-id="cc78d-1194">Name</span></span>|<span data-ttu-id="cc78d-1195">类型</span><span class="sxs-lookup"><span data-stu-id="cc78d-1195">Type</span></span>|<span data-ttu-id="cc78d-1196">属性</span><span class="sxs-lookup"><span data-stu-id="cc78d-1196">Attributes</span></span>|<span data-ttu-id="cc78d-1197">说明</span><span class="sxs-lookup"><span data-stu-id="cc78d-1197">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="cc78d-1198">Object</span><span class="sxs-lookup"><span data-stu-id="cc78d-1198">Object</span></span>|<span data-ttu-id="cc78d-1199">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="cc78d-1199">&lt;optional&gt;</span></span>|<span data-ttu-id="cc78d-1200">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="cc78d-1200">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="cc78d-1201">Object</span><span class="sxs-lookup"><span data-stu-id="cc78d-1201">Object</span></span>|<span data-ttu-id="cc78d-1202">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="cc78d-1202">&lt;optional&gt;</span></span>|<span data-ttu-id="cc78d-1203">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="cc78d-1203">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="cc78d-1204">函数</span><span class="sxs-lookup"><span data-stu-id="cc78d-1204">function</span></span>|<span data-ttu-id="cc78d-1205">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="cc78d-1205">&lt;optional&gt;</span></span>|<span data-ttu-id="cc78d-1206">方法完成后，使用单个参数 `callback`（一个 [`asyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `AsyncResult` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="cc78d-1206">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="cc78d-1207">如果成功, 初始化数据在`asyncResult.value`属性中提供为字符串。</span><span class="sxs-lookup"><span data-stu-id="cc78d-1207">On success, the initialization data is provided in the `asyncResult.value` property as a string.</span></span><br/><span data-ttu-id="cc78d-1208">如果没有初始化上下文, 该`asyncResult`对象将包含其`Error` `code`属性设置为`9020`的对象及其`name`属性设置为。 `GenericResponseError`</span><span class="sxs-lookup"><span data-stu-id="cc78d-1208">If there is no initialization context, the `asyncResult` object will contain an `Error` object with its `code` property set to `9020` and its `name` property set to `GenericResponseError`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="cc78d-1209">Requirements</span><span class="sxs-lookup"><span data-stu-id="cc78d-1209">Requirements</span></span>

|<span data-ttu-id="cc78d-1210">要求</span><span class="sxs-lookup"><span data-stu-id="cc78d-1210">Requirement</span></span>|<span data-ttu-id="cc78d-1211">值</span><span class="sxs-lookup"><span data-stu-id="cc78d-1211">Value</span></span>|
|---|---|
|[<span data-ttu-id="cc78d-1212">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="cc78d-1212">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="cc78d-1213">预览</span><span class="sxs-lookup"><span data-stu-id="cc78d-1213">Preview</span></span>|
|[<span data-ttu-id="cc78d-1214">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="cc78d-1214">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="cc78d-1215">ReadItem</span><span class="sxs-lookup"><span data-stu-id="cc78d-1215">ReadItem</span></span>|
|[<span data-ttu-id="cc78d-1216">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="cc78d-1216">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="cc78d-1217">阅读</span><span class="sxs-lookup"><span data-stu-id="cc78d-1217">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="cc78d-1218">示例</span><span class="sxs-lookup"><span data-stu-id="cc78d-1218">Example</span></span>

```javascript
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

---
---

#### <a name="getregexmatches--object"></a><span data-ttu-id="cc78d-1219">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="cc78d-1219">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="cc78d-1220">返回所选项目中匹配在清单 XML 文件中定义的正则表达式的字符串值。</span><span class="sxs-lookup"><span data-stu-id="cc78d-1220">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="cc78d-1221">在 Outlook for iOS 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="cc78d-1221">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="cc78d-p165">`getRegExMatches` 方法返回匹配在清单 XML 文件中的每个 `ItemHasRegularExpressionMatch` 或 `ItemHasKnownEntity` 规则元素中定义的正则表达式的字符串。对于 `ItemHasRegularExpressionMatch` 规则，匹配字符串必须发生在该规则指定的项目的属性中。`PropertyName` 简单类型定义支持的属性。</span><span class="sxs-lookup"><span data-stu-id="cc78d-p165">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="cc78d-1225">例如，考虑一个外接程序清单具有以下 `Rule` 元素：</span><span class="sxs-lookup"><span data-stu-id="cc78d-1225">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="cc78d-1226">从 `getRegExMatches` 返回的对象应有两个属性：`fruits` 和 `veggies`。</span><span class="sxs-lookup"><span data-stu-id="cc78d-1226">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="cc78d-p166">如果在项目的正文属性上指定 `ItemHasRegularExpressionMatch` 规则，则正则表达式应进一步筛选正文，不应尝试返回该项目的整个正文。使用正则表达式（如 `.*`）获取项目的整个正文并不总是返回预期的结果。而是使用 [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) 方法检索整个正文。</span><span class="sxs-lookup"><span data-stu-id="cc78d-p166">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="cc78d-1230">Requirements</span><span class="sxs-lookup"><span data-stu-id="cc78d-1230">Requirements</span></span>

|<span data-ttu-id="cc78d-1231">要求</span><span class="sxs-lookup"><span data-stu-id="cc78d-1231">Requirement</span></span>|<span data-ttu-id="cc78d-1232">值</span><span class="sxs-lookup"><span data-stu-id="cc78d-1232">Value</span></span>|
|---|---|
|[<span data-ttu-id="cc78d-1233">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="cc78d-1233">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="cc78d-1234">1.0</span><span class="sxs-lookup"><span data-stu-id="cc78d-1234">1.0</span></span>|
|[<span data-ttu-id="cc78d-1235">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="cc78d-1235">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="cc78d-1236">ReadItem</span><span class="sxs-lookup"><span data-stu-id="cc78d-1236">ReadItem</span></span>|
|[<span data-ttu-id="cc78d-1237">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="cc78d-1237">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="cc78d-1238">阅读</span><span class="sxs-lookup"><span data-stu-id="cc78d-1238">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="cc78d-1239">返回：</span><span class="sxs-lookup"><span data-stu-id="cc78d-1239">Returns:</span></span>

<span data-ttu-id="cc78d-p167">一个包含匹配在清单 XML 文件中定义的正则表达式的字符串数组的对象。每个数组的名称等于匹配 `ItemHasRegularExpressionMatch` 规则的 `RegExName` 属性或匹配 `ItemHasKnownEntity` 规则的 `FilterName` 属性的相应值。</span><span class="sxs-lookup"><span data-stu-id="cc78d-p167">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<dl class="param-type"><span data-ttu-id="cc78d-1242">

<dt>
类型</dt>


</span><span class="sxs-lookup"><span data-stu-id="cc78d-1242">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="cc78d-1243">对象</span><span class="sxs-lookup"><span data-stu-id="cc78d-1243">Object</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="cc78d-1244">示例</span><span class="sxs-lookup"><span data-stu-id="cc78d-1244">Example</span></span>

<span data-ttu-id="cc78d-1245">以下示例显示了如何访问正则表达式规则元素 `fruits` 和 `veggies` 的匹配项的数组，这些元素在清单中指定。</span><span class="sxs-lookup"><span data-stu-id="cc78d-1245">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```javascript
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veggies = allMatches.veggies;
```

---
---

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="cc78d-1246">getRegExMatchesByName(name) → (nullable) {Array.<String>}</span><span class="sxs-lookup"><span data-stu-id="cc78d-1246">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="cc78d-1247">返回匹配在清单 XML 文件中定义的命名正则表达式的所选项目中的字符串值。</span><span class="sxs-lookup"><span data-stu-id="cc78d-1247">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="cc78d-1248">在 Outlook for iOS 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="cc78d-1248">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="cc78d-1249">`getRegExMatchesByName` 方法返回匹配在具有指定 `RegExName` 元素值的清单 XML 文件中的 `ItemHasRegularExpressionMatch` 规则元素中定义的正则表达式的字符串。</span><span class="sxs-lookup"><span data-stu-id="cc78d-1249">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="cc78d-p168">如果在项目的正文属性上指定 `ItemHasRegularExpressionMatch` 规则，则正则表达式应进一步筛选正文，不应尝试返回该项目的整个正文。使用正则表达式（如 `.*`）获取项目的整个正文并不总是返回预期的结果。</span><span class="sxs-lookup"><span data-stu-id="cc78d-p168">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="cc78d-1252">参数</span><span class="sxs-lookup"><span data-stu-id="cc78d-1252">Parameters</span></span>

|<span data-ttu-id="cc78d-1253">名称</span><span class="sxs-lookup"><span data-stu-id="cc78d-1253">Name</span></span>|<span data-ttu-id="cc78d-1254">类型</span><span class="sxs-lookup"><span data-stu-id="cc78d-1254">Type</span></span>|<span data-ttu-id="cc78d-1255">描述</span><span class="sxs-lookup"><span data-stu-id="cc78d-1255">Description</span></span>|
|---|---|---|
|`name`|<span data-ttu-id="cc78d-1256">字符串</span><span class="sxs-lookup"><span data-stu-id="cc78d-1256">String</span></span>|<span data-ttu-id="cc78d-1257">定义筛选器匹配的 `ItemHasRegularExpressionMatch` 规则元素的名称。</span><span class="sxs-lookup"><span data-stu-id="cc78d-1257">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="cc78d-1258">Requirements</span><span class="sxs-lookup"><span data-stu-id="cc78d-1258">Requirements</span></span>

|<span data-ttu-id="cc78d-1259">要求</span><span class="sxs-lookup"><span data-stu-id="cc78d-1259">Requirement</span></span>|<span data-ttu-id="cc78d-1260">值</span><span class="sxs-lookup"><span data-stu-id="cc78d-1260">Value</span></span>|
|---|---|
|[<span data-ttu-id="cc78d-1261">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="cc78d-1261">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="cc78d-1262">1.0</span><span class="sxs-lookup"><span data-stu-id="cc78d-1262">1.0</span></span>|
|[<span data-ttu-id="cc78d-1263">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="cc78d-1263">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="cc78d-1264">ReadItem</span><span class="sxs-lookup"><span data-stu-id="cc78d-1264">ReadItem</span></span>|
|[<span data-ttu-id="cc78d-1265">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="cc78d-1265">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="cc78d-1266">阅读</span><span class="sxs-lookup"><span data-stu-id="cc78d-1266">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="cc78d-1267">返回：</span><span class="sxs-lookup"><span data-stu-id="cc78d-1267">Returns:</span></span>

<span data-ttu-id="cc78d-1268">一个包含匹配在清单 XML 文件中定义的正则表达式的字符串的数组。</span><span class="sxs-lookup"><span data-stu-id="cc78d-1268">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<dl class="param-type"><span data-ttu-id="cc78d-1269">

<dt>类型</dt>

</span><span class="sxs-lookup"><span data-stu-id="cc78d-1269">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="cc78d-1270">Array.< String ></span><span class="sxs-lookup"><span data-stu-id="cc78d-1270">Array.< String ></span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="cc78d-1271">示例</span><span class="sxs-lookup"><span data-stu-id="cc78d-1271">Example</span></span>

```javascript
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

---
---

####  <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="cc78d-1272">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="cc78d-1272">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="cc78d-1273">以异步方式返回邮件的主题或正文中选定的数据。</span><span class="sxs-lookup"><span data-stu-id="cc78d-1273">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="cc78d-p169">如果没有选定内容，但光标位于正文或主题中，此方法将会为所选数据返回 null。如果选定的是字段，而不是正文或主题，则此方法返回 `InvalidSelection` 错误。</span><span class="sxs-lookup"><span data-stu-id="cc78d-p169">If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="cc78d-1276">参数</span><span class="sxs-lookup"><span data-stu-id="cc78d-1276">Parameters</span></span>

|<span data-ttu-id="cc78d-1277">名称</span><span class="sxs-lookup"><span data-stu-id="cc78d-1277">Name</span></span>|<span data-ttu-id="cc78d-1278">类型</span><span class="sxs-lookup"><span data-stu-id="cc78d-1278">Type</span></span>|<span data-ttu-id="cc78d-1279">属性</span><span class="sxs-lookup"><span data-stu-id="cc78d-1279">Attributes</span></span>|<span data-ttu-id="cc78d-1280">说明</span><span class="sxs-lookup"><span data-stu-id="cc78d-1280">Description</span></span>|
|---|---|---|---|
|`coercionType`|[<span data-ttu-id="cc78d-1281">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="cc78d-1281">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="cc78d-p170">请求数据的格式。如果为文本，则此方法返回纯文本作为字符串，删除任何显示的 HTML 标记。如果为 HTML，则此方法返回所选文本，不论是纯文本还是 HTML。</span><span class="sxs-lookup"><span data-stu-id="cc78d-p170">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`|<span data-ttu-id="cc78d-1285">Object</span><span class="sxs-lookup"><span data-stu-id="cc78d-1285">Object</span></span>|<span data-ttu-id="cc78d-1286">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="cc78d-1286">&lt;optional&gt;</span></span>|<span data-ttu-id="cc78d-1287">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="cc78d-1287">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="cc78d-1288">Object</span><span class="sxs-lookup"><span data-stu-id="cc78d-1288">Object</span></span>|<span data-ttu-id="cc78d-1289">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="cc78d-1289">&lt;optional&gt;</span></span>|<span data-ttu-id="cc78d-1290">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="cc78d-1290">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="cc78d-1291">函数</span><span class="sxs-lookup"><span data-stu-id="cc78d-1291">function</span></span>||<span data-ttu-id="cc78d-1292">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="cc78d-1292">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="cc78d-1293">若要从回调方法访问所选数据，请调用 `asyncResult.value.data`。</span><span class="sxs-lookup"><span data-stu-id="cc78d-1293">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="cc78d-1294">若要访问选定内容的源属性，请调用 `asyncResult.value.sourceProperty`，这将为 `body` 或 `subject`。</span><span class="sxs-lookup"><span data-stu-id="cc78d-1294">To access the source property that the selection comes from, call `asyncResult.value.sourceProperty`, which will be either `body` or `subject`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="cc78d-1295">要求</span><span class="sxs-lookup"><span data-stu-id="cc78d-1295">Requirements</span></span>

|<span data-ttu-id="cc78d-1296">要求</span><span class="sxs-lookup"><span data-stu-id="cc78d-1296">Requirement</span></span>|<span data-ttu-id="cc78d-1297">值</span><span class="sxs-lookup"><span data-stu-id="cc78d-1297">Value</span></span>|
|---|---|
|[<span data-ttu-id="cc78d-1298">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="cc78d-1298">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="cc78d-1299">1.2</span><span class="sxs-lookup"><span data-stu-id="cc78d-1299">1.2</span></span>|
|[<span data-ttu-id="cc78d-1300">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="cc78d-1300">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="cc78d-1301">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="cc78d-1301">ReadWriteItem</span></span>|
|[<span data-ttu-id="cc78d-1302">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="cc78d-1302">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="cc78d-1303">撰写</span><span class="sxs-lookup"><span data-stu-id="cc78d-1303">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="cc78d-1304">返回：</span><span class="sxs-lookup"><span data-stu-id="cc78d-1304">Returns:</span></span>

<span data-ttu-id="cc78d-1305">作为字符串的所选数据的格式由 `coercionType` 确定。</span><span class="sxs-lookup"><span data-stu-id="cc78d-1305">The selected data as a string with format determined by `coercionType`.</span></span>

<dl class="param-type"><span data-ttu-id="cc78d-1306">

<dt>
类型</dt>


</span><span class="sxs-lookup"><span data-stu-id="cc78d-1306">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="cc78d-1307">字符串</span><span class="sxs-lookup"><span data-stu-id="cc78d-1307">String</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="cc78d-1308">示例</span><span class="sxs-lookup"><span data-stu-id="cc78d-1308">Example</span></span>

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

---
---

#### <a name="getselectedentities--entitiesjavascriptapioutlookofficeentities"></a><span data-ttu-id="cc78d-1309">getSelectedEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="cc78d-1309">getSelectedEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span></span>

<span data-ttu-id="cc78d-1310">获取在用户已选择的突出显示匹配项中找到的实体。</span><span class="sxs-lookup"><span data-stu-id="cc78d-1310">Gets the entities found in a highlighted match a user has selected.</span></span> <span data-ttu-id="cc78d-1311">突出显示匹配项适用于[上下文外接程序](/outlook/add-ins/contextual-outlook-add-ins)。</span><span class="sxs-lookup"><span data-stu-id="cc78d-1311">Highlighted matches apply to [contextual add-ins](/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="cc78d-1312">在 Outlook for iOS 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="cc78d-1312">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="cc78d-1313">Requirements</span><span class="sxs-lookup"><span data-stu-id="cc78d-1313">Requirements</span></span>

|<span data-ttu-id="cc78d-1314">要求</span><span class="sxs-lookup"><span data-stu-id="cc78d-1314">Requirement</span></span>|<span data-ttu-id="cc78d-1315">值</span><span class="sxs-lookup"><span data-stu-id="cc78d-1315">Value</span></span>|
|---|---|
|[<span data-ttu-id="cc78d-1316">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="cc78d-1316">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="cc78d-1317">1.6</span><span class="sxs-lookup"><span data-stu-id="cc78d-1317">1.6</span></span>|
|[<span data-ttu-id="cc78d-1318">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="cc78d-1318">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="cc78d-1319">ReadItem</span><span class="sxs-lookup"><span data-stu-id="cc78d-1319">ReadItem</span></span>|
|[<span data-ttu-id="cc78d-1320">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="cc78d-1320">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="cc78d-1321">阅读</span><span class="sxs-lookup"><span data-stu-id="cc78d-1321">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="cc78d-1322">返回：</span><span class="sxs-lookup"><span data-stu-id="cc78d-1322">Returns:</span></span>

<span data-ttu-id="cc78d-1323">类型：[Entities](/javascript/api/outlook/office.entities)</span><span class="sxs-lookup"><span data-stu-id="cc78d-1323">Type: [Entities](/javascript/api/outlook/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="cc78d-1324">示例</span><span class="sxs-lookup"><span data-stu-id="cc78d-1324">Example</span></span>

<span data-ttu-id="cc78d-1325">以下示例访问用户选择的突出显示匹配项中的地址实体。</span><span class="sxs-lookup"><span data-stu-id="cc78d-1325">The following example accesses the addresses entities in the highlighted match selected by the user.</span></span>

```javascript
var contacts = Office.context.mailbox.item.getSelectedEntities().addresses;
```

---
---

#### <a name="getselectedregexmatches--object"></a><span data-ttu-id="cc78d-1326">getSelectedRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="cc78d-1326">getSelectedRegExMatches() → {Object}</span></span>

<span data-ttu-id="cc78d-p173">返回突出显示匹配项中匹配在清单 XML 文件中定义的正则表达式的字符串值。突出显示匹配项适用于[上下文外接程序](/outlook/add-ins/contextual-outlook-add-ins)。</span><span class="sxs-lookup"><span data-stu-id="cc78d-p173">Returns string values in a highlighted match that match the regular expressions defined in the manifest XML file. Highlighted matches apply to [contextual add-ins](/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="cc78d-1329">在 Outlook for iOS 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="cc78d-1329">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="cc78d-p174">`getSelectedRegExMatches` 方法返回匹配在清单 XML 文件中的每个 `ItemHasRegularExpressionMatch` 或 `ItemHasKnownEntity` 规则元素中定义的正则表达式的字符串。对于 `ItemHasRegularExpressionMatch` 规则，匹配字符串必须发生在该规则指定的项目的属性中。`PropertyName` 简单类型定义支持的属性。</span><span class="sxs-lookup"><span data-stu-id="cc78d-p174">The `getSelectedRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="cc78d-1333">例如，考虑一个外接程序清单具有以下 `Rule` 元素：</span><span class="sxs-lookup"><span data-stu-id="cc78d-1333">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="cc78d-1334">从 `getRegExMatches` 返回的对象应有两个属性：`fruits` 和 `veggies`。</span><span class="sxs-lookup"><span data-stu-id="cc78d-1334">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="cc78d-p175">如果在项目的正文属性上指定 `ItemHasRegularExpressionMatch` 规则，则正则表达式应进一步筛选正文，不应尝试返回该项目的整个正文。使用正则表达式（如 `.*`）获取项目的整个正文并不总是返回预期的结果。而是使用 [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) 方法检索整个正文。</span><span class="sxs-lookup"><span data-stu-id="cc78d-p175">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="cc78d-1338">Requirements</span><span class="sxs-lookup"><span data-stu-id="cc78d-1338">Requirements</span></span>

|<span data-ttu-id="cc78d-1339">要求</span><span class="sxs-lookup"><span data-stu-id="cc78d-1339">Requirement</span></span>|<span data-ttu-id="cc78d-1340">值</span><span class="sxs-lookup"><span data-stu-id="cc78d-1340">Value</span></span>|
|---|---|
|[<span data-ttu-id="cc78d-1341">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="cc78d-1341">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="cc78d-1342">1.6</span><span class="sxs-lookup"><span data-stu-id="cc78d-1342">1.6</span></span>|
|[<span data-ttu-id="cc78d-1343">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="cc78d-1343">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="cc78d-1344">ReadItem</span><span class="sxs-lookup"><span data-stu-id="cc78d-1344">ReadItem</span></span>|
|[<span data-ttu-id="cc78d-1345">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="cc78d-1345">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="cc78d-1346">阅读</span><span class="sxs-lookup"><span data-stu-id="cc78d-1346">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="cc78d-1347">返回：</span><span class="sxs-lookup"><span data-stu-id="cc78d-1347">Returns:</span></span>

<span data-ttu-id="cc78d-p176">一个包含匹配在清单 XML 文件中定义的正则表达式的字符串数组的对象。每个数组的名称等于匹配 `ItemHasRegularExpressionMatch` 规则的 `RegExName` 属性或匹配 `ItemHasKnownEntity` 规则的 `FilterName` 属性的相应值。</span><span class="sxs-lookup"><span data-stu-id="cc78d-p176">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

##### <a name="example"></a><span data-ttu-id="cc78d-1350">示例</span><span class="sxs-lookup"><span data-stu-id="cc78d-1350">Example</span></span>

<span data-ttu-id="cc78d-1351">以下示例显示了如何访问正则表达式规则元素 `fruits` 和 `veggies` 的匹配项的数组，这些元素在清单中指定。</span><span class="sxs-lookup"><span data-stu-id="cc78d-1351">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```javascript
var selectedMatches = Office.context.mailbox.item.getSelectedRegExMatches();
var fruits = selectedMatches.fruits;
var veggies = selectedMatches.veggies;
```

---
---

#### <a name="getsharedpropertiesasyncoptions-callback"></a><span data-ttu-id="cc78d-1352">getSharedPropertiesAsync ([options], 回拨)</span><span class="sxs-lookup"><span data-stu-id="cc78d-1352">getSharedPropertiesAsync([options], callback)</span></span>

<span data-ttu-id="cc78d-1353">获取共享文件夹、日历或邮箱中的所选约会或邮件的属性。</span><span class="sxs-lookup"><span data-stu-id="cc78d-1353">Gets the properties of the selected appointment or message in a shared folder, calendar, or mailbox.</span></span>

##### <a name="parameters"></a><span data-ttu-id="cc78d-1354">参数</span><span class="sxs-lookup"><span data-stu-id="cc78d-1354">Parameters</span></span>

|<span data-ttu-id="cc78d-1355">名称</span><span class="sxs-lookup"><span data-stu-id="cc78d-1355">Name</span></span>|<span data-ttu-id="cc78d-1356">类型</span><span class="sxs-lookup"><span data-stu-id="cc78d-1356">Type</span></span>|<span data-ttu-id="cc78d-1357">属性</span><span class="sxs-lookup"><span data-stu-id="cc78d-1357">Attributes</span></span>|<span data-ttu-id="cc78d-1358">说明</span><span class="sxs-lookup"><span data-stu-id="cc78d-1358">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="cc78d-1359">Object</span><span class="sxs-lookup"><span data-stu-id="cc78d-1359">Object</span></span>|<span data-ttu-id="cc78d-1360">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="cc78d-1360">&lt;optional&gt;</span></span>|<span data-ttu-id="cc78d-1361">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="cc78d-1361">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="cc78d-1362">Object</span><span class="sxs-lookup"><span data-stu-id="cc78d-1362">Object</span></span>|<span data-ttu-id="cc78d-1363">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="cc78d-1363">&lt;optional&gt;</span></span>|<span data-ttu-id="cc78d-1364">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="cc78d-1364">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="cc78d-1365">函数</span><span class="sxs-lookup"><span data-stu-id="cc78d-1365">function</span></span>||<span data-ttu-id="cc78d-1366">方法完成后，使用单个参数 `callback`（一个 [`asyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `AsyncResult` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="cc78d-1366">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="cc78d-1367">共享属性作为[`SharedProperties`](/javascript/api/outlook/office.sharedproperties) `asyncResult.value`属性中的对象提供。</span><span class="sxs-lookup"><span data-stu-id="cc78d-1367">The shared properties are provided as a [`SharedProperties`](/javascript/api/outlook/office.sharedproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="cc78d-1368">此对象可用于获取项目的共享属性。</span><span class="sxs-lookup"><span data-stu-id="cc78d-1368">This object can be used to get the item's shared properties.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="cc78d-1369">Requirements</span><span class="sxs-lookup"><span data-stu-id="cc78d-1369">Requirements</span></span>

|<span data-ttu-id="cc78d-1370">要求</span><span class="sxs-lookup"><span data-stu-id="cc78d-1370">Requirement</span></span>|<span data-ttu-id="cc78d-1371">值</span><span class="sxs-lookup"><span data-stu-id="cc78d-1371">Value</span></span>|
|---|---|
|[<span data-ttu-id="cc78d-1372">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="cc78d-1372">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="cc78d-1373">预览</span><span class="sxs-lookup"><span data-stu-id="cc78d-1373">Preview</span></span>|
|[<span data-ttu-id="cc78d-1374">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="cc78d-1374">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="cc78d-1375">ReadItem</span><span class="sxs-lookup"><span data-stu-id="cc78d-1375">ReadItem</span></span>|
|[<span data-ttu-id="cc78d-1376">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="cc78d-1376">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="cc78d-1377">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="cc78d-1377">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="cc78d-1378">示例</span><span class="sxs-lookup"><span data-stu-id="cc78d-1378">Example</span></span>

```javascript
Office.context.mailbox.item.getSharedPropertiesAsync(callback);

function callback (asyncResult) {
  var context = asyncResult.context;
  var sharedProperties = asyncResult.value;
}
```

---
---

####  <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="cc78d-1379">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="cc78d-1379">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="cc78d-1380">异步加载所选项目上此外接程序的自定义属性。</span><span class="sxs-lookup"><span data-stu-id="cc78d-1380">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="cc78d-p178">自定义属性基于每个应用、每个项目存储为键/值对。此方法在回调中返回 `CustomProperties` 对象，该回调提供访问特定于当前项目和当前外接程序的自定义属性的方法。自定义属性未在项目上加密，因此这不应用作安全存储。</span><span class="sxs-lookup"><span data-stu-id="cc78d-p178">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="cc78d-1384">参数</span><span class="sxs-lookup"><span data-stu-id="cc78d-1384">Parameters</span></span>

|<span data-ttu-id="cc78d-1385">名称</span><span class="sxs-lookup"><span data-stu-id="cc78d-1385">Name</span></span>|<span data-ttu-id="cc78d-1386">类型</span><span class="sxs-lookup"><span data-stu-id="cc78d-1386">Type</span></span>|<span data-ttu-id="cc78d-1387">属性</span><span class="sxs-lookup"><span data-stu-id="cc78d-1387">Attributes</span></span>|<span data-ttu-id="cc78d-1388">说明</span><span class="sxs-lookup"><span data-stu-id="cc78d-1388">Description</span></span>|
|---|---|---|---|
|`callback`|<span data-ttu-id="cc78d-1389">函数</span><span class="sxs-lookup"><span data-stu-id="cc78d-1389">function</span></span>||<span data-ttu-id="cc78d-1390">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="cc78d-1390">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="cc78d-1391">自定义属性作为 `asyncResult.value` 属性中的 [`CustomProperties`](/javascript/api/outlook/office.customproperties) 对象提供。</span><span class="sxs-lookup"><span data-stu-id="cc78d-1391">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook/office.customproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="cc78d-1392">此对象可用于获取、设置以及从项目中删除自定义属性，并将自定义属性集的更改重新保存到服务器。</span><span class="sxs-lookup"><span data-stu-id="cc78d-1392">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`|<span data-ttu-id="cc78d-1393">Object</span><span class="sxs-lookup"><span data-stu-id="cc78d-1393">Object</span></span>|<span data-ttu-id="cc78d-1394">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="cc78d-1394">&lt;optional&gt;</span></span>|<span data-ttu-id="cc78d-1395">开发人员可以提供他们想要在回调函数中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="cc78d-1395">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="cc78d-1396">此对象可以通过回调函数中的 `asyncResult.asyncContext` 属性进行访问。</span><span class="sxs-lookup"><span data-stu-id="cc78d-1396">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="cc78d-1397">Requirements</span><span class="sxs-lookup"><span data-stu-id="cc78d-1397">Requirements</span></span>

|<span data-ttu-id="cc78d-1398">要求</span><span class="sxs-lookup"><span data-stu-id="cc78d-1398">Requirement</span></span>|<span data-ttu-id="cc78d-1399">值</span><span class="sxs-lookup"><span data-stu-id="cc78d-1399">Value</span></span>|
|---|---|
|[<span data-ttu-id="cc78d-1400">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="cc78d-1400">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="cc78d-1401">1.0</span><span class="sxs-lookup"><span data-stu-id="cc78d-1401">1.0</span></span>|
|[<span data-ttu-id="cc78d-1402">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="cc78d-1402">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="cc78d-1403">ReadItem</span><span class="sxs-lookup"><span data-stu-id="cc78d-1403">ReadItem</span></span>|
|[<span data-ttu-id="cc78d-1404">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="cc78d-1404">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="cc78d-1405">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="cc78d-1405">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="cc78d-1406">示例</span><span class="sxs-lookup"><span data-stu-id="cc78d-1406">Example</span></span>

<span data-ttu-id="cc78d-p181">以下代码示例显示了如何使用 `loadCustomPropertiesAsync` 方法异步加载特定于当前项目的自定义属性。该示例还显示了如何使用 `CustomProperties.saveAsync` 方法将这些属性重新保存到服务器。加载自定义属性后，该代码示例将使用 `CustomProperties.get` 方法读取自定义属性 `myProp`，使用 `CustomProperties.set` 方法写入自定义属性 `otherProp`，最后调用 `saveAsync` 方法保存这些自定义属性。</span><span class="sxs-lookup"><span data-stu-id="cc78d-p181">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

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

---
---

####  <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="cc78d-1410">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="cc78d-1410">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="cc78d-1411">将附件从邮件或约会中删除。</span><span class="sxs-lookup"><span data-stu-id="cc78d-1411">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="cc78d-1412">`removeAttachmentAsync` 方法删除项目中带指定标识符的附件。</span><span class="sxs-lookup"><span data-stu-id="cc78d-1412">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item.</span></span> <span data-ttu-id="cc78d-1413">最佳做法是，仅当同一个邮件应用程序在同一会话中添加了一个附件时，你才应使用该附件标识符来删除该附件。</span><span class="sxs-lookup"><span data-stu-id="cc78d-1413">As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session.</span></span> <span data-ttu-id="cc78d-1414">在 Outlook Web App 和适用于设备的 OWA 中，附件标识符只在同一个会话中才有效。</span><span class="sxs-lookup"><span data-stu-id="cc78d-1414">In Outlook Web App and OWA for Devices, the attachment identifier is valid only within the same session.</span></span> <span data-ttu-id="cc78d-1415">当用户关闭应用程序时, 或者如果用户开始撰写内嵌窗体, 随后弹出窗体以继续在单独的窗口中。</span><span class="sxs-lookup"><span data-stu-id="cc78d-1415">A session is over when the user closes the app, or if the user starts composing an inline form then subsequently pops out the form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="cc78d-1416">参数</span><span class="sxs-lookup"><span data-stu-id="cc78d-1416">Parameters</span></span>

|<span data-ttu-id="cc78d-1417">名称</span><span class="sxs-lookup"><span data-stu-id="cc78d-1417">Name</span></span>|<span data-ttu-id="cc78d-1418">类型</span><span class="sxs-lookup"><span data-stu-id="cc78d-1418">Type</span></span>|<span data-ttu-id="cc78d-1419">属性</span><span class="sxs-lookup"><span data-stu-id="cc78d-1419">Attributes</span></span>|<span data-ttu-id="cc78d-1420">描述</span><span class="sxs-lookup"><span data-stu-id="cc78d-1420">Description</span></span>|
|---|---|---|---|
|`attachmentId`|<span data-ttu-id="cc78d-1421">字符串</span><span class="sxs-lookup"><span data-stu-id="cc78d-1421">String</span></span>||<span data-ttu-id="cc78d-1422">要删除的附件的标识符。</span><span class="sxs-lookup"><span data-stu-id="cc78d-1422">The identifier of the attachment to remove.</span></span>|
|`options`|<span data-ttu-id="cc78d-1423">Object</span><span class="sxs-lookup"><span data-stu-id="cc78d-1423">Object</span></span>|<span data-ttu-id="cc78d-1424">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="cc78d-1424">&lt;optional&gt;</span></span>|<span data-ttu-id="cc78d-1425">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="cc78d-1425">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="cc78d-1426">Object</span><span class="sxs-lookup"><span data-stu-id="cc78d-1426">Object</span></span>|<span data-ttu-id="cc78d-1427">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="cc78d-1427">&lt;optional&gt;</span></span>|<span data-ttu-id="cc78d-1428">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="cc78d-1428">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="cc78d-1429">函数</span><span class="sxs-lookup"><span data-stu-id="cc78d-1429">function</span></span>|<span data-ttu-id="cc78d-1430">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="cc78d-1430">&lt;optional&gt;</span></span>|<span data-ttu-id="cc78d-1431">方法完成后，使用单个参数 `callback`（一个 [`asyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `AsyncResult` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="cc78d-1431">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="cc78d-1432">如果删除附件失败，`asyncResult.error` 属性将包含一个说明失败原因的错误代码。</span><span class="sxs-lookup"><span data-stu-id="cc78d-1432">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="cc78d-1433">错误</span><span class="sxs-lookup"><span data-stu-id="cc78d-1433">Errors</span></span>

|<span data-ttu-id="cc78d-1434">错误代码</span><span class="sxs-lookup"><span data-stu-id="cc78d-1434">Error code</span></span>|<span data-ttu-id="cc78d-1435">说明</span><span class="sxs-lookup"><span data-stu-id="cc78d-1435">Description</span></span>|
|------------|-------------|
|`InvalidAttachmentId`|<span data-ttu-id="cc78d-1436">附件标识符不存在。</span><span class="sxs-lookup"><span data-stu-id="cc78d-1436">The attachment identifier does not exist.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="cc78d-1437">要求</span><span class="sxs-lookup"><span data-stu-id="cc78d-1437">Requirements</span></span>

|<span data-ttu-id="cc78d-1438">要求</span><span class="sxs-lookup"><span data-stu-id="cc78d-1438">Requirement</span></span>|<span data-ttu-id="cc78d-1439">值</span><span class="sxs-lookup"><span data-stu-id="cc78d-1439">Value</span></span>|
|---|---|
|[<span data-ttu-id="cc78d-1440">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="cc78d-1440">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="cc78d-1441">1.1</span><span class="sxs-lookup"><span data-stu-id="cc78d-1441">1.1</span></span>|
|[<span data-ttu-id="cc78d-1442">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="cc78d-1442">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="cc78d-1443">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="cc78d-1443">ReadWriteItem</span></span>|
|[<span data-ttu-id="cc78d-1444">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="cc78d-1444">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="cc78d-1445">撰写</span><span class="sxs-lookup"><span data-stu-id="cc78d-1445">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="cc78d-1446">示例</span><span class="sxs-lookup"><span data-stu-id="cc78d-1446">Example</span></span>

<span data-ttu-id="cc78d-1447">以下代码删除包含标识符 '0' 的附件。</span><span class="sxs-lookup"><span data-stu-id="cc78d-1447">The following code removes an attachment with an identifier of '0'.</span></span>

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

---
---

####  <a name="removehandlerasynceventtype-options-callback"></a><span data-ttu-id="cc78d-1448">removeHandlerAsync(eventType, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="cc78d-1448">removeHandlerAsync(eventType, [options], [callback])</span></span>

<span data-ttu-id="cc78d-1449">删除受支持事件类型的事件处理程序。</span><span class="sxs-lookup"><span data-stu-id="cc78d-1449">Removes the event handlers for a supported event type.</span></span>

<span data-ttu-id="cc78d-1450">目前, 受支持的事件`Office.EventType.AttachmentsChanged`类型`Office.EventType.AppointmentTimeChanged`是`Office.EventType.EnhancedLocationsChanged`、 `Office.EventType.RecipientsChanged`、、 `Office.EventType.RecurrenceChanged`和。</span><span class="sxs-lookup"><span data-stu-id="cc78d-1450">Currently the supported event types are `Office.EventType.AttachmentsChanged`, `Office.EventType.AppointmentTimeChanged`, `Office.EventType.EnhancedLocationsChanged`, `Office.EventType.RecipientsChanged`, and `Office.EventType.RecurrenceChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="cc78d-1451">参数</span><span class="sxs-lookup"><span data-stu-id="cc78d-1451">Parameters</span></span>

| <span data-ttu-id="cc78d-1452">名称</span><span class="sxs-lookup"><span data-stu-id="cc78d-1452">Name</span></span> | <span data-ttu-id="cc78d-1453">类型</span><span class="sxs-lookup"><span data-stu-id="cc78d-1453">Type</span></span> | <span data-ttu-id="cc78d-1454">属性</span><span class="sxs-lookup"><span data-stu-id="cc78d-1454">Attributes</span></span> | <span data-ttu-id="cc78d-1455">说明</span><span class="sxs-lookup"><span data-stu-id="cc78d-1455">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="cc78d-1456">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="cc78d-1456">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="cc78d-1457">应撤销处理程序的事件。</span><span class="sxs-lookup"><span data-stu-id="cc78d-1457">The event that should revoke the handler.</span></span> |
| `options` | <span data-ttu-id="cc78d-1458">对象</span><span class="sxs-lookup"><span data-stu-id="cc78d-1458">Object</span></span> | <span data-ttu-id="cc78d-1459">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="cc78d-1459">&lt;optional&gt;</span></span> | <span data-ttu-id="cc78d-1460">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="cc78d-1460">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="cc78d-1461">Object</span><span class="sxs-lookup"><span data-stu-id="cc78d-1461">Object</span></span> | <span data-ttu-id="cc78d-1462">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="cc78d-1462">&lt;optional&gt;</span></span> | <span data-ttu-id="cc78d-1463">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="cc78d-1463">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="cc78d-1464">函数</span><span class="sxs-lookup"><span data-stu-id="cc78d-1464">function</span></span>| <span data-ttu-id="cc78d-1465">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="cc78d-1465">&lt;optional&gt;</span></span>|<span data-ttu-id="cc78d-1466">方法完成后，使用单个参数 `callback`（一个 [`asyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `AsyncResult` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="cc78d-1466">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="cc78d-1467">Requirements</span><span class="sxs-lookup"><span data-stu-id="cc78d-1467">Requirements</span></span>

|<span data-ttu-id="cc78d-1468">要求</span><span class="sxs-lookup"><span data-stu-id="cc78d-1468">Requirement</span></span>| <span data-ttu-id="cc78d-1469">值</span><span class="sxs-lookup"><span data-stu-id="cc78d-1469">Value</span></span>|
|---|---|
|[<span data-ttu-id="cc78d-1470">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="cc78d-1470">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="cc78d-1471">1.7</span><span class="sxs-lookup"><span data-stu-id="cc78d-1471">1.7</span></span> |
|[<span data-ttu-id="cc78d-1472">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="cc78d-1472">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="cc78d-1473">ReadItem</span><span class="sxs-lookup"><span data-stu-id="cc78d-1473">ReadItem</span></span> |
|[<span data-ttu-id="cc78d-1474">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="cc78d-1474">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="cc78d-1475">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="cc78d-1475">Compose or Read</span></span> |

---
---

####  <a name="saveasyncoptions-callback"></a><span data-ttu-id="cc78d-1476">saveAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="cc78d-1476">saveAsync([options], callback)</span></span>

<span data-ttu-id="cc78d-1477">异步保存项目。</span><span class="sxs-lookup"><span data-stu-id="cc78d-1477">Asynchronously saves an item.</span></span>

<span data-ttu-id="cc78d-p183">调用时，此方法将当前邮件保存为草稿，并通过回调方法返回项目 ID。在 Outlook Web App 或 Outlook 联机模式下，该项目被保存到服务器中。在 Outlook 缓存模式下，该项目被保存到本地缓存中。</span><span class="sxs-lookup"><span data-stu-id="cc78d-p183">When invoked, this method saves the current message as a draft and returns the item id via the callback method. In Outlook Web App or Outlook in online mode, the item is saved to the server. In Outlook in cached mode, the item is saved to the local cache.</span></span>

> [!NOTE]
> <span data-ttu-id="cc78d-1481">如果加载项在撰写模式下对某个项目调用 `saveAsync` 来获得 `itemId`，以便与 EWS 或 REST API 一同使用，请注意，当 Outlook 处于高速缓存模式时，可能需要一段时间项目才能真正同步到服务器。</span><span class="sxs-lookup"><span data-stu-id="cc78d-1481">If your add-in calls `saveAsync` on an item in compose mode in order to get an `itemId` to use with EWS or the REST API, be aware that when Outlook is in cached mode, it may take some time before the item is actually synced to the server.</span></span> <span data-ttu-id="cc78d-1482">在项目同步前，使用 `itemId` 将返回一个错误。</span><span class="sxs-lookup"><span data-stu-id="cc78d-1482">Until the item is synced, using the `itemId` will return an error.</span></span>

<span data-ttu-id="cc78d-p185">由于约会没有草稿状态，如果以撰写模式在约会中调用 `saveAsync`，则该项将被保存为用户日历中的正常约会。对于之前未保存过的新约会，则不会发送邀请。保存现有约会将向添加或删除的与会者发送更新。</span><span class="sxs-lookup"><span data-stu-id="cc78d-p185">Since appointments have no draft state, if `saveAsync` is called on an appointment in compose mode, the item will be saved as a normal appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent. Saving an existing appointment will send an update to added or removed attendees.</span></span>

> [!NOTE]
> <span data-ttu-id="cc78d-1486">以下客户端在撰写模式下对约会上的 `saveAsync` 具有不同的行为：</span><span class="sxs-lookup"><span data-stu-id="cc78d-1486">The following clients have different behavior for `saveAsync` on appointments in compose mode:</span></span>
>
> - <span data-ttu-id="cc78d-1487">Mac Outlook 不支持在撰写模式下对会议执行 `saveAsync` 操作。</span><span class="sxs-lookup"><span data-stu-id="cc78d-1487">Mac Outlook does not support `saveAsync` on a meeting in compose mode.</span></span> <span data-ttu-id="cc78d-1488">对 Mac Outlook 中的会议调用 `saveAsync` 将会返回一个错误。</span><span class="sxs-lookup"><span data-stu-id="cc78d-1488">Calling `saveAsync` on a meeting in Mac Outlook will return an error.</span></span>
> - <span data-ttu-id="cc78d-1489">在撰写模式下的约会上调用 `saveAsync` 时，Outlook 网页版始终发送邀请或更新。</span><span class="sxs-lookup"><span data-stu-id="cc78d-1489">Outlook on the web always sends an invitation or update when `saveAsync` is called on an appointment in compose mode.</span></span>

##### <a name="parameters"></a><span data-ttu-id="cc78d-1490">参数</span><span class="sxs-lookup"><span data-stu-id="cc78d-1490">Parameters</span></span>

|<span data-ttu-id="cc78d-1491">名称</span><span class="sxs-lookup"><span data-stu-id="cc78d-1491">Name</span></span>|<span data-ttu-id="cc78d-1492">类型</span><span class="sxs-lookup"><span data-stu-id="cc78d-1492">Type</span></span>|<span data-ttu-id="cc78d-1493">属性</span><span class="sxs-lookup"><span data-stu-id="cc78d-1493">Attributes</span></span>|<span data-ttu-id="cc78d-1494">说明</span><span class="sxs-lookup"><span data-stu-id="cc78d-1494">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="cc78d-1495">对象</span><span class="sxs-lookup"><span data-stu-id="cc78d-1495">Object</span></span>|<span data-ttu-id="cc78d-1496">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="cc78d-1496">&lt;optional&gt;</span></span>|<span data-ttu-id="cc78d-1497">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="cc78d-1497">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="cc78d-1498">Object</span><span class="sxs-lookup"><span data-stu-id="cc78d-1498">Object</span></span>|<span data-ttu-id="cc78d-1499">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="cc78d-1499">&lt;optional&gt;</span></span>|<span data-ttu-id="cc78d-1500">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="cc78d-1500">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="cc78d-1501">函数</span><span class="sxs-lookup"><span data-stu-id="cc78d-1501">function</span></span>||<span data-ttu-id="cc78d-1502">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="cc78d-1502">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="cc78d-1503">如果成功，该项目标识符将在 `asyncResult.value` 属性中提供。</span><span class="sxs-lookup"><span data-stu-id="cc78d-1503">On success, the item identifier is provided in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="cc78d-1504">要求</span><span class="sxs-lookup"><span data-stu-id="cc78d-1504">Requirements</span></span>

|<span data-ttu-id="cc78d-1505">要求</span><span class="sxs-lookup"><span data-stu-id="cc78d-1505">Requirement</span></span>|<span data-ttu-id="cc78d-1506">值</span><span class="sxs-lookup"><span data-stu-id="cc78d-1506">Value</span></span>|
|---|---|
|[<span data-ttu-id="cc78d-1507">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="cc78d-1507">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="cc78d-1508">1.3</span><span class="sxs-lookup"><span data-stu-id="cc78d-1508">1.3</span></span>|
|[<span data-ttu-id="cc78d-1509">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="cc78d-1509">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="cc78d-1510">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="cc78d-1510">ReadWriteItem</span></span>|
|[<span data-ttu-id="cc78d-1511">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="cc78d-1511">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="cc78d-1512">撰写</span><span class="sxs-lookup"><span data-stu-id="cc78d-1512">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="cc78d-1513">示例</span><span class="sxs-lookup"><span data-stu-id="cc78d-1513">Examples</span></span>

```javascript
Office.context.mailbox.item.saveAsync(
  function callback(result) {
    // Process the result.
  });
```

<span data-ttu-id="cc78d-p187">下面是传递给回调函数的 `result` 参数的示例。`value` 属性包含的项目的项目 ID。</span><span class="sxs-lookup"><span data-stu-id="cc78d-p187">The following is an example of the `result` parameter passed to the callback function. The `value` property contains the item ID of the item.</span></span>

```json
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

---
---

####  <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="cc78d-1516">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="cc78d-1516">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="cc78d-1517">以异步方式将数据插入到邮件的正文或主题中。</span><span class="sxs-lookup"><span data-stu-id="cc78d-1517">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="cc78d-p188">`setSelectedDataAsync` 方法将指定字符串插入到项目主题或正文的光标位置，或者，如果在编辑器中已选择文本，则该方法将替换选择的文本。如果光标不在正文或主题字段中，则返回错误。插入之后，光标会位于插入内容的末尾。</span><span class="sxs-lookup"><span data-stu-id="cc78d-p188">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="cc78d-1521">参数</span><span class="sxs-lookup"><span data-stu-id="cc78d-1521">Parameters</span></span>

|<span data-ttu-id="cc78d-1522">名称</span><span class="sxs-lookup"><span data-stu-id="cc78d-1522">Name</span></span>|<span data-ttu-id="cc78d-1523">类型</span><span class="sxs-lookup"><span data-stu-id="cc78d-1523">Type</span></span>|<span data-ttu-id="cc78d-1524">属性</span><span class="sxs-lookup"><span data-stu-id="cc78d-1524">Attributes</span></span>|<span data-ttu-id="cc78d-1525">描述</span><span class="sxs-lookup"><span data-stu-id="cc78d-1525">Description</span></span>|
|---|---|---|---|
|`data`|<span data-ttu-id="cc78d-1526">字符串</span><span class="sxs-lookup"><span data-stu-id="cc78d-1526">String</span></span>||<span data-ttu-id="cc78d-p189">要插入的数据。数据不得超过 1,000,000 个字符。如果传入的数据超过 1,000,000 个字符，则会引发 `ArgumentOutOfRange` 异常。</span><span class="sxs-lookup"><span data-stu-id="cc78d-p189">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`|<span data-ttu-id="cc78d-1530">Object</span><span class="sxs-lookup"><span data-stu-id="cc78d-1530">Object</span></span>|<span data-ttu-id="cc78d-1531">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="cc78d-1531">&lt;optional&gt;</span></span>|<span data-ttu-id="cc78d-1532">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="cc78d-1532">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="cc78d-1533">Object</span><span class="sxs-lookup"><span data-stu-id="cc78d-1533">Object</span></span>|<span data-ttu-id="cc78d-1534">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="cc78d-1534">&lt;optional&gt;</span></span>|<span data-ttu-id="cc78d-1535">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="cc78d-1535">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`|[<span data-ttu-id="cc78d-1536">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="cc78d-1536">Office.CoercionType</span></span>](office.md#coerciontype-string)|<span data-ttu-id="cc78d-1537">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="cc78d-1537">&lt;optional&gt;</span></span>|<span data-ttu-id="cc78d-p190">如果为 `text`，则在 Outlook Web App 和 Outlook 中应用当前样式。如果该字段是 HTML 编辑器，则仅插入文本数据，即使数据为 HTML。</span><span class="sxs-lookup"><span data-stu-id="cc78d-p190">If `text`, the current style is applied in Outlook Web App and Outlook. If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="cc78d-p191">如果 `html` 和该字段支持 HTML（主题不支持），则在 Outlook Web App 中应用当前样式，而在 Outlook 中应用默认样式。如果该字段是文本字段，则返回 `InvalidDataFormat` 错误。</span><span class="sxs-lookup"><span data-stu-id="cc78d-p191">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook Web App and the default style is applied in Outlook. If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="cc78d-1542">如果未设置 `coercionType`，则结果取决于该字段：如果该字段是 HTML，则使用 HTML；如果该字段是文本，则使用纯文本。</span><span class="sxs-lookup"><span data-stu-id="cc78d-1542">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`|<span data-ttu-id="cc78d-1543">function</span><span class="sxs-lookup"><span data-stu-id="cc78d-1543">function</span></span>||<span data-ttu-id="cc78d-1544">方法完成后，使用单个参数 `callback`（一个 [`asyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `AsyncResult` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="cc78d-1544">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="cc78d-1545">Requirements</span><span class="sxs-lookup"><span data-stu-id="cc78d-1545">Requirements</span></span>

|<span data-ttu-id="cc78d-1546">要求</span><span class="sxs-lookup"><span data-stu-id="cc78d-1546">Requirement</span></span>|<span data-ttu-id="cc78d-1547">值</span><span class="sxs-lookup"><span data-stu-id="cc78d-1547">Value</span></span>|
|---|---|
|[<span data-ttu-id="cc78d-1548">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="cc78d-1548">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="cc78d-1549">1.2</span><span class="sxs-lookup"><span data-stu-id="cc78d-1549">1.2</span></span>|
|[<span data-ttu-id="cc78d-1550">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="cc78d-1550">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="cc78d-1551">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="cc78d-1551">ReadWriteItem</span></span>|
|[<span data-ttu-id="cc78d-1552">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="cc78d-1552">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="cc78d-1553">撰写</span><span class="sxs-lookup"><span data-stu-id="cc78d-1553">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="cc78d-1554">示例</span><span class="sxs-lookup"><span data-stu-id="cc78d-1554">Example</span></span>

```javascript
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```
