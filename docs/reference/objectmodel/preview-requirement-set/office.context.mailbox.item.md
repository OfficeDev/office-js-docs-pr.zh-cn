---
title: "\"Context.subname\"-\"邮箱\"-预览要求集"
description: ''
ms.date: 06/20/2019
localization_priority: Normal
ms.openlocfilehash: 9e3a969540d84891eba49b98c34b69396ee702a8
ms.sourcegitcommit: 382e2735a1295da914f2bfc38883e518070cec61
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 06/21/2019
ms.locfileid: "35128494"
---
# <a name="item"></a><span data-ttu-id="07787-102">item</span><span class="sxs-lookup"><span data-stu-id="07787-102">item</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmditem"></a><span data-ttu-id="07787-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span><span class="sxs-lookup"><span data-stu-id="07787-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span></span>

<span data-ttu-id="07787-p101">`item` 命名空间用于访问当前选定的邮件、会议请求或约会。可以通过使用 [itemType](#itemtype-officemailboxenumsitemtype) 属性确定 `item` 的类型。</span><span class="sxs-lookup"><span data-stu-id="07787-p101">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="07787-106">Requirements</span><span class="sxs-lookup"><span data-stu-id="07787-106">Requirements</span></span>

|<span data-ttu-id="07787-107">要求</span><span class="sxs-lookup"><span data-stu-id="07787-107">Requirement</span></span>|<span data-ttu-id="07787-108">值</span><span class="sxs-lookup"><span data-stu-id="07787-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="07787-109">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="07787-109">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="07787-110">1.0</span><span class="sxs-lookup"><span data-stu-id="07787-110">1.0</span></span>|
|[<span data-ttu-id="07787-111">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="07787-111">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="07787-112">受限</span><span class="sxs-lookup"><span data-stu-id="07787-112">Restricted</span></span>|
|[<span data-ttu-id="07787-113">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="07787-113">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="07787-114">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="07787-114">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="07787-115">成员和方法</span><span class="sxs-lookup"><span data-stu-id="07787-115">Members and methods</span></span>

| <span data-ttu-id="07787-116">成员</span><span class="sxs-lookup"><span data-stu-id="07787-116">Member</span></span> | <span data-ttu-id="07787-117">类型</span><span class="sxs-lookup"><span data-stu-id="07787-117">Type</span></span> |
|--------|------|
| [<span data-ttu-id="07787-118">attachments</span><span class="sxs-lookup"><span data-stu-id="07787-118">attachments</span></span>](#attachments-arrayattachmentdetails) | <span data-ttu-id="07787-119">成员</span><span class="sxs-lookup"><span data-stu-id="07787-119">Member</span></span> |
| [<span data-ttu-id="07787-120">bcc</span><span class="sxs-lookup"><span data-stu-id="07787-120">bcc</span></span>](#bcc-recipients) | <span data-ttu-id="07787-121">成员</span><span class="sxs-lookup"><span data-stu-id="07787-121">Member</span></span> |
| [<span data-ttu-id="07787-122">body</span><span class="sxs-lookup"><span data-stu-id="07787-122">body</span></span>](#body-body) | <span data-ttu-id="07787-123">成员</span><span class="sxs-lookup"><span data-stu-id="07787-123">Member</span></span> |
| [<span data-ttu-id="07787-124">种类</span><span class="sxs-lookup"><span data-stu-id="07787-124">categories</span></span>](#categories-categories) | <span data-ttu-id="07787-125">成员</span><span class="sxs-lookup"><span data-stu-id="07787-125">Member</span></span> |
| [<span data-ttu-id="07787-126">cc</span><span class="sxs-lookup"><span data-stu-id="07787-126">cc</span></span>](#cc-arrayemailaddressdetailsrecipients) | <span data-ttu-id="07787-127">成员</span><span class="sxs-lookup"><span data-stu-id="07787-127">Member</span></span> |
| [<span data-ttu-id="07787-128">conversationId</span><span class="sxs-lookup"><span data-stu-id="07787-128">conversationId</span></span>](#nullable-conversationid-string) | <span data-ttu-id="07787-129">成员</span><span class="sxs-lookup"><span data-stu-id="07787-129">Member</span></span> |
| [<span data-ttu-id="07787-130">dateTimeCreated</span><span class="sxs-lookup"><span data-stu-id="07787-130">dateTimeCreated</span></span>](#datetimecreated-date) | <span data-ttu-id="07787-131">成员</span><span class="sxs-lookup"><span data-stu-id="07787-131">Member</span></span> |
| [<span data-ttu-id="07787-132">dateTimeModified</span><span class="sxs-lookup"><span data-stu-id="07787-132">dateTimeModified</span></span>](#datetimemodified-date) | <span data-ttu-id="07787-133">成员</span><span class="sxs-lookup"><span data-stu-id="07787-133">Member</span></span> |
| [<span data-ttu-id="07787-134">end</span><span class="sxs-lookup"><span data-stu-id="07787-134">end</span></span>](#end-datetime) | <span data-ttu-id="07787-135">成员</span><span class="sxs-lookup"><span data-stu-id="07787-135">Member</span></span> |
| [<span data-ttu-id="07787-136">enhancedLocation</span><span class="sxs-lookup"><span data-stu-id="07787-136">enhancedLocation</span></span>](#enhancedlocation-enhancedlocation) | <span data-ttu-id="07787-137">成员</span><span class="sxs-lookup"><span data-stu-id="07787-137">Member</span></span> |
| [<span data-ttu-id="07787-138">from</span><span class="sxs-lookup"><span data-stu-id="07787-138">from</span></span>](#from-emailaddressdetailsfrom) | <span data-ttu-id="07787-139">成员</span><span class="sxs-lookup"><span data-stu-id="07787-139">Member</span></span> |
| [<span data-ttu-id="07787-140">internetHeaders</span><span class="sxs-lookup"><span data-stu-id="07787-140">internetHeaders</span></span>](#internetheaders-internetheaders) | <span data-ttu-id="07787-141">成员</span><span class="sxs-lookup"><span data-stu-id="07787-141">Member</span></span> |
| [<span data-ttu-id="07787-142">internetMessageId</span><span class="sxs-lookup"><span data-stu-id="07787-142">internetMessageId</span></span>](#internetmessageid-string) | <span data-ttu-id="07787-143">成员</span><span class="sxs-lookup"><span data-stu-id="07787-143">Member</span></span> |
| [<span data-ttu-id="07787-144">itemClass</span><span class="sxs-lookup"><span data-stu-id="07787-144">itemClass</span></span>](#itemclass-string) | <span data-ttu-id="07787-145">成员</span><span class="sxs-lookup"><span data-stu-id="07787-145">Member</span></span> |
| [<span data-ttu-id="07787-146">itemId</span><span class="sxs-lookup"><span data-stu-id="07787-146">itemId</span></span>](#nullable-itemid-string) | <span data-ttu-id="07787-147">成员</span><span class="sxs-lookup"><span data-stu-id="07787-147">Member</span></span> |
| [<span data-ttu-id="07787-148">itemType</span><span class="sxs-lookup"><span data-stu-id="07787-148">itemType</span></span>](#itemtype-officemailboxenumsitemtype) | <span data-ttu-id="07787-149">成员</span><span class="sxs-lookup"><span data-stu-id="07787-149">Member</span></span> |
| [<span data-ttu-id="07787-150">location</span><span class="sxs-lookup"><span data-stu-id="07787-150">location</span></span>](#location-stringlocation) | <span data-ttu-id="07787-151">成员</span><span class="sxs-lookup"><span data-stu-id="07787-151">Member</span></span> |
| [<span data-ttu-id="07787-152">normalizedSubject</span><span class="sxs-lookup"><span data-stu-id="07787-152">normalizedSubject</span></span>](#normalizedsubject-string) | <span data-ttu-id="07787-153">成员</span><span class="sxs-lookup"><span data-stu-id="07787-153">Member</span></span> |
| [<span data-ttu-id="07787-154">notificationMessages</span><span class="sxs-lookup"><span data-stu-id="07787-154">notificationMessages</span></span>](#notificationmessages-notificationmessages) | <span data-ttu-id="07787-155">Member</span><span class="sxs-lookup"><span data-stu-id="07787-155">Member</span></span> |
| [<span data-ttu-id="07787-156">optionalAttendees</span><span class="sxs-lookup"><span data-stu-id="07787-156">optionalAttendees</span></span>](#optionalattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="07787-157">成员</span><span class="sxs-lookup"><span data-stu-id="07787-157">Member</span></span> |
| [<span data-ttu-id="07787-158">organizer</span><span class="sxs-lookup"><span data-stu-id="07787-158">organizer</span></span>](#organizer-emailaddressdetailsorganizer) | <span data-ttu-id="07787-159">成员</span><span class="sxs-lookup"><span data-stu-id="07787-159">Member</span></span> |
| [<span data-ttu-id="07787-160">recurrence</span><span class="sxs-lookup"><span data-stu-id="07787-160">recurrence</span></span>](#nullable-recurrence-recurrence) | <span data-ttu-id="07787-161">成员</span><span class="sxs-lookup"><span data-stu-id="07787-161">Member</span></span> |
| [<span data-ttu-id="07787-162">requiredAttendees</span><span class="sxs-lookup"><span data-stu-id="07787-162">requiredAttendees</span></span>](#requiredattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="07787-163">成员</span><span class="sxs-lookup"><span data-stu-id="07787-163">Member</span></span> |
| [<span data-ttu-id="07787-164">sender</span><span class="sxs-lookup"><span data-stu-id="07787-164">sender</span></span>](#sender-emailaddressdetails) | <span data-ttu-id="07787-165">Member</span><span class="sxs-lookup"><span data-stu-id="07787-165">Member</span></span> |
| [<span data-ttu-id="07787-166">Webcasts&seriesid</span><span class="sxs-lookup"><span data-stu-id="07787-166">seriesId</span></span>](#nullable-seriesid-string) | <span data-ttu-id="07787-167">Member</span><span class="sxs-lookup"><span data-stu-id="07787-167">Member</span></span> |
| [<span data-ttu-id="07787-168">start</span><span class="sxs-lookup"><span data-stu-id="07787-168">start</span></span>](#start-datetime) | <span data-ttu-id="07787-169">Member</span><span class="sxs-lookup"><span data-stu-id="07787-169">Member</span></span> |
| [<span data-ttu-id="07787-170">subject</span><span class="sxs-lookup"><span data-stu-id="07787-170">subject</span></span>](#subject-stringsubject) | <span data-ttu-id="07787-171">成员</span><span class="sxs-lookup"><span data-stu-id="07787-171">Member</span></span> |
| [<span data-ttu-id="07787-172">to</span><span class="sxs-lookup"><span data-stu-id="07787-172">to</span></span>](#to-arrayemailaddressdetailsrecipients) | <span data-ttu-id="07787-173">成员</span><span class="sxs-lookup"><span data-stu-id="07787-173">Member</span></span> |
| [<span data-ttu-id="07787-174">addFileAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="07787-174">addFileAttachmentAsync</span></span>](#addfileattachmentasyncuri-attachmentname-options-callback) | <span data-ttu-id="07787-175">方法</span><span class="sxs-lookup"><span data-stu-id="07787-175">Method</span></span> |
| [<span data-ttu-id="07787-176">addFileAttachmentFromBase64Async</span><span class="sxs-lookup"><span data-stu-id="07787-176">addFileAttachmentFromBase64Async</span></span>](#addfileattachmentfrombase64asyncbase64file-attachmentname-options-callback) | <span data-ttu-id="07787-177">方法</span><span class="sxs-lookup"><span data-stu-id="07787-177">Method</span></span> |
| [<span data-ttu-id="07787-178">addHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="07787-178">addHandlerAsync</span></span>](#addhandlerasynceventtype-handler-options-callback) | <span data-ttu-id="07787-179">方法</span><span class="sxs-lookup"><span data-stu-id="07787-179">Method</span></span> |
| [<span data-ttu-id="07787-180">addItemAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="07787-180">addItemAttachmentAsync</span></span>](#additemattachmentasyncitemid-attachmentname-options-callback) | <span data-ttu-id="07787-181">方法</span><span class="sxs-lookup"><span data-stu-id="07787-181">Method</span></span> |
| [<span data-ttu-id="07787-182">close</span><span class="sxs-lookup"><span data-stu-id="07787-182">close</span></span>](#close) | <span data-ttu-id="07787-183">方法</span><span class="sxs-lookup"><span data-stu-id="07787-183">Method</span></span> |
| [<span data-ttu-id="07787-184">displayReplyAllForm</span><span class="sxs-lookup"><span data-stu-id="07787-184">displayReplyAllForm</span></span>](#displayreplyallformformdata-callback) | <span data-ttu-id="07787-185">方法</span><span class="sxs-lookup"><span data-stu-id="07787-185">Method</span></span> |
| [<span data-ttu-id="07787-186">displayReplyForm</span><span class="sxs-lookup"><span data-stu-id="07787-186">displayReplyForm</span></span>](#displayreplyformformdata-callback) | <span data-ttu-id="07787-187">方法</span><span class="sxs-lookup"><span data-stu-id="07787-187">Method</span></span> |
| [<span data-ttu-id="07787-188">getAttachmentContentAsync</span><span class="sxs-lookup"><span data-stu-id="07787-188">getAttachmentContentAsync</span></span>](#getattachmentcontentasyncattachmentid-options-callback--attachmentcontent) | <span data-ttu-id="07787-189">方法</span><span class="sxs-lookup"><span data-stu-id="07787-189">Method</span></span> |
| [<span data-ttu-id="07787-190">getAttachmentsAsync</span><span class="sxs-lookup"><span data-stu-id="07787-190">getAttachmentsAsync</span></span>](#getattachmentsasyncoptions-callback--arrayattachmentdetails) | <span data-ttu-id="07787-191">方法</span><span class="sxs-lookup"><span data-stu-id="07787-191">Method</span></span> |
| [<span data-ttu-id="07787-192">getEntities</span><span class="sxs-lookup"><span data-stu-id="07787-192">getEntities</span></span>](#getentities--entities) | <span data-ttu-id="07787-193">方法</span><span class="sxs-lookup"><span data-stu-id="07787-193">Method</span></span> |
| [<span data-ttu-id="07787-194">getEntitiesByType</span><span class="sxs-lookup"><span data-stu-id="07787-194">getEntitiesByType</span></span>](#getentitiesbytypeentitytype--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="07787-195">方法</span><span class="sxs-lookup"><span data-stu-id="07787-195">Method</span></span> |
| [<span data-ttu-id="07787-196">getFilteredEntitiesByName</span><span class="sxs-lookup"><span data-stu-id="07787-196">getFilteredEntitiesByName</span></span>](#getfilteredentitiesbynamename--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="07787-197">方法</span><span class="sxs-lookup"><span data-stu-id="07787-197">Method</span></span> |
| [<span data-ttu-id="07787-198">Office.context.mailbox.item.getinitializationcontextasync</span><span class="sxs-lookup"><span data-stu-id="07787-198">getInitializationContextAsync</span></span>](#getinitializationcontextasyncoptions-callback) | <span data-ttu-id="07787-199">方法</span><span class="sxs-lookup"><span data-stu-id="07787-199">Method</span></span> |
| [<span data-ttu-id="07787-200">getItemIdAsync</span><span class="sxs-lookup"><span data-stu-id="07787-200">getItemIdAsync</span></span>](#getitemidasyncoptions-callback) | <span data-ttu-id="07787-201">方法</span><span class="sxs-lookup"><span data-stu-id="07787-201">Method</span></span> |
| [<span data-ttu-id="07787-202">getRegExMatches</span><span class="sxs-lookup"><span data-stu-id="07787-202">getRegExMatches</span></span>](#getregexmatches--object) | <span data-ttu-id="07787-203">方法</span><span class="sxs-lookup"><span data-stu-id="07787-203">Method</span></span> |
| [<span data-ttu-id="07787-204">getRegExMatchesByName</span><span class="sxs-lookup"><span data-stu-id="07787-204">getRegExMatchesByName</span></span>](#getregexmatchesbynamename--nullable-array-string-) | <span data-ttu-id="07787-205">方法</span><span class="sxs-lookup"><span data-stu-id="07787-205">Method</span></span> |
| [<span data-ttu-id="07787-206">getSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="07787-206">getSelectedDataAsync</span></span>](#getselecteddataasynccoerciontype-options-callback--string) | <span data-ttu-id="07787-207">方法</span><span class="sxs-lookup"><span data-stu-id="07787-207">Method</span></span> |
| [<span data-ttu-id="07787-208">Office.context.mailbox.item.getselectedentities</span><span class="sxs-lookup"><span data-stu-id="07787-208">getSelectedEntities</span></span>](#getselectedentities--entities) | <span data-ttu-id="07787-209">方法</span><span class="sxs-lookup"><span data-stu-id="07787-209">Method</span></span> |
| [<span data-ttu-id="07787-210">Office.context.mailbox.item.getselectedregexmatches</span><span class="sxs-lookup"><span data-stu-id="07787-210">getSelectedRegExMatches</span></span>](#getselectedregexmatches--object) | <span data-ttu-id="07787-211">方法</span><span class="sxs-lookup"><span data-stu-id="07787-211">Method</span></span> |
| [<span data-ttu-id="07787-212">getSharedPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="07787-212">getSharedPropertiesAsync</span></span>](#getsharedpropertiesasyncoptions-callback) | <span data-ttu-id="07787-213">方法</span><span class="sxs-lookup"><span data-stu-id="07787-213">Method</span></span> |
| [<span data-ttu-id="07787-214">loadCustomPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="07787-214">loadCustomPropertiesAsync</span></span>](#loadcustompropertiesasynccallback-usercontext) | <span data-ttu-id="07787-215">方法</span><span class="sxs-lookup"><span data-stu-id="07787-215">Method</span></span> |
| [<span data-ttu-id="07787-216">removeAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="07787-216">removeAttachmentAsync</span></span>](#removeattachmentasyncattachmentid-options-callback) | <span data-ttu-id="07787-217">方法</span><span class="sxs-lookup"><span data-stu-id="07787-217">Method</span></span> |
| [<span data-ttu-id="07787-218">removeHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="07787-218">removeHandlerAsync</span></span>](#removehandlerasynceventtype-options-callback) | <span data-ttu-id="07787-219">方法</span><span class="sxs-lookup"><span data-stu-id="07787-219">Method</span></span> |
| [<span data-ttu-id="07787-220">saveAsync</span><span class="sxs-lookup"><span data-stu-id="07787-220">saveAsync</span></span>](#saveasyncoptions-callback) | <span data-ttu-id="07787-221">方法</span><span class="sxs-lookup"><span data-stu-id="07787-221">Method</span></span> |
| [<span data-ttu-id="07787-222">setSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="07787-222">setSelectedDataAsync</span></span>](#setselecteddataasyncdata-options-callback) | <span data-ttu-id="07787-223">方法</span><span class="sxs-lookup"><span data-stu-id="07787-223">Method</span></span> |

### <a name="example"></a><span data-ttu-id="07787-224">示例</span><span class="sxs-lookup"><span data-stu-id="07787-224">Example</span></span>

<span data-ttu-id="07787-225">以下 JavaScript 代码示例显示了如何访问 Outlook 中当前项目的 `subject` 属性。</span><span class="sxs-lookup"><span data-stu-id="07787-225">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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

### <a name="members"></a><span data-ttu-id="07787-226">成员</span><span class="sxs-lookup"><span data-stu-id="07787-226">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetails"></a><span data-ttu-id="07787-227">附件: Array. <[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="07787-227">attachments: Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span></span>

<span data-ttu-id="07787-228">以数组的形式获取项目的附件。</span><span class="sxs-lookup"><span data-stu-id="07787-228">Gets the item's attachments as an array.</span></span> <span data-ttu-id="07787-229">仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="07787-229">Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="07787-230">由于潜在的安全问题，某些类型的文件会受到 Outlook 阻止，并且不会返回。</span><span class="sxs-lookup"><span data-stu-id="07787-230">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="07787-231">如需了解更多信息，请参阅 [Outlook 中阻止的附件](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519)。</span><span class="sxs-lookup"><span data-stu-id="07787-231">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="07787-232">类型</span><span class="sxs-lookup"><span data-stu-id="07787-232">Type</span></span>

*   <span data-ttu-id="07787-233">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="07787-233">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span></span>

##### <a name="requirements"></a><span data-ttu-id="07787-234">要求</span><span class="sxs-lookup"><span data-stu-id="07787-234">Requirements</span></span>

|<span data-ttu-id="07787-235">要求</span><span class="sxs-lookup"><span data-stu-id="07787-235">Requirement</span></span>|<span data-ttu-id="07787-236">值</span><span class="sxs-lookup"><span data-stu-id="07787-236">Value</span></span>|
|---|---|
|[<span data-ttu-id="07787-237">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="07787-237">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="07787-238">1.0</span><span class="sxs-lookup"><span data-stu-id="07787-238">1.0</span></span>|
|[<span data-ttu-id="07787-239">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="07787-239">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="07787-240">ReadItem</span><span class="sxs-lookup"><span data-stu-id="07787-240">ReadItem</span></span>|
|[<span data-ttu-id="07787-241">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="07787-241">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="07787-242">阅读</span><span class="sxs-lookup"><span data-stu-id="07787-242">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="07787-243">示例</span><span class="sxs-lookup"><span data-stu-id="07787-243">Example</span></span>

<span data-ttu-id="07787-244">以下代码使用当前项目上所有附件的详细信息构成 HTML 字符串。</span><span class="sxs-lookup"><span data-stu-id="07787-244">The following code builds an HTML string with details of all attachments on the current item.</span></span>

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

#### <a name="bcc-recipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="07787-245">密件抄送:[收件人](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="07787-245">bcc: [Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="07787-246">获取对象，该对象提供用于获取或更新邮件的密件抄送 (Bcc) 行上的收件人的方法。</span><span class="sxs-lookup"><span data-stu-id="07787-246">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="07787-247">仅限撰写模式。</span><span class="sxs-lookup"><span data-stu-id="07787-247">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="07787-248">类型</span><span class="sxs-lookup"><span data-stu-id="07787-248">Type</span></span>

*   [<span data-ttu-id="07787-249">收件人</span><span class="sxs-lookup"><span data-stu-id="07787-249">Recipients</span></span>](/javascript/api/outlook/office.recipients)

##### <a name="requirements"></a><span data-ttu-id="07787-250">要求</span><span class="sxs-lookup"><span data-stu-id="07787-250">Requirements</span></span>

|<span data-ttu-id="07787-251">要求</span><span class="sxs-lookup"><span data-stu-id="07787-251">Requirement</span></span>|<span data-ttu-id="07787-252">值</span><span class="sxs-lookup"><span data-stu-id="07787-252">Value</span></span>|
|---|---|
|[<span data-ttu-id="07787-253">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="07787-253">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="07787-254">1.1</span><span class="sxs-lookup"><span data-stu-id="07787-254">1.1</span></span>|
|[<span data-ttu-id="07787-255">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="07787-255">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="07787-256">ReadItem</span><span class="sxs-lookup"><span data-stu-id="07787-256">ReadItem</span></span>|
|[<span data-ttu-id="07787-257">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="07787-257">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="07787-258">撰写</span><span class="sxs-lookup"><span data-stu-id="07787-258">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="07787-259">示例</span><span class="sxs-lookup"><span data-stu-id="07787-259">Example</span></span>

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

#### <a name="body-bodyjavascriptapioutlookofficebody"></a><span data-ttu-id="07787-260">正文:[正文](/javascript/api/outlook/office.body)</span><span class="sxs-lookup"><span data-stu-id="07787-260">body: [Body](/javascript/api/outlook/office.body)</span></span>

<span data-ttu-id="07787-261">获取一个提供用于处理项目正文的方法的对象。</span><span class="sxs-lookup"><span data-stu-id="07787-261">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="07787-262">类型</span><span class="sxs-lookup"><span data-stu-id="07787-262">Type</span></span>

*   [<span data-ttu-id="07787-263">Body</span><span class="sxs-lookup"><span data-stu-id="07787-263">Body</span></span>](/javascript/api/outlook/office.body)

##### <a name="requirements"></a><span data-ttu-id="07787-264">要求</span><span class="sxs-lookup"><span data-stu-id="07787-264">Requirements</span></span>

|<span data-ttu-id="07787-265">要求</span><span class="sxs-lookup"><span data-stu-id="07787-265">Requirement</span></span>|<span data-ttu-id="07787-266">值</span><span class="sxs-lookup"><span data-stu-id="07787-266">Value</span></span>|
|---|---|
|[<span data-ttu-id="07787-267">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="07787-267">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="07787-268">1.1</span><span class="sxs-lookup"><span data-stu-id="07787-268">1.1</span></span>|
|[<span data-ttu-id="07787-269">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="07787-269">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="07787-270">ReadItem</span><span class="sxs-lookup"><span data-stu-id="07787-270">ReadItem</span></span>|
|[<span data-ttu-id="07787-271">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="07787-271">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="07787-272">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="07787-272">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="07787-273">示例</span><span class="sxs-lookup"><span data-stu-id="07787-273">Example</span></span>

<span data-ttu-id="07787-274">本示例获取纯文本格式的邮件正文。</span><span class="sxs-lookup"><span data-stu-id="07787-274">This example gets the body of the message in plain text.</span></span>

```javascript
Office.context.mailbox.item.body.getAsync(
  "text",
  { asyncContext: "This is passed to the callback" },
  function callback(result) {
    // Do something with the result.
  });

```

<span data-ttu-id="07787-275">以下是传递到回调函数的结果参数的示例。</span><span class="sxs-lookup"><span data-stu-id="07787-275">The following is an example of the result parameter passed to the callback function.</span></span>

```json
{
  "value": "TEXT of whole body (including threads below)",
  "status": "succeeded",
  "asyncContext": "This is passed to the callback"
}
```

---
---

#### <a name="categories-categoriesjavascriptapioutlookofficecategories"></a><span data-ttu-id="07787-276">类别:[类别](/javascript/api/outlook/office.categories)</span><span class="sxs-lookup"><span data-stu-id="07787-276">categories: [Categories](/javascript/api/outlook/office.categories)</span></span>

<span data-ttu-id="07787-277">获取一个对象, 该对象提供用于管理项的类别的方法。</span><span class="sxs-lookup"><span data-stu-id="07787-277">Gets an object that provides methods for managing the item's categories.</span></span>

> [!NOTE]
> <span data-ttu-id="07787-278">IOS 或 Android 上的 Outlook 不支持此成员。</span><span class="sxs-lookup"><span data-stu-id="07787-278">This member is not supported in Outlook on iOS or Android.</span></span>

##### <a name="type"></a><span data-ttu-id="07787-279">类型</span><span class="sxs-lookup"><span data-stu-id="07787-279">Type</span></span>

*   [<span data-ttu-id="07787-280">Categories</span><span class="sxs-lookup"><span data-stu-id="07787-280">Categories</span></span>](/javascript/api/outlook/office.categories)

##### <a name="requirements"></a><span data-ttu-id="07787-281">要求</span><span class="sxs-lookup"><span data-stu-id="07787-281">Requirements</span></span>

|<span data-ttu-id="07787-282">要求</span><span class="sxs-lookup"><span data-stu-id="07787-282">Requirement</span></span>|<span data-ttu-id="07787-283">值</span><span class="sxs-lookup"><span data-stu-id="07787-283">Value</span></span>|
|---|---|
|[<span data-ttu-id="07787-284">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="07787-284">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="07787-285">预览</span><span class="sxs-lookup"><span data-stu-id="07787-285">Preview</span></span>|
|[<span data-ttu-id="07787-286">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="07787-286">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="07787-287">ReadItem</span><span class="sxs-lookup"><span data-stu-id="07787-287">ReadItem</span></span>|
|[<span data-ttu-id="07787-288">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="07787-288">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="07787-289">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="07787-289">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="07787-290">示例</span><span class="sxs-lookup"><span data-stu-id="07787-290">Example</span></span>

<span data-ttu-id="07787-291">此示例获取项的类别。</span><span class="sxs-lookup"><span data-stu-id="07787-291">This example gets the item's categories.</span></span>

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

#### <a name="cc-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="07787-292"><[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[收件人](/javascript/api/outlook/office.recipients)的抄送: Array</span><span class="sxs-lookup"><span data-stu-id="07787-292">cc: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="07787-293">提供对邮件的抄送 (Cc) 收件人的访问权限。</span><span class="sxs-lookup"><span data-stu-id="07787-293">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="07787-294">对象的类型和访问级别取决于当前项目的模式。</span><span class="sxs-lookup"><span data-stu-id="07787-294">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="07787-295">阅读模式</span><span class="sxs-lookup"><span data-stu-id="07787-295">Read mode</span></span>

<span data-ttu-id="07787-p106">`cc` 属性返回包含邮件的**抄送**行上所列的每个收件人的 `EmailAddressDetails` 对象的数组。集合上限为 100 个成员。</span><span class="sxs-lookup"><span data-stu-id="07787-p106">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.</span></span>

```javascript
console.log(JSON.stringify(Office.context.mailbox.item.cc));
```

##### <a name="compose-mode"></a><span data-ttu-id="07787-298">撰写模式</span><span class="sxs-lookup"><span data-stu-id="07787-298">Compose mode</span></span>

<span data-ttu-id="07787-299">`cc` 属性返回一个 `Recipients` 对象，该对象提供用于获取或更新邮件的“**抄送**”行上收件人的方法。</span><span class="sxs-lookup"><span data-stu-id="07787-299">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span>

```javascript
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="07787-300">类型</span><span class="sxs-lookup"><span data-stu-id="07787-300">Type</span></span>

*   <span data-ttu-id="07787-301">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="07787-301">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="07787-302">要求</span><span class="sxs-lookup"><span data-stu-id="07787-302">Requirements</span></span>

|<span data-ttu-id="07787-303">要求</span><span class="sxs-lookup"><span data-stu-id="07787-303">Requirement</span></span>|<span data-ttu-id="07787-304">值</span><span class="sxs-lookup"><span data-stu-id="07787-304">Value</span></span>|
|---|---|
|[<span data-ttu-id="07787-305">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="07787-305">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="07787-306">1.0</span><span class="sxs-lookup"><span data-stu-id="07787-306">1.0</span></span>|
|[<span data-ttu-id="07787-307">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="07787-307">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="07787-308">ReadItem</span><span class="sxs-lookup"><span data-stu-id="07787-308">ReadItem</span></span>|
|[<span data-ttu-id="07787-309">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="07787-309">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="07787-310">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="07787-310">Compose or Read</span></span>|

---
---

#### <a name="nullable-conversationid-string"></a><span data-ttu-id="07787-311">(可以为 null) conversationId: String</span><span class="sxs-lookup"><span data-stu-id="07787-311">(nullable) conversationId: String</span></span>

<span data-ttu-id="07787-312">获取包含特定消息的电子邮件会话的标识符。</span><span class="sxs-lookup"><span data-stu-id="07787-312">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="07787-p107">如果在阅读窗体或撰写窗体的回复中激活邮件应用程序，则此属性可以获得一个整数值。如果用户随后更改了回复邮件的主题（若发送回复），则该邮件的对话 ID 将改变且之前获取的值将不适用。</span><span class="sxs-lookup"><span data-stu-id="07787-p107">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="07787-p108">对于撰写窗体的新项目，此属性获得一个 null 值。如果用户设置一个主题并保存该项目，`conversationId` 属性将返回一个值。</span><span class="sxs-lookup"><span data-stu-id="07787-p108">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="07787-317">Type</span><span class="sxs-lookup"><span data-stu-id="07787-317">Type</span></span>

*   <span data-ttu-id="07787-318">String</span><span class="sxs-lookup"><span data-stu-id="07787-318">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="07787-319">要求</span><span class="sxs-lookup"><span data-stu-id="07787-319">Requirements</span></span>

|<span data-ttu-id="07787-320">要求</span><span class="sxs-lookup"><span data-stu-id="07787-320">Requirement</span></span>|<span data-ttu-id="07787-321">值</span><span class="sxs-lookup"><span data-stu-id="07787-321">Value</span></span>|
|---|---|
|[<span data-ttu-id="07787-322">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="07787-322">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="07787-323">1.0</span><span class="sxs-lookup"><span data-stu-id="07787-323">1.0</span></span>|
|[<span data-ttu-id="07787-324">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="07787-324">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="07787-325">ReadItem</span><span class="sxs-lookup"><span data-stu-id="07787-325">ReadItem</span></span>|
|[<span data-ttu-id="07787-326">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="07787-326">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="07787-327">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="07787-327">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="07787-328">示例</span><span class="sxs-lookup"><span data-stu-id="07787-328">Example</span></span>

```javascript
var conversationId = Office.context.mailbox.item.conversationId;
console.log("conversationId: " + conversationId);
```

---
---

#### <a name="datetimecreated-date"></a><span data-ttu-id="07787-329">dateTimeCreated: Date</span><span class="sxs-lookup"><span data-stu-id="07787-329">dateTimeCreated: Date</span></span>

<span data-ttu-id="07787-p109">获取项目创建的日期和时间。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="07787-p109">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="07787-332">类型</span><span class="sxs-lookup"><span data-stu-id="07787-332">Type</span></span>

*   <span data-ttu-id="07787-333">日期</span><span class="sxs-lookup"><span data-stu-id="07787-333">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="07787-334">要求</span><span class="sxs-lookup"><span data-stu-id="07787-334">Requirements</span></span>

|<span data-ttu-id="07787-335">要求</span><span class="sxs-lookup"><span data-stu-id="07787-335">Requirement</span></span>|<span data-ttu-id="07787-336">值</span><span class="sxs-lookup"><span data-stu-id="07787-336">Value</span></span>|
|---|---|
|[<span data-ttu-id="07787-337">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="07787-337">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="07787-338">1.0</span><span class="sxs-lookup"><span data-stu-id="07787-338">1.0</span></span>|
|[<span data-ttu-id="07787-339">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="07787-339">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="07787-340">ReadItem</span><span class="sxs-lookup"><span data-stu-id="07787-340">ReadItem</span></span>|
|[<span data-ttu-id="07787-341">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="07787-341">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="07787-342">阅读</span><span class="sxs-lookup"><span data-stu-id="07787-342">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="07787-343">示例</span><span class="sxs-lookup"><span data-stu-id="07787-343">Example</span></span>

```javascript
var dateTimeCreated = Office.context.mailbox.item.dateTimeCreated;
console.log("Date and time created: " + dateTimeCreated);
```

---
---

#### <a name="datetimemodified-date"></a><span data-ttu-id="07787-344">dateTimeModified: Date</span><span class="sxs-lookup"><span data-stu-id="07787-344">dateTimeModified: Date</span></span>

<span data-ttu-id="07787-345">获取项目最近一次修改的日期和时间。</span><span class="sxs-lookup"><span data-stu-id="07787-345">Gets the date and time that an item was last modified.</span></span> <span data-ttu-id="07787-346">仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="07787-346">Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="07787-347">IOS 或 Android 上的 Outlook 不支持此成员。</span><span class="sxs-lookup"><span data-stu-id="07787-347">This member is not supported in Outlook on iOS or Android.</span></span>

##### <a name="type"></a><span data-ttu-id="07787-348">类型</span><span class="sxs-lookup"><span data-stu-id="07787-348">Type</span></span>

*   <span data-ttu-id="07787-349">日期</span><span class="sxs-lookup"><span data-stu-id="07787-349">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="07787-350">要求</span><span class="sxs-lookup"><span data-stu-id="07787-350">Requirements</span></span>

|<span data-ttu-id="07787-351">要求</span><span class="sxs-lookup"><span data-stu-id="07787-351">Requirement</span></span>|<span data-ttu-id="07787-352">值</span><span class="sxs-lookup"><span data-stu-id="07787-352">Value</span></span>|
|---|---|
|[<span data-ttu-id="07787-353">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="07787-353">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="07787-354">1.0</span><span class="sxs-lookup"><span data-stu-id="07787-354">1.0</span></span>|
|[<span data-ttu-id="07787-355">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="07787-355">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="07787-356">ReadItem</span><span class="sxs-lookup"><span data-stu-id="07787-356">ReadItem</span></span>|
|[<span data-ttu-id="07787-357">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="07787-357">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="07787-358">阅读</span><span class="sxs-lookup"><span data-stu-id="07787-358">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="07787-359">示例</span><span class="sxs-lookup"><span data-stu-id="07787-359">Example</span></span>

```javascript
var dateTimeModified = Office.context.mailbox.item.dateTimeModified;
console.log("Date and time modified: " + dateTimeModified);
```

---
---

#### <a name="end-datetimejavascriptapioutlookofficetime"></a><span data-ttu-id="07787-360">结束: 日期 |[时间](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="07787-360">end: Date|[Time](/javascript/api/outlook/office.time)</span></span>

<span data-ttu-id="07787-361">获取或设置约会结束的日期和时间。</span><span class="sxs-lookup"><span data-stu-id="07787-361">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="07787-p111">将 `end` 属性表示为协调世界时 (UTC) 的日期和时间值。可使用 [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) 方法将 end 属性值转换为客户端的本地日期和时间。</span><span class="sxs-lookup"><span data-stu-id="07787-p111">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="07787-364">阅读模式</span><span class="sxs-lookup"><span data-stu-id="07787-364">Read mode</span></span>

<span data-ttu-id="07787-365">`end` 属性返回 `Date` 对象。</span><span class="sxs-lookup"><span data-stu-id="07787-365">The `end` property returns a `Date` object.</span></span>

```javascript
var end = Office.context.mailbox.item.end;
console.log("Appointment end: " + end);
```

##### <a name="compose-mode"></a><span data-ttu-id="07787-366">撰写模式</span><span class="sxs-lookup"><span data-stu-id="07787-366">Compose mode</span></span>

<span data-ttu-id="07787-367">`end` 属性返回 `Time` 对象。</span><span class="sxs-lookup"><span data-stu-id="07787-367">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="07787-368">使用 [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) 方法设置结束时间时，应使用 [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) 方法将客户端的本地时间转换为服务器的 UTC。</span><span class="sxs-lookup"><span data-stu-id="07787-368">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="07787-369">以下示例使用 `Time` 对象的 [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) 方法设置约会的结束时间。</span><span class="sxs-lookup"><span data-stu-id="07787-369">The following example sets the end time of an appointment by using the [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="07787-370">类型</span><span class="sxs-lookup"><span data-stu-id="07787-370">Type</span></span>

*   <span data-ttu-id="07787-371">Date | [Time](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="07787-371">Date | [Time](/javascript/api/outlook/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="07787-372">要求</span><span class="sxs-lookup"><span data-stu-id="07787-372">Requirements</span></span>

|<span data-ttu-id="07787-373">要求</span><span class="sxs-lookup"><span data-stu-id="07787-373">Requirement</span></span>|<span data-ttu-id="07787-374">值</span><span class="sxs-lookup"><span data-stu-id="07787-374">Value</span></span>|
|---|---|
|[<span data-ttu-id="07787-375">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="07787-375">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="07787-376">1.0</span><span class="sxs-lookup"><span data-stu-id="07787-376">1.0</span></span>|
|[<span data-ttu-id="07787-377">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="07787-377">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="07787-378">ReadItem</span><span class="sxs-lookup"><span data-stu-id="07787-378">ReadItem</span></span>|
|[<span data-ttu-id="07787-379">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="07787-379">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="07787-380">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="07787-380">Compose or Read</span></span>|

---
---

#### <a name="enhancedlocation-enhancedlocationjavascriptapioutlookofficeenhancedlocation"></a><span data-ttu-id="07787-381">enhancedLocation: [enhancedLocation](/javascript/api/outlook/office.enhancedlocation)</span><span class="sxs-lookup"><span data-stu-id="07787-381">enhancedLocation: [EnhancedLocation](/javascript/api/outlook/office.enhancedlocation)</span></span>

<span data-ttu-id="07787-382">获取或设置约会的位置。</span><span class="sxs-lookup"><span data-stu-id="07787-382">Gets or sets the locations of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="07787-383">阅读模式</span><span class="sxs-lookup"><span data-stu-id="07787-383">Read mode</span></span>

<span data-ttu-id="07787-384">该`enhancedLocation`属性返回一个[EnhancedLocation](/javascript/api/outlook/office.enhancedlocation)对象, 该对象允许您获取与约会关联的一组位置 (每个由[LocationDetails](/javascript/api/outlook/office.locationdetails)对象表示)。</span><span class="sxs-lookup"><span data-stu-id="07787-384">The `enhancedLocation` property returns an [EnhancedLocation](/javascript/api/outlook/office.enhancedlocation) object that allows you to get the set of locations (each represented by a [LocationDetails](/javascript/api/outlook/office.locationdetails) object) associated with the appointment.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="07787-385">撰写模式</span><span class="sxs-lookup"><span data-stu-id="07787-385">Compose mode</span></span>

<span data-ttu-id="07787-386">该`enhancedLocation`属性返回一个[EnhancedLocation](/javascript/api/outlook/office.enhancedlocation)对象, 该对象提供用于获取、删除或添加约会位置的方法。</span><span class="sxs-lookup"><span data-stu-id="07787-386">The `enhancedLocation` property returns an [EnhancedLocation](/javascript/api/outlook/office.enhancedlocation) object that provides methods to get, remove, or add locations on an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="07787-387">类型</span><span class="sxs-lookup"><span data-stu-id="07787-387">Type</span></span>

*   [<span data-ttu-id="07787-388">EnhancedLocation</span><span class="sxs-lookup"><span data-stu-id="07787-388">EnhancedLocation</span></span>](/javascript/api/outlook/office.enhancedlocation)

##### <a name="requirements"></a><span data-ttu-id="07787-389">要求</span><span class="sxs-lookup"><span data-stu-id="07787-389">Requirements</span></span>

|<span data-ttu-id="07787-390">要求</span><span class="sxs-lookup"><span data-stu-id="07787-390">Requirement</span></span>|<span data-ttu-id="07787-391">值</span><span class="sxs-lookup"><span data-stu-id="07787-391">Value</span></span>|
|---|---|
|[<span data-ttu-id="07787-392">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="07787-392">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="07787-393">预览</span><span class="sxs-lookup"><span data-stu-id="07787-393">Preview</span></span>|
|[<span data-ttu-id="07787-394">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="07787-394">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="07787-395">ReadItem</span><span class="sxs-lookup"><span data-stu-id="07787-395">ReadItem</span></span>|
|[<span data-ttu-id="07787-396">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="07787-396">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="07787-397">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="07787-397">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="07787-398">示例</span><span class="sxs-lookup"><span data-stu-id="07787-398">Example</span></span>

<span data-ttu-id="07787-399">下面的示例将获取与约会相关联的当前位置。</span><span class="sxs-lookup"><span data-stu-id="07787-399">The following example gets the current locations associated with the appointment.</span></span>

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

#### <a name="from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsfromjavascriptapioutlookofficefrom"></a><span data-ttu-id="07787-400">发件人: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[From](/javascript/api/outlook/office.from)</span><span class="sxs-lookup"><span data-stu-id="07787-400">from: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[From](/javascript/api/outlook/office.from)</span></span>

<span data-ttu-id="07787-401">获取邮件发件人的电子邮件地址。</span><span class="sxs-lookup"><span data-stu-id="07787-401">Gets the email address of the sender of a message.</span></span>

<span data-ttu-id="07787-p112">`from` 和 [`sender`](#sender-emailaddressdetails) 属性表示同一个人，邮件由代理人发送的除外。在此情况下，`from` 属性表示代理程序，而 sender 属性表示代理人。</span><span class="sxs-lookup"><span data-stu-id="07787-p112">The `from` and [`sender`](#sender-emailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="07787-404">`from` 属性中 `EmailAddressDetails` 对象的 `recipientType` 属性为 `undefined`。</span><span class="sxs-lookup"><span data-stu-id="07787-404">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="07787-405">阅读模式</span><span class="sxs-lookup"><span data-stu-id="07787-405">Read mode</span></span>

<span data-ttu-id="07787-406">`from`属性返回一个`EmailAddressDetails`对象。</span><span class="sxs-lookup"><span data-stu-id="07787-406">The `from` property returns an `EmailAddressDetails` object.</span></span>

```javascript
var from = Office.context.mailbox.item.from;
console.log("From " + from);
```

##### <a name="compose-mode"></a><span data-ttu-id="07787-407">撰写模式</span><span class="sxs-lookup"><span data-stu-id="07787-407">Compose mode</span></span>

<span data-ttu-id="07787-408">`from`属性返回一个`From`对象, 该对象提供用于获取 "起始" 值的方法。</span><span class="sxs-lookup"><span data-stu-id="07787-408">The `from` property returns a `From` object that provides a method to get the from value.</span></span>

```javascript
Office.context.mailbox.item.from.getAsync(callback);

function callback(asyncResult) {
  var from = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="07787-409">类型</span><span class="sxs-lookup"><span data-stu-id="07787-409">Type</span></span>

*   <span data-ttu-id="07787-410">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [](/javascript/api/outlook/office.from)</span><span class="sxs-lookup"><span data-stu-id="07787-410">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [From](/javascript/api/outlook/office.from)</span></span>

##### <a name="requirements"></a><span data-ttu-id="07787-411">要求</span><span class="sxs-lookup"><span data-stu-id="07787-411">Requirements</span></span>

|<span data-ttu-id="07787-412">要求</span><span class="sxs-lookup"><span data-stu-id="07787-412">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="07787-413">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="07787-413">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="07787-414">1.0</span><span class="sxs-lookup"><span data-stu-id="07787-414">1.0</span></span>|<span data-ttu-id="07787-415">1.7</span><span class="sxs-lookup"><span data-stu-id="07787-415">1.7</span></span>|
|[<span data-ttu-id="07787-416">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="07787-416">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="07787-417">ReadItem</span><span class="sxs-lookup"><span data-stu-id="07787-417">ReadItem</span></span>|<span data-ttu-id="07787-418">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="07787-418">ReadWriteItem</span></span>|
|[<span data-ttu-id="07787-419">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="07787-419">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="07787-420">阅读</span><span class="sxs-lookup"><span data-stu-id="07787-420">Read</span></span>|<span data-ttu-id="07787-421">撰写</span><span class="sxs-lookup"><span data-stu-id="07787-421">Compose</span></span>|

---
---

#### <a name="internetheaders-internetheadersjavascriptapioutlookofficeinternetheaders"></a><span data-ttu-id="07787-422">internetHeaders: [internetHeaders](/javascript/api/outlook/office.internetheaders)</span><span class="sxs-lookup"><span data-stu-id="07787-422">internetHeaders: [InternetHeaders](/javascript/api/outlook/office.internetheaders)</span></span>

<span data-ttu-id="07787-423">获取或设置邮件的 internet 邮件头。</span><span class="sxs-lookup"><span data-stu-id="07787-423">Gets or sets the internet headers of a message.</span></span>

##### <a name="type"></a><span data-ttu-id="07787-424">类型</span><span class="sxs-lookup"><span data-stu-id="07787-424">Type</span></span>

*   [<span data-ttu-id="07787-425">InternetHeaders</span><span class="sxs-lookup"><span data-stu-id="07787-425">InternetHeaders</span></span>](/javascript/api/outlook/office.internetheaders)

##### <a name="requirements"></a><span data-ttu-id="07787-426">要求</span><span class="sxs-lookup"><span data-stu-id="07787-426">Requirements</span></span>

|<span data-ttu-id="07787-427">要求</span><span class="sxs-lookup"><span data-stu-id="07787-427">Requirement</span></span>|<span data-ttu-id="07787-428">值</span><span class="sxs-lookup"><span data-stu-id="07787-428">Value</span></span>|
|---|---|
|[<span data-ttu-id="07787-429">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="07787-429">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="07787-430">预览</span><span class="sxs-lookup"><span data-stu-id="07787-430">Preview</span></span>|
|[<span data-ttu-id="07787-431">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="07787-431">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="07787-432">ReadItem</span><span class="sxs-lookup"><span data-stu-id="07787-432">ReadItem</span></span>|
|[<span data-ttu-id="07787-433">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="07787-433">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="07787-434">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="07787-434">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="07787-435">示例</span><span class="sxs-lookup"><span data-stu-id="07787-435">Example</span></span>

```javascript
Office.context.mailbox.item.internetHeaders.getAsync(["header1", "header2"], callback);

function callback(asyncResult) {
  var dictionary = asyncResult.value;
  var header1_value = dictionary["header1"];
}
```

---
---

#### <a name="internetmessageid-string"></a><span data-ttu-id="07787-436">internetMessageId: String</span><span class="sxs-lookup"><span data-stu-id="07787-436">internetMessageId: String</span></span>

<span data-ttu-id="07787-p113">获取电子邮件的 Internet 消息标识符。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="07787-p113">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="07787-439">Type</span><span class="sxs-lookup"><span data-stu-id="07787-439">Type</span></span>

*   <span data-ttu-id="07787-440">String</span><span class="sxs-lookup"><span data-stu-id="07787-440">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="07787-441">要求</span><span class="sxs-lookup"><span data-stu-id="07787-441">Requirements</span></span>

|<span data-ttu-id="07787-442">要求</span><span class="sxs-lookup"><span data-stu-id="07787-442">Requirement</span></span>|<span data-ttu-id="07787-443">值</span><span class="sxs-lookup"><span data-stu-id="07787-443">Value</span></span>|
|---|---|
|[<span data-ttu-id="07787-444">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="07787-444">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="07787-445">1.0</span><span class="sxs-lookup"><span data-stu-id="07787-445">1.0</span></span>|
|[<span data-ttu-id="07787-446">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="07787-446">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="07787-447">ReadItem</span><span class="sxs-lookup"><span data-stu-id="07787-447">ReadItem</span></span>|
|[<span data-ttu-id="07787-448">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="07787-448">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="07787-449">阅读</span><span class="sxs-lookup"><span data-stu-id="07787-449">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="07787-450">示例</span><span class="sxs-lookup"><span data-stu-id="07787-450">Example</span></span>

```javascript
var internetMessageId = Office.context.mailbox.item.internetMessageId;
console.log("internetMessageId: " + internetMessageId);
```

---
---

#### <a name="itemclass-string"></a><span data-ttu-id="07787-451">itemClass: String</span><span class="sxs-lookup"><span data-stu-id="07787-451">itemClass: String</span></span>

<span data-ttu-id="07787-p114">获取选定项目的 Exchange Web 服务项目类。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="07787-p114">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="07787-p115">`itemClass` 属性指定所选项目的邮件类别。以下是邮件或约会项目的默认邮件类别。</span><span class="sxs-lookup"><span data-stu-id="07787-p115">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

|<span data-ttu-id="07787-456">类型</span><span class="sxs-lookup"><span data-stu-id="07787-456">Type</span></span>|<span data-ttu-id="07787-457">说明</span><span class="sxs-lookup"><span data-stu-id="07787-457">Description</span></span>|<span data-ttu-id="07787-458">项目类</span><span class="sxs-lookup"><span data-stu-id="07787-458">item class</span></span>|
|---|---|---|
|<span data-ttu-id="07787-459">约会项目</span><span class="sxs-lookup"><span data-stu-id="07787-459">Appointment items</span></span>|<span data-ttu-id="07787-460">这些是项目类 `IPM.Appointment` 或 `IPM.Appointment.Occurrence` 的日历项目。</span><span class="sxs-lookup"><span data-stu-id="07787-460">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurrence`.</span></span>|`IPM.Appointment`<br />`IPM.Appointment.Occurrence`|
|<span data-ttu-id="07787-461">邮件项目</span><span class="sxs-lookup"><span data-stu-id="07787-461">Message items</span></span>|<span data-ttu-id="07787-462">这些项目包括具有默认邮件类别 `IPM.Note` 的电子邮件，以及使用 `IPM.Schedule.Meeting` 作为基础邮件类别的会议请求、响应和取消。</span><span class="sxs-lookup"><span data-stu-id="07787-462">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span>|`IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled`|

<span data-ttu-id="07787-463">你可以创建用于扩展默认邮件类别的自定义邮件类别，例如，自定义约会邮件类别 `IPM.Appointment.Contoso`。</span><span class="sxs-lookup"><span data-stu-id="07787-463">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="07787-464">类型</span><span class="sxs-lookup"><span data-stu-id="07787-464">Type</span></span>

*   <span data-ttu-id="07787-465">String</span><span class="sxs-lookup"><span data-stu-id="07787-465">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="07787-466">要求</span><span class="sxs-lookup"><span data-stu-id="07787-466">Requirements</span></span>

|<span data-ttu-id="07787-467">要求</span><span class="sxs-lookup"><span data-stu-id="07787-467">Requirement</span></span>|<span data-ttu-id="07787-468">值</span><span class="sxs-lookup"><span data-stu-id="07787-468">Value</span></span>|
|---|---|
|[<span data-ttu-id="07787-469">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="07787-469">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="07787-470">1.0</span><span class="sxs-lookup"><span data-stu-id="07787-470">1.0</span></span>|
|[<span data-ttu-id="07787-471">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="07787-471">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="07787-472">ReadItem</span><span class="sxs-lookup"><span data-stu-id="07787-472">ReadItem</span></span>|
|[<span data-ttu-id="07787-473">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="07787-473">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="07787-474">阅读</span><span class="sxs-lookup"><span data-stu-id="07787-474">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="07787-475">示例</span><span class="sxs-lookup"><span data-stu-id="07787-475">Example</span></span>

```javascript
var itemClass = Office.context.mailbox.item.itemClass;
console.log("Item class: " + itemClass);
```

---
---

#### <a name="nullable-itemid-string"></a><span data-ttu-id="07787-476">(可以为 null) itemId: String</span><span class="sxs-lookup"><span data-stu-id="07787-476">(nullable) itemId: String</span></span>

<span data-ttu-id="07787-p116">获取当前项目的 Exchange Web 服务项目标识符。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="07787-p116">Gets the Exchange Web Services item identifier for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="07787-479">`itemId` 属性返回的标识符与 Exchange Web 服务项目标识符相同。</span><span class="sxs-lookup"><span data-stu-id="07787-479">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="07787-480">`itemId` 属性与 Outlook 条目 ID 或 Outlook REST API 使用的 ID 不同。</span><span class="sxs-lookup"><span data-stu-id="07787-480">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="07787-481">使用此值进行 REST API 调用前，应使用 [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string) 对它进行转换。</span><span class="sxs-lookup"><span data-stu-id="07787-481">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="07787-482">有关详细信息，请参阅[从 Outlook 加载项使用 Outlook REST API](/outlook/add-ins/use-rest-api#get-the-item-id)。</span><span class="sxs-lookup"><span data-stu-id="07787-482">For more details, see [Use the Outlook REST APIs from an Outlook add-in](/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

<span data-ttu-id="07787-p118">`itemId` 属性在撰写模式下不可用。如果需要项目标识符，[`saveAsync`](#saveasyncoptions-callback) 方法可用于将项目保存到存储，这将在回调函数的 [`AsyncResult.value`](/javascript/api/office/office.asyncresult) 参数中返回项目标识符。</span><span class="sxs-lookup"><span data-stu-id="07787-p118">The `itemId` property is not available in compose mode. If an item identifier is required, the [`saveAsync`](#saveasyncoptions-callback) method can be used to save the item to the store, which will return the item identifier in the [`AsyncResult.value`](/javascript/api/office/office.asyncresult) parameter in the callback function.</span></span>

##### <a name="type"></a><span data-ttu-id="07787-485">Type</span><span class="sxs-lookup"><span data-stu-id="07787-485">Type</span></span>

*   <span data-ttu-id="07787-486">String</span><span class="sxs-lookup"><span data-stu-id="07787-486">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="07787-487">要求</span><span class="sxs-lookup"><span data-stu-id="07787-487">Requirements</span></span>

|<span data-ttu-id="07787-488">要求</span><span class="sxs-lookup"><span data-stu-id="07787-488">Requirement</span></span>|<span data-ttu-id="07787-489">值</span><span class="sxs-lookup"><span data-stu-id="07787-489">Value</span></span>|
|---|---|
|[<span data-ttu-id="07787-490">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="07787-490">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="07787-491">1.0</span><span class="sxs-lookup"><span data-stu-id="07787-491">1.0</span></span>|
|[<span data-ttu-id="07787-492">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="07787-492">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="07787-493">ReadItem</span><span class="sxs-lookup"><span data-stu-id="07787-493">ReadItem</span></span>|
|[<span data-ttu-id="07787-494">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="07787-494">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="07787-495">阅读</span><span class="sxs-lookup"><span data-stu-id="07787-495">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="07787-496">示例</span><span class="sxs-lookup"><span data-stu-id="07787-496">Example</span></span>

<span data-ttu-id="07787-p119">以下代码检查项目标识符是否存在。如果 `itemId` 属性返回 `null` 或 `undefined`，则将项目保存到存储，并从异步结果中获取项目标识符。</span><span class="sxs-lookup"><span data-stu-id="07787-p119">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

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

#### <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtype"></a><span data-ttu-id="07787-499">itemType: [MailboxEnums](/javascript/api/outlook/office.mailboxenums.itemtype)</span><span class="sxs-lookup"><span data-stu-id="07787-499">itemType: [Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype)</span></span>

<span data-ttu-id="07787-500">获取实例表示的项的类型。</span><span class="sxs-lookup"><span data-stu-id="07787-500">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="07787-501">`itemType` 属性返回其中一个 `ItemType` 枚举值，指示 `item` 对象实例是邮件还是约会。</span><span class="sxs-lookup"><span data-stu-id="07787-501">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="07787-502">类型</span><span class="sxs-lookup"><span data-stu-id="07787-502">Type</span></span>

*   [<span data-ttu-id="07787-503">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="07787-503">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook/office.mailboxenums.itemtype)

##### <a name="requirements"></a><span data-ttu-id="07787-504">要求</span><span class="sxs-lookup"><span data-stu-id="07787-504">Requirements</span></span>

|<span data-ttu-id="07787-505">要求</span><span class="sxs-lookup"><span data-stu-id="07787-505">Requirement</span></span>|<span data-ttu-id="07787-506">值</span><span class="sxs-lookup"><span data-stu-id="07787-506">Value</span></span>|
|---|---|
|[<span data-ttu-id="07787-507">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="07787-507">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="07787-508">1.0</span><span class="sxs-lookup"><span data-stu-id="07787-508">1.0</span></span>|
|[<span data-ttu-id="07787-509">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="07787-509">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="07787-510">ReadItem</span><span class="sxs-lookup"><span data-stu-id="07787-510">ReadItem</span></span>|
|[<span data-ttu-id="07787-511">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="07787-511">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="07787-512">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="07787-512">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="07787-513">示例</span><span class="sxs-lookup"><span data-stu-id="07787-513">Example</span></span>

```javascript
if (Office.context.mailbox.item.itemType === Office.MailboxEnums.ItemType.Message) {
  // Do something.
} else {
  // Do something else.
}
```

---
---

#### <a name="location-stringlocationjavascriptapioutlookofficelocation"></a><span data-ttu-id="07787-514">位置: 字符串 |[位置](/javascript/api/outlook/office.location)</span><span class="sxs-lookup"><span data-stu-id="07787-514">location: String|[Location](/javascript/api/outlook/office.location)</span></span>

<span data-ttu-id="07787-515">获取或设置约会的位置。</span><span class="sxs-lookup"><span data-stu-id="07787-515">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="07787-516">阅读模式</span><span class="sxs-lookup"><span data-stu-id="07787-516">Read mode</span></span>

<span data-ttu-id="07787-517">`location` 属性返回一个包含约会位置的字符串。</span><span class="sxs-lookup"><span data-stu-id="07787-517">The `location` property returns a string that contains the location of the appointment.</span></span>

```javascript
var location = Office.context.mailbox.item.location;
console.log("location: " + location);
```

##### <a name="compose-mode"></a><span data-ttu-id="07787-518">撰写模式</span><span class="sxs-lookup"><span data-stu-id="07787-518">Compose mode</span></span>

<span data-ttu-id="07787-519">`location` 属性返回一个 `Location` 对象，该对象提供用于获取和设置约会位置的方法。</span><span class="sxs-lookup"><span data-stu-id="07787-519">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

```javascript
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="07787-520">类型</span><span class="sxs-lookup"><span data-stu-id="07787-520">Type</span></span>

*   <span data-ttu-id="07787-521">String | [Location](/javascript/api/outlook/office.location)</span><span class="sxs-lookup"><span data-stu-id="07787-521">String | [Location](/javascript/api/outlook/office.location)</span></span>

##### <a name="requirements"></a><span data-ttu-id="07787-522">要求</span><span class="sxs-lookup"><span data-stu-id="07787-522">Requirements</span></span>

|<span data-ttu-id="07787-523">要求</span><span class="sxs-lookup"><span data-stu-id="07787-523">Requirement</span></span>|<span data-ttu-id="07787-524">值</span><span class="sxs-lookup"><span data-stu-id="07787-524">Value</span></span>|
|---|---|
|[<span data-ttu-id="07787-525">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="07787-525">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="07787-526">1.0</span><span class="sxs-lookup"><span data-stu-id="07787-526">1.0</span></span>|
|[<span data-ttu-id="07787-527">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="07787-527">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="07787-528">ReadItem</span><span class="sxs-lookup"><span data-stu-id="07787-528">ReadItem</span></span>|
|[<span data-ttu-id="07787-529">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="07787-529">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="07787-530">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="07787-530">Compose or Read</span></span>|

---
---

#### <a name="normalizedsubject-string"></a><span data-ttu-id="07787-531">normalizedSubject: String</span><span class="sxs-lookup"><span data-stu-id="07787-531">normalizedSubject: String</span></span>

<span data-ttu-id="07787-p120">获取删除了所有前缀（包括 `RE:` 和 `FWD:`）的项目主题。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="07787-p120">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="07787-p121">normalizedSubject 属性获取包含由电子邮件程序添加的任何标准前缀（如 `RE:` 和 `FW:`）的项目主题。若要获取包含完整前缀的项目主题，请使用 [`subject`](#subject-stringsubject) 属性。</span><span class="sxs-lookup"><span data-stu-id="07787-p121">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="07787-536">Type</span><span class="sxs-lookup"><span data-stu-id="07787-536">Type</span></span>

*   <span data-ttu-id="07787-537">String</span><span class="sxs-lookup"><span data-stu-id="07787-537">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="07787-538">要求</span><span class="sxs-lookup"><span data-stu-id="07787-538">Requirements</span></span>

|<span data-ttu-id="07787-539">要求</span><span class="sxs-lookup"><span data-stu-id="07787-539">Requirement</span></span>|<span data-ttu-id="07787-540">值</span><span class="sxs-lookup"><span data-stu-id="07787-540">Value</span></span>|
|---|---|
|[<span data-ttu-id="07787-541">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="07787-541">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="07787-542">1.0</span><span class="sxs-lookup"><span data-stu-id="07787-542">1.0</span></span>|
|[<span data-ttu-id="07787-543">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="07787-543">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="07787-544">ReadItem</span><span class="sxs-lookup"><span data-stu-id="07787-544">ReadItem</span></span>|
|[<span data-ttu-id="07787-545">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="07787-545">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="07787-546">阅读</span><span class="sxs-lookup"><span data-stu-id="07787-546">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="07787-547">示例</span><span class="sxs-lookup"><span data-stu-id="07787-547">Example</span></span>

```javascript
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
console.log("Normalized subject: " + normalizedSubject);
```

---
---

#### <a name="notificationmessages-notificationmessagesjavascriptapioutlookofficenotificationmessages"></a><span data-ttu-id="07787-548">notificationMessages: [notificationMessages](/javascript/api/outlook/office.notificationmessages)</span><span class="sxs-lookup"><span data-stu-id="07787-548">notificationMessages: [NotificationMessages](/javascript/api/outlook/office.notificationmessages)</span></span>

<span data-ttu-id="07787-549">获取项目的通知邮件。</span><span class="sxs-lookup"><span data-stu-id="07787-549">Gets the notification messages for an item.</span></span>

##### <a name="type"></a><span data-ttu-id="07787-550">类型</span><span class="sxs-lookup"><span data-stu-id="07787-550">Type</span></span>

*   [<span data-ttu-id="07787-551">NotificationMessages</span><span class="sxs-lookup"><span data-stu-id="07787-551">NotificationMessages</span></span>](/javascript/api/outlook/office.notificationmessages)

##### <a name="requirements"></a><span data-ttu-id="07787-552">要求</span><span class="sxs-lookup"><span data-stu-id="07787-552">Requirements</span></span>

|<span data-ttu-id="07787-553">要求</span><span class="sxs-lookup"><span data-stu-id="07787-553">Requirement</span></span>|<span data-ttu-id="07787-554">值</span><span class="sxs-lookup"><span data-stu-id="07787-554">Value</span></span>|
|---|---|
|[<span data-ttu-id="07787-555">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="07787-555">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="07787-556">1.3</span><span class="sxs-lookup"><span data-stu-id="07787-556">1.3</span></span>|
|[<span data-ttu-id="07787-557">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="07787-557">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="07787-558">ReadItem</span><span class="sxs-lookup"><span data-stu-id="07787-558">ReadItem</span></span>|
|[<span data-ttu-id="07787-559">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="07787-559">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="07787-560">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="07787-560">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="07787-561">示例</span><span class="sxs-lookup"><span data-stu-id="07787-561">Example</span></span>

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

#### <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="07787-562">optionalAttendees: <[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[收件人](/javascript/api/outlook/office.recipients)的数组</span><span class="sxs-lookup"><span data-stu-id="07787-562">optionalAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="07787-563">提供对事件的可选与会者的访问权限。</span><span class="sxs-lookup"><span data-stu-id="07787-563">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="07787-564">对象的类型和访问级别取决于当前项目的模式。</span><span class="sxs-lookup"><span data-stu-id="07787-564">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="07787-565">阅读模式</span><span class="sxs-lookup"><span data-stu-id="07787-565">Read mode</span></span>

<span data-ttu-id="07787-566">`optionalAttendees` 属性返回一个数组，其中包含每个可选与会者的 `EmailAddressDetails` 对象。</span><span class="sxs-lookup"><span data-stu-id="07787-566">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span>

```javascript
var optionalAttendees = Office.context.mailbox.item.optionalAttendees;
console.log("Optional attendees: " + JSON.stringify(optionalAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="07787-567">撰写模式</span><span class="sxs-lookup"><span data-stu-id="07787-567">Compose mode</span></span>

<span data-ttu-id="07787-568">`optionalAttendees` 属性返回一个 `Recipients` 对象，该对象提供用于获取或更新可选与会者的方法。</span><span class="sxs-lookup"><span data-stu-id="07787-568">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span>

```javascript
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="07787-569">类型</span><span class="sxs-lookup"><span data-stu-id="07787-569">Type</span></span>

*   <span data-ttu-id="07787-570">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="07787-570">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="07787-571">要求</span><span class="sxs-lookup"><span data-stu-id="07787-571">Requirements</span></span>

|<span data-ttu-id="07787-572">要求</span><span class="sxs-lookup"><span data-stu-id="07787-572">Requirement</span></span>|<span data-ttu-id="07787-573">值</span><span class="sxs-lookup"><span data-stu-id="07787-573">Value</span></span>|
|---|---|
|[<span data-ttu-id="07787-574">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="07787-574">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="07787-575">1.0</span><span class="sxs-lookup"><span data-stu-id="07787-575">1.0</span></span>|
|[<span data-ttu-id="07787-576">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="07787-576">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="07787-577">ReadItem</span><span class="sxs-lookup"><span data-stu-id="07787-577">ReadItem</span></span>|
|[<span data-ttu-id="07787-578">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="07787-578">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="07787-579">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="07787-579">Compose or Read</span></span>|

---
---

#### <a name="organizer-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsorganizerjavascriptapioutlookofficeorganizer"></a><span data-ttu-id="07787-580">组织者: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[组织者](/javascript/api/outlook/office.organizer)</span><span class="sxs-lookup"><span data-stu-id="07787-580">organizer: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[Organizer](/javascript/api/outlook/office.organizer)</span></span>

<span data-ttu-id="07787-581">获取指定会议的组织者的电子邮件地址。</span><span class="sxs-lookup"><span data-stu-id="07787-581">Gets the email address of the organizer for a specified meeting.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="07787-582">阅读模式</span><span class="sxs-lookup"><span data-stu-id="07787-582">Read mode</span></span>

<span data-ttu-id="07787-583">该`organizer`属性返回一个[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)对象, 该对象代表会议组织者。</span><span class="sxs-lookup"><span data-stu-id="07787-583">The `organizer` property returns an [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) object that represents the meeting organizer.</span></span>

```javascript
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
console.log("Organizer: " + organizerName + " (" + organizerAddress + ")");
```

##### <a name="compose-mode"></a><span data-ttu-id="07787-584">撰写模式</span><span class="sxs-lookup"><span data-stu-id="07787-584">Compose mode</span></span>

<span data-ttu-id="07787-585">该`organizer`属性返回一个[管理](/javascript/api/outlook/office.organizer)器对象, 该对象提供获取组织者值的方法。</span><span class="sxs-lookup"><span data-stu-id="07787-585">The `organizer` property returns an [Organizer](/javascript/api/outlook/office.organizer) object that provides a method to get the organizer value.</span></span>

```javascript
Office.context.mailbox.item.organizer.getAsync(
  function(asyncResult) {
    console.log(JSON.stringify(asyncResult));
  }
);
```

##### <a name="type"></a><span data-ttu-id="07787-586">类型</span><span class="sxs-lookup"><span data-stu-id="07787-586">Type</span></span>

*   <span data-ttu-id="07787-587">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [组织者](/javascript/api/outlook/office.organizer)</span><span class="sxs-lookup"><span data-stu-id="07787-587">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [Organizer](/javascript/api/outlook/office.organizer)</span></span>

##### <a name="requirements"></a><span data-ttu-id="07787-588">要求</span><span class="sxs-lookup"><span data-stu-id="07787-588">Requirements</span></span>

|<span data-ttu-id="07787-589">要求</span><span class="sxs-lookup"><span data-stu-id="07787-589">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="07787-590">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="07787-590">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="07787-591">1.0</span><span class="sxs-lookup"><span data-stu-id="07787-591">1.0</span></span>|<span data-ttu-id="07787-592">1.7</span><span class="sxs-lookup"><span data-stu-id="07787-592">1.7</span></span>|
|[<span data-ttu-id="07787-593">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="07787-593">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="07787-594">ReadItem</span><span class="sxs-lookup"><span data-stu-id="07787-594">ReadItem</span></span>|<span data-ttu-id="07787-595">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="07787-595">ReadWriteItem</span></span>|
|[<span data-ttu-id="07787-596">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="07787-596">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="07787-597">阅读</span><span class="sxs-lookup"><span data-stu-id="07787-597">Read</span></span>|<span data-ttu-id="07787-598">撰写</span><span class="sxs-lookup"><span data-stu-id="07787-598">Compose</span></span>|

---
---

#### <a name="nullable-recurrence-recurrencejavascriptapioutlookofficerecurrence"></a><span data-ttu-id="07787-599">(可以为 null) 定期:[定期](/javascript/api/outlook/office.recurrence)</span><span class="sxs-lookup"><span data-stu-id="07787-599">(nullable) recurrence: [Recurrence](/javascript/api/outlook/office.recurrence)</span></span>

<span data-ttu-id="07787-600">获取或设置约会的定期模式。</span><span class="sxs-lookup"><span data-stu-id="07787-600">Gets or sets the recurrence pattern of an appointment.</span></span> <span data-ttu-id="07787-601">获取会议请求的定期模式。</span><span class="sxs-lookup"><span data-stu-id="07787-601">Gets the recurrence pattern of a meeting request.</span></span> <span data-ttu-id="07787-602">约会项目的阅读和撰写模式。</span><span class="sxs-lookup"><span data-stu-id="07787-602">Read and compose modes for appointment items.</span></span> <span data-ttu-id="07787-603">会议请求项目的阅读模式。</span><span class="sxs-lookup"><span data-stu-id="07787-603">Read mode for meeting request items.</span></span>

<span data-ttu-id="07787-604">如果`recurrence`项目是系列中的一个系列或一个实例, 则该属性返回定期约会或会议请求的[定期](/javascript/api/outlook/office.recurrence)对象。</span><span class="sxs-lookup"><span data-stu-id="07787-604">The `recurrence` property returns a [recurrence](/javascript/api/outlook/office.recurrence) object for recurring appointments or meetings requests if an item is a series or an instance in a series.</span></span> <span data-ttu-id="07787-605">`null`返回单个约会的单个约会和会议请求。</span><span class="sxs-lookup"><span data-stu-id="07787-605">`null` is returned for single appointments and meeting requests of single appointments.</span></span> <span data-ttu-id="07787-606">`undefined`对于不是会议请求的邮件, 将返回。</span><span class="sxs-lookup"><span data-stu-id="07787-606">`undefined` is returned for messages that are not meeting requests.</span></span>

> <span data-ttu-id="07787-607">注意: 会议请求的`itemClass`值为 IPM。Schedule. 会议请求。</span><span class="sxs-lookup"><span data-stu-id="07787-607">Note: Meeting requests have an `itemClass` value of IPM.Schedule.Meeting.Request.</span></span>

> <span data-ttu-id="07787-608">注意: 如果定期对象为`null`, 则表示该对象是单个约会的单个约会或会议请求, 而不是某个系列的一部分。</span><span class="sxs-lookup"><span data-stu-id="07787-608">Note: If the recurrence object is `null`, this indicates that the object is a single appointment or a meeting request of a single appointment and NOT a part of a series.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="07787-609">阅读模式</span><span class="sxs-lookup"><span data-stu-id="07787-609">Read mode</span></span>

<span data-ttu-id="07787-610">该`recurrence`属性返回一个代表约会定期的[定期](/javascript/api/outlook/office.recurrence)对象。</span><span class="sxs-lookup"><span data-stu-id="07787-610">The `recurrence` property returns a [Recurrence](/javascript/api/outlook/office.recurrence) object that represents the appointment recurrence.</span></span> <span data-ttu-id="07787-611">此功能适用于约会和会议请求。</span><span class="sxs-lookup"><span data-stu-id="07787-611">This is available for appointments and meeting requests.</span></span>

```javascript
var recurrence = Office.context.mailbox.item.recurrence;
console.log("Recurrence: " + JSON.stringify(recurrence));
```

##### <a name="compose-mode"></a><span data-ttu-id="07787-612">撰写模式</span><span class="sxs-lookup"><span data-stu-id="07787-612">Compose mode</span></span>

<span data-ttu-id="07787-613">该`recurrence`属性返回一个[定期](/javascript/api/outlook/office.recurrence)对象, 该对象提供用于管理约会周期的方法。</span><span class="sxs-lookup"><span data-stu-id="07787-613">The `recurrence` property returns a [Recurrence](/javascript/api/outlook/office.recurrence) object that provides methods to manage the appointment recurrence.</span></span> <span data-ttu-id="07787-614">这可用于约会。</span><span class="sxs-lookup"><span data-stu-id="07787-614">This is available for appointments.</span></span>

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

##### <a name="type"></a><span data-ttu-id="07787-615">类型</span><span class="sxs-lookup"><span data-stu-id="07787-615">Type</span></span>

* [<span data-ttu-id="07787-616">循环</span><span class="sxs-lookup"><span data-stu-id="07787-616">Recurrence</span></span>](/javascript/api/outlook/office.recurrence)

|<span data-ttu-id="07787-617">要求</span><span class="sxs-lookup"><span data-stu-id="07787-617">Requirement</span></span>|<span data-ttu-id="07787-618">值</span><span class="sxs-lookup"><span data-stu-id="07787-618">Value</span></span>|
|---|---|
|[<span data-ttu-id="07787-619">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="07787-619">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="07787-620">1.7</span><span class="sxs-lookup"><span data-stu-id="07787-620">1.7</span></span>|
|[<span data-ttu-id="07787-621">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="07787-621">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="07787-622">ReadItem</span><span class="sxs-lookup"><span data-stu-id="07787-622">ReadItem</span></span>|
|[<span data-ttu-id="07787-623">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="07787-623">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="07787-624">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="07787-624">Compose or Read</span></span>|

---
---

#### <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="07787-625">requiredAttendees: <[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[收件人](/javascript/api/outlook/office.recipients)的数组</span><span class="sxs-lookup"><span data-stu-id="07787-625">requiredAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="07787-626">提供对事件的必需与会者的访问权限。</span><span class="sxs-lookup"><span data-stu-id="07787-626">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="07787-627">对象的类型和访问级别取决于当前项目的模式。</span><span class="sxs-lookup"><span data-stu-id="07787-627">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="07787-628">阅读模式</span><span class="sxs-lookup"><span data-stu-id="07787-628">Read mode</span></span>

<span data-ttu-id="07787-629">`requiredAttendees` 属性返回一个数组，其中包含每个必需与会者的 `EmailAddressDetails` 对象。</span><span class="sxs-lookup"><span data-stu-id="07787-629">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span>

```javascript
var requiredAttendees = Office.context.mailbox.item.requiredAttendees;
console.log("Required attendees: " + JSON.stringify(requiredAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="07787-630">撰写模式</span><span class="sxs-lookup"><span data-stu-id="07787-630">Compose mode</span></span>

<span data-ttu-id="07787-631">`requiredAttendees` 属性返回一个 `Recipients` 对象，该对象提供用于获取或更新必需与会者的方法。</span><span class="sxs-lookup"><span data-stu-id="07787-631">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span>

```javascript
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
  console.log(JSON.stringify(arrayOfRequiredAttendeesRecipients));
}
```

##### <a name="type"></a><span data-ttu-id="07787-632">类型</span><span class="sxs-lookup"><span data-stu-id="07787-632">Type</span></span>

*   <span data-ttu-id="07787-633">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="07787-633">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="07787-634">要求</span><span class="sxs-lookup"><span data-stu-id="07787-634">Requirements</span></span>

|<span data-ttu-id="07787-635">要求</span><span class="sxs-lookup"><span data-stu-id="07787-635">Requirement</span></span>|<span data-ttu-id="07787-636">值</span><span class="sxs-lookup"><span data-stu-id="07787-636">Value</span></span>|
|---|---|
|[<span data-ttu-id="07787-637">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="07787-637">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="07787-638">1.0</span><span class="sxs-lookup"><span data-stu-id="07787-638">1.0</span></span>|
|[<span data-ttu-id="07787-639">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="07787-639">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="07787-640">ReadItem</span><span class="sxs-lookup"><span data-stu-id="07787-640">ReadItem</span></span>|
|[<span data-ttu-id="07787-641">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="07787-641">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="07787-642">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="07787-642">Compose or Read</span></span>|

---
---

#### <a name="sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetails"></a><span data-ttu-id="07787-643">发件人: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="07787-643">sender: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)</span></span>

<span data-ttu-id="07787-p128">获取电子邮件发件人的电子邮件地址。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="07787-p128">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="07787-p129">[`from`](#from-emailaddressdetailsfrom) 和 `sender` 属性表示同一个人，邮件由代理人发送的除外。在此情况下，`from` 属性表示代理程序，而 sender 属性表示代理人。</span><span class="sxs-lookup"><span data-stu-id="07787-p129">The [`from`](#from-emailaddressdetailsfrom) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="07787-648">`sender` 属性中 `EmailAddressDetails` 对象的 `recipientType` 属性为 `undefined`。</span><span class="sxs-lookup"><span data-stu-id="07787-648">The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="07787-649">类型</span><span class="sxs-lookup"><span data-stu-id="07787-649">Type</span></span>

*   [<span data-ttu-id="07787-650">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="07787-650">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="07787-651">要求</span><span class="sxs-lookup"><span data-stu-id="07787-651">Requirements</span></span>

|<span data-ttu-id="07787-652">要求</span><span class="sxs-lookup"><span data-stu-id="07787-652">Requirement</span></span>|<span data-ttu-id="07787-653">值</span><span class="sxs-lookup"><span data-stu-id="07787-653">Value</span></span>|
|---|---|
|[<span data-ttu-id="07787-654">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="07787-654">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="07787-655">1.0</span><span class="sxs-lookup"><span data-stu-id="07787-655">1.0</span></span>|
|[<span data-ttu-id="07787-656">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="07787-656">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="07787-657">ReadItem</span><span class="sxs-lookup"><span data-stu-id="07787-657">ReadItem</span></span>|
|[<span data-ttu-id="07787-658">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="07787-658">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="07787-659">阅读</span><span class="sxs-lookup"><span data-stu-id="07787-659">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="07787-660">示例</span><span class="sxs-lookup"><span data-stu-id="07787-660">Example</span></span>

```javascript
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
console.log("Sender: " + senderName + " (" + senderAddress + ")");
```

---
---

#### <a name="nullable-seriesid-string"></a><span data-ttu-id="07787-661">(可以为 null) Webcasts&seriesid: String</span><span class="sxs-lookup"><span data-stu-id="07787-661">(nullable) seriesId: String</span></span>

<span data-ttu-id="07787-662">获取实例所属的系列的 id。</span><span class="sxs-lookup"><span data-stu-id="07787-662">Gets the id of the series that an instance belongs to.</span></span>

<span data-ttu-id="07787-663">在 web 上的 Outlook 和桌面客户端中`seriesId` , 返回此项所属的父 (系列) 项的 Exchange web 服务 (EWS) ID。</span><span class="sxs-lookup"><span data-stu-id="07787-663">In Outlook on the web and desktop clients, the `seriesId` returns the Exchange Web Services (EWS) ID of the parent (series) item that this item belongs to.</span></span> <span data-ttu-id="07787-664">但是, 在 iOS 和 Android 中, `seriesId`将返回父项的 REST ID。</span><span class="sxs-lookup"><span data-stu-id="07787-664">However, in iOS and Android, the `seriesId` returns the REST ID of the parent item.</span></span>

> [!NOTE]
> <span data-ttu-id="07787-665">`seriesId` 属性返回的标识符与 Exchange Web 服务项目标识符相同。</span><span class="sxs-lookup"><span data-stu-id="07787-665">The identifier returned by the `seriesId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="07787-666">`seriesId`属性与 OUTLOOK REST API 使用的 outlook id 不相同。</span><span class="sxs-lookup"><span data-stu-id="07787-666">The `seriesId` property is not identical to the Outlook IDs used by the Outlook REST API.</span></span> <span data-ttu-id="07787-667">使用此值进行 REST API 调用前，应使用 [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string) 对它进行转换。</span><span class="sxs-lookup"><span data-stu-id="07787-667">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="07787-668">有关详细信息，请参阅[从 Outlook 加载项使用 Outlook REST API](/outlook/add-ins/use-rest-api)。</span><span class="sxs-lookup"><span data-stu-id="07787-668">For more details, see [Use the Outlook REST APIs from an Outlook add-in](/outlook/add-ins/use-rest-api).</span></span>

<span data-ttu-id="07787-669">对于`seriesId`不包含`null`父项 (如单个约会、系列项或会议请求) 的项, 该属性将返回, `undefined`对于不是会议请求的任何其他项, 该属性返回。</span><span class="sxs-lookup"><span data-stu-id="07787-669">The `seriesId` property returns `null` for items that do not have parent items such as single appointments, series items, or meeting requests and returns `undefined` for any other items that are not meeting requests.</span></span>

##### <a name="type"></a><span data-ttu-id="07787-670">Type</span><span class="sxs-lookup"><span data-stu-id="07787-670">Type</span></span>

* <span data-ttu-id="07787-671">String</span><span class="sxs-lookup"><span data-stu-id="07787-671">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="07787-672">要求</span><span class="sxs-lookup"><span data-stu-id="07787-672">Requirements</span></span>

|<span data-ttu-id="07787-673">要求</span><span class="sxs-lookup"><span data-stu-id="07787-673">Requirement</span></span>|<span data-ttu-id="07787-674">值</span><span class="sxs-lookup"><span data-stu-id="07787-674">Value</span></span>|
|---|---|
|[<span data-ttu-id="07787-675">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="07787-675">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="07787-676">1.7</span><span class="sxs-lookup"><span data-stu-id="07787-676">1.7</span></span>|
|[<span data-ttu-id="07787-677">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="07787-677">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="07787-678">ReadItem</span><span class="sxs-lookup"><span data-stu-id="07787-678">ReadItem</span></span>|
|[<span data-ttu-id="07787-679">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="07787-679">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="07787-680">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="07787-680">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="07787-681">示例</span><span class="sxs-lookup"><span data-stu-id="07787-681">Example</span></span>

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

#### <a name="start-datetimejavascriptapioutlookofficetime"></a><span data-ttu-id="07787-682">开始日期: 日期 |[时间](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="07787-682">start: Date|[Time](/javascript/api/outlook/office.time)</span></span>

<span data-ttu-id="07787-683">获取或设置约会开始的日期和时间。</span><span class="sxs-lookup"><span data-stu-id="07787-683">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="07787-p132">将 `start` 属性表示为协调世界时 (UTC) 的日期和时间值。可使用 [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) 方法将该值转换为客户端的本地日期和时间。</span><span class="sxs-lookup"><span data-stu-id="07787-p132">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="07787-686">阅读模式</span><span class="sxs-lookup"><span data-stu-id="07787-686">Read mode</span></span>

<span data-ttu-id="07787-687">`start` 属性返回 `Date` 对象。</span><span class="sxs-lookup"><span data-stu-id="07787-687">The `start` property returns a `Date` object.</span></span>

```javascript
var start = Office.context.mailbox.item.start;
console.log("Appointment start: " + JSON.stringify(start));
```

##### <a name="compose-mode"></a><span data-ttu-id="07787-688">撰写模式</span><span class="sxs-lookup"><span data-stu-id="07787-688">Compose mode</span></span>

<span data-ttu-id="07787-689">`start` 属性返回 `Time` 对象。</span><span class="sxs-lookup"><span data-stu-id="07787-689">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="07787-690">使用 [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) 方法设置开始时间时，应使用 [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) 方法将客户端的本地时间转换为服务器的 UTC。</span><span class="sxs-lookup"><span data-stu-id="07787-690">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="07787-691">以下示例通过使用 `Time` 对象的 [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) 方法，设置撰写模式下约会的开始时间。</span><span class="sxs-lookup"><span data-stu-id="07787-691">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="07787-692">类型</span><span class="sxs-lookup"><span data-stu-id="07787-692">Type</span></span>

*   <span data-ttu-id="07787-693">Date | [Time](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="07787-693">Date | [Time](/javascript/api/outlook/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="07787-694">要求</span><span class="sxs-lookup"><span data-stu-id="07787-694">Requirements</span></span>

|<span data-ttu-id="07787-695">要求</span><span class="sxs-lookup"><span data-stu-id="07787-695">Requirement</span></span>|<span data-ttu-id="07787-696">值</span><span class="sxs-lookup"><span data-stu-id="07787-696">Value</span></span>|
|---|---|
|[<span data-ttu-id="07787-697">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="07787-697">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="07787-698">1.0</span><span class="sxs-lookup"><span data-stu-id="07787-698">1.0</span></span>|
|[<span data-ttu-id="07787-699">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="07787-699">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="07787-700">ReadItem</span><span class="sxs-lookup"><span data-stu-id="07787-700">ReadItem</span></span>|
|[<span data-ttu-id="07787-701">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="07787-701">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="07787-702">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="07787-702">Compose or Read</span></span>|

---
---

#### <a name="subject-stringsubjectjavascriptapioutlookofficesubject"></a><span data-ttu-id="07787-703">subject: String |[主题](/javascript/api/outlook/office.subject)</span><span class="sxs-lookup"><span data-stu-id="07787-703">subject: String|[Subject](/javascript/api/outlook/office.subject)</span></span>

<span data-ttu-id="07787-704">获取或设置显示在项目的主题字段中的说明。</span><span class="sxs-lookup"><span data-stu-id="07787-704">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="07787-705">`subject` 属性获取或设置由电子邮件服务器发送项目时的整个主题。</span><span class="sxs-lookup"><span data-stu-id="07787-705">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="07787-706">阅读模式</span><span class="sxs-lookup"><span data-stu-id="07787-706">Read mode</span></span>

<span data-ttu-id="07787-p133">`subject` 属性返回一个字符串。使用 [`normalizedSubject`](#normalizedsubject-string) 属性获取不带任何前导前缀（如 `RE:` 和 `FW:`）的主题。</span><span class="sxs-lookup"><span data-stu-id="07787-p133">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

<span data-ttu-id="07787-709">以下 JavaScript 代码示例显示了如何访问 Outlook 中当前项目的 `subject` 属性。</span><span class="sxs-lookup"><span data-stu-id="07787-709">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

```javascript
var subject = Office.context.mailbox.item.subject;
console.log(subject);
```

##### <a name="compose-mode"></a><span data-ttu-id="07787-710">撰写模式</span><span class="sxs-lookup"><span data-stu-id="07787-710">Compose mode</span></span>
<span data-ttu-id="07787-711">`subject` 属性返回一个 `Subject` 对象，该对象提供用于获取和设置主题的方法。</span><span class="sxs-lookup"><span data-stu-id="07787-711">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```javascript
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
  console.log(subject);
}
```

##### <a name="type"></a><span data-ttu-id="07787-712">类型</span><span class="sxs-lookup"><span data-stu-id="07787-712">Type</span></span>

*   <span data-ttu-id="07787-713">String | [Subject](/javascript/api/outlook/office.subject)</span><span class="sxs-lookup"><span data-stu-id="07787-713">String | [Subject](/javascript/api/outlook/office.subject)</span></span>

##### <a name="requirements"></a><span data-ttu-id="07787-714">要求</span><span class="sxs-lookup"><span data-stu-id="07787-714">Requirements</span></span>

|<span data-ttu-id="07787-715">要求</span><span class="sxs-lookup"><span data-stu-id="07787-715">Requirement</span></span>|<span data-ttu-id="07787-716">值</span><span class="sxs-lookup"><span data-stu-id="07787-716">Value</span></span>|
|---|---|
|[<span data-ttu-id="07787-717">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="07787-717">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="07787-718">1.0</span><span class="sxs-lookup"><span data-stu-id="07787-718">1.0</span></span>|
|[<span data-ttu-id="07787-719">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="07787-719">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="07787-720">ReadItem</span><span class="sxs-lookup"><span data-stu-id="07787-720">ReadItem</span></span>|
|[<span data-ttu-id="07787-721">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="07787-721">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="07787-722">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="07787-722">Compose or Read</span></span>|

---
---

#### <a name="to-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="07787-723">to: <[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[收件人](/javascript/api/outlook/office.recipients)的数组</span><span class="sxs-lookup"><span data-stu-id="07787-723">to: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="07787-724">提供对邮件的“**收件人**”行上的收件人的访问权限。</span><span class="sxs-lookup"><span data-stu-id="07787-724">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="07787-725">对象的类型和访问级别取决于当前项目的模式。</span><span class="sxs-lookup"><span data-stu-id="07787-725">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="07787-726">阅读模式</span><span class="sxs-lookup"><span data-stu-id="07787-726">Read mode</span></span>

<span data-ttu-id="07787-p135">`to` 属性返回包含邮件的**收件人**行上所列的每个收件人的 `EmailAddressDetails` 对象的数组。集合上限为 100 个成员。</span><span class="sxs-lookup"><span data-stu-id="07787-p135">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.</span></span>

```javascript
console.log(JSON.stringify(Office.context.mailbox.item.to));
```

##### <a name="compose-mode"></a><span data-ttu-id="07787-729">撰写模式</span><span class="sxs-lookup"><span data-stu-id="07787-729">Compose mode</span></span>

<span data-ttu-id="07787-730">`to` 属性返回一个 `Recipients` 对象，该对象提供用于获取或更新邮件的“**收件人**”行上收件人的方法。</span><span class="sxs-lookup"><span data-stu-id="07787-730">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span>

```javascript
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="07787-731">类型</span><span class="sxs-lookup"><span data-stu-id="07787-731">Type</span></span>

*   <span data-ttu-id="07787-732">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="07787-732">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="07787-733">要求</span><span class="sxs-lookup"><span data-stu-id="07787-733">Requirements</span></span>

|<span data-ttu-id="07787-734">要求</span><span class="sxs-lookup"><span data-stu-id="07787-734">Requirement</span></span>|<span data-ttu-id="07787-735">值</span><span class="sxs-lookup"><span data-stu-id="07787-735">Value</span></span>|
|---|---|
|[<span data-ttu-id="07787-736">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="07787-736">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="07787-737">1.0</span><span class="sxs-lookup"><span data-stu-id="07787-737">1.0</span></span>|
|[<span data-ttu-id="07787-738">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="07787-738">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="07787-739">ReadItem</span><span class="sxs-lookup"><span data-stu-id="07787-739">ReadItem</span></span>|
|[<span data-ttu-id="07787-740">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="07787-740">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="07787-741">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="07787-741">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="07787-742">方法</span><span class="sxs-lookup"><span data-stu-id="07787-742">Methods</span></span>

#### <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="07787-743">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="07787-743">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="07787-744">将文件作为附件添加到邮件或约会。</span><span class="sxs-lookup"><span data-stu-id="07787-744">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="07787-745">`addFileAttachmentAsync` 方法在指定的 URI 上载文件并将其附加到撰写窗体中的项目。</span><span class="sxs-lookup"><span data-stu-id="07787-745">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="07787-746">随后可以将该标识符与 [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) 方法一同使用，以删除同一个会话中的附件。</span><span class="sxs-lookup"><span data-stu-id="07787-746">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="07787-747">参数</span><span class="sxs-lookup"><span data-stu-id="07787-747">Parameters</span></span>
|<span data-ttu-id="07787-748">名称</span><span class="sxs-lookup"><span data-stu-id="07787-748">Name</span></span>|<span data-ttu-id="07787-749">类型</span><span class="sxs-lookup"><span data-stu-id="07787-749">Type</span></span>|<span data-ttu-id="07787-750">属性</span><span class="sxs-lookup"><span data-stu-id="07787-750">Attributes</span></span>|<span data-ttu-id="07787-751">说明</span><span class="sxs-lookup"><span data-stu-id="07787-751">Description</span></span>|
|---|---|---|---|
|`uri`|<span data-ttu-id="07787-752">String</span><span class="sxs-lookup"><span data-stu-id="07787-752">String</span></span>||<span data-ttu-id="07787-p136">提供附加到邮件或约会的文件的位置的 URI。最大长度为 2048 个字符。</span><span class="sxs-lookup"><span data-stu-id="07787-p136">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`|<span data-ttu-id="07787-755">字符串</span><span class="sxs-lookup"><span data-stu-id="07787-755">String</span></span>||<span data-ttu-id="07787-p137">在附件上载过程中显示的附件名称。最大长度为 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="07787-p137">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="07787-758">Object</span><span class="sxs-lookup"><span data-stu-id="07787-758">Object</span></span>|<span data-ttu-id="07787-759">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="07787-759">&lt;optional&gt;</span></span>|<span data-ttu-id="07787-760">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="07787-760">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="07787-761">对象</span><span class="sxs-lookup"><span data-stu-id="07787-761">Object</span></span>|<span data-ttu-id="07787-762">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="07787-762">&lt;optional&gt;</span></span>|<span data-ttu-id="07787-763">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="07787-763">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.isInline`|<span data-ttu-id="07787-764">布尔值</span><span class="sxs-lookup"><span data-stu-id="07787-764">Boolean</span></span>|<span data-ttu-id="07787-765">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="07787-765">&lt;optional&gt;</span></span>|<span data-ttu-id="07787-766">如果为 `true`，则表示附件将在邮件正文中内联显示，并且不应显示在附件列表中。</span><span class="sxs-lookup"><span data-stu-id="07787-766">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`callback`|<span data-ttu-id="07787-767">函数</span><span class="sxs-lookup"><span data-stu-id="07787-767">function</span></span>|<span data-ttu-id="07787-768">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="07787-768">&lt;optional&gt;</span></span>|<span data-ttu-id="07787-769">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="07787-769">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="07787-770">如果成功，附件标识符将在 `asyncResult.value` 属性中提供。</span><span class="sxs-lookup"><span data-stu-id="07787-770">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="07787-771">如果上传附件失败，`asyncResult` 对象将包含一个提供错误说明的 `Error` 对象。</span><span class="sxs-lookup"><span data-stu-id="07787-771">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="07787-772">错误</span><span class="sxs-lookup"><span data-stu-id="07787-772">Errors</span></span>

|<span data-ttu-id="07787-773">错误代码</span><span class="sxs-lookup"><span data-stu-id="07787-773">Error code</span></span>|<span data-ttu-id="07787-774">说明</span><span class="sxs-lookup"><span data-stu-id="07787-774">Description</span></span>|
|------------|-------------|
|`AttachmentSizeExceeded`|<span data-ttu-id="07787-775">附件大小超过了允许的大小。</span><span class="sxs-lookup"><span data-stu-id="07787-775">The attachment is larger than allowed.</span></span>|
|`FileTypeNotSupported`|<span data-ttu-id="07787-776">该附件的扩展名不是允许的扩展名。</span><span class="sxs-lookup"><span data-stu-id="07787-776">The attachment has an extension that is not allowed.</span></span>|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="07787-777">邮件或约会具有的附件过多。</span><span class="sxs-lookup"><span data-stu-id="07787-777">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="07787-778">要求</span><span class="sxs-lookup"><span data-stu-id="07787-778">Requirements</span></span>

|<span data-ttu-id="07787-779">要求</span><span class="sxs-lookup"><span data-stu-id="07787-779">Requirement</span></span>|<span data-ttu-id="07787-780">值</span><span class="sxs-lookup"><span data-stu-id="07787-780">Value</span></span>|
|---|---|
|[<span data-ttu-id="07787-781">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="07787-781">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="07787-782">1.1</span><span class="sxs-lookup"><span data-stu-id="07787-782">1.1</span></span>|
|[<span data-ttu-id="07787-783">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="07787-783">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="07787-784">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="07787-784">ReadWriteItem</span></span>|
|[<span data-ttu-id="07787-785">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="07787-785">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="07787-786">撰写</span><span class="sxs-lookup"><span data-stu-id="07787-786">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="07787-787">示例</span><span class="sxs-lookup"><span data-stu-id="07787-787">Examples</span></span>

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

<span data-ttu-id="07787-788">以下示例将图像文件添加为内联附件，并在邮件正文中引用该附件。</span><span class="sxs-lookup"><span data-stu-id="07787-788">The following example adds an image file as an inline attachment and references the attachment in the message body.</span></span>

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

#### <a name="addfileattachmentfrombase64asyncbase64file-attachmentname-options-callback"></a><span data-ttu-id="07787-789">addFileAttachmentFromBase64Async (base64File, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="07787-789">addFileAttachmentFromBase64Async(base64File, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="07787-790">将 base64 编码中的文件作为附件添加到邮件或约会中。</span><span class="sxs-lookup"><span data-stu-id="07787-790">Adds a file from the base64 encoding to a message or appointment as an attachment.</span></span>

<span data-ttu-id="07787-791">该`addFileAttachmentFromBase64Async`方法从 base64 编码中上载文件, 并将其附加到撰写窗体中的项目。</span><span class="sxs-lookup"><span data-stu-id="07787-791">The `addFileAttachmentFromBase64Async` method uploads the file from the base64 encoding and attaches it to the item in the compose form.</span></span> <span data-ttu-id="07787-792">此方法返回 AsyncResult 对象中的附件标识符。</span><span class="sxs-lookup"><span data-stu-id="07787-792">This method returns the attachment identifier in the AsyncResult.value object.</span></span>

<span data-ttu-id="07787-793">随后可以将该标识符与 [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) 方法一同使用，以删除同一个会话中的附件。</span><span class="sxs-lookup"><span data-stu-id="07787-793">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="07787-794">参数</span><span class="sxs-lookup"><span data-stu-id="07787-794">Parameters</span></span>

|<span data-ttu-id="07787-795">名称</span><span class="sxs-lookup"><span data-stu-id="07787-795">Name</span></span>|<span data-ttu-id="07787-796">类型</span><span class="sxs-lookup"><span data-stu-id="07787-796">Type</span></span>|<span data-ttu-id="07787-797">属性</span><span class="sxs-lookup"><span data-stu-id="07787-797">Attributes</span></span>|<span data-ttu-id="07787-798">说明</span><span class="sxs-lookup"><span data-stu-id="07787-798">Description</span></span>|
|---|---|---|---|
|`base64File`|<span data-ttu-id="07787-799">String</span><span class="sxs-lookup"><span data-stu-id="07787-799">String</span></span>||<span data-ttu-id="07787-800">要添加到电子邮件或事件的图像或文件的 base64 编码内容。</span><span class="sxs-lookup"><span data-stu-id="07787-800">The base64 encoded content of an image or file to be added to an email or event.</span></span>|
|`attachmentName`|<span data-ttu-id="07787-801">String</span><span class="sxs-lookup"><span data-stu-id="07787-801">String</span></span>||<span data-ttu-id="07787-p139">在附件上载过程中显示的附件名称。最大长度为 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="07787-p139">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="07787-804">对象</span><span class="sxs-lookup"><span data-stu-id="07787-804">Object</span></span>|<span data-ttu-id="07787-805">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="07787-805">&lt;optional&gt;</span></span>|<span data-ttu-id="07787-806">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="07787-806">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="07787-807">对象</span><span class="sxs-lookup"><span data-stu-id="07787-807">Object</span></span>|<span data-ttu-id="07787-808">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="07787-808">&lt;optional&gt;</span></span>|<span data-ttu-id="07787-809">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="07787-809">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.isInline`|<span data-ttu-id="07787-810">布尔值</span><span class="sxs-lookup"><span data-stu-id="07787-810">Boolean</span></span>|<span data-ttu-id="07787-811">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="07787-811">&lt;optional&gt;</span></span>|<span data-ttu-id="07787-812">如果为 `true`，则表示附件将在邮件正文中内联显示，并且不应显示在附件列表中。</span><span class="sxs-lookup"><span data-stu-id="07787-812">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`callback`|<span data-ttu-id="07787-813">函数</span><span class="sxs-lookup"><span data-stu-id="07787-813">function</span></span>|<span data-ttu-id="07787-814">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="07787-814">&lt;optional&gt;</span></span>|<span data-ttu-id="07787-815">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="07787-815">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="07787-816">如果成功，附件标识符将在 `asyncResult.value` 属性中提供。</span><span class="sxs-lookup"><span data-stu-id="07787-816">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="07787-817">如果上传附件失败，`asyncResult` 对象将包含一个提供错误说明的 `Error` 对象。</span><span class="sxs-lookup"><span data-stu-id="07787-817">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="07787-818">错误</span><span class="sxs-lookup"><span data-stu-id="07787-818">Errors</span></span>

|<span data-ttu-id="07787-819">错误代码</span><span class="sxs-lookup"><span data-stu-id="07787-819">Error code</span></span>|<span data-ttu-id="07787-820">说明</span><span class="sxs-lookup"><span data-stu-id="07787-820">Description</span></span>|
|------------|-------------|
|`AttachmentSizeExceeded`|<span data-ttu-id="07787-821">附件大小超过了允许的大小。</span><span class="sxs-lookup"><span data-stu-id="07787-821">The attachment is larger than allowed.</span></span>|
|`FileTypeNotSupported`|<span data-ttu-id="07787-822">该附件的扩展名不是允许的扩展名。</span><span class="sxs-lookup"><span data-stu-id="07787-822">The attachment has an extension that is not allowed.</span></span>|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="07787-823">邮件或约会具有的附件过多。</span><span class="sxs-lookup"><span data-stu-id="07787-823">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="07787-824">要求</span><span class="sxs-lookup"><span data-stu-id="07787-824">Requirements</span></span>

|<span data-ttu-id="07787-825">要求</span><span class="sxs-lookup"><span data-stu-id="07787-825">Requirement</span></span>|<span data-ttu-id="07787-826">值</span><span class="sxs-lookup"><span data-stu-id="07787-826">Value</span></span>|
|---|---|
|[<span data-ttu-id="07787-827">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="07787-827">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="07787-828">预览</span><span class="sxs-lookup"><span data-stu-id="07787-828">Preview</span></span>|
|[<span data-ttu-id="07787-829">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="07787-829">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="07787-830">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="07787-830">ReadWriteItem</span></span>|
|[<span data-ttu-id="07787-831">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="07787-831">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="07787-832">撰写</span><span class="sxs-lookup"><span data-stu-id="07787-832">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="07787-833">示例</span><span class="sxs-lookup"><span data-stu-id="07787-833">Examples</span></span>

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

#### <a name="addhandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="07787-834">addHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="07787-834">addHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="07787-835">添加支持事件的事件处理程序。</span><span class="sxs-lookup"><span data-stu-id="07787-835">Adds an event handler for a supported event.</span></span>

<span data-ttu-id="07787-836">目前, 受支持的事件`Office.EventType.AttachmentsChanged`类型`Office.EventType.AppointmentTimeChanged`是`Office.EventType.EnhancedLocationsChanged`、 `Office.EventType.RecipientsChanged`、、 `Office.EventType.RecurrenceChanged`和。</span><span class="sxs-lookup"><span data-stu-id="07787-836">Currently the supported event types are `Office.EventType.AttachmentsChanged`, `Office.EventType.AppointmentTimeChanged`, `Office.EventType.EnhancedLocationsChanged`, `Office.EventType.RecipientsChanged`, and `Office.EventType.RecurrenceChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="07787-837">参数</span><span class="sxs-lookup"><span data-stu-id="07787-837">Parameters</span></span>

| <span data-ttu-id="07787-838">名称</span><span class="sxs-lookup"><span data-stu-id="07787-838">Name</span></span> | <span data-ttu-id="07787-839">类型</span><span class="sxs-lookup"><span data-stu-id="07787-839">Type</span></span> | <span data-ttu-id="07787-840">属性</span><span class="sxs-lookup"><span data-stu-id="07787-840">Attributes</span></span> | <span data-ttu-id="07787-841">说明</span><span class="sxs-lookup"><span data-stu-id="07787-841">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="07787-842">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="07787-842">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="07787-843">应调用处理程序的事件。</span><span class="sxs-lookup"><span data-stu-id="07787-843">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="07787-844">函数</span><span class="sxs-lookup"><span data-stu-id="07787-844">Function</span></span> || <span data-ttu-id="07787-p140">用于处理事件的函数。此函数必须接受一个参数，即对象文本。参数上的 `type` 属性将匹配传递给 `addHandlerAsync` 的 `eventType` 参数。</span><span class="sxs-lookup"><span data-stu-id="07787-p140">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `addHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="07787-848">Object</span><span class="sxs-lookup"><span data-stu-id="07787-848">Object</span></span> | <span data-ttu-id="07787-849">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="07787-849">&lt;optional&gt;</span></span> | <span data-ttu-id="07787-850">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="07787-850">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="07787-851">对象</span><span class="sxs-lookup"><span data-stu-id="07787-851">Object</span></span> | <span data-ttu-id="07787-852">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="07787-852">&lt;optional&gt;</span></span> | <span data-ttu-id="07787-853">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="07787-853">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="07787-854">函数</span><span class="sxs-lookup"><span data-stu-id="07787-854">function</span></span>| <span data-ttu-id="07787-855">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="07787-855">&lt;optional&gt;</span></span>|<span data-ttu-id="07787-856">方法完成后，使用单个参数 `callback`（一个 [`asyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `AsyncResult` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="07787-856">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="07787-857">Requirements</span><span class="sxs-lookup"><span data-stu-id="07787-857">Requirements</span></span>

|<span data-ttu-id="07787-858">要求</span><span class="sxs-lookup"><span data-stu-id="07787-858">Requirement</span></span>| <span data-ttu-id="07787-859">值</span><span class="sxs-lookup"><span data-stu-id="07787-859">Value</span></span>|
|---|---|
|[<span data-ttu-id="07787-860">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="07787-860">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="07787-861">1.7</span><span class="sxs-lookup"><span data-stu-id="07787-861">1.7</span></span> |
|[<span data-ttu-id="07787-862">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="07787-862">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="07787-863">ReadItem</span><span class="sxs-lookup"><span data-stu-id="07787-863">ReadItem</span></span> |
|[<span data-ttu-id="07787-864">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="07787-864">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="07787-865">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="07787-865">Compose or Read</span></span> |

##### <a name="example"></a><span data-ttu-id="07787-866">示例</span><span class="sxs-lookup"><span data-stu-id="07787-866">Example</span></span>

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

#### <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="07787-867">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="07787-867">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="07787-868">将 Exchange 项目（如邮件）作为附件添加到邮件或约会。</span><span class="sxs-lookup"><span data-stu-id="07787-868">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="07787-p141">`addItemAttachmentAsync` 方法将包含指定 Exchange 标识符的项目附加到撰写窗体中的项目。如果指定一个回调方法，此方法使用 `asyncResult` 参数调用，该参数包含一个附件标识符或代码，指示附加项目过程中出现的任何错误。可以使用 `options` 参数将状态信息传递给回调方法（如果需要）。</span><span class="sxs-lookup"><span data-stu-id="07787-p141">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="07787-872">随后可以将该标识符与 [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) 方法一同使用，以删除同一个会话中的附件。</span><span class="sxs-lookup"><span data-stu-id="07787-872">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="07787-873">如果 Office 外接程序在 web 上的 Outlook 中运行, 则该`addItemAttachmentAsync`方法可以将项目附加到您正在编辑的项目之外的项目中;但是, 不支持这种情况, 建议不要这样做。</span><span class="sxs-lookup"><span data-stu-id="07787-873">If your Office Add-in is running in Outlook on the web, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="07787-874">参数</span><span class="sxs-lookup"><span data-stu-id="07787-874">Parameters</span></span>

|<span data-ttu-id="07787-875">名称</span><span class="sxs-lookup"><span data-stu-id="07787-875">Name</span></span>|<span data-ttu-id="07787-876">类型</span><span class="sxs-lookup"><span data-stu-id="07787-876">Type</span></span>|<span data-ttu-id="07787-877">属性</span><span class="sxs-lookup"><span data-stu-id="07787-877">Attributes</span></span>|<span data-ttu-id="07787-878">说明</span><span class="sxs-lookup"><span data-stu-id="07787-878">Description</span></span>|
|---|---|---|---|
|`itemId`|<span data-ttu-id="07787-879">字符串</span><span class="sxs-lookup"><span data-stu-id="07787-879">String</span></span>||<span data-ttu-id="07787-p142">要附加的项目的 Exchange 标识符。最大长度为 100 个字符。</span><span class="sxs-lookup"><span data-stu-id="07787-p142">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`|<span data-ttu-id="07787-882">String</span><span class="sxs-lookup"><span data-stu-id="07787-882">String</span></span>||<span data-ttu-id="07787-883">要附加的项目的主题。</span><span class="sxs-lookup"><span data-stu-id="07787-883">The subject of the item to be attached.</span></span> <span data-ttu-id="07787-884">最大长度为 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="07787-884">The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="07787-885">对象</span><span class="sxs-lookup"><span data-stu-id="07787-885">Object</span></span>|<span data-ttu-id="07787-886">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="07787-886">&lt;optional&gt;</span></span>|<span data-ttu-id="07787-887">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="07787-887">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="07787-888">对象</span><span class="sxs-lookup"><span data-stu-id="07787-888">Object</span></span>|<span data-ttu-id="07787-889">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="07787-889">&lt;optional&gt;</span></span>|<span data-ttu-id="07787-890">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="07787-890">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="07787-891">函数</span><span class="sxs-lookup"><span data-stu-id="07787-891">function</span></span>|<span data-ttu-id="07787-892">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="07787-892">&lt;optional&gt;</span></span>|<span data-ttu-id="07787-893">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="07787-893">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="07787-894">如果成功，附件标识符将在 `asyncResult.value` 属性中提供。</span><span class="sxs-lookup"><span data-stu-id="07787-894">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="07787-895">如果添加附件失败，`asyncResult` 对象将包含一个提供错误说明的 `Error` 对象。</span><span class="sxs-lookup"><span data-stu-id="07787-895">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="07787-896">错误</span><span class="sxs-lookup"><span data-stu-id="07787-896">Errors</span></span>

|<span data-ttu-id="07787-897">错误代码</span><span class="sxs-lookup"><span data-stu-id="07787-897">Error code</span></span>|<span data-ttu-id="07787-898">说明</span><span class="sxs-lookup"><span data-stu-id="07787-898">Description</span></span>|
|------------|-------------|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="07787-899">邮件或约会具有的附件过多。</span><span class="sxs-lookup"><span data-stu-id="07787-899">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="07787-900">要求</span><span class="sxs-lookup"><span data-stu-id="07787-900">Requirements</span></span>

|<span data-ttu-id="07787-901">要求</span><span class="sxs-lookup"><span data-stu-id="07787-901">Requirement</span></span>|<span data-ttu-id="07787-902">值</span><span class="sxs-lookup"><span data-stu-id="07787-902">Value</span></span>|
|---|---|
|[<span data-ttu-id="07787-903">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="07787-903">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="07787-904">1.1</span><span class="sxs-lookup"><span data-stu-id="07787-904">1.1</span></span>|
|[<span data-ttu-id="07787-905">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="07787-905">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="07787-906">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="07787-906">ReadWriteItem</span></span>|
|[<span data-ttu-id="07787-907">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="07787-907">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="07787-908">撰写</span><span class="sxs-lookup"><span data-stu-id="07787-908">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="07787-909">示例</span><span class="sxs-lookup"><span data-stu-id="07787-909">Example</span></span>

<span data-ttu-id="07787-910">以下示例将现有的 Outlook 项目添加为名为 `My Attachment` 的附件。</span><span class="sxs-lookup"><span data-stu-id="07787-910">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

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

#### <a name="close"></a><span data-ttu-id="07787-911">close()</span><span class="sxs-lookup"><span data-stu-id="07787-911">close()</span></span>

<span data-ttu-id="07787-912">关闭当前正在撰写的项目。</span><span class="sxs-lookup"><span data-stu-id="07787-912">Closes the current item that is being composed.</span></span>

<span data-ttu-id="07787-p144">
            \`close\` 方法的行为取决于要撰写的项目的当前状态。如果项目具有未保存的更改，客户端将提示用户保存、放弃或取消关闭操作。</span><span class="sxs-lookup"><span data-stu-id="07787-p144">The behavior of the `close` method depends on the current state of the item being composed. If the item has unsaved changes, the client prompts the user to save, discard, or cancel the close action.</span></span>

> [!NOTE]
> <span data-ttu-id="07787-915">在 Outlook 网页版中，如果该项目是约会并且之前已使用 `saveAsync` 保存，则即使自上次保存项目后未发生任何更改，也会提示用户保存、放弃或取消。</span><span class="sxs-lookup"><span data-stu-id="07787-915">In Outlook on the web, if the item is an appointment and it has previously been saved using `saveAsync`, the user is prompted to save, discard, or cancel even if no changes have occurred since the item was last saved.</span></span>

<span data-ttu-id="07787-916">在 Outlook 桌面客户端中，如果邮件是内联答复，`close` 方法不起作用。</span><span class="sxs-lookup"><span data-stu-id="07787-916">In the Outlook desktop client, if the message is an inline reply, the `close` method has no effect.</span></span>

##### <a name="requirements"></a><span data-ttu-id="07787-917">要求</span><span class="sxs-lookup"><span data-stu-id="07787-917">Requirements</span></span>

|<span data-ttu-id="07787-918">要求</span><span class="sxs-lookup"><span data-stu-id="07787-918">Requirement</span></span>|<span data-ttu-id="07787-919">值</span><span class="sxs-lookup"><span data-stu-id="07787-919">Value</span></span>|
|---|---|
|[<span data-ttu-id="07787-920">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="07787-920">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="07787-921">1.3</span><span class="sxs-lookup"><span data-stu-id="07787-921">1.3</span></span>|
|[<span data-ttu-id="07787-922">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="07787-922">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="07787-923">受限</span><span class="sxs-lookup"><span data-stu-id="07787-923">Restricted</span></span>|
|[<span data-ttu-id="07787-924">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="07787-924">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="07787-925">撰写</span><span class="sxs-lookup"><span data-stu-id="07787-925">Compose</span></span>|

---
---

#### <a name="displayreplyallformformdata-callback"></a><span data-ttu-id="07787-926">displayReplyAllForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="07787-926">displayReplyAllForm(formData, [callback])</span></span>

<span data-ttu-id="07787-927">显示答复窗体，其中包括所选邮件的发件人和所有收件人或所选约会的组织者和所有与会者。</span><span class="sxs-lookup"><span data-stu-id="07787-927">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="07787-928">IOS 或 Android 上的 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="07787-928">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="07787-929">在 web 上的 Outlook 中, 答复窗体显示为3列视图中的弹出窗体和2列或1列视图中的弹出式窗体。</span><span class="sxs-lookup"><span data-stu-id="07787-929">In Outlook on the web, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="07787-930">如果任意字符串参数超出其限制，`displayReplyAllForm` 将引发异常。</span><span class="sxs-lookup"><span data-stu-id="07787-930">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="07787-931">如果在`formData.attachments`参数中指定了附件, 则 web 上的 Outlook 和桌面客户端将尝试下载所有附件并将其附加到答复窗体。</span><span class="sxs-lookup"><span data-stu-id="07787-931">When attachments are specified in the `formData.attachments` parameter, Outlook on the web and desktop clients attempt to download all attachments and attach them to the reply form.</span></span> <span data-ttu-id="07787-932">如果无法添加任何附件，则在窗体 UI 中显示错误。</span><span class="sxs-lookup"><span data-stu-id="07787-932">If any attachments fail to be added, an error is shown in the form UI.</span></span> <span data-ttu-id="07787-933">如果这不可能，则不引发错误消息。</span><span class="sxs-lookup"><span data-stu-id="07787-933">If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="07787-934">参数</span><span class="sxs-lookup"><span data-stu-id="07787-934">Parameters</span></span>

|<span data-ttu-id="07787-935">名称</span><span class="sxs-lookup"><span data-stu-id="07787-935">Name</span></span>|<span data-ttu-id="07787-936">类型</span><span class="sxs-lookup"><span data-stu-id="07787-936">Type</span></span>|<span data-ttu-id="07787-937">属性</span><span class="sxs-lookup"><span data-stu-id="07787-937">Attributes</span></span>|<span data-ttu-id="07787-938">说明</span><span class="sxs-lookup"><span data-stu-id="07787-938">Description</span></span>|
|---|---|---|---|
|`formData`|<span data-ttu-id="07787-939">字符串 &#124; 对象</span><span class="sxs-lookup"><span data-stu-id="07787-939">String &#124; Object</span></span>||<span data-ttu-id="07787-p146">一个包含文本和 HTML 且表示答复窗体的正文的字符串。字符串限制为 32 KB。</span><span class="sxs-lookup"><span data-stu-id="07787-p146">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="07787-942">**或**</span><span class="sxs-lookup"><span data-stu-id="07787-942">**OR**</span></span><br/><span data-ttu-id="07787-p147">包含正文或附件数据和回调函数的对象。对象定义如下。</span><span class="sxs-lookup"><span data-stu-id="07787-p147">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span>|
|`formData.htmlBody`|<span data-ttu-id="07787-945">String</span><span class="sxs-lookup"><span data-stu-id="07787-945">String</span></span>|<span data-ttu-id="07787-946">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="07787-946">&lt;optional&gt;</span></span>|<span data-ttu-id="07787-p148">一个包含文本和 HTML 且表示答复窗体的正文的字符串。字符串限制为 32 KB。</span><span class="sxs-lookup"><span data-stu-id="07787-p148">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
|`formData.attachments`|<span data-ttu-id="07787-949">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="07787-949">Array.&lt;Object&gt;</span></span>|<span data-ttu-id="07787-950">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="07787-950">&lt;optional&gt;</span></span>|<span data-ttu-id="07787-951">JSON 对象（是文件或项目附件）的数组。</span><span class="sxs-lookup"><span data-stu-id="07787-951">An array of JSON objects that are either file or item attachments.</span></span>|
|`formData.attachments.type`|<span data-ttu-id="07787-952">String</span><span class="sxs-lookup"><span data-stu-id="07787-952">String</span></span>||<span data-ttu-id="07787-p149">指示附件的类型。必须是文件附件的 `file` 或项目附件的 `item`。</span><span class="sxs-lookup"><span data-stu-id="07787-p149">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span>|
|`formData.attachments.name`|<span data-ttu-id="07787-955">字符串</span><span class="sxs-lookup"><span data-stu-id="07787-955">String</span></span>||<span data-ttu-id="07787-956">一个包含附件的名称的字符串，最多包含 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="07787-956">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
|`formData.attachments.url`|<span data-ttu-id="07787-957">字符串</span><span class="sxs-lookup"><span data-stu-id="07787-957">String</span></span>||<span data-ttu-id="07787-p150">仅在将 `type` 设置为 `file` 时使用。文件的位置的 URI。</span><span class="sxs-lookup"><span data-stu-id="07787-p150">Only used if `type` is set to `file`. The URI of the location for the file.</span></span>|
|`formData.attachments.isInline`|<span data-ttu-id="07787-960">布尔</span><span class="sxs-lookup"><span data-stu-id="07787-960">Boolean</span></span>||<span data-ttu-id="07787-p151">仅在将 `type` 设置为 `file` 时使用。如果为 `true`，则表示附件将在邮件正文中内联显示，并且不应显示在附件列表中。</span><span class="sxs-lookup"><span data-stu-id="07787-p151">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`formData.attachments.itemId`|<span data-ttu-id="07787-963">String</span><span class="sxs-lookup"><span data-stu-id="07787-963">String</span></span>||<span data-ttu-id="07787-p152">仅在将 `type` 设置为 `item` 时使用。附件的 EWS 项目 ID。字符串最长为 100 个字符。</span><span class="sxs-lookup"><span data-stu-id="07787-p152">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span>|
|`callback`|<span data-ttu-id="07787-967">函数</span><span class="sxs-lookup"><span data-stu-id="07787-967">function</span></span>|<span data-ttu-id="07787-968">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="07787-968">&lt;optional&gt;</span></span>|<span data-ttu-id="07787-969">方法完成后，使用单个参数 `asyncResult`（一个 [AsyncResult](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="07787-969">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="07787-970">要求</span><span class="sxs-lookup"><span data-stu-id="07787-970">Requirements</span></span>

|<span data-ttu-id="07787-971">要求</span><span class="sxs-lookup"><span data-stu-id="07787-971">Requirement</span></span>|<span data-ttu-id="07787-972">值</span><span class="sxs-lookup"><span data-stu-id="07787-972">Value</span></span>|
|---|---|
|[<span data-ttu-id="07787-973">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="07787-973">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="07787-974">1.0</span><span class="sxs-lookup"><span data-stu-id="07787-974">1.0</span></span>|
|[<span data-ttu-id="07787-975">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="07787-975">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="07787-976">ReadItem</span><span class="sxs-lookup"><span data-stu-id="07787-976">ReadItem</span></span>|
|[<span data-ttu-id="07787-977">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="07787-977">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="07787-978">阅读</span><span class="sxs-lookup"><span data-stu-id="07787-978">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="07787-979">示例</span><span class="sxs-lookup"><span data-stu-id="07787-979">Examples</span></span>

<span data-ttu-id="07787-980">以下代码将一个字符串传递到 `displayReplyAllForm` 函数。</span><span class="sxs-lookup"><span data-stu-id="07787-980">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="07787-981">使用空白正文答复。</span><span class="sxs-lookup"><span data-stu-id="07787-981">Reply with an empty body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="07787-982">仅使用正文答复。</span><span class="sxs-lookup"><span data-stu-id="07787-982">Reply with just a body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="07787-983">使用正文和文件附件答复。</span><span class="sxs-lookup"><span data-stu-id="07787-983">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="07787-984">使用正文和项目附件答复。</span><span class="sxs-lookup"><span data-stu-id="07787-984">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="07787-985">使用正文、文件附件、项目附件和回调答复。</span><span class="sxs-lookup"><span data-stu-id="07787-985">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="displayreplyformformdata-callback"></a><span data-ttu-id="07787-986">displayReplyForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="07787-986">displayReplyForm(formData, [callback])</span></span>

<span data-ttu-id="07787-987">显示答复窗体，其中仅包括所选邮件的发件人或所选约会的组织者。</span><span class="sxs-lookup"><span data-stu-id="07787-987">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="07787-988">IOS 或 Android 上的 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="07787-988">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="07787-989">在 web 上的 Outlook 中, 答复窗体显示为3列视图中的弹出窗体和2列或1列视图中的弹出式窗体。</span><span class="sxs-lookup"><span data-stu-id="07787-989">In Outlook on the web, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="07787-990">如果任意字符串参数超出其限制，`displayReplyForm` 将引发异常。</span><span class="sxs-lookup"><span data-stu-id="07787-990">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="07787-991">如果在`formData.attachments`参数中指定了附件, 则 web 上的 Outlook 和桌面客户端将尝试下载所有附件并将其附加到答复窗体。</span><span class="sxs-lookup"><span data-stu-id="07787-991">When attachments are specified in the `formData.attachments` parameter, Outlook on the web and desktop clients attempt to download all attachments and attach them to the reply form.</span></span> <span data-ttu-id="07787-992">如果无法添加任何附件，则在窗体 UI 中显示错误。</span><span class="sxs-lookup"><span data-stu-id="07787-992">If any attachments fail to be added, an error is shown in the form UI.</span></span> <span data-ttu-id="07787-993">如果这不可能，则不引发错误消息。</span><span class="sxs-lookup"><span data-stu-id="07787-993">If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="07787-994">参数</span><span class="sxs-lookup"><span data-stu-id="07787-994">Parameters</span></span>

|<span data-ttu-id="07787-995">名称</span><span class="sxs-lookup"><span data-stu-id="07787-995">Name</span></span>|<span data-ttu-id="07787-996">类型</span><span class="sxs-lookup"><span data-stu-id="07787-996">Type</span></span>|<span data-ttu-id="07787-997">属性</span><span class="sxs-lookup"><span data-stu-id="07787-997">Attributes</span></span>|<span data-ttu-id="07787-998">说明</span><span class="sxs-lookup"><span data-stu-id="07787-998">Description</span></span>|
|---|---|---|---|
|`formData`|<span data-ttu-id="07787-999">字符串 &#124; 对象</span><span class="sxs-lookup"><span data-stu-id="07787-999">String &#124; Object</span></span>||<span data-ttu-id="07787-p154">一个包含文本和 HTML 且表示答复窗体的正文的字符串。字符串限制为 32 KB。</span><span class="sxs-lookup"><span data-stu-id="07787-p154">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="07787-1002">**或**</span><span class="sxs-lookup"><span data-stu-id="07787-1002">**OR**</span></span><br/><span data-ttu-id="07787-p155">包含正文或附件数据和回调函数的对象。对象定义如下。</span><span class="sxs-lookup"><span data-stu-id="07787-p155">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span>|
|`formData.htmlBody`|<span data-ttu-id="07787-1005">String</span><span class="sxs-lookup"><span data-stu-id="07787-1005">String</span></span>|<span data-ttu-id="07787-1006">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="07787-1006">&lt;optional&gt;</span></span>|<span data-ttu-id="07787-p156">一个包含文本和 HTML 且表示答复窗体的正文的字符串。字符串限制为 32 KB。</span><span class="sxs-lookup"><span data-stu-id="07787-p156">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
|`formData.attachments`|<span data-ttu-id="07787-1009">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="07787-1009">Array.&lt;Object&gt;</span></span>|<span data-ttu-id="07787-1010">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="07787-1010">&lt;optional&gt;</span></span>|<span data-ttu-id="07787-1011">JSON 对象（是文件或项目附件）的数组。</span><span class="sxs-lookup"><span data-stu-id="07787-1011">An array of JSON objects that are either file or item attachments.</span></span>|
|`formData.attachments.type`|<span data-ttu-id="07787-1012">String</span><span class="sxs-lookup"><span data-stu-id="07787-1012">String</span></span>||<span data-ttu-id="07787-p157">指示附件的类型。必须是文件附件的 `file` 或项目附件的 `item`。</span><span class="sxs-lookup"><span data-stu-id="07787-p157">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span>|
|`formData.attachments.name`|<span data-ttu-id="07787-1015">字符串</span><span class="sxs-lookup"><span data-stu-id="07787-1015">String</span></span>||<span data-ttu-id="07787-1016">一个包含附件的名称的字符串，最多包含 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="07787-1016">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
|`formData.attachments.url`|<span data-ttu-id="07787-1017">String</span><span class="sxs-lookup"><span data-stu-id="07787-1017">String</span></span>||<span data-ttu-id="07787-p158">仅在将 `type` 设置为 `file` 时使用。文件的位置的 URI。</span><span class="sxs-lookup"><span data-stu-id="07787-p158">Only used if `type` is set to `file`. The URI of the location for the file.</span></span>|
|`formData.attachments.isInline`|<span data-ttu-id="07787-1020">布尔</span><span class="sxs-lookup"><span data-stu-id="07787-1020">Boolean</span></span>||<span data-ttu-id="07787-p159">仅在将 `type` 设置为 `file` 时使用。如果为 `true`，则表示附件将在邮件正文中内联显示，并且不应显示在附件列表中。</span><span class="sxs-lookup"><span data-stu-id="07787-p159">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`formData.attachments.itemId`|<span data-ttu-id="07787-1023">String</span><span class="sxs-lookup"><span data-stu-id="07787-1023">String</span></span>||<span data-ttu-id="07787-p160">仅在将 `type` 设置为 `item` 时使用。附件的 EWS 项目 ID。字符串最长为 100 个字符。</span><span class="sxs-lookup"><span data-stu-id="07787-p160">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span>|
|`callback`|<span data-ttu-id="07787-1027">函数</span><span class="sxs-lookup"><span data-stu-id="07787-1027">function</span></span>|<span data-ttu-id="07787-1028">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="07787-1028">&lt;optional&gt;</span></span>|<span data-ttu-id="07787-1029">方法完成后，使用单个参数 `asyncResult`（一个 [AsyncResult](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="07787-1029">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="07787-1030">要求</span><span class="sxs-lookup"><span data-stu-id="07787-1030">Requirements</span></span>

|<span data-ttu-id="07787-1031">要求</span><span class="sxs-lookup"><span data-stu-id="07787-1031">Requirement</span></span>|<span data-ttu-id="07787-1032">值</span><span class="sxs-lookup"><span data-stu-id="07787-1032">Value</span></span>|
|---|---|
|[<span data-ttu-id="07787-1033">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="07787-1033">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="07787-1034">1.0</span><span class="sxs-lookup"><span data-stu-id="07787-1034">1.0</span></span>|
|[<span data-ttu-id="07787-1035">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="07787-1035">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="07787-1036">ReadItem</span><span class="sxs-lookup"><span data-stu-id="07787-1036">ReadItem</span></span>|
|[<span data-ttu-id="07787-1037">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="07787-1037">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="07787-1038">阅读</span><span class="sxs-lookup"><span data-stu-id="07787-1038">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="07787-1039">示例</span><span class="sxs-lookup"><span data-stu-id="07787-1039">Examples</span></span>

<span data-ttu-id="07787-1040">以下代码将一个字符串传递到 `displayReplyForm` 函数。</span><span class="sxs-lookup"><span data-stu-id="07787-1040">The following code passes a string to the `displayReplyForm` function.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="07787-1041">使用空白正文答复。</span><span class="sxs-lookup"><span data-stu-id="07787-1041">Reply with an empty body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="07787-1042">仅使用正文答复。</span><span class="sxs-lookup"><span data-stu-id="07787-1042">Reply with just a body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="07787-1043">使用正文和文件附件答复。</span><span class="sxs-lookup"><span data-stu-id="07787-1043">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="07787-1044">使用正文和项目附件答复。</span><span class="sxs-lookup"><span data-stu-id="07787-1044">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="07787-1045">使用正文、文件附件、项目附件和回调答复。</span><span class="sxs-lookup"><span data-stu-id="07787-1045">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="getattachmentcontentasyncattachmentid-options-callback--attachmentcontentjavascriptapioutlookofficeattachmentcontent"></a><span data-ttu-id="07787-1046">getAttachmentContentAsync (attachmentId, [options], [callback]) → [AttachmentContent](/javascript/api/outlook/office.attachmentcontent)</span><span class="sxs-lookup"><span data-stu-id="07787-1046">getAttachmentContentAsync(attachmentId, [options], [callback]) → [AttachmentContent](/javascript/api/outlook/office.attachmentcontent)</span></span>

<span data-ttu-id="07787-1047">从邮件或约会中获取指定附件并将其作为`AttachmentContent`对象返回。</span><span class="sxs-lookup"><span data-stu-id="07787-1047">Gets the specified attachment from a message or appointment and returns it as an `AttachmentContent` object.</span></span>

<span data-ttu-id="07787-1048">该`getAttachmentContentAsync`方法从项目中获取具有指定标识符的附件。</span><span class="sxs-lookup"><span data-stu-id="07787-1048">The `getAttachmentContentAsync` method gets the attachment with the specified identifier from the item.</span></span> <span data-ttu-id="07787-1049">作为一种最佳做法, 您应使用标识符在与`getAttachmentsAsync` or `item.attachments`调用一起检索到会话的同一会话中检索附件。</span><span class="sxs-lookup"><span data-stu-id="07787-1049">As a best practice, you should use the identifier to retrieve an attachment in the same session that the attachmentIds were retrieved with the `getAttachmentsAsync` or `item.attachments` call.</span></span> <span data-ttu-id="07787-1050">在 web 和移动设备上的 Outlook 中, 附件标识符仅在同一个会话中有效。</span><span class="sxs-lookup"><span data-stu-id="07787-1050">In Outlook on the web and mobile devices, the attachment identifier is valid only within the same session.</span></span> <span data-ttu-id="07787-1051">当用户关闭应用程序时, 或者如果用户开始撰写内嵌窗体, 随后弹出窗体以继续在单独的窗口中。</span><span class="sxs-lookup"><span data-stu-id="07787-1051">A session is over when the user closes the app, or if the user starts composing an inline form then subsequently pops out the form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="07787-1052">参数</span><span class="sxs-lookup"><span data-stu-id="07787-1052">Parameters</span></span>

|<span data-ttu-id="07787-1053">名称</span><span class="sxs-lookup"><span data-stu-id="07787-1053">Name</span></span>|<span data-ttu-id="07787-1054">类型</span><span class="sxs-lookup"><span data-stu-id="07787-1054">Type</span></span>|<span data-ttu-id="07787-1055">属性</span><span class="sxs-lookup"><span data-stu-id="07787-1055">Attributes</span></span>|<span data-ttu-id="07787-1056">说明</span><span class="sxs-lookup"><span data-stu-id="07787-1056">Description</span></span>|
|---|---|---|---|
|`attachmentId`|<span data-ttu-id="07787-1057">String</span><span class="sxs-lookup"><span data-stu-id="07787-1057">String</span></span>||<span data-ttu-id="07787-1058">要获取的附件的标识符。</span><span class="sxs-lookup"><span data-stu-id="07787-1058">The identifier of the attachment you want to get.</span></span>|
|`options`|<span data-ttu-id="07787-1059">对象</span><span class="sxs-lookup"><span data-stu-id="07787-1059">Object</span></span>|<span data-ttu-id="07787-1060">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="07787-1060">&lt;optional&gt;</span></span>|<span data-ttu-id="07787-1061">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="07787-1061">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="07787-1062">对象</span><span class="sxs-lookup"><span data-stu-id="07787-1062">Object</span></span>|<span data-ttu-id="07787-1063">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="07787-1063">&lt;optional&gt;</span></span>|<span data-ttu-id="07787-1064">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="07787-1064">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="07787-1065">函数</span><span class="sxs-lookup"><span data-stu-id="07787-1065">function</span></span>|<span data-ttu-id="07787-1066">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="07787-1066">&lt;optional&gt;</span></span>|<span data-ttu-id="07787-1067">方法完成后，使用单个参数 `asyncResult`（一个 [AsyncResult](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="07787-1067">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="07787-1068">要求</span><span class="sxs-lookup"><span data-stu-id="07787-1068">Requirements</span></span>

|<span data-ttu-id="07787-1069">要求</span><span class="sxs-lookup"><span data-stu-id="07787-1069">Requirement</span></span>|<span data-ttu-id="07787-1070">值</span><span class="sxs-lookup"><span data-stu-id="07787-1070">Value</span></span>|
|---|---|
|[<span data-ttu-id="07787-1071">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="07787-1071">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="07787-1072">预览</span><span class="sxs-lookup"><span data-stu-id="07787-1072">Preview</span></span>|
|[<span data-ttu-id="07787-1073">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="07787-1073">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="07787-1074">ReadItem</span><span class="sxs-lookup"><span data-stu-id="07787-1074">ReadItem</span></span>|
|[<span data-ttu-id="07787-1075">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="07787-1075">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="07787-1076">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="07787-1076">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="07787-1077">返回：</span><span class="sxs-lookup"><span data-stu-id="07787-1077">Returns:</span></span>

<span data-ttu-id="07787-1078">类型: [AttachmentContent](/javascript/api/outlook/office.attachmentcontent)</span><span class="sxs-lookup"><span data-stu-id="07787-1078">Type: [AttachmentContent](/javascript/api/outlook/office.attachmentcontent)</span></span>

##### <a name="example"></a><span data-ttu-id="07787-1079">示例</span><span class="sxs-lookup"><span data-stu-id="07787-1079">Example</span></span>

```javascript
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
  if (result.value.format === Office.MailboxEnums.AttachmentContentFormat.Base64) {
    // Handle file attachment.
  } else if (result.value.format === Office.MailboxEnums.AttachmentContentFormat.Eml) {
    // Handle email item attachment.
  } else if (result.value.format === Office.MailboxEnums.AttachmentContentFormat.ICalendar) {
    // Handle .icalender attachment.
  } else if (result.value.format === Office.MailboxEnums.AttachmentContentFormat.Url) {
    // Handle cloud attachment.
  } else {
    // Handle attachment formats that are not supported.
  }
}
```

---
---

#### <a name="getattachmentsasyncoptions-callback--arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetails"></a><span data-ttu-id="07787-1080">getAttachmentsAsync ([options], [callback]) → Array. <[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="07787-1080">getAttachmentsAsync([options], [callback]) → Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span></span>

<span data-ttu-id="07787-1081">以数组的形式获取项目的附件。</span><span class="sxs-lookup"><span data-stu-id="07787-1081">Gets the item's attachments as an array.</span></span> <span data-ttu-id="07787-1082">仅限撰写模式。</span><span class="sxs-lookup"><span data-stu-id="07787-1082">Compose mode only.</span></span>

##### <a name="parameters"></a><span data-ttu-id="07787-1083">参数</span><span class="sxs-lookup"><span data-stu-id="07787-1083">Parameters</span></span>

|<span data-ttu-id="07787-1084">名称</span><span class="sxs-lookup"><span data-stu-id="07787-1084">Name</span></span>|<span data-ttu-id="07787-1085">类型</span><span class="sxs-lookup"><span data-stu-id="07787-1085">Type</span></span>|<span data-ttu-id="07787-1086">属性</span><span class="sxs-lookup"><span data-stu-id="07787-1086">Attributes</span></span>|<span data-ttu-id="07787-1087">说明</span><span class="sxs-lookup"><span data-stu-id="07787-1087">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="07787-1088">对象</span><span class="sxs-lookup"><span data-stu-id="07787-1088">Object</span></span>|<span data-ttu-id="07787-1089">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="07787-1089">&lt;optional&gt;</span></span>|<span data-ttu-id="07787-1090">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="07787-1090">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="07787-1091">对象</span><span class="sxs-lookup"><span data-stu-id="07787-1091">Object</span></span>|<span data-ttu-id="07787-1092">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="07787-1092">&lt;optional&gt;</span></span>|<span data-ttu-id="07787-1093">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="07787-1093">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="07787-1094">函数</span><span class="sxs-lookup"><span data-stu-id="07787-1094">function</span></span>|<span data-ttu-id="07787-1095">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="07787-1095">&lt;optional&gt;</span></span>|<span data-ttu-id="07787-1096">方法完成后，使用单个参数 `asyncResult`（一个 [AsyncResult](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="07787-1096">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="07787-1097">要求</span><span class="sxs-lookup"><span data-stu-id="07787-1097">Requirements</span></span>

|<span data-ttu-id="07787-1098">要求</span><span class="sxs-lookup"><span data-stu-id="07787-1098">Requirement</span></span>|<span data-ttu-id="07787-1099">值</span><span class="sxs-lookup"><span data-stu-id="07787-1099">Value</span></span>|
|---|---|
|[<span data-ttu-id="07787-1100">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="07787-1100">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="07787-1101">预览</span><span class="sxs-lookup"><span data-stu-id="07787-1101">Preview</span></span>|
|[<span data-ttu-id="07787-1102">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="07787-1102">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="07787-1103">ReadItem</span><span class="sxs-lookup"><span data-stu-id="07787-1103">ReadItem</span></span>|
|[<span data-ttu-id="07787-1104">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="07787-1104">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="07787-1105">撰写</span><span class="sxs-lookup"><span data-stu-id="07787-1105">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="07787-1106">返回：</span><span class="sxs-lookup"><span data-stu-id="07787-1106">Returns:</span></span>

<span data-ttu-id="07787-1107">类型: Array. <[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="07787-1107">Type: Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span></span>

##### <a name="example"></a><span data-ttu-id="07787-1108">示例</span><span class="sxs-lookup"><span data-stu-id="07787-1108">Example</span></span>

<span data-ttu-id="07787-1109">下面的示例将生成一个 HTML 字符串, 其中包含当前项目上所有附件的详细信息。</span><span class="sxs-lookup"><span data-stu-id="07787-1109">The following example builds an HTML string with details of all attachments on the current item.</span></span>

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

#### <a name="getentities--entitiesjavascriptapioutlookofficeentities"></a><span data-ttu-id="07787-1110">getEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="07787-1110">getEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span></span>

<span data-ttu-id="07787-1111">获取在所选项目的正文中找到的实体。</span><span class="sxs-lookup"><span data-stu-id="07787-1111">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="07787-1112">IOS 或 Android 上的 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="07787-1112">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="07787-1113">要求</span><span class="sxs-lookup"><span data-stu-id="07787-1113">Requirements</span></span>

|<span data-ttu-id="07787-1114">要求</span><span class="sxs-lookup"><span data-stu-id="07787-1114">Requirement</span></span>|<span data-ttu-id="07787-1115">值</span><span class="sxs-lookup"><span data-stu-id="07787-1115">Value</span></span>|
|---|---|
|[<span data-ttu-id="07787-1116">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="07787-1116">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="07787-1117">1.0</span><span class="sxs-lookup"><span data-stu-id="07787-1117">1.0</span></span>|
|[<span data-ttu-id="07787-1118">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="07787-1118">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="07787-1119">ReadItem</span><span class="sxs-lookup"><span data-stu-id="07787-1119">ReadItem</span></span>|
|[<span data-ttu-id="07787-1120">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="07787-1120">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="07787-1121">阅读</span><span class="sxs-lookup"><span data-stu-id="07787-1121">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="07787-1122">返回：</span><span class="sxs-lookup"><span data-stu-id="07787-1122">Returns:</span></span>

<span data-ttu-id="07787-1123">类型：[Entities](/javascript/api/outlook/office.entities)</span><span class="sxs-lookup"><span data-stu-id="07787-1123">Type: [Entities](/javascript/api/outlook/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="07787-1124">示例</span><span class="sxs-lookup"><span data-stu-id="07787-1124">Example</span></span>

<span data-ttu-id="07787-1125">以下示例访问当前项目的正文中的联系人实体。</span><span class="sxs-lookup"><span data-stu-id="07787-1125">The following example accesses the contacts entities in the current item's body.</span></span>

```javascript
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

---
---

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlookofficecontactmeetingsuggestionjavascriptapioutlookofficemeetingsuggestionphonenumberjavascriptapioutlookofficephonenumbertasksuggestionjavascriptapioutlookofficetasksuggestion"></a><span data-ttu-id="07787-1126">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="07787-1126">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span></span>

<span data-ttu-id="07787-1127">获取所选项目的正文中找到的指定实体类型的所有实体的数组。</span><span class="sxs-lookup"><span data-stu-id="07787-1127">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="07787-1128">IOS 或 Android 上的 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="07787-1128">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="07787-1129">参数</span><span class="sxs-lookup"><span data-stu-id="07787-1129">Parameters</span></span>

|<span data-ttu-id="07787-1130">名称</span><span class="sxs-lookup"><span data-stu-id="07787-1130">Name</span></span>|<span data-ttu-id="07787-1131">类型</span><span class="sxs-lookup"><span data-stu-id="07787-1131">Type</span></span>|<span data-ttu-id="07787-1132">说明</span><span class="sxs-lookup"><span data-stu-id="07787-1132">Description</span></span>|
|---|---|---|
|`entityType`|[<span data-ttu-id="07787-1133">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="07787-1133">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook/office.mailboxenums.entitytype)|<span data-ttu-id="07787-1134">EntityType 枚举值之一。</span><span class="sxs-lookup"><span data-stu-id="07787-1134">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="07787-1135">要求</span><span class="sxs-lookup"><span data-stu-id="07787-1135">Requirements</span></span>

|<span data-ttu-id="07787-1136">要求</span><span class="sxs-lookup"><span data-stu-id="07787-1136">Requirement</span></span>|<span data-ttu-id="07787-1137">值</span><span class="sxs-lookup"><span data-stu-id="07787-1137">Value</span></span>|
|---|---|
|[<span data-ttu-id="07787-1138">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="07787-1138">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="07787-1139">1.0</span><span class="sxs-lookup"><span data-stu-id="07787-1139">1.0</span></span>|
|[<span data-ttu-id="07787-1140">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="07787-1140">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="07787-1141">受限</span><span class="sxs-lookup"><span data-stu-id="07787-1141">Restricted</span></span>|
|[<span data-ttu-id="07787-1142">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="07787-1142">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="07787-1143">阅读</span><span class="sxs-lookup"><span data-stu-id="07787-1143">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="07787-1144">返回：</span><span class="sxs-lookup"><span data-stu-id="07787-1144">Returns:</span></span>

<span data-ttu-id="07787-1145">如果在 `entityType` 中传递的值不是 `EntityType` 枚举的有效成员，该方法返回 null。</span><span class="sxs-lookup"><span data-stu-id="07787-1145">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="07787-1146">如果指定类型的任何实体都不存在于该项目的正文中，该方法将返回空数组。</span><span class="sxs-lookup"><span data-stu-id="07787-1146">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="07787-1147">否则，返回的数组中对象的类型取决于 `entityType` 参数中请求实体的类型。</span><span class="sxs-lookup"><span data-stu-id="07787-1147">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="07787-1148">当使用此方法的最低权限级别**受限**时，某些实体类型需要 **ReadItem** 才能进行访问，如下表中所指定。</span><span class="sxs-lookup"><span data-stu-id="07787-1148">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

|<span data-ttu-id="07787-1149">`entityType` 的值</span><span class="sxs-lookup"><span data-stu-id="07787-1149">Value of `entityType`</span></span>|<span data-ttu-id="07787-1150">返回的数组中对象的类型</span><span class="sxs-lookup"><span data-stu-id="07787-1150">Type of objects in returned array</span></span>|<span data-ttu-id="07787-1151">所需权限级别</span><span class="sxs-lookup"><span data-stu-id="07787-1151">Required Permission Level</span></span>|
|---|---|---|
|`Address`|<span data-ttu-id="07787-1152">字符串</span><span class="sxs-lookup"><span data-stu-id="07787-1152">String</span></span>|<span data-ttu-id="07787-1153">**受限**</span><span class="sxs-lookup"><span data-stu-id="07787-1153">**Restricted**</span></span>|
|`Contact`|<span data-ttu-id="07787-1154">Contact</span><span class="sxs-lookup"><span data-stu-id="07787-1154">Contact</span></span>|<span data-ttu-id="07787-1155">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="07787-1155">**ReadItem**</span></span>|
|`EmailAddress`|<span data-ttu-id="07787-1156">String</span><span class="sxs-lookup"><span data-stu-id="07787-1156">String</span></span>|<span data-ttu-id="07787-1157">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="07787-1157">**ReadItem**</span></span>|
|`MeetingSuggestion`|<span data-ttu-id="07787-1158">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="07787-1158">MeetingSuggestion</span></span>|<span data-ttu-id="07787-1159">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="07787-1159">**ReadItem**</span></span>|
|`PhoneNumber`|<span data-ttu-id="07787-1160">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="07787-1160">PhoneNumber</span></span>|<span data-ttu-id="07787-1161">**受限**</span><span class="sxs-lookup"><span data-stu-id="07787-1161">**Restricted**</span></span>|
|`TaskSuggestion`|<span data-ttu-id="07787-1162">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="07787-1162">TaskSuggestion</span></span>|<span data-ttu-id="07787-1163">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="07787-1163">**ReadItem**</span></span>|
|`URL`|<span data-ttu-id="07787-1164">字符串</span><span class="sxs-lookup"><span data-stu-id="07787-1164">String</span></span>|<span data-ttu-id="07787-1165">**受限**</span><span class="sxs-lookup"><span data-stu-id="07787-1165">**Restricted**</span></span>|

<span data-ttu-id="07787-1166">类型：Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="07787-1166">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span></span>

##### <a name="example"></a><span data-ttu-id="07787-1167">示例</span><span class="sxs-lookup"><span data-stu-id="07787-1167">Example</span></span>

<span data-ttu-id="07787-1168">以下示例显示了如何访问表示当前项目的正文中的邮政地址的字符串数组。</span><span class="sxs-lookup"><span data-stu-id="07787-1168">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlookofficecontactmeetingsuggestionjavascriptapioutlookofficemeetingsuggestionphonenumberjavascriptapioutlookofficephonenumbertasksuggestionjavascriptapioutlookofficetasksuggestion"></a><span data-ttu-id="07787-1169">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="07787-1169">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span></span>

<span data-ttu-id="07787-1170">返回传递清单 XML 文件中定义的命名筛选器的所选项目中的已知实体。</span><span class="sxs-lookup"><span data-stu-id="07787-1170">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="07787-1171">IOS 或 Android 上的 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="07787-1171">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="07787-1172">`getFilteredEntitiesByName` 方法返回匹配在具有指定 `FilterName` 元素值的清单 XML 文件中的 [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) 规则元素中定义的正则表达式的实体。</span><span class="sxs-lookup"><span data-stu-id="07787-1172">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="07787-1173">参数</span><span class="sxs-lookup"><span data-stu-id="07787-1173">Parameters</span></span>

|<span data-ttu-id="07787-1174">名称</span><span class="sxs-lookup"><span data-stu-id="07787-1174">Name</span></span>|<span data-ttu-id="07787-1175">类型</span><span class="sxs-lookup"><span data-stu-id="07787-1175">Type</span></span>|<span data-ttu-id="07787-1176">说明</span><span class="sxs-lookup"><span data-stu-id="07787-1176">Description</span></span>|
|---|---|---|
|`name`|<span data-ttu-id="07787-1177">字符串</span><span class="sxs-lookup"><span data-stu-id="07787-1177">String</span></span>|<span data-ttu-id="07787-1178">定义筛选器匹配的 `ItemHasKnownEntity` 规则元素的名称。</span><span class="sxs-lookup"><span data-stu-id="07787-1178">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="07787-1179">要求</span><span class="sxs-lookup"><span data-stu-id="07787-1179">Requirements</span></span>

|<span data-ttu-id="07787-1180">要求</span><span class="sxs-lookup"><span data-stu-id="07787-1180">Requirement</span></span>|<span data-ttu-id="07787-1181">值</span><span class="sxs-lookup"><span data-stu-id="07787-1181">Value</span></span>|
|---|---|
|[<span data-ttu-id="07787-1182">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="07787-1182">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="07787-1183">1.0</span><span class="sxs-lookup"><span data-stu-id="07787-1183">1.0</span></span>|
|[<span data-ttu-id="07787-1184">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="07787-1184">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="07787-1185">ReadItem</span><span class="sxs-lookup"><span data-stu-id="07787-1185">ReadItem</span></span>|
|[<span data-ttu-id="07787-1186">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="07787-1186">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="07787-1187">阅读</span><span class="sxs-lookup"><span data-stu-id="07787-1187">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="07787-1188">返回：</span><span class="sxs-lookup"><span data-stu-id="07787-1188">Returns:</span></span>

<span data-ttu-id="07787-p164">如果具有匹配 `name` 参数的 `FilterName` 元素值的清单中没有任何 `ItemHasKnownEntity` 元素，则该方法返回 `null`。如果 `name` 参数匹配清单中的 `ItemHasKnownEntity` 元素，但在匹配的当前项目中没有实体，则该方法返回一个空数组。</span><span class="sxs-lookup"><span data-stu-id="07787-p164">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="07787-1191">类型：Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="07787-1191">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span></span>

---
---

#### <a name="getinitializationcontextasyncoptions-callback"></a><span data-ttu-id="07787-1192">Office.context.mailbox.item.getinitializationcontextasync ([options], [callback])</span><span class="sxs-lookup"><span data-stu-id="07787-1192">getInitializationContextAsync([options], [callback])</span></span>

<span data-ttu-id="07787-1193">获取[通过可操作邮件激活](/outlook/actionable-messages/invoke-add-in-from-actionable-message)加载项时传递的初始化数据。</span><span class="sxs-lookup"><span data-stu-id="07787-1193">Gets initialization data passed when the add-in is [activated by an actionable message](/outlook/actionable-messages/invoke-add-in-from-actionable-message).</span></span>

> [!NOTE]
> <span data-ttu-id="07787-1194">仅 Outlook 2016 或更高版本 (高于16.0.8413.1000 的即点即用版本) 和适用于 Office 365 的 Outlook 网页版支持此方法。</span><span class="sxs-lookup"><span data-stu-id="07787-1194">This method is only supported by Outlook 2016 or later on Windows (Click-to-Run versions later than 16.0.8413.1000) and Outlook on the web for Office 365.</span></span>

##### <a name="parameters"></a><span data-ttu-id="07787-1195">参数</span><span class="sxs-lookup"><span data-stu-id="07787-1195">Parameters</span></span>

|<span data-ttu-id="07787-1196">名称</span><span class="sxs-lookup"><span data-stu-id="07787-1196">Name</span></span>|<span data-ttu-id="07787-1197">类型</span><span class="sxs-lookup"><span data-stu-id="07787-1197">Type</span></span>|<span data-ttu-id="07787-1198">属性</span><span class="sxs-lookup"><span data-stu-id="07787-1198">Attributes</span></span>|<span data-ttu-id="07787-1199">说明</span><span class="sxs-lookup"><span data-stu-id="07787-1199">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="07787-1200">对象</span><span class="sxs-lookup"><span data-stu-id="07787-1200">Object</span></span>|<span data-ttu-id="07787-1201">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="07787-1201">&lt;optional&gt;</span></span>|<span data-ttu-id="07787-1202">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="07787-1202">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="07787-1203">对象</span><span class="sxs-lookup"><span data-stu-id="07787-1203">Object</span></span>|<span data-ttu-id="07787-1204">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="07787-1204">&lt;optional&gt;</span></span>|<span data-ttu-id="07787-1205">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="07787-1205">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="07787-1206">函数</span><span class="sxs-lookup"><span data-stu-id="07787-1206">function</span></span>|<span data-ttu-id="07787-1207">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="07787-1207">&lt;optional&gt;</span></span>|<span data-ttu-id="07787-1208">方法完成后，使用单个参数 `callback`（一个 [`asyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `AsyncResult` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="07787-1208">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="07787-1209">如果成功, 初始化数据在`asyncResult.value`属性中提供为字符串。</span><span class="sxs-lookup"><span data-stu-id="07787-1209">On success, the initialization data is provided in the `asyncResult.value` property as a string.</span></span><br/><span data-ttu-id="07787-1210">如果没有初始化上下文, 该`asyncResult`对象将包含其`Error` `code`属性设置为`9020`的对象及其`name`属性设置为。 `GenericResponseError`</span><span class="sxs-lookup"><span data-stu-id="07787-1210">If there is no initialization context, the `asyncResult` object will contain an `Error` object with its `code` property set to `9020` and its `name` property set to `GenericResponseError`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="07787-1211">要求</span><span class="sxs-lookup"><span data-stu-id="07787-1211">Requirements</span></span>

|<span data-ttu-id="07787-1212">要求</span><span class="sxs-lookup"><span data-stu-id="07787-1212">Requirement</span></span>|<span data-ttu-id="07787-1213">值</span><span class="sxs-lookup"><span data-stu-id="07787-1213">Value</span></span>|
|---|---|
|[<span data-ttu-id="07787-1214">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="07787-1214">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="07787-1215">预览</span><span class="sxs-lookup"><span data-stu-id="07787-1215">Preview</span></span>|
|[<span data-ttu-id="07787-1216">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="07787-1216">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="07787-1217">ReadItem</span><span class="sxs-lookup"><span data-stu-id="07787-1217">ReadItem</span></span>|
|[<span data-ttu-id="07787-1218">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="07787-1218">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="07787-1219">阅读</span><span class="sxs-lookup"><span data-stu-id="07787-1219">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="07787-1220">示例</span><span class="sxs-lookup"><span data-stu-id="07787-1220">Example</span></span>

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

#### <a name="getitemidasyncoptions-callback"></a><span data-ttu-id="07787-1221">getItemIdAsync ([options], 回拨)</span><span class="sxs-lookup"><span data-stu-id="07787-1221">getItemIdAsync([options], callback)</span></span>

<span data-ttu-id="07787-1222">异步获取已保存项的 ID。</span><span class="sxs-lookup"><span data-stu-id="07787-1222">Asynchronously gets the ID of a saved item.</span></span> <span data-ttu-id="07787-1223">仅限撰写模式。</span><span class="sxs-lookup"><span data-stu-id="07787-1223">Compose mode only.</span></span>

<span data-ttu-id="07787-1224">调用此方法时, 此方法通过回调方法返回项 ID。</span><span class="sxs-lookup"><span data-stu-id="07787-1224">When invoked, this method returns the item ID via the callback method.</span></span>

> [!NOTE]
> <span data-ttu-id="07787-1225">如果你的外接程序`getItemIdAsync`对撰写模式中的项 (例如, 要获取`itemId`使用 EWS 或 REST API 的使用) 调用, 请注意, 当 Outlook 处于缓存模式下时, 可能需要一段时间才能将项目同步到服务器。</span><span class="sxs-lookup"><span data-stu-id="07787-1225">If your add-in calls `getItemIdAsync` on an item in compose mode (e.g., to get an `itemId` to use with EWS or the REST API), be aware that when Outlook is in cached mode, it may take some time before the item is synced to the server.</span></span> <span data-ttu-id="07787-1226">在同步项目之前, 无法识别`itemId`该项目并使用它将返回错误。</span><span class="sxs-lookup"><span data-stu-id="07787-1226">Until the item is synced, the `itemId` is not recognized and using it returns an error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="07787-1227">参数</span><span class="sxs-lookup"><span data-stu-id="07787-1227">Parameters</span></span>

|<span data-ttu-id="07787-1228">名称</span><span class="sxs-lookup"><span data-stu-id="07787-1228">Name</span></span>|<span data-ttu-id="07787-1229">类型</span><span class="sxs-lookup"><span data-stu-id="07787-1229">Type</span></span>|<span data-ttu-id="07787-1230">属性</span><span class="sxs-lookup"><span data-stu-id="07787-1230">Attributes</span></span>|<span data-ttu-id="07787-1231">说明</span><span class="sxs-lookup"><span data-stu-id="07787-1231">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="07787-1232">对象</span><span class="sxs-lookup"><span data-stu-id="07787-1232">Object</span></span>|<span data-ttu-id="07787-1233">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="07787-1233">&lt;optional&gt;</span></span>|<span data-ttu-id="07787-1234">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="07787-1234">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="07787-1235">对象</span><span class="sxs-lookup"><span data-stu-id="07787-1235">Object</span></span>|<span data-ttu-id="07787-1236">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="07787-1236">&lt;optional&gt;</span></span>|<span data-ttu-id="07787-1237">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="07787-1237">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="07787-1238">函数</span><span class="sxs-lookup"><span data-stu-id="07787-1238">function</span></span>||<span data-ttu-id="07787-1239">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="07787-1239">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="07787-1240">如果成功，该项目标识符将在 `asyncResult.value` 属性中提供。</span><span class="sxs-lookup"><span data-stu-id="07787-1240">On success, the item identifier is provided in the `asyncResult.value` property.</span></span>|

##### <a name="errors"></a><span data-ttu-id="07787-1241">错误</span><span class="sxs-lookup"><span data-stu-id="07787-1241">Errors</span></span>

|<span data-ttu-id="07787-1242">错误代码</span><span class="sxs-lookup"><span data-stu-id="07787-1242">Error code</span></span>|<span data-ttu-id="07787-1243">说明</span><span class="sxs-lookup"><span data-stu-id="07787-1243">Description</span></span>|
|------------|-------------|
|`ItemNotSaved`|<span data-ttu-id="07787-1244">在保存项目之前, 无法检索此 id。</span><span class="sxs-lookup"><span data-stu-id="07787-1244">The id can't be retrieved until the item is saved.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="07787-1245">要求</span><span class="sxs-lookup"><span data-stu-id="07787-1245">Requirements</span></span>

|<span data-ttu-id="07787-1246">要求</span><span class="sxs-lookup"><span data-stu-id="07787-1246">Requirement</span></span>|<span data-ttu-id="07787-1247">值</span><span class="sxs-lookup"><span data-stu-id="07787-1247">Value</span></span>|
|---|---|
|[<span data-ttu-id="07787-1248">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="07787-1248">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="07787-1249">预览</span><span class="sxs-lookup"><span data-stu-id="07787-1249">Preview</span></span>|
|[<span data-ttu-id="07787-1250">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="07787-1250">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="07787-1251">ReadItem</span><span class="sxs-lookup"><span data-stu-id="07787-1251">ReadItem</span></span>|
|[<span data-ttu-id="07787-1252">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="07787-1252">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="07787-1253">撰写</span><span class="sxs-lookup"><span data-stu-id="07787-1253">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="07787-1254">示例</span><span class="sxs-lookup"><span data-stu-id="07787-1254">Examples</span></span>

```javascript
Office.context.mailbox.item.getItemIdAsync(
  function callback(result) {
    // Process the result.
  });
```

<span data-ttu-id="07787-1255">下面的示例演示传递给回调函数`result`的参数的结构。</span><span class="sxs-lookup"><span data-stu-id="07787-1255">The following example shows the structure of the `result` parameter that's passed to the callback function.</span></span> <span data-ttu-id="07787-1256">`value`属性包含项 ID。</span><span class="sxs-lookup"><span data-stu-id="07787-1256">The `value` property contains the item ID.</span></span>

```json
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

---
---

#### <a name="getregexmatches--object"></a><span data-ttu-id="07787-1257">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="07787-1257">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="07787-1258">返回所选项目中匹配在清单 XML 文件中定义的正则表达式的字符串值。</span><span class="sxs-lookup"><span data-stu-id="07787-1258">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="07787-1259">IOS 或 Android 上的 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="07787-1259">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="07787-p168">`getRegExMatches` 方法返回匹配在清单 XML 文件中的每个 `ItemHasRegularExpressionMatch` 或 `ItemHasKnownEntity` 规则元素中定义的正则表达式的字符串。对于 `ItemHasRegularExpressionMatch` 规则，匹配字符串必须发生在该规则指定的项目的属性中。`PropertyName` 简单类型定义支持的属性。</span><span class="sxs-lookup"><span data-stu-id="07787-p168">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="07787-1263">例如，考虑一个外接程序清单具有以下 `Rule` 元素：</span><span class="sxs-lookup"><span data-stu-id="07787-1263">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="07787-1264">从 `getRegExMatches` 返回的对象应有两个属性：`fruits` 和 `veggies`。</span><span class="sxs-lookup"><span data-stu-id="07787-1264">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="07787-p169">如果在项目的正文属性上指定 `ItemHasRegularExpressionMatch` 规则，则正则表达式应进一步筛选正文，不应尝试返回该项目的整个正文。使用正则表达式（如 `.*`）获取项目的整个正文并不总是返回预期的结果。而是使用 [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) 方法检索整个正文。</span><span class="sxs-lookup"><span data-stu-id="07787-p169">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="07787-1268">Requirements</span><span class="sxs-lookup"><span data-stu-id="07787-1268">Requirements</span></span>

|<span data-ttu-id="07787-1269">要求</span><span class="sxs-lookup"><span data-stu-id="07787-1269">Requirement</span></span>|<span data-ttu-id="07787-1270">值</span><span class="sxs-lookup"><span data-stu-id="07787-1270">Value</span></span>|
|---|---|
|[<span data-ttu-id="07787-1271">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="07787-1271">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="07787-1272">1.0</span><span class="sxs-lookup"><span data-stu-id="07787-1272">1.0</span></span>|
|[<span data-ttu-id="07787-1273">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="07787-1273">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="07787-1274">ReadItem</span><span class="sxs-lookup"><span data-stu-id="07787-1274">ReadItem</span></span>|
|[<span data-ttu-id="07787-1275">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="07787-1275">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="07787-1276">阅读</span><span class="sxs-lookup"><span data-stu-id="07787-1276">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="07787-1277">返回：</span><span class="sxs-lookup"><span data-stu-id="07787-1277">Returns:</span></span>

<span data-ttu-id="07787-p170">一个包含匹配在清单 XML 文件中定义的正则表达式的字符串数组的对象。每个数组的名称等于匹配 `ItemHasRegularExpressionMatch` 规则的 `RegExName` 属性或匹配 `ItemHasKnownEntity` 规则的 `FilterName` 属性的相应值。</span><span class="sxs-lookup"><span data-stu-id="07787-p170">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<dl class="param-type"><span data-ttu-id="07787-1280">

<dt>
类型</dt>


</span><span class="sxs-lookup"><span data-stu-id="07787-1280">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="07787-1281">对象</span><span class="sxs-lookup"><span data-stu-id="07787-1281">Object</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="07787-1282">示例</span><span class="sxs-lookup"><span data-stu-id="07787-1282">Example</span></span>

<span data-ttu-id="07787-1283">以下示例显示了如何访问正则表达式规则元素 `fruits` 和 `veggies` 的匹配项的数组，这些元素在清单中指定。</span><span class="sxs-lookup"><span data-stu-id="07787-1283">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```javascript
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veggies = allMatches.veggies;
```

---
---

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="07787-1284">getRegExMatchesByName(name) → (nullable) {Array.<String>}</span><span class="sxs-lookup"><span data-stu-id="07787-1284">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="07787-1285">返回匹配在清单 XML 文件中定义的命名正则表达式的所选项目中的字符串值。</span><span class="sxs-lookup"><span data-stu-id="07787-1285">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="07787-1286">IOS 或 Android 上的 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="07787-1286">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="07787-1287">`getRegExMatchesByName` 方法返回匹配在具有指定 `RegExName` 元素值的清单 XML 文件中的 `ItemHasRegularExpressionMatch` 规则元素中定义的正则表达式的字符串。</span><span class="sxs-lookup"><span data-stu-id="07787-1287">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="07787-p171">如果在项目的正文属性上指定 `ItemHasRegularExpressionMatch` 规则，则正则表达式应进一步筛选正文，不应尝试返回该项目的整个正文。使用正则表达式（如 `.*`）获取项目的整个正文并不总是返回预期的结果。</span><span class="sxs-lookup"><span data-stu-id="07787-p171">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="07787-1290">参数</span><span class="sxs-lookup"><span data-stu-id="07787-1290">Parameters</span></span>

|<span data-ttu-id="07787-1291">名称</span><span class="sxs-lookup"><span data-stu-id="07787-1291">Name</span></span>|<span data-ttu-id="07787-1292">类型</span><span class="sxs-lookup"><span data-stu-id="07787-1292">Type</span></span>|<span data-ttu-id="07787-1293">说明</span><span class="sxs-lookup"><span data-stu-id="07787-1293">Description</span></span>|
|---|---|---|
|`name`|<span data-ttu-id="07787-1294">String</span><span class="sxs-lookup"><span data-stu-id="07787-1294">String</span></span>|<span data-ttu-id="07787-1295">定义筛选器匹配的 `ItemHasRegularExpressionMatch` 规则元素的名称。</span><span class="sxs-lookup"><span data-stu-id="07787-1295">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="07787-1296">要求</span><span class="sxs-lookup"><span data-stu-id="07787-1296">Requirements</span></span>

|<span data-ttu-id="07787-1297">要求</span><span class="sxs-lookup"><span data-stu-id="07787-1297">Requirement</span></span>|<span data-ttu-id="07787-1298">值</span><span class="sxs-lookup"><span data-stu-id="07787-1298">Value</span></span>|
|---|---|
|[<span data-ttu-id="07787-1299">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="07787-1299">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="07787-1300">1.0</span><span class="sxs-lookup"><span data-stu-id="07787-1300">1.0</span></span>|
|[<span data-ttu-id="07787-1301">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="07787-1301">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="07787-1302">ReadItem</span><span class="sxs-lookup"><span data-stu-id="07787-1302">ReadItem</span></span>|
|[<span data-ttu-id="07787-1303">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="07787-1303">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="07787-1304">阅读</span><span class="sxs-lookup"><span data-stu-id="07787-1304">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="07787-1305">返回：</span><span class="sxs-lookup"><span data-stu-id="07787-1305">Returns:</span></span>

<span data-ttu-id="07787-1306">一个包含匹配在清单 XML 文件中定义的正则表达式的字符串的数组。</span><span class="sxs-lookup"><span data-stu-id="07787-1306">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<dl class="param-type"><span data-ttu-id="07787-1307">

<dt>类型</dt>

</span><span class="sxs-lookup"><span data-stu-id="07787-1307">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="07787-1308">Array.< String ></span><span class="sxs-lookup"><span data-stu-id="07787-1308">Array.< String ></span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="07787-1309">示例</span><span class="sxs-lookup"><span data-stu-id="07787-1309">Example</span></span>

```javascript
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

---
---

#### <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="07787-1310">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="07787-1310">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="07787-1311">以异步方式返回邮件的主题或正文中选定的数据。</span><span class="sxs-lookup"><span data-stu-id="07787-1311">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="07787-p172">如果没有选定内容，但光标位于正文或主题中，此方法将会为所选数据返回 null。如果选定的是字段，而不是正文或主题，则此方法返回 `InvalidSelection` 错误。</span><span class="sxs-lookup"><span data-stu-id="07787-p172">If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="07787-1314">参数</span><span class="sxs-lookup"><span data-stu-id="07787-1314">Parameters</span></span>

|<span data-ttu-id="07787-1315">名称</span><span class="sxs-lookup"><span data-stu-id="07787-1315">Name</span></span>|<span data-ttu-id="07787-1316">类型</span><span class="sxs-lookup"><span data-stu-id="07787-1316">Type</span></span>|<span data-ttu-id="07787-1317">属性</span><span class="sxs-lookup"><span data-stu-id="07787-1317">Attributes</span></span>|<span data-ttu-id="07787-1318">说明</span><span class="sxs-lookup"><span data-stu-id="07787-1318">Description</span></span>|
|---|---|---|---|
|`coercionType`|[<span data-ttu-id="07787-1319">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="07787-1319">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="07787-p173">请求数据的格式。如果为文本，则此方法返回纯文本作为字符串，删除任何显示的 HTML 标记。如果为 HTML，则此方法返回所选文本，不论是纯文本还是 HTML。</span><span class="sxs-lookup"><span data-stu-id="07787-p173">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`|<span data-ttu-id="07787-1323">对象</span><span class="sxs-lookup"><span data-stu-id="07787-1323">Object</span></span>|<span data-ttu-id="07787-1324">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="07787-1324">&lt;optional&gt;</span></span>|<span data-ttu-id="07787-1325">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="07787-1325">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="07787-1326">Object</span><span class="sxs-lookup"><span data-stu-id="07787-1326">Object</span></span>|<span data-ttu-id="07787-1327">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="07787-1327">&lt;optional&gt;</span></span>|<span data-ttu-id="07787-1328">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="07787-1328">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="07787-1329">函数</span><span class="sxs-lookup"><span data-stu-id="07787-1329">function</span></span>||<span data-ttu-id="07787-1330">方法完成后，使用单个参数 `callback`（一个 [`asyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `AsyncResult` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="07787-1330">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="07787-1331">若要从回调方法访问所选数据，请调用 `asyncResult.value.data`。</span><span class="sxs-lookup"><span data-stu-id="07787-1331">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="07787-1332">若要访问选定内容的源属性，请调用 `asyncResult.value.sourceProperty`，这将为 `body` 或 `subject`。</span><span class="sxs-lookup"><span data-stu-id="07787-1332">To access the source property that the selection comes from, call `asyncResult.value.sourceProperty`, which will be either `body` or `subject`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="07787-1333">要求</span><span class="sxs-lookup"><span data-stu-id="07787-1333">Requirements</span></span>

|<span data-ttu-id="07787-1334">要求</span><span class="sxs-lookup"><span data-stu-id="07787-1334">Requirement</span></span>|<span data-ttu-id="07787-1335">值</span><span class="sxs-lookup"><span data-stu-id="07787-1335">Value</span></span>|
|---|---|
|[<span data-ttu-id="07787-1336">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="07787-1336">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="07787-1337">1.2</span><span class="sxs-lookup"><span data-stu-id="07787-1337">1.2</span></span>|
|[<span data-ttu-id="07787-1338">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="07787-1338">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="07787-1339">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="07787-1339">ReadWriteItem</span></span>|
|[<span data-ttu-id="07787-1340">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="07787-1340">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="07787-1341">撰写</span><span class="sxs-lookup"><span data-stu-id="07787-1341">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="07787-1342">返回：</span><span class="sxs-lookup"><span data-stu-id="07787-1342">Returns:</span></span>

<span data-ttu-id="07787-1343">作为字符串的所选数据的格式由 `coercionType` 确定。</span><span class="sxs-lookup"><span data-stu-id="07787-1343">The selected data as a string with format determined by `coercionType`.</span></span>

<dl class="param-type"><span data-ttu-id="07787-1344">

<dt>
类型</dt>


</span><span class="sxs-lookup"><span data-stu-id="07787-1344">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="07787-1345">字符串</span><span class="sxs-lookup"><span data-stu-id="07787-1345">String</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="07787-1346">示例</span><span class="sxs-lookup"><span data-stu-id="07787-1346">Example</span></span>

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

#### <a name="getselectedentities--entitiesjavascriptapioutlookofficeentities"></a><span data-ttu-id="07787-1347">getSelectedEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="07787-1347">getSelectedEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span></span>

<span data-ttu-id="07787-1348">获取在用户已选择的突出显示匹配项中找到的实体。</span><span class="sxs-lookup"><span data-stu-id="07787-1348">Gets the entities found in a highlighted match a user has selected.</span></span> <span data-ttu-id="07787-1349">突出显示匹配项适用于[上下文外接程序](/outlook/add-ins/contextual-outlook-add-ins)。</span><span class="sxs-lookup"><span data-stu-id="07787-1349">Highlighted matches apply to [contextual add-ins](/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="07787-1350">IOS 或 Android 上的 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="07787-1350">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="07787-1351">要求</span><span class="sxs-lookup"><span data-stu-id="07787-1351">Requirements</span></span>

|<span data-ttu-id="07787-1352">要求</span><span class="sxs-lookup"><span data-stu-id="07787-1352">Requirement</span></span>|<span data-ttu-id="07787-1353">值</span><span class="sxs-lookup"><span data-stu-id="07787-1353">Value</span></span>|
|---|---|
|[<span data-ttu-id="07787-1354">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="07787-1354">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="07787-1355">1.6</span><span class="sxs-lookup"><span data-stu-id="07787-1355">1.6</span></span>|
|[<span data-ttu-id="07787-1356">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="07787-1356">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="07787-1357">ReadItem</span><span class="sxs-lookup"><span data-stu-id="07787-1357">ReadItem</span></span>|
|[<span data-ttu-id="07787-1358">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="07787-1358">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="07787-1359">阅读</span><span class="sxs-lookup"><span data-stu-id="07787-1359">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="07787-1360">返回：</span><span class="sxs-lookup"><span data-stu-id="07787-1360">Returns:</span></span>

<span data-ttu-id="07787-1361">类型：[Entities](/javascript/api/outlook/office.entities)</span><span class="sxs-lookup"><span data-stu-id="07787-1361">Type: [Entities](/javascript/api/outlook/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="07787-1362">示例</span><span class="sxs-lookup"><span data-stu-id="07787-1362">Example</span></span>

<span data-ttu-id="07787-1363">以下示例访问用户选择的突出显示匹配项中的地址实体。</span><span class="sxs-lookup"><span data-stu-id="07787-1363">The following example accesses the addresses entities in the highlighted match selected by the user.</span></span>

```javascript
var contacts = Office.context.mailbox.item.getSelectedEntities().addresses;
```

---
---

#### <a name="getselectedregexmatches--object"></a><span data-ttu-id="07787-1364">getSelectedRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="07787-1364">getSelectedRegExMatches() → {Object}</span></span>

<span data-ttu-id="07787-p176">返回突出显示匹配项中匹配在清单 XML 文件中定义的正则表达式的字符串值。突出显示匹配项适用于[上下文外接程序](/outlook/add-ins/contextual-outlook-add-ins)。</span><span class="sxs-lookup"><span data-stu-id="07787-p176">Returns string values in a highlighted match that match the regular expressions defined in the manifest XML file. Highlighted matches apply to [contextual add-ins](/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="07787-1367">IOS 或 Android 上的 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="07787-1367">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="07787-p177">`getSelectedRegExMatches` 方法返回匹配在清单 XML 文件中的每个 `ItemHasRegularExpressionMatch` 或 `ItemHasKnownEntity` 规则元素中定义的正则表达式的字符串。对于 `ItemHasRegularExpressionMatch` 规则，匹配字符串必须发生在该规则指定的项目的属性中。`PropertyName` 简单类型定义支持的属性。</span><span class="sxs-lookup"><span data-stu-id="07787-p177">The `getSelectedRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="07787-1371">例如，考虑一个外接程序清单具有以下 `Rule` 元素：</span><span class="sxs-lookup"><span data-stu-id="07787-1371">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="07787-1372">从 `getRegExMatches` 返回的对象应有两个属性：`fruits` 和 `veggies`。</span><span class="sxs-lookup"><span data-stu-id="07787-1372">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="07787-p178">如果在项目的正文属性上指定 `ItemHasRegularExpressionMatch` 规则，则正则表达式应进一步筛选正文，不应尝试返回该项目的整个正文。使用正则表达式（如 `.*`）获取项目的整个正文并不总是返回预期的结果。而是使用 [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) 方法检索整个正文。</span><span class="sxs-lookup"><span data-stu-id="07787-p178">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="07787-1376">Requirements</span><span class="sxs-lookup"><span data-stu-id="07787-1376">Requirements</span></span>

|<span data-ttu-id="07787-1377">要求</span><span class="sxs-lookup"><span data-stu-id="07787-1377">Requirement</span></span>|<span data-ttu-id="07787-1378">值</span><span class="sxs-lookup"><span data-stu-id="07787-1378">Value</span></span>|
|---|---|
|[<span data-ttu-id="07787-1379">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="07787-1379">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="07787-1380">1.6</span><span class="sxs-lookup"><span data-stu-id="07787-1380">1.6</span></span>|
|[<span data-ttu-id="07787-1381">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="07787-1381">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="07787-1382">ReadItem</span><span class="sxs-lookup"><span data-stu-id="07787-1382">ReadItem</span></span>|
|[<span data-ttu-id="07787-1383">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="07787-1383">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="07787-1384">阅读</span><span class="sxs-lookup"><span data-stu-id="07787-1384">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="07787-1385">返回：</span><span class="sxs-lookup"><span data-stu-id="07787-1385">Returns:</span></span>

<span data-ttu-id="07787-p179">一个包含匹配在清单 XML 文件中定义的正则表达式的字符串数组的对象。每个数组的名称等于匹配 `ItemHasRegularExpressionMatch` 规则的 `RegExName` 属性或匹配 `ItemHasKnownEntity` 规则的 `FilterName` 属性的相应值。</span><span class="sxs-lookup"><span data-stu-id="07787-p179">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

##### <a name="example"></a><span data-ttu-id="07787-1388">示例</span><span class="sxs-lookup"><span data-stu-id="07787-1388">Example</span></span>

<span data-ttu-id="07787-1389">以下示例显示了如何访问正则表达式规则元素 `fruits` 和 `veggies` 的匹配项的数组，这些元素在清单中指定。</span><span class="sxs-lookup"><span data-stu-id="07787-1389">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```javascript
var selectedMatches = Office.context.mailbox.item.getSelectedRegExMatches();
var fruits = selectedMatches.fruits;
var veggies = selectedMatches.veggies;
```

---
---

#### <a name="getsharedpropertiesasyncoptions-callback"></a><span data-ttu-id="07787-1390">getSharedPropertiesAsync ([options], 回拨)</span><span class="sxs-lookup"><span data-stu-id="07787-1390">getSharedPropertiesAsync([options], callback)</span></span>

<span data-ttu-id="07787-1391">获取共享文件夹、日历或邮箱中的所选约会或邮件的属性。</span><span class="sxs-lookup"><span data-stu-id="07787-1391">Gets the properties of the selected appointment or message in a shared folder, calendar, or mailbox.</span></span>

##### <a name="parameters"></a><span data-ttu-id="07787-1392">参数</span><span class="sxs-lookup"><span data-stu-id="07787-1392">Parameters</span></span>

|<span data-ttu-id="07787-1393">名称</span><span class="sxs-lookup"><span data-stu-id="07787-1393">Name</span></span>|<span data-ttu-id="07787-1394">类型</span><span class="sxs-lookup"><span data-stu-id="07787-1394">Type</span></span>|<span data-ttu-id="07787-1395">属性</span><span class="sxs-lookup"><span data-stu-id="07787-1395">Attributes</span></span>|<span data-ttu-id="07787-1396">说明</span><span class="sxs-lookup"><span data-stu-id="07787-1396">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="07787-1397">对象</span><span class="sxs-lookup"><span data-stu-id="07787-1397">Object</span></span>|<span data-ttu-id="07787-1398">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="07787-1398">&lt;optional&gt;</span></span>|<span data-ttu-id="07787-1399">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="07787-1399">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="07787-1400">对象</span><span class="sxs-lookup"><span data-stu-id="07787-1400">Object</span></span>|<span data-ttu-id="07787-1401">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="07787-1401">&lt;optional&gt;</span></span>|<span data-ttu-id="07787-1402">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="07787-1402">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="07787-1403">function</span><span class="sxs-lookup"><span data-stu-id="07787-1403">function</span></span>||<span data-ttu-id="07787-1404">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="07787-1404">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="07787-1405">共享属性作为[`SharedProperties`](/javascript/api/outlook/office.sharedproperties) `asyncResult.value`属性中的对象提供。</span><span class="sxs-lookup"><span data-stu-id="07787-1405">The shared properties are provided as a [`SharedProperties`](/javascript/api/outlook/office.sharedproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="07787-1406">此对象可用于获取项目的共享属性。</span><span class="sxs-lookup"><span data-stu-id="07787-1406">This object can be used to get the item's shared properties.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="07787-1407">要求</span><span class="sxs-lookup"><span data-stu-id="07787-1407">Requirements</span></span>

|<span data-ttu-id="07787-1408">要求</span><span class="sxs-lookup"><span data-stu-id="07787-1408">Requirement</span></span>|<span data-ttu-id="07787-1409">值</span><span class="sxs-lookup"><span data-stu-id="07787-1409">Value</span></span>|
|---|---|
|[<span data-ttu-id="07787-1410">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="07787-1410">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="07787-1411">预览</span><span class="sxs-lookup"><span data-stu-id="07787-1411">Preview</span></span>|
|[<span data-ttu-id="07787-1412">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="07787-1412">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="07787-1413">ReadItem</span><span class="sxs-lookup"><span data-stu-id="07787-1413">ReadItem</span></span>|
|[<span data-ttu-id="07787-1414">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="07787-1414">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="07787-1415">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="07787-1415">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="07787-1416">示例</span><span class="sxs-lookup"><span data-stu-id="07787-1416">Example</span></span>

```javascript
Office.context.mailbox.item.getSharedPropertiesAsync(callback);

function callback (asyncResult) {
  var context = asyncResult.context;
  var sharedProperties = asyncResult.value;
}
```

---
---

#### <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="07787-1417">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="07787-1417">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="07787-1418">异步加载所选项目上此外接程序的自定义属性。</span><span class="sxs-lookup"><span data-stu-id="07787-1418">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="07787-p181">自定义属性基于每个应用、每个项目存储为键/值对。此方法在回调中返回 `CustomProperties` 对象，该回调提供访问特定于当前项目和当前外接程序的自定义属性的方法。自定义属性未在项目上加密，因此这不应用作安全存储。</span><span class="sxs-lookup"><span data-stu-id="07787-p181">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="07787-1422">参数</span><span class="sxs-lookup"><span data-stu-id="07787-1422">Parameters</span></span>

|<span data-ttu-id="07787-1423">名称</span><span class="sxs-lookup"><span data-stu-id="07787-1423">Name</span></span>|<span data-ttu-id="07787-1424">类型</span><span class="sxs-lookup"><span data-stu-id="07787-1424">Type</span></span>|<span data-ttu-id="07787-1425">属性</span><span class="sxs-lookup"><span data-stu-id="07787-1425">Attributes</span></span>|<span data-ttu-id="07787-1426">说明</span><span class="sxs-lookup"><span data-stu-id="07787-1426">Description</span></span>|
|---|---|---|---|
|`callback`|<span data-ttu-id="07787-1427">函数</span><span class="sxs-lookup"><span data-stu-id="07787-1427">function</span></span>||<span data-ttu-id="07787-1428">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="07787-1428">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="07787-1429">自定义属性作为 `asyncResult.value` 属性中的 [`CustomProperties`](/javascript/api/outlook/office.customproperties) 对象提供。</span><span class="sxs-lookup"><span data-stu-id="07787-1429">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook/office.customproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="07787-1430">此对象可用于获取、设置以及从项目中删除自定义属性，并将自定义属性集的更改重新保存到服务器。</span><span class="sxs-lookup"><span data-stu-id="07787-1430">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`|<span data-ttu-id="07787-1431">对象</span><span class="sxs-lookup"><span data-stu-id="07787-1431">Object</span></span>|<span data-ttu-id="07787-1432">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="07787-1432">&lt;optional&gt;</span></span>|<span data-ttu-id="07787-1433">开发人员可以提供他们想要在回调函数中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="07787-1433">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="07787-1434">此对象可以通过回调函数中的 `asyncResult.asyncContext` 属性进行访问。</span><span class="sxs-lookup"><span data-stu-id="07787-1434">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="07787-1435">要求</span><span class="sxs-lookup"><span data-stu-id="07787-1435">Requirements</span></span>

|<span data-ttu-id="07787-1436">要求</span><span class="sxs-lookup"><span data-stu-id="07787-1436">Requirement</span></span>|<span data-ttu-id="07787-1437">值</span><span class="sxs-lookup"><span data-stu-id="07787-1437">Value</span></span>|
|---|---|
|[<span data-ttu-id="07787-1438">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="07787-1438">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="07787-1439">1.0</span><span class="sxs-lookup"><span data-stu-id="07787-1439">1.0</span></span>|
|[<span data-ttu-id="07787-1440">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="07787-1440">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="07787-1441">ReadItem</span><span class="sxs-lookup"><span data-stu-id="07787-1441">ReadItem</span></span>|
|[<span data-ttu-id="07787-1442">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="07787-1442">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="07787-1443">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="07787-1443">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="07787-1444">示例</span><span class="sxs-lookup"><span data-stu-id="07787-1444">Example</span></span>

<span data-ttu-id="07787-p184">以下代码示例显示了如何使用 `loadCustomPropertiesAsync` 方法异步加载特定于当前项目的自定义属性。该示例还显示了如何使用 `CustomProperties.saveAsync` 方法将这些属性重新保存到服务器。加载自定义属性后，该代码示例将使用 `CustomProperties.get` 方法读取自定义属性 `myProp`，使用 `CustomProperties.set` 方法写入自定义属性 `otherProp`，最后调用 `saveAsync` 方法保存这些自定义属性。</span><span class="sxs-lookup"><span data-stu-id="07787-p184">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

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

#### <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="07787-1448">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="07787-1448">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="07787-1449">将附件从邮件或约会中删除。</span><span class="sxs-lookup"><span data-stu-id="07787-1449">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="07787-1450">`removeAttachmentAsync` 方法删除项目中带指定标识符的附件。</span><span class="sxs-lookup"><span data-stu-id="07787-1450">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item.</span></span> <span data-ttu-id="07787-1451">最佳做法是，仅当同一个邮件应用程序在同一会话中添加了一个附件时，你才应使用该附件标识符来删除该附件。</span><span class="sxs-lookup"><span data-stu-id="07787-1451">As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session.</span></span> <span data-ttu-id="07787-1452">在 web 和移动设备上的 Outlook 中, 附件标识符仅在同一个会话中有效。</span><span class="sxs-lookup"><span data-stu-id="07787-1452">In Outlook on the web and mobile devices, the attachment identifier is valid only within the same session.</span></span> <span data-ttu-id="07787-1453">当用户关闭应用程序时, 或者如果用户开始撰写内嵌窗体, 随后弹出窗体以继续在单独的窗口中。</span><span class="sxs-lookup"><span data-stu-id="07787-1453">A session is over when the user closes the app, or if the user starts composing an inline form then subsequently pops out the form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="07787-1454">参数</span><span class="sxs-lookup"><span data-stu-id="07787-1454">Parameters</span></span>

|<span data-ttu-id="07787-1455">名称</span><span class="sxs-lookup"><span data-stu-id="07787-1455">Name</span></span>|<span data-ttu-id="07787-1456">类型</span><span class="sxs-lookup"><span data-stu-id="07787-1456">Type</span></span>|<span data-ttu-id="07787-1457">属性</span><span class="sxs-lookup"><span data-stu-id="07787-1457">Attributes</span></span>|<span data-ttu-id="07787-1458">说明</span><span class="sxs-lookup"><span data-stu-id="07787-1458">Description</span></span>|
|---|---|---|---|
|`attachmentId`|<span data-ttu-id="07787-1459">字符串</span><span class="sxs-lookup"><span data-stu-id="07787-1459">String</span></span>||<span data-ttu-id="07787-1460">要删除的附件的标识符。</span><span class="sxs-lookup"><span data-stu-id="07787-1460">The identifier of the attachment to remove.</span></span>|
|`options`|<span data-ttu-id="07787-1461">对象</span><span class="sxs-lookup"><span data-stu-id="07787-1461">Object</span></span>|<span data-ttu-id="07787-1462">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="07787-1462">&lt;optional&gt;</span></span>|<span data-ttu-id="07787-1463">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="07787-1463">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="07787-1464">对象</span><span class="sxs-lookup"><span data-stu-id="07787-1464">Object</span></span>|<span data-ttu-id="07787-1465">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="07787-1465">&lt;optional&gt;</span></span>|<span data-ttu-id="07787-1466">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="07787-1466">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="07787-1467">函数</span><span class="sxs-lookup"><span data-stu-id="07787-1467">function</span></span>|<span data-ttu-id="07787-1468">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="07787-1468">&lt;optional&gt;</span></span>|<span data-ttu-id="07787-1469">方法完成后，使用单个参数 `callback`（一个 [`asyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `AsyncResult` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="07787-1469">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="07787-1470">如果删除附件失败，`asyncResult.error` 属性将包含一个说明失败原因的错误代码。</span><span class="sxs-lookup"><span data-stu-id="07787-1470">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="07787-1471">错误</span><span class="sxs-lookup"><span data-stu-id="07787-1471">Errors</span></span>

|<span data-ttu-id="07787-1472">错误代码</span><span class="sxs-lookup"><span data-stu-id="07787-1472">Error code</span></span>|<span data-ttu-id="07787-1473">说明</span><span class="sxs-lookup"><span data-stu-id="07787-1473">Description</span></span>|
|------------|-------------|
|`InvalidAttachmentId`|<span data-ttu-id="07787-1474">附件标识符不存在。</span><span class="sxs-lookup"><span data-stu-id="07787-1474">The attachment identifier does not exist.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="07787-1475">要求</span><span class="sxs-lookup"><span data-stu-id="07787-1475">Requirements</span></span>

|<span data-ttu-id="07787-1476">要求</span><span class="sxs-lookup"><span data-stu-id="07787-1476">Requirement</span></span>|<span data-ttu-id="07787-1477">值</span><span class="sxs-lookup"><span data-stu-id="07787-1477">Value</span></span>|
|---|---|
|[<span data-ttu-id="07787-1478">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="07787-1478">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="07787-1479">1.1</span><span class="sxs-lookup"><span data-stu-id="07787-1479">1.1</span></span>|
|[<span data-ttu-id="07787-1480">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="07787-1480">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="07787-1481">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="07787-1481">ReadWriteItem</span></span>|
|[<span data-ttu-id="07787-1482">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="07787-1482">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="07787-1483">撰写</span><span class="sxs-lookup"><span data-stu-id="07787-1483">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="07787-1484">示例</span><span class="sxs-lookup"><span data-stu-id="07787-1484">Example</span></span>

<span data-ttu-id="07787-1485">以下代码删除包含标识符 '0' 的附件。</span><span class="sxs-lookup"><span data-stu-id="07787-1485">The following code removes an attachment with an identifier of '0'.</span></span>

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

#### <a name="removehandlerasynceventtype-options-callback"></a><span data-ttu-id="07787-1486">removeHandlerAsync(eventType, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="07787-1486">removeHandlerAsync(eventType, [options], [callback])</span></span>

<span data-ttu-id="07787-1487">删除受支持事件类型的事件处理程序。</span><span class="sxs-lookup"><span data-stu-id="07787-1487">Removes the event handlers for a supported event type.</span></span>

<span data-ttu-id="07787-1488">目前, 受支持的事件`Office.EventType.AttachmentsChanged`类型`Office.EventType.AppointmentTimeChanged`是`Office.EventType.EnhancedLocationsChanged`、 `Office.EventType.RecipientsChanged`、、 `Office.EventType.RecurrenceChanged`和。</span><span class="sxs-lookup"><span data-stu-id="07787-1488">Currently the supported event types are `Office.EventType.AttachmentsChanged`, `Office.EventType.AppointmentTimeChanged`, `Office.EventType.EnhancedLocationsChanged`, `Office.EventType.RecipientsChanged`, and `Office.EventType.RecurrenceChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="07787-1489">参数</span><span class="sxs-lookup"><span data-stu-id="07787-1489">Parameters</span></span>

| <span data-ttu-id="07787-1490">名称</span><span class="sxs-lookup"><span data-stu-id="07787-1490">Name</span></span> | <span data-ttu-id="07787-1491">类型</span><span class="sxs-lookup"><span data-stu-id="07787-1491">Type</span></span> | <span data-ttu-id="07787-1492">属性</span><span class="sxs-lookup"><span data-stu-id="07787-1492">Attributes</span></span> | <span data-ttu-id="07787-1493">说明</span><span class="sxs-lookup"><span data-stu-id="07787-1493">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="07787-1494">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="07787-1494">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="07787-1495">应撤销处理程序的事件。</span><span class="sxs-lookup"><span data-stu-id="07787-1495">The event that should revoke the handler.</span></span> |
| `options` | <span data-ttu-id="07787-1496">对象</span><span class="sxs-lookup"><span data-stu-id="07787-1496">Object</span></span> | <span data-ttu-id="07787-1497">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="07787-1497">&lt;optional&gt;</span></span> | <span data-ttu-id="07787-1498">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="07787-1498">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="07787-1499">对象</span><span class="sxs-lookup"><span data-stu-id="07787-1499">Object</span></span> | <span data-ttu-id="07787-1500">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="07787-1500">&lt;optional&gt;</span></span> | <span data-ttu-id="07787-1501">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="07787-1501">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="07787-1502">函数</span><span class="sxs-lookup"><span data-stu-id="07787-1502">function</span></span>| <span data-ttu-id="07787-1503">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="07787-1503">&lt;optional&gt;</span></span>|<span data-ttu-id="07787-1504">方法完成后，使用单个参数 `callback`（一个 [`asyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `AsyncResult` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="07787-1504">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="07787-1505">Requirements</span><span class="sxs-lookup"><span data-stu-id="07787-1505">Requirements</span></span>

|<span data-ttu-id="07787-1506">要求</span><span class="sxs-lookup"><span data-stu-id="07787-1506">Requirement</span></span>| <span data-ttu-id="07787-1507">值</span><span class="sxs-lookup"><span data-stu-id="07787-1507">Value</span></span>|
|---|---|
|[<span data-ttu-id="07787-1508">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="07787-1508">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="07787-1509">1.7</span><span class="sxs-lookup"><span data-stu-id="07787-1509">1.7</span></span> |
|[<span data-ttu-id="07787-1510">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="07787-1510">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="07787-1511">ReadItem</span><span class="sxs-lookup"><span data-stu-id="07787-1511">ReadItem</span></span> |
|[<span data-ttu-id="07787-1512">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="07787-1512">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="07787-1513">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="07787-1513">Compose or Read</span></span> |

---
---

#### <a name="saveasyncoptions-callback"></a><span data-ttu-id="07787-1514">saveAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="07787-1514">saveAsync([options], callback)</span></span>

<span data-ttu-id="07787-1515">异步保存项目。</span><span class="sxs-lookup"><span data-stu-id="07787-1515">Asynchronously saves an item.</span></span>

<span data-ttu-id="07787-1516">调用时，此方法将当前邮件保存为草稿，并通过回调方法返回项目 ID。</span><span class="sxs-lookup"><span data-stu-id="07787-1516">When invoked, this method saves the current message as a draft and returns the item id via the callback method.</span></span> <span data-ttu-id="07787-1517">在 Outlook 网页或 Outlook 的联机模式中, 将项目保存到服务器。</span><span class="sxs-lookup"><span data-stu-id="07787-1517">In Outlook on the web or Outlook in online mode, the item is saved to the server.</span></span> <span data-ttu-id="07787-1518">在 Outlook 缓存模式下，该项目被保存到本地缓存中。</span><span class="sxs-lookup"><span data-stu-id="07787-1518">In Outlook in cached mode, the item is saved to the local cache.</span></span>

> [!NOTE]
> <span data-ttu-id="07787-1519">如果加载项在撰写模式下对某个项目调用 `saveAsync` 来获得 `itemId`，以便与 EWS 或 REST API 一同使用，请注意，当 Outlook 处于高速缓存模式时，可能需要一段时间项目才能真正同步到服务器。</span><span class="sxs-lookup"><span data-stu-id="07787-1519">If your add-in calls `saveAsync` on an item in compose mode in order to get an `itemId` to use with EWS or the REST API, be aware that when Outlook is in cached mode, it may take some time before the item is actually synced to the server.</span></span> <span data-ttu-id="07787-1520">在项目同步前，使用 `itemId` 将返回一个错误。</span><span class="sxs-lookup"><span data-stu-id="07787-1520">Until the item is synced, using the `itemId` will return an error.</span></span>

<span data-ttu-id="07787-p188">由于约会没有草稿状态，如果以撰写模式在约会中调用 `saveAsync`，则该项将被保存为用户日历中的正常约会。对于之前未保存过的新约会，则不会发送邀请。保存现有约会将向添加或删除的与会者发送更新。</span><span class="sxs-lookup"><span data-stu-id="07787-p188">Since appointments have no draft state, if `saveAsync` is called on an appointment in compose mode, the item will be saved as a normal appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent. Saving an existing appointment will send an update to added or removed attendees.</span></span>

> [!NOTE]
> <span data-ttu-id="07787-1524">以下客户端在撰写模式下对约会上的 `saveAsync` 具有不同的行为：</span><span class="sxs-lookup"><span data-stu-id="07787-1524">The following clients have different behavior for `saveAsync` on appointments in compose mode:</span></span>
>
> - <span data-ttu-id="07787-1525">Mac 上的 Outlook 不支持保存会议。</span><span class="sxs-lookup"><span data-stu-id="07787-1525">Outlook on Mac does not support saving a meeting.</span></span> <span data-ttu-id="07787-1526">在`saveAsync`撰写模式下从会议中调用时, 此方法将失败。</span><span class="sxs-lookup"><span data-stu-id="07787-1526">The `saveAsync` method fails when called from a meeting in compose mode.</span></span> <span data-ttu-id="07787-1527">若要解决此问题, 请参阅[使用 OFFICE JS API 将会议保存为 Outlook For Mac 中的草稿](https://support.microsoft.com/help/4505745)。</span><span class="sxs-lookup"><span data-stu-id="07787-1527">See [Cannot save a meeting as a draft in Outlook for Mac by using Office JS API](https://support.microsoft.com/help/4505745) for a workaround.</span></span>
> - <span data-ttu-id="07787-1528">在撰写模式下的约会上调用 `saveAsync` 时，Outlook 网页版始终发送邀请或更新。</span><span class="sxs-lookup"><span data-stu-id="07787-1528">Outlook on the web always sends an invitation or update when `saveAsync` is called on an appointment in compose mode.</span></span>

##### <a name="parameters"></a><span data-ttu-id="07787-1529">参数</span><span class="sxs-lookup"><span data-stu-id="07787-1529">Parameters</span></span>

|<span data-ttu-id="07787-1530">名称</span><span class="sxs-lookup"><span data-stu-id="07787-1530">Name</span></span>|<span data-ttu-id="07787-1531">类型</span><span class="sxs-lookup"><span data-stu-id="07787-1531">Type</span></span>|<span data-ttu-id="07787-1532">属性</span><span class="sxs-lookup"><span data-stu-id="07787-1532">Attributes</span></span>|<span data-ttu-id="07787-1533">说明</span><span class="sxs-lookup"><span data-stu-id="07787-1533">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="07787-1534">对象</span><span class="sxs-lookup"><span data-stu-id="07787-1534">Object</span></span>|<span data-ttu-id="07787-1535">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="07787-1535">&lt;optional&gt;</span></span>|<span data-ttu-id="07787-1536">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="07787-1536">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="07787-1537">对象</span><span class="sxs-lookup"><span data-stu-id="07787-1537">Object</span></span>|<span data-ttu-id="07787-1538">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="07787-1538">&lt;optional&gt;</span></span>|<span data-ttu-id="07787-1539">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="07787-1539">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="07787-1540">函数</span><span class="sxs-lookup"><span data-stu-id="07787-1540">function</span></span>||<span data-ttu-id="07787-1541">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="07787-1541">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="07787-1542">如果成功，该项目标识符将在 `asyncResult.value` 属性中提供。</span><span class="sxs-lookup"><span data-stu-id="07787-1542">On success, the item identifier is provided in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="07787-1543">要求</span><span class="sxs-lookup"><span data-stu-id="07787-1543">Requirements</span></span>

|<span data-ttu-id="07787-1544">要求</span><span class="sxs-lookup"><span data-stu-id="07787-1544">Requirement</span></span>|<span data-ttu-id="07787-1545">值</span><span class="sxs-lookup"><span data-stu-id="07787-1545">Value</span></span>|
|---|---|
|[<span data-ttu-id="07787-1546">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="07787-1546">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="07787-1547">1.3</span><span class="sxs-lookup"><span data-stu-id="07787-1547">1.3</span></span>|
|[<span data-ttu-id="07787-1548">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="07787-1548">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="07787-1549">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="07787-1549">ReadWriteItem</span></span>|
|[<span data-ttu-id="07787-1550">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="07787-1550">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="07787-1551">撰写</span><span class="sxs-lookup"><span data-stu-id="07787-1551">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="07787-1552">示例</span><span class="sxs-lookup"><span data-stu-id="07787-1552">Examples</span></span>

```javascript
Office.context.mailbox.item.saveAsync(
  function callback(result) {
    // Process the result.
  });
```

<span data-ttu-id="07787-p190">下面是传递给回调函数的 `result` 参数的示例。`value` 属性包含的项目的项目 ID。</span><span class="sxs-lookup"><span data-stu-id="07787-p190">The following is an example of the `result` parameter passed to the callback function. The `value` property contains the item ID of the item.</span></span>

```json
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

---
---

#### <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="07787-1555">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="07787-1555">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="07787-1556">以异步方式将数据插入到邮件的正文或主题中。</span><span class="sxs-lookup"><span data-stu-id="07787-1556">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="07787-p191">`setSelectedDataAsync` 方法将指定字符串插入到项目主题或正文的光标位置，或者，如果在编辑器中已选择文本，则该方法将替换选择的文本。如果光标不在正文或主题字段中，则返回错误。插入之后，光标会位于插入内容的末尾。</span><span class="sxs-lookup"><span data-stu-id="07787-p191">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="07787-1560">参数</span><span class="sxs-lookup"><span data-stu-id="07787-1560">Parameters</span></span>

|<span data-ttu-id="07787-1561">名称</span><span class="sxs-lookup"><span data-stu-id="07787-1561">Name</span></span>|<span data-ttu-id="07787-1562">类型</span><span class="sxs-lookup"><span data-stu-id="07787-1562">Type</span></span>|<span data-ttu-id="07787-1563">属性</span><span class="sxs-lookup"><span data-stu-id="07787-1563">Attributes</span></span>|<span data-ttu-id="07787-1564">说明</span><span class="sxs-lookup"><span data-stu-id="07787-1564">Description</span></span>|
|---|---|---|---|
|`data`|<span data-ttu-id="07787-1565">字符串</span><span class="sxs-lookup"><span data-stu-id="07787-1565">String</span></span>||<span data-ttu-id="07787-p192">要插入的数据。数据不得超过 1,000,000 个字符。如果传入的数据超过 1,000,000 个字符，则会引发 `ArgumentOutOfRange` 异常。</span><span class="sxs-lookup"><span data-stu-id="07787-p192">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`|<span data-ttu-id="07787-1569">Object</span><span class="sxs-lookup"><span data-stu-id="07787-1569">Object</span></span>|<span data-ttu-id="07787-1570">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="07787-1570">&lt;optional&gt;</span></span>|<span data-ttu-id="07787-1571">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="07787-1571">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="07787-1572">对象</span><span class="sxs-lookup"><span data-stu-id="07787-1572">Object</span></span>|<span data-ttu-id="07787-1573">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="07787-1573">&lt;optional&gt;</span></span>|<span data-ttu-id="07787-1574">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="07787-1574">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`|[<span data-ttu-id="07787-1575">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="07787-1575">Office.CoercionType</span></span>](office.md#coerciontype-string)|<span data-ttu-id="07787-1576">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="07787-1576">&lt;optional&gt;</span></span>|<span data-ttu-id="07787-1577">如果`text`为, 则当前样式应用于 web 上的 Outlook 和桌面客户端。</span><span class="sxs-lookup"><span data-stu-id="07787-1577">If `text`, the current style is applied in Outlook on the web and desktop clients.</span></span> <span data-ttu-id="07787-1578">如果该字段是 HTML 编辑器，则仅插入文本数据，即使数据为 HTML。</span><span class="sxs-lookup"><span data-stu-id="07787-1578">If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="07787-1579">如果`html`和字段支持 HTML (主题不), 则当前样式应用于 web 上的 outlook, 并且在 outlook 桌面客户端中应用了默认样式。</span><span class="sxs-lookup"><span data-stu-id="07787-1579">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook on the web and the default style is applied in Outlook desktop clients.</span></span> <span data-ttu-id="07787-1580">如果该字段是文本字段，则返回 `InvalidDataFormat` 错误。</span><span class="sxs-lookup"><span data-stu-id="07787-1580">If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="07787-1581">如果未设置 `coercionType`，则结果取决于该字段：如果该字段是 HTML，则使用 HTML；如果该字段是文本，则使用纯文本。</span><span class="sxs-lookup"><span data-stu-id="07787-1581">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`|<span data-ttu-id="07787-1582">function</span><span class="sxs-lookup"><span data-stu-id="07787-1582">function</span></span>||<span data-ttu-id="07787-1583">方法完成后，使用单个参数 `callback`（一个 [`asyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `AsyncResult` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="07787-1583">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="07787-1584">Requirements</span><span class="sxs-lookup"><span data-stu-id="07787-1584">Requirements</span></span>

|<span data-ttu-id="07787-1585">要求</span><span class="sxs-lookup"><span data-stu-id="07787-1585">Requirement</span></span>|<span data-ttu-id="07787-1586">值</span><span class="sxs-lookup"><span data-stu-id="07787-1586">Value</span></span>|
|---|---|
|[<span data-ttu-id="07787-1587">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="07787-1587">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="07787-1588">1.2</span><span class="sxs-lookup"><span data-stu-id="07787-1588">1.2</span></span>|
|[<span data-ttu-id="07787-1589">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="07787-1589">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="07787-1590">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="07787-1590">ReadWriteItem</span></span>|
|[<span data-ttu-id="07787-1591">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="07787-1591">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="07787-1592">撰写</span><span class="sxs-lookup"><span data-stu-id="07787-1592">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="07787-1593">示例</span><span class="sxs-lookup"><span data-stu-id="07787-1593">Example</span></span>

```javascript
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```
