---
title: "\"context.subname\"-\"邮箱\"-预览要求集"
description: ''
ms.date: 02/15/2019
localization_priority: Normal
ms.openlocfilehash: ed3926e7e77550f43b87306cf27cf1e96341bd82
ms.sourcegitcommit: f26778b596b6b022814c39601485ff676ed4e2fa
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 02/16/2019
ms.locfileid: "30068243"
---
# <a name="item"></a><span data-ttu-id="44017-102">item</span><span class="sxs-lookup"><span data-stu-id="44017-102">item</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmditem"></a><span data-ttu-id="44017-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span><span class="sxs-lookup"><span data-stu-id="44017-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span></span>

<span data-ttu-id="44017-p101">`item` 命名空间用于访问当前选定的邮件、会议请求或约会。可以通过使用 [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtype) 属性确定 `item` 的类型。</span><span class="sxs-lookup"><span data-stu-id="44017-p101">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="44017-106">Requirements</span><span class="sxs-lookup"><span data-stu-id="44017-106">Requirements</span></span>

|<span data-ttu-id="44017-107">要求</span><span class="sxs-lookup"><span data-stu-id="44017-107">Requirement</span></span>|<span data-ttu-id="44017-108">值</span><span class="sxs-lookup"><span data-stu-id="44017-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="44017-109">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="44017-109">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="44017-110">1.0</span><span class="sxs-lookup"><span data-stu-id="44017-110">1.0</span></span>|
|[<span data-ttu-id="44017-111">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="44017-111">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="44017-112">受限</span><span class="sxs-lookup"><span data-stu-id="44017-112">Restricted</span></span>|
|[<span data-ttu-id="44017-113">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="44017-113">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="44017-114">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="44017-114">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="44017-115">成员和方法</span><span class="sxs-lookup"><span data-stu-id="44017-115">Members and methods</span></span>

| <span data-ttu-id="44017-116">成员</span><span class="sxs-lookup"><span data-stu-id="44017-116">Member</span></span> | <span data-ttu-id="44017-117">类型</span><span class="sxs-lookup"><span data-stu-id="44017-117">Type</span></span> |
|--------|------|
| [<span data-ttu-id="44017-118">attachments</span><span class="sxs-lookup"><span data-stu-id="44017-118">attachments</span></span>](#attachments-arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetails) | <span data-ttu-id="44017-119">成员</span><span class="sxs-lookup"><span data-stu-id="44017-119">Member</span></span> |
| [<span data-ttu-id="44017-120">bcc</span><span class="sxs-lookup"><span data-stu-id="44017-120">bcc</span></span>](#bcc-recipientsjavascriptapioutlookofficerecipients) | <span data-ttu-id="44017-121">成员</span><span class="sxs-lookup"><span data-stu-id="44017-121">Member</span></span> |
| [<span data-ttu-id="44017-122">body</span><span class="sxs-lookup"><span data-stu-id="44017-122">body</span></span>](#body-bodyjavascriptapioutlookofficebody) | <span data-ttu-id="44017-123">成员</span><span class="sxs-lookup"><span data-stu-id="44017-123">Member</span></span> |
| [<span data-ttu-id="44017-124">cc</span><span class="sxs-lookup"><span data-stu-id="44017-124">cc</span></span>](#cc-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients) | <span data-ttu-id="44017-125">成员</span><span class="sxs-lookup"><span data-stu-id="44017-125">Member</span></span> |
| [<span data-ttu-id="44017-126">conversationId</span><span class="sxs-lookup"><span data-stu-id="44017-126">conversationId</span></span>](#nullable-conversationid-string) | <span data-ttu-id="44017-127">成员</span><span class="sxs-lookup"><span data-stu-id="44017-127">Member</span></span> |
| [<span data-ttu-id="44017-128">dateTimeCreated</span><span class="sxs-lookup"><span data-stu-id="44017-128">dateTimeCreated</span></span>](#datetimecreated-date) | <span data-ttu-id="44017-129">成员</span><span class="sxs-lookup"><span data-stu-id="44017-129">Member</span></span> |
| [<span data-ttu-id="44017-130">dateTimeModified</span><span class="sxs-lookup"><span data-stu-id="44017-130">dateTimeModified</span></span>](#datetimemodified-date) | <span data-ttu-id="44017-131">成员</span><span class="sxs-lookup"><span data-stu-id="44017-131">Member</span></span> |
| [<span data-ttu-id="44017-132">end</span><span class="sxs-lookup"><span data-stu-id="44017-132">end</span></span>](#end-datetimejavascriptapioutlookofficetime) | <span data-ttu-id="44017-133">成员</span><span class="sxs-lookup"><span data-stu-id="44017-133">Member</span></span> |
| [<span data-ttu-id="44017-134">enhancedLocation</span><span class="sxs-lookup"><span data-stu-id="44017-134">enhancedLocation</span></span>](#enhancedlocation-enhancedlocationjavascriptapioutlookofficeenhancedlocation) | <span data-ttu-id="44017-135">成员</span><span class="sxs-lookup"><span data-stu-id="44017-135">Member</span></span> |
| [<span data-ttu-id="44017-136">from</span><span class="sxs-lookup"><span data-stu-id="44017-136">from</span></span>](#from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsfromjavascriptapioutlookofficefrom) | <span data-ttu-id="44017-137">成员</span><span class="sxs-lookup"><span data-stu-id="44017-137">Member</span></span> |
| [<span data-ttu-id="44017-138">internetHeaders</span><span class="sxs-lookup"><span data-stu-id="44017-138">internetHeaders</span></span>](#internetheaders-internetheadersjavascriptapioutlookofficeinternetheaders) | <span data-ttu-id="44017-139">成员</span><span class="sxs-lookup"><span data-stu-id="44017-139">Member</span></span> |
| [<span data-ttu-id="44017-140">internetMessageId</span><span class="sxs-lookup"><span data-stu-id="44017-140">internetMessageId</span></span>](#internetmessageid-string) | <span data-ttu-id="44017-141">成员</span><span class="sxs-lookup"><span data-stu-id="44017-141">Member</span></span> |
| [<span data-ttu-id="44017-142">itemClass</span><span class="sxs-lookup"><span data-stu-id="44017-142">itemClass</span></span>](#itemclass-string) | <span data-ttu-id="44017-143">成员</span><span class="sxs-lookup"><span data-stu-id="44017-143">Member</span></span> |
| [<span data-ttu-id="44017-144">itemId</span><span class="sxs-lookup"><span data-stu-id="44017-144">itemId</span></span>](#nullable-itemid-string) | <span data-ttu-id="44017-145">成员</span><span class="sxs-lookup"><span data-stu-id="44017-145">Member</span></span> |
| [<span data-ttu-id="44017-146">itemType</span><span class="sxs-lookup"><span data-stu-id="44017-146">itemType</span></span>](#itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtype) | <span data-ttu-id="44017-147">成员</span><span class="sxs-lookup"><span data-stu-id="44017-147">Member</span></span> |
| [<span data-ttu-id="44017-148">location</span><span class="sxs-lookup"><span data-stu-id="44017-148">location</span></span>](#location-stringlocationjavascriptapioutlookofficelocation) | <span data-ttu-id="44017-149">成员</span><span class="sxs-lookup"><span data-stu-id="44017-149">Member</span></span> |
| [<span data-ttu-id="44017-150">normalizedSubject</span><span class="sxs-lookup"><span data-stu-id="44017-150">normalizedSubject</span></span>](#normalizedsubject-string) | <span data-ttu-id="44017-151">成员</span><span class="sxs-lookup"><span data-stu-id="44017-151">Member</span></span> |
| [<span data-ttu-id="44017-152">notificationMessages</span><span class="sxs-lookup"><span data-stu-id="44017-152">notificationMessages</span></span>](#notificationmessages-notificationmessagesjavascriptapioutlookofficenotificationmessages) | <span data-ttu-id="44017-153">成员</span><span class="sxs-lookup"><span data-stu-id="44017-153">Member</span></span> |
| [<span data-ttu-id="44017-154">optionalAttendees</span><span class="sxs-lookup"><span data-stu-id="44017-154">optionalAttendees</span></span>](#optionalattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients) | <span data-ttu-id="44017-155">Member</span><span class="sxs-lookup"><span data-stu-id="44017-155">Member</span></span> |
| [<span data-ttu-id="44017-156">organizer</span><span class="sxs-lookup"><span data-stu-id="44017-156">organizer</span></span>](#organizer-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsorganizerjavascriptapioutlookofficeorganizer) | <span data-ttu-id="44017-157">成员</span><span class="sxs-lookup"><span data-stu-id="44017-157">Member</span></span> |
| [<span data-ttu-id="44017-158">recurrence</span><span class="sxs-lookup"><span data-stu-id="44017-158">recurrence</span></span>](#nullable-recurrence-recurrencejavascriptapioutlookofficerecurrence) | <span data-ttu-id="44017-159">成员</span><span class="sxs-lookup"><span data-stu-id="44017-159">Member</span></span> |
| [<span data-ttu-id="44017-160">requiredAttendees</span><span class="sxs-lookup"><span data-stu-id="44017-160">requiredAttendees</span></span>](#requiredattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients) | <span data-ttu-id="44017-161">成员</span><span class="sxs-lookup"><span data-stu-id="44017-161">Member</span></span> |
| [<span data-ttu-id="44017-162">sender</span><span class="sxs-lookup"><span data-stu-id="44017-162">sender</span></span>](#sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetails) | <span data-ttu-id="44017-163">成员</span><span class="sxs-lookup"><span data-stu-id="44017-163">Member</span></span> |
| [<span data-ttu-id="44017-164">seriesId</span><span class="sxs-lookup"><span data-stu-id="44017-164">seriesId</span></span>](#nullable-seriesid-string) | <span data-ttu-id="44017-165">Member</span><span class="sxs-lookup"><span data-stu-id="44017-165">Member</span></span> |
| [<span data-ttu-id="44017-166">start</span><span class="sxs-lookup"><span data-stu-id="44017-166">start</span></span>](#start-datetimejavascriptapioutlookofficetime) | <span data-ttu-id="44017-167">成员</span><span class="sxs-lookup"><span data-stu-id="44017-167">Member</span></span> |
| [<span data-ttu-id="44017-168">subject</span><span class="sxs-lookup"><span data-stu-id="44017-168">subject</span></span>](#subject-stringsubjectjavascriptapioutlookofficesubject) | <span data-ttu-id="44017-169">成员</span><span class="sxs-lookup"><span data-stu-id="44017-169">Member</span></span> |
| [<span data-ttu-id="44017-170">to</span><span class="sxs-lookup"><span data-stu-id="44017-170">to</span></span>](#to-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients) | <span data-ttu-id="44017-171">成员</span><span class="sxs-lookup"><span data-stu-id="44017-171">Member</span></span> |
| [<span data-ttu-id="44017-172">addFileAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="44017-172">addFileAttachmentAsync</span></span>](#addfileattachmentasyncuri-attachmentname-options-callback) | <span data-ttu-id="44017-173">方法</span><span class="sxs-lookup"><span data-stu-id="44017-173">Method</span></span> |
| [<span data-ttu-id="44017-174">addFileAttachmentFromBase64Async</span><span class="sxs-lookup"><span data-stu-id="44017-174">addFileAttachmentFromBase64Async</span></span>](#addfileattachmentfrombase64asyncbase64file-attachmentname-options-callback) | <span data-ttu-id="44017-175">方法</span><span class="sxs-lookup"><span data-stu-id="44017-175">Method</span></span> |
| [<span data-ttu-id="44017-176">addHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="44017-176">addHandlerAsync</span></span>](#addhandlerasynceventtype-handler-options-callback) | <span data-ttu-id="44017-177">方法</span><span class="sxs-lookup"><span data-stu-id="44017-177">Method</span></span> |
| [<span data-ttu-id="44017-178">addItemAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="44017-178">addItemAttachmentAsync</span></span>](#additemattachmentasyncitemid-attachmentname-options-callback) | <span data-ttu-id="44017-179">方法</span><span class="sxs-lookup"><span data-stu-id="44017-179">Method</span></span> |
| [<span data-ttu-id="44017-180">close</span><span class="sxs-lookup"><span data-stu-id="44017-180">close</span></span>](#close) | <span data-ttu-id="44017-181">方法</span><span class="sxs-lookup"><span data-stu-id="44017-181">Method</span></span> |
| [<span data-ttu-id="44017-182">displayReplyAllForm</span><span class="sxs-lookup"><span data-stu-id="44017-182">displayReplyAllForm</span></span>](#displayreplyallformformdata-callback) | <span data-ttu-id="44017-183">方法</span><span class="sxs-lookup"><span data-stu-id="44017-183">Method</span></span> |
| [<span data-ttu-id="44017-184">displayReplyForm</span><span class="sxs-lookup"><span data-stu-id="44017-184">displayReplyForm</span></span>](#displayreplyformformdata-callback) | <span data-ttu-id="44017-185">方法</span><span class="sxs-lookup"><span data-stu-id="44017-185">Method</span></span> |
| [<span data-ttu-id="44017-186">getAttachmentContentAsync</span><span class="sxs-lookup"><span data-stu-id="44017-186">getAttachmentContentAsync</span></span>](#getattachmentcontentasyncattachmentid-options-callback--attachmentcontentjavascriptapioutlookofficeattachmentcontent) | <span data-ttu-id="44017-187">方法</span><span class="sxs-lookup"><span data-stu-id="44017-187">Method</span></span> |
| [<span data-ttu-id="44017-188">getAttachmentsAsync</span><span class="sxs-lookup"><span data-stu-id="44017-188">getAttachmentsAsync</span></span>](#getattachmentsasyncoptions-callback--arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetails) | <span data-ttu-id="44017-189">方法</span><span class="sxs-lookup"><span data-stu-id="44017-189">Method</span></span> |
| [<span data-ttu-id="44017-190">getEntities</span><span class="sxs-lookup"><span data-stu-id="44017-190">getEntities</span></span>](#getentities--entitiesjavascriptapioutlookofficeentities) | <span data-ttu-id="44017-191">方法</span><span class="sxs-lookup"><span data-stu-id="44017-191">Method</span></span> |
| [<span data-ttu-id="44017-192">getEntitiesByType</span><span class="sxs-lookup"><span data-stu-id="44017-192">getEntitiesByType</span></span>](#getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlookofficecontactmeetingsuggestionjavascriptapioutlookofficemeetingsuggestionphonenumberjavascriptapioutlookofficephonenumbertasksuggestionjavascriptapioutlookofficetasksuggestion) | <span data-ttu-id="44017-193">方法</span><span class="sxs-lookup"><span data-stu-id="44017-193">Method</span></span> |
| [<span data-ttu-id="44017-194">getFilteredEntitiesByName</span><span class="sxs-lookup"><span data-stu-id="44017-194">getFilteredEntitiesByName</span></span>](#getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlookofficecontactmeetingsuggestionjavascriptapioutlookofficemeetingsuggestionphonenumberjavascriptapioutlookofficephonenumbertasksuggestionjavascriptapioutlookofficetasksuggestion) | <span data-ttu-id="44017-195">方法</span><span class="sxs-lookup"><span data-stu-id="44017-195">Method</span></span> |
| [<span data-ttu-id="44017-196">getInitializationContextAsync</span><span class="sxs-lookup"><span data-stu-id="44017-196">getInitializationContextAsync</span></span>](#getinitializationcontextasyncoptions-callback) | <span data-ttu-id="44017-197">方法</span><span class="sxs-lookup"><span data-stu-id="44017-197">Method</span></span> |
| [<span data-ttu-id="44017-198">getRegExMatches</span><span class="sxs-lookup"><span data-stu-id="44017-198">getRegExMatches</span></span>](#getregexmatches--object) | <span data-ttu-id="44017-199">方法</span><span class="sxs-lookup"><span data-stu-id="44017-199">Method</span></span> |
| [<span data-ttu-id="44017-200">getRegExMatchesByName</span><span class="sxs-lookup"><span data-stu-id="44017-200">getRegExMatchesByName</span></span>](#getregexmatchesbynamename--nullable-array-string-) | <span data-ttu-id="44017-201">方法</span><span class="sxs-lookup"><span data-stu-id="44017-201">Method</span></span> |
| [<span data-ttu-id="44017-202">getSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="44017-202">getSelectedDataAsync</span></span>](#getselecteddataasynccoerciontype-options-callback--string) | <span data-ttu-id="44017-203">方法</span><span class="sxs-lookup"><span data-stu-id="44017-203">Method</span></span> |
| [<span data-ttu-id="44017-204">getSelectedEntities</span><span class="sxs-lookup"><span data-stu-id="44017-204">getSelectedEntities</span></span>](#getselectedentities--entitiesjavascriptapioutlookofficeentities) | <span data-ttu-id="44017-205">方法</span><span class="sxs-lookup"><span data-stu-id="44017-205">Method</span></span> |
| [<span data-ttu-id="44017-206">getSelectedRegExMatches</span><span class="sxs-lookup"><span data-stu-id="44017-206">getSelectedRegExMatches</span></span>](#getselectedregexmatches--object) | <span data-ttu-id="44017-207">方法</span><span class="sxs-lookup"><span data-stu-id="44017-207">Method</span></span> |
| [<span data-ttu-id="44017-208">getSharedPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="44017-208">getSharedPropertiesAsync</span></span>](#getsharedpropertiesasyncoptions-callback) | <span data-ttu-id="44017-209">方法</span><span class="sxs-lookup"><span data-stu-id="44017-209">Method</span></span> |
| [<span data-ttu-id="44017-210">loadCustomPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="44017-210">loadCustomPropertiesAsync</span></span>](#loadcustompropertiesasynccallback-usercontext) | <span data-ttu-id="44017-211">方法</span><span class="sxs-lookup"><span data-stu-id="44017-211">Method</span></span> |
| [<span data-ttu-id="44017-212">removeAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="44017-212">removeAttachmentAsync</span></span>](#removeattachmentasyncattachmentid-options-callback) | <span data-ttu-id="44017-213">方法</span><span class="sxs-lookup"><span data-stu-id="44017-213">Method</span></span> |
| [<span data-ttu-id="44017-214">removeHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="44017-214">removeHandlerAsync</span></span>](#removehandlerasynceventtype-options-callback) | <span data-ttu-id="44017-215">方法</span><span class="sxs-lookup"><span data-stu-id="44017-215">Method</span></span> |
| [<span data-ttu-id="44017-216">saveAsync</span><span class="sxs-lookup"><span data-stu-id="44017-216">saveAsync</span></span>](#saveasyncoptions-callback) | <span data-ttu-id="44017-217">方法</span><span class="sxs-lookup"><span data-stu-id="44017-217">Method</span></span> |
| [<span data-ttu-id="44017-218">setSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="44017-218">setSelectedDataAsync</span></span>](#setselecteddataasyncdata-options-callback) | <span data-ttu-id="44017-219">方法</span><span class="sxs-lookup"><span data-stu-id="44017-219">Method</span></span> |

### <a name="example"></a><span data-ttu-id="44017-220">示例</span><span class="sxs-lookup"><span data-stu-id="44017-220">Example</span></span>

<span data-ttu-id="44017-221">以下 JavaScript 代码示例显示了如何访问 Outlook 中当前项目的 `subject` 属性。</span><span class="sxs-lookup"><span data-stu-id="44017-221">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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

### <a name="members"></a><span data-ttu-id="44017-222">成员</span><span class="sxs-lookup"><span data-stu-id="44017-222">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetails"></a><span data-ttu-id="44017-223">attachments :Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="44017-223">attachments :Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span></span>

<span data-ttu-id="44017-224">获取项目的附件作为数组。</span><span class="sxs-lookup"><span data-stu-id="44017-224">Gets the item's attachments as an array.</span></span> <span data-ttu-id="44017-225">仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="44017-225">Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="44017-226">由于潜在的安全问题，某些类型的文件会受到 Outlook 阻止，并且不会返回。</span><span class="sxs-lookup"><span data-stu-id="44017-226">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="44017-227">如需了解更多信息，请参阅 [Outlook 中阻止的附件](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519)。</span><span class="sxs-lookup"><span data-stu-id="44017-227">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="44017-228">类型</span><span class="sxs-lookup"><span data-stu-id="44017-228">Type</span></span>

*   <span data-ttu-id="44017-229">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="44017-229">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span></span>

##### <a name="requirements"></a><span data-ttu-id="44017-230">要求</span><span class="sxs-lookup"><span data-stu-id="44017-230">Requirements</span></span>

|<span data-ttu-id="44017-231">要求</span><span class="sxs-lookup"><span data-stu-id="44017-231">Requirement</span></span>|<span data-ttu-id="44017-232">值</span><span class="sxs-lookup"><span data-stu-id="44017-232">Value</span></span>|
|---|---|
|[<span data-ttu-id="44017-233">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="44017-233">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="44017-234">1.0</span><span class="sxs-lookup"><span data-stu-id="44017-234">1.0</span></span>|
|[<span data-ttu-id="44017-235">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="44017-235">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="44017-236">ReadItem</span><span class="sxs-lookup"><span data-stu-id="44017-236">ReadItem</span></span>|
|[<span data-ttu-id="44017-237">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="44017-237">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="44017-238">阅读</span><span class="sxs-lookup"><span data-stu-id="44017-238">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="44017-239">示例</span><span class="sxs-lookup"><span data-stu-id="44017-239">Example</span></span>

<span data-ttu-id="44017-240">以下代码使用当前项目上所有附件的详细信息构成 HTML 字符串。</span><span class="sxs-lookup"><span data-stu-id="44017-240">The following code builds an HTML string with details of all attachments on the current item.</span></span>

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

####  <a name="bcc-recipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="44017-241">bcc :[Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="44017-241">bcc :[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="44017-242">获取对象，该对象提供用于获取或更新邮件的密件抄送 (Bcc) 行上的收件人的方法。</span><span class="sxs-lookup"><span data-stu-id="44017-242">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="44017-243">仅限撰写模式。</span><span class="sxs-lookup"><span data-stu-id="44017-243">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="44017-244">类型</span><span class="sxs-lookup"><span data-stu-id="44017-244">Type</span></span>

*   [<span data-ttu-id="44017-245">收件人</span><span class="sxs-lookup"><span data-stu-id="44017-245">Recipients</span></span>](/javascript/api/outlook/office.recipients)

##### <a name="requirements"></a><span data-ttu-id="44017-246">要求</span><span class="sxs-lookup"><span data-stu-id="44017-246">Requirements</span></span>

|<span data-ttu-id="44017-247">要求</span><span class="sxs-lookup"><span data-stu-id="44017-247">Requirement</span></span>|<span data-ttu-id="44017-248">值</span><span class="sxs-lookup"><span data-stu-id="44017-248">Value</span></span>|
|---|---|
|[<span data-ttu-id="44017-249">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="44017-249">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="44017-250">1.1</span><span class="sxs-lookup"><span data-stu-id="44017-250">1.1</span></span>|
|[<span data-ttu-id="44017-251">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="44017-251">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="44017-252">ReadItem</span><span class="sxs-lookup"><span data-stu-id="44017-252">ReadItem</span></span>|
|[<span data-ttu-id="44017-253">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="44017-253">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="44017-254">撰写</span><span class="sxs-lookup"><span data-stu-id="44017-254">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="44017-255">示例</span><span class="sxs-lookup"><span data-stu-id="44017-255">Example</span></span>

```javascript
Office.context.mailbox.item.bcc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.bcc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.bcc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfBccRecipients = asyncResult.value;
}
```

####  <a name="body-bodyjavascriptapioutlookofficebody"></a><span data-ttu-id="44017-256">body :[Body](/javascript/api/outlook/office.body)</span><span class="sxs-lookup"><span data-stu-id="44017-256">body :[Body](/javascript/api/outlook/office.body)</span></span>

<span data-ttu-id="44017-257">获取一个提供用于处理项目正文的方法的对象。</span><span class="sxs-lookup"><span data-stu-id="44017-257">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="44017-258">类型</span><span class="sxs-lookup"><span data-stu-id="44017-258">Type</span></span>

*   [<span data-ttu-id="44017-259">Body</span><span class="sxs-lookup"><span data-stu-id="44017-259">Body</span></span>](/javascript/api/outlook/office.body)

##### <a name="requirements"></a><span data-ttu-id="44017-260">要求</span><span class="sxs-lookup"><span data-stu-id="44017-260">Requirements</span></span>

|<span data-ttu-id="44017-261">要求</span><span class="sxs-lookup"><span data-stu-id="44017-261">Requirement</span></span>|<span data-ttu-id="44017-262">值</span><span class="sxs-lookup"><span data-stu-id="44017-262">Value</span></span>|
|---|---|
|[<span data-ttu-id="44017-263">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="44017-263">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="44017-264">1.1</span><span class="sxs-lookup"><span data-stu-id="44017-264">1.1</span></span>|
|[<span data-ttu-id="44017-265">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="44017-265">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="44017-266">ReadItem</span><span class="sxs-lookup"><span data-stu-id="44017-266">ReadItem</span></span>|
|[<span data-ttu-id="44017-267">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="44017-267">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="44017-268">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="44017-268">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="44017-269">示例</span><span class="sxs-lookup"><span data-stu-id="44017-269">Example</span></span>

<span data-ttu-id="44017-270">本示例获取纯文本格式的邮件正文。</span><span class="sxs-lookup"><span data-stu-id="44017-270">This example gets the body of the message in plain text.</span></span>

```javascript
Office.context.mailbox.item.body.getAsync(
  "text",
  { asyncContext: "This is passed to the callback" },
  function callback(result) {
    // Do something with the result.
  });

```

<span data-ttu-id="44017-271">下面是传递给回调函数的 result 参数的示例。</span><span class="sxs-lookup"><span data-stu-id="44017-271">The following is an example of the result parameter passed to the callback function.</span></span>

```json
{
  "value": "TEXT of whole body (including threads below)",
  "status": "succeeded",
  "asyncContext": "This is passed to the callback"
}
```

####  <a name="cc-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="44017-272">cc :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="44017-272">cc :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="44017-273">提供对邮件的抄送 (Cc) 收件人的访问权限。</span><span class="sxs-lookup"><span data-stu-id="44017-273">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="44017-274">对象的类型和访问级别取决于当前项目的模式。</span><span class="sxs-lookup"><span data-stu-id="44017-274">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="44017-275">读取模式</span><span class="sxs-lookup"><span data-stu-id="44017-275">Read mode</span></span>

<span data-ttu-id="44017-p106">`cc` 属性返回包含邮件的**抄送**行上所列的每个收件人的 `EmailAddressDetails` 对象的数组。集合上限为 100 个成员。</span><span class="sxs-lookup"><span data-stu-id="44017-p106">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.</span></span>

```javascript
console.log(JSON.stringify(Office.context.mailbox.item.cc));
```

##### <a name="compose-mode"></a><span data-ttu-id="44017-278">撰写模式</span><span class="sxs-lookup"><span data-stu-id="44017-278">Compose mode</span></span>

<span data-ttu-id="44017-279">`cc` 属性返回一个 `Recipients` 对象，该对象提供用于获取或更新邮件的“**抄送**”行上收件人的方法。</span><span class="sxs-lookup"><span data-stu-id="44017-279">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span>

```javascript
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="44017-280">类型</span><span class="sxs-lookup"><span data-stu-id="44017-280">Type</span></span>

*   <span data-ttu-id="44017-281">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="44017-281">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="44017-282">Requirements</span><span class="sxs-lookup"><span data-stu-id="44017-282">Requirements</span></span>

|<span data-ttu-id="44017-283">要求</span><span class="sxs-lookup"><span data-stu-id="44017-283">Requirement</span></span>|<span data-ttu-id="44017-284">值</span><span class="sxs-lookup"><span data-stu-id="44017-284">Value</span></span>|
|---|---|
|[<span data-ttu-id="44017-285">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="44017-285">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="44017-286">1.0</span><span class="sxs-lookup"><span data-stu-id="44017-286">1.0</span></span>|
|[<span data-ttu-id="44017-287">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="44017-287">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="44017-288">ReadItem</span><span class="sxs-lookup"><span data-stu-id="44017-288">ReadItem</span></span>|
|[<span data-ttu-id="44017-289">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="44017-289">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="44017-290">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="44017-290">Compose or Read</span></span>|

####  <a name="nullable-conversationid-string"></a><span data-ttu-id="44017-291">(nullable) conversationId :String</span><span class="sxs-lookup"><span data-stu-id="44017-291">(nullable) conversationId :String</span></span>

<span data-ttu-id="44017-292">获取包含特定消息的电子邮件会话的标识符。</span><span class="sxs-lookup"><span data-stu-id="44017-292">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="44017-p107">如果在阅读窗体或撰写窗体的回复中激活邮件应用程序，则此属性可以获得一个整数值。如果用户随后更改了回复邮件的主题（若发送回复），则该邮件的对话 ID 将改变且之前获取的值将不适用。</span><span class="sxs-lookup"><span data-stu-id="44017-p107">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="44017-p108">对于撰写窗体的新项目，此属性获得一个 null 值。如果用户设置一个主题并保存该项目，`conversationId` 属性将返回一个值。</span><span class="sxs-lookup"><span data-stu-id="44017-p108">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="44017-297">Type</span><span class="sxs-lookup"><span data-stu-id="44017-297">Type</span></span>

*   <span data-ttu-id="44017-298">String</span><span class="sxs-lookup"><span data-stu-id="44017-298">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="44017-299">要求</span><span class="sxs-lookup"><span data-stu-id="44017-299">Requirements</span></span>

|<span data-ttu-id="44017-300">要求</span><span class="sxs-lookup"><span data-stu-id="44017-300">Requirement</span></span>|<span data-ttu-id="44017-301">值</span><span class="sxs-lookup"><span data-stu-id="44017-301">Value</span></span>|
|---|---|
|[<span data-ttu-id="44017-302">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="44017-302">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="44017-303">1.0</span><span class="sxs-lookup"><span data-stu-id="44017-303">1.0</span></span>|
|[<span data-ttu-id="44017-304">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="44017-304">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="44017-305">ReadItem</span><span class="sxs-lookup"><span data-stu-id="44017-305">ReadItem</span></span>|
|[<span data-ttu-id="44017-306">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="44017-306">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="44017-307">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="44017-307">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="44017-308">示例</span><span class="sxs-lookup"><span data-stu-id="44017-308">Example</span></span>

```javascript
var conversationId = Office.context.mailbox.item.conversationId;
console.log("conversationId: " + conversationId);
```

#### <a name="datetimecreated-date"></a><span data-ttu-id="44017-309">dateTimeCreated :Date</span><span class="sxs-lookup"><span data-stu-id="44017-309">dateTimeCreated :Date</span></span>

<span data-ttu-id="44017-p109">获取项目创建的日期和时间。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="44017-p109">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="44017-312">类型</span><span class="sxs-lookup"><span data-stu-id="44017-312">Type</span></span>

*   <span data-ttu-id="44017-313">日期</span><span class="sxs-lookup"><span data-stu-id="44017-313">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="44017-314">Requirements</span><span class="sxs-lookup"><span data-stu-id="44017-314">Requirements</span></span>

|<span data-ttu-id="44017-315">要求</span><span class="sxs-lookup"><span data-stu-id="44017-315">Requirement</span></span>|<span data-ttu-id="44017-316">值</span><span class="sxs-lookup"><span data-stu-id="44017-316">Value</span></span>|
|---|---|
|[<span data-ttu-id="44017-317">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="44017-317">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="44017-318">1.0</span><span class="sxs-lookup"><span data-stu-id="44017-318">1.0</span></span>|
|[<span data-ttu-id="44017-319">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="44017-319">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="44017-320">ReadItem</span><span class="sxs-lookup"><span data-stu-id="44017-320">ReadItem</span></span>|
|[<span data-ttu-id="44017-321">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="44017-321">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="44017-322">阅读</span><span class="sxs-lookup"><span data-stu-id="44017-322">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="44017-323">示例</span><span class="sxs-lookup"><span data-stu-id="44017-323">Example</span></span>

```javascript
var dateTimeCreated = Office.context.mailbox.item.dateTimeCreated;
console.log("Date and time created: " + dateTimeCreated);
```

#### <a name="datetimemodified-date"></a><span data-ttu-id="44017-324">dateTimeModified :Date</span><span class="sxs-lookup"><span data-stu-id="44017-324">dateTimeModified :Date</span></span>

<span data-ttu-id="44017-p110">获取项目最近一次修改的日期和时间。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="44017-p110">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="44017-327">在 Outlook for iOS 或 Outlook for Android 中不支持此成员。</span><span class="sxs-lookup"><span data-stu-id="44017-327">This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="type"></a><span data-ttu-id="44017-328">类型</span><span class="sxs-lookup"><span data-stu-id="44017-328">Type</span></span>

*   <span data-ttu-id="44017-329">日期</span><span class="sxs-lookup"><span data-stu-id="44017-329">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="44017-330">Requirements</span><span class="sxs-lookup"><span data-stu-id="44017-330">Requirements</span></span>

|<span data-ttu-id="44017-331">要求</span><span class="sxs-lookup"><span data-stu-id="44017-331">Requirement</span></span>|<span data-ttu-id="44017-332">值</span><span class="sxs-lookup"><span data-stu-id="44017-332">Value</span></span>|
|---|---|
|[<span data-ttu-id="44017-333">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="44017-333">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="44017-334">1.0</span><span class="sxs-lookup"><span data-stu-id="44017-334">1.0</span></span>|
|[<span data-ttu-id="44017-335">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="44017-335">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="44017-336">ReadItem</span><span class="sxs-lookup"><span data-stu-id="44017-336">ReadItem</span></span>|
|[<span data-ttu-id="44017-337">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="44017-337">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="44017-338">阅读</span><span class="sxs-lookup"><span data-stu-id="44017-338">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="44017-339">示例</span><span class="sxs-lookup"><span data-stu-id="44017-339">Example</span></span>

```javascript
var dateTimeModified = Office.context.mailbox.item.dateTimeModified;
console.log("Date and time modified: " + dateTimeModified);
```

####  <a name="end-datetimejavascriptapioutlookofficetime"></a><span data-ttu-id="44017-340">end :Date|[Time](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="44017-340">end :Date|[Time](/javascript/api/outlook/office.time)</span></span>

<span data-ttu-id="44017-341">获取或设置约会结束的日期和时间。</span><span class="sxs-lookup"><span data-stu-id="44017-341">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="44017-p111">将 `end` 属性表示为协调世界时 (UTC) 的日期和时间值。可使用 [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttime) 方法将 end 属性值转换为客户端的本地日期和时间。</span><span class="sxs-lookup"><span data-stu-id="44017-p111">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="44017-344">阅读模式</span><span class="sxs-lookup"><span data-stu-id="44017-344">Read mode</span></span>

<span data-ttu-id="44017-345">`end` 属性返回 `Date` 对象。</span><span class="sxs-lookup"><span data-stu-id="44017-345">The `end` property returns a `Date` object.</span></span>

```javascript
var end = Office.context.mailbox.item.end;
console.log("Appointment end: " + end);
```

##### <a name="compose-mode"></a><span data-ttu-id="44017-346">撰写模式</span><span class="sxs-lookup"><span data-stu-id="44017-346">Compose mode</span></span>

<span data-ttu-id="44017-347">`end` 属性返回 `Time` 对象。</span><span class="sxs-lookup"><span data-stu-id="44017-347">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="44017-348">使用 [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) 方法设置结束时间时，应使用 [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) 方法将客户端的本地时间转换为服务器的 UTC。</span><span class="sxs-lookup"><span data-stu-id="44017-348">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="44017-349">下面的示例使用[`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) `Time`对象的方法设置约会的结束时间。</span><span class="sxs-lookup"><span data-stu-id="44017-349">The following example sets the end time of an appointment by using the [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="44017-350">类型</span><span class="sxs-lookup"><span data-stu-id="44017-350">Type</span></span>

*   <span data-ttu-id="44017-351">Date | [Time](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="44017-351">Date | [Time](/javascript/api/outlook/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="44017-352">要求</span><span class="sxs-lookup"><span data-stu-id="44017-352">Requirements</span></span>

|<span data-ttu-id="44017-353">要求</span><span class="sxs-lookup"><span data-stu-id="44017-353">Requirement</span></span>|<span data-ttu-id="44017-354">值</span><span class="sxs-lookup"><span data-stu-id="44017-354">Value</span></span>|
|---|---|
|[<span data-ttu-id="44017-355">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="44017-355">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="44017-356">1.0</span><span class="sxs-lookup"><span data-stu-id="44017-356">1.0</span></span>|
|[<span data-ttu-id="44017-357">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="44017-357">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="44017-358">ReadItem</span><span class="sxs-lookup"><span data-stu-id="44017-358">ReadItem</span></span>|
|[<span data-ttu-id="44017-359">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="44017-359">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="44017-360">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="44017-360">Compose or Read</span></span>|

#### <a name="enhancedlocation-enhancedlocationjavascriptapioutlookofficeenhancedlocation"></a><span data-ttu-id="44017-361">enhancedLocation:[enhancedLocation](/javascript/api/outlook/office.enhancedlocation)</span><span class="sxs-lookup"><span data-stu-id="44017-361">enhancedLocation :[EnhancedLocation](/javascript/api/outlook/office.enhancedlocation)</span></span>

<span data-ttu-id="44017-362">获取或设置约会的位置。</span><span class="sxs-lookup"><span data-stu-id="44017-362">Gets or sets the locations of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="44017-363">阅读模式</span><span class="sxs-lookup"><span data-stu-id="44017-363">Read mode</span></span>

<span data-ttu-id="44017-364">该`enhancedLocation`属性返回一个[EnhancedLocation](/javascript/api/outlook/office.enhancedlocation)对象, 该对象允许您获取与约会关联的一组位置 (每个由[LocationDetails](/javascript/api/outlook/office.locationdetails)对象表示)。</span><span class="sxs-lookup"><span data-stu-id="44017-364">The `enhancedLocation` property returns an [EnhancedLocation](/javascript/api/outlook/office.enhancedlocation) object that allows you to get the set of locations (each represented by a [LocationDetails](/javascript/api/outlook/office.locationdetails) object) associated with the appointment.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="44017-365">撰写模式</span><span class="sxs-lookup"><span data-stu-id="44017-365">Compose mode</span></span>

<span data-ttu-id="44017-366">该`enhancedLocation`属性返回一个[EnhancedLocation](/javascript/api/outlook/office.enhancedlocation)对象, 该对象提供用于获取、删除或添加约会位置的方法。</span><span class="sxs-lookup"><span data-stu-id="44017-366">The `enhancedLocation` property returns an [EnhancedLocation](/javascript/api/outlook/office.enhancedlocation) object that provides methods to get, remove, or add locations on an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="44017-367">类型</span><span class="sxs-lookup"><span data-stu-id="44017-367">Type</span></span>

*   [<span data-ttu-id="44017-368">EnhancedLocation</span><span class="sxs-lookup"><span data-stu-id="44017-368">EnhancedLocation</span></span>](/javascript/api/outlook/office.enhancedlocation)

##### <a name="requirements"></a><span data-ttu-id="44017-369">Requirements</span><span class="sxs-lookup"><span data-stu-id="44017-369">Requirements</span></span>

|<span data-ttu-id="44017-370">要求</span><span class="sxs-lookup"><span data-stu-id="44017-370">Requirement</span></span>|<span data-ttu-id="44017-371">值</span><span class="sxs-lookup"><span data-stu-id="44017-371">Value</span></span>|
|---|---|
|[<span data-ttu-id="44017-372">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="44017-372">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="44017-373">预览</span><span class="sxs-lookup"><span data-stu-id="44017-373">Preview</span></span>|
|[<span data-ttu-id="44017-374">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="44017-374">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="44017-375">ReadItem</span><span class="sxs-lookup"><span data-stu-id="44017-375">ReadItem</span></span>|
|[<span data-ttu-id="44017-376">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="44017-376">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="44017-377">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="44017-377">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="44017-378">示例</span><span class="sxs-lookup"><span data-stu-id="44017-378">Example</span></span>

<span data-ttu-id="44017-379">下面的示例将获取与约会相关联的当前位置。</span><span class="sxs-lookup"><span data-stu-id="44017-379">The following example gets the current locations associated with the appointment.</span></span>

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

#### <a name="from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsfromjavascriptapioutlookofficefrom"></a><span data-ttu-id="44017-380">from :[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[From](/javascript/api/outlook/office.from)</span><span class="sxs-lookup"><span data-stu-id="44017-380">from :[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[From](/javascript/api/outlook/office.from)</span></span>

<span data-ttu-id="44017-381">获取邮件发件人的电子邮件地址。</span><span class="sxs-lookup"><span data-stu-id="44017-381">Gets the email address of the sender of a message.</span></span>

<span data-ttu-id="44017-p112">`from` 和 [`sender`](#sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetails) 属性表示同一个人，邮件由代理人发送的除外。在此情况下，`from` 属性表示代理程序，而 sender 属性表示代理人。</span><span class="sxs-lookup"><span data-stu-id="44017-p112">The `from` and [`sender`](#sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="44017-384">`from` 属性中 `EmailAddressDetails` 对象的 `recipientType` 属性为 `undefined`。</span><span class="sxs-lookup"><span data-stu-id="44017-384">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="44017-385">阅读模式</span><span class="sxs-lookup"><span data-stu-id="44017-385">Read mode</span></span>

<span data-ttu-id="44017-386">`from` 属性返回一个 `EmailAddressDetails` 对象。</span><span class="sxs-lookup"><span data-stu-id="44017-386">The `from` property returns an `EmailAddressDetails` object.</span></span>

```javascript
var from = Office.context.mailbox.item.from;
console.log("From " + from);
```

##### <a name="compose-mode"></a><span data-ttu-id="44017-387">撰写模式</span><span class="sxs-lookup"><span data-stu-id="44017-387">Compose mode</span></span>

<span data-ttu-id="44017-388">`from` 属性返回一个 `From` 对象，该对象提供从值中进行获取的方法。</span><span class="sxs-lookup"><span data-stu-id="44017-388">The `from` property returns a `From` object that provides a method to get the from value.</span></span>

```javascript
Office.context.mailbox.item.from.getAsync(callback);

function callback(asyncResult) {
  var from = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="44017-389">类型</span><span class="sxs-lookup"><span data-stu-id="44017-389">Type</span></span>

*   <span data-ttu-id="44017-390">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [From](/javascript/api/outlook/office.from)</span><span class="sxs-lookup"><span data-stu-id="44017-390">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [From](/javascript/api/outlook/office.from)</span></span>

##### <a name="requirements"></a><span data-ttu-id="44017-391">要求</span><span class="sxs-lookup"><span data-stu-id="44017-391">Requirements</span></span>

|<span data-ttu-id="44017-392">要求</span><span class="sxs-lookup"><span data-stu-id="44017-392">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="44017-393">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="44017-393">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="44017-394">1.0</span><span class="sxs-lookup"><span data-stu-id="44017-394">1.0</span></span>|<span data-ttu-id="44017-395">1.7</span><span class="sxs-lookup"><span data-stu-id="44017-395">1.7</span></span>|
|[<span data-ttu-id="44017-396">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="44017-396">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="44017-397">ReadItem</span><span class="sxs-lookup"><span data-stu-id="44017-397">ReadItem</span></span>|<span data-ttu-id="44017-398">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="44017-398">ReadWriteItem</span></span>|
|[<span data-ttu-id="44017-399">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="44017-399">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="44017-400">阅读</span><span class="sxs-lookup"><span data-stu-id="44017-400">Read</span></span>|<span data-ttu-id="44017-401">撰写</span><span class="sxs-lookup"><span data-stu-id="44017-401">Compose</span></span>|

#### <a name="internetheaders-internetheadersjavascriptapioutlookofficeinternetheaders"></a><span data-ttu-id="44017-402">internetHeaders :[InternetHeaders](/javascript/api/outlook/office.internetheaders)</span><span class="sxs-lookup"><span data-stu-id="44017-402">internetHeaders :[InternetHeaders](/javascript/api/outlook/office.internetheaders)</span></span>

<span data-ttu-id="44017-403">获取或设置消息的 Internet 标头。</span><span class="sxs-lookup"><span data-stu-id="44017-403">Gets or sets the internet headers of a message.</span></span>

##### <a name="type"></a><span data-ttu-id="44017-404">类型</span><span class="sxs-lookup"><span data-stu-id="44017-404">Type</span></span>

*   [<span data-ttu-id="44017-405">InternetHeaders</span><span class="sxs-lookup"><span data-stu-id="44017-405">InternetHeaders</span></span>](/javascript/api/outlook/office.internetheaders)

##### <a name="requirements"></a><span data-ttu-id="44017-406">要求</span><span class="sxs-lookup"><span data-stu-id="44017-406">Requirements</span></span>

|<span data-ttu-id="44017-407">要求</span><span class="sxs-lookup"><span data-stu-id="44017-407">Requirement</span></span>|<span data-ttu-id="44017-408">值</span><span class="sxs-lookup"><span data-stu-id="44017-408">Value</span></span>|
|---|---|
|[<span data-ttu-id="44017-409">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="44017-409">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="44017-410">预览</span><span class="sxs-lookup"><span data-stu-id="44017-410">Preview</span></span>|
|[<span data-ttu-id="44017-411">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="44017-411">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="44017-412">ReadItem</span><span class="sxs-lookup"><span data-stu-id="44017-412">ReadItem</span></span>|
|[<span data-ttu-id="44017-413">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="44017-413">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="44017-414">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="44017-414">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="44017-415">示例</span><span class="sxs-lookup"><span data-stu-id="44017-415">Example</span></span>

```javascript
Office.context.mailbox.item.internetHeaders.getAsync(["header1", "header2"], callback);

function callback(asyncResult) {
  var dictionary = asyncResult.value;
  var header1_value = dictionary["header1"];
}
```

#### <a name="internetmessageid-string"></a><span data-ttu-id="44017-416">internetMessageId :String</span><span class="sxs-lookup"><span data-stu-id="44017-416">internetMessageId :String</span></span>

<span data-ttu-id="44017-p113">获取电子邮件的 Internet 消息标识符。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="44017-p113">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="44017-419">Type</span><span class="sxs-lookup"><span data-stu-id="44017-419">Type</span></span>

*   <span data-ttu-id="44017-420">String</span><span class="sxs-lookup"><span data-stu-id="44017-420">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="44017-421">要求</span><span class="sxs-lookup"><span data-stu-id="44017-421">Requirements</span></span>

|<span data-ttu-id="44017-422">要求</span><span class="sxs-lookup"><span data-stu-id="44017-422">Requirement</span></span>|<span data-ttu-id="44017-423">值</span><span class="sxs-lookup"><span data-stu-id="44017-423">Value</span></span>|
|---|---|
|[<span data-ttu-id="44017-424">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="44017-424">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="44017-425">1.0</span><span class="sxs-lookup"><span data-stu-id="44017-425">1.0</span></span>|
|[<span data-ttu-id="44017-426">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="44017-426">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="44017-427">ReadItem</span><span class="sxs-lookup"><span data-stu-id="44017-427">ReadItem</span></span>|
|[<span data-ttu-id="44017-428">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="44017-428">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="44017-429">阅读</span><span class="sxs-lookup"><span data-stu-id="44017-429">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="44017-430">示例</span><span class="sxs-lookup"><span data-stu-id="44017-430">Example</span></span>

```javascript
var internetMessageId = Office.context.mailbox.item.internetMessageId;
console.log("internetMessageId: " + internetMessageId);
```

#### <a name="itemclass-string"></a><span data-ttu-id="44017-431">itemClass :String</span><span class="sxs-lookup"><span data-stu-id="44017-431">itemClass :String</span></span>

<span data-ttu-id="44017-p114">获取选定项目的 Exchange Web 服务项目类。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="44017-p114">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="44017-p115">`itemClass` 属性指定所选项目的邮件类别。以下是邮件或约会项目的默认邮件类别。</span><span class="sxs-lookup"><span data-stu-id="44017-p115">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

|<span data-ttu-id="44017-436">类型</span><span class="sxs-lookup"><span data-stu-id="44017-436">Type</span></span>|<span data-ttu-id="44017-437">说明</span><span class="sxs-lookup"><span data-stu-id="44017-437">Description</span></span>|<span data-ttu-id="44017-438">项目类</span><span class="sxs-lookup"><span data-stu-id="44017-438">item class</span></span>|
|---|---|---|
|<span data-ttu-id="44017-439">约会项目</span><span class="sxs-lookup"><span data-stu-id="44017-439">Appointment items</span></span>|<span data-ttu-id="44017-440">这些是项目类 `IPM.Appointment` 或 `IPM.Appointment.Occurrence` 的日历项目。</span><span class="sxs-lookup"><span data-stu-id="44017-440">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurrence`.</span></span>|`IPM.Appointment`<br />`IPM.Appointment.Occurrence`|
|<span data-ttu-id="44017-441">邮件项目</span><span class="sxs-lookup"><span data-stu-id="44017-441">Message items</span></span>|<span data-ttu-id="44017-442">这些项目包括具有默认邮件类别 `IPM.Note` 的电子邮件，以及使用 `IPM.Schedule.Meeting` 作为基础邮件类别的会议请求、响应和取消。</span><span class="sxs-lookup"><span data-stu-id="44017-442">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span>|`IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled`|

<span data-ttu-id="44017-443">你可以创建用于扩展默认邮件类别的自定义邮件类别，例如，自定义约会邮件类别 `IPM.Appointment.Contoso`。</span><span class="sxs-lookup"><span data-stu-id="44017-443">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="44017-444">Type</span><span class="sxs-lookup"><span data-stu-id="44017-444">Type</span></span>

*   <span data-ttu-id="44017-445">String</span><span class="sxs-lookup"><span data-stu-id="44017-445">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="44017-446">要求</span><span class="sxs-lookup"><span data-stu-id="44017-446">Requirements</span></span>

|<span data-ttu-id="44017-447">要求</span><span class="sxs-lookup"><span data-stu-id="44017-447">Requirement</span></span>|<span data-ttu-id="44017-448">值</span><span class="sxs-lookup"><span data-stu-id="44017-448">Value</span></span>|
|---|---|
|[<span data-ttu-id="44017-449">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="44017-449">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="44017-450">1.0</span><span class="sxs-lookup"><span data-stu-id="44017-450">1.0</span></span>|
|[<span data-ttu-id="44017-451">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="44017-451">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="44017-452">ReadItem</span><span class="sxs-lookup"><span data-stu-id="44017-452">ReadItem</span></span>|
|[<span data-ttu-id="44017-453">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="44017-453">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="44017-454">阅读</span><span class="sxs-lookup"><span data-stu-id="44017-454">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="44017-455">示例</span><span class="sxs-lookup"><span data-stu-id="44017-455">Example</span></span>

```javascript
var itemClass = Office.context.mailbox.item.itemClass;
console.log("Item class: " + itemClass);
```

#### <a name="nullable-itemid-string"></a><span data-ttu-id="44017-456">(nullable) itemId :String</span><span class="sxs-lookup"><span data-stu-id="44017-456">(nullable) itemId :String</span></span>

<span data-ttu-id="44017-p116">获取当前项目的 Exchange Web 服务项目标识符。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="44017-p116">Gets the Exchange Web Services item identifier for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="44017-459">`itemId` 属性返回的标识符与 Exchange Web 服务项目标识符相同。</span><span class="sxs-lookup"><span data-stu-id="44017-459">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="44017-460">`itemId` 属性与 Outlook 条目 ID 或 Outlook REST API 使用的 ID 不同。</span><span class="sxs-lookup"><span data-stu-id="44017-460">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="44017-461">使用此值进行 REST API 调用前，应使用 [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string) 对它进行转换。</span><span class="sxs-lookup"><span data-stu-id="44017-461">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="44017-462">有关详细信息，请参阅[从 Outlook 加载项使用 Outlook REST API](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id)。</span><span class="sxs-lookup"><span data-stu-id="44017-462">For more details, see [Use the Outlook REST APIs from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

<span data-ttu-id="44017-p118">`itemId` 属性在撰写模式下不可用。如果需要项目标识符，[`saveAsync`](#saveasyncoptions-callback) 方法可用于将项目保存到存储，这将在回调函数的 [`AsyncResult.value`](/javascript/api/office/office.asyncresult) 参数中返回项目标识符。</span><span class="sxs-lookup"><span data-stu-id="44017-p118">The `itemId` property is not available in compose mode. If an item identifier is required, the [`saveAsync`](#saveasyncoptions-callback) method can be used to save the item to the store, which will return the item identifier in the [`AsyncResult.value`](/javascript/api/office/office.asyncresult) parameter in the callback function.</span></span>

##### <a name="type"></a><span data-ttu-id="44017-465">Type</span><span class="sxs-lookup"><span data-stu-id="44017-465">Type</span></span>

*   <span data-ttu-id="44017-466">String</span><span class="sxs-lookup"><span data-stu-id="44017-466">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="44017-467">要求</span><span class="sxs-lookup"><span data-stu-id="44017-467">Requirements</span></span>

|<span data-ttu-id="44017-468">要求</span><span class="sxs-lookup"><span data-stu-id="44017-468">Requirement</span></span>|<span data-ttu-id="44017-469">值</span><span class="sxs-lookup"><span data-stu-id="44017-469">Value</span></span>|
|---|---|
|[<span data-ttu-id="44017-470">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="44017-470">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="44017-471">1.0</span><span class="sxs-lookup"><span data-stu-id="44017-471">1.0</span></span>|
|[<span data-ttu-id="44017-472">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="44017-472">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="44017-473">ReadItem</span><span class="sxs-lookup"><span data-stu-id="44017-473">ReadItem</span></span>|
|[<span data-ttu-id="44017-474">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="44017-474">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="44017-475">阅读</span><span class="sxs-lookup"><span data-stu-id="44017-475">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="44017-476">示例</span><span class="sxs-lookup"><span data-stu-id="44017-476">Example</span></span>

<span data-ttu-id="44017-p119">以下代码检查项目标识符是否存在。如果 `itemId` 属性返回 `null` 或 `undefined`，则将项目保存到存储，并从异步结果中获取项目标识符。</span><span class="sxs-lookup"><span data-stu-id="44017-p119">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

```javascript
var itemId = Office.context.mailbox.item.itemId;
if (itemId === null || itemId == undefined) {
  Office.context.mailbox.item.saveAsync(function(result) {
    itemId = result.value;
  });
}
```

####  <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtype"></a><span data-ttu-id="44017-479">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype)</span><span class="sxs-lookup"><span data-stu-id="44017-479">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype)</span></span>

<span data-ttu-id="44017-480">获取实例表示的项的类型。</span><span class="sxs-lookup"><span data-stu-id="44017-480">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="44017-481">`itemType` 属性返回其中一个 `ItemType` 枚举值，指示 `item` 对象实例是邮件还是约会。</span><span class="sxs-lookup"><span data-stu-id="44017-481">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="44017-482">类型</span><span class="sxs-lookup"><span data-stu-id="44017-482">Type</span></span>

*   [<span data-ttu-id="44017-483">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="44017-483">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook/office.mailboxenums.itemtype)

##### <a name="requirements"></a><span data-ttu-id="44017-484">Requirements</span><span class="sxs-lookup"><span data-stu-id="44017-484">Requirements</span></span>

|<span data-ttu-id="44017-485">要求</span><span class="sxs-lookup"><span data-stu-id="44017-485">Requirement</span></span>|<span data-ttu-id="44017-486">值</span><span class="sxs-lookup"><span data-stu-id="44017-486">Value</span></span>|
|---|---|
|[<span data-ttu-id="44017-487">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="44017-487">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="44017-488">1.0</span><span class="sxs-lookup"><span data-stu-id="44017-488">1.0</span></span>|
|[<span data-ttu-id="44017-489">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="44017-489">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="44017-490">ReadItem</span><span class="sxs-lookup"><span data-stu-id="44017-490">ReadItem</span></span>|
|[<span data-ttu-id="44017-491">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="44017-491">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="44017-492">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="44017-492">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="44017-493">示例</span><span class="sxs-lookup"><span data-stu-id="44017-493">Example</span></span>

```javascript
if (Office.context.mailbox.item.itemType === Office.MailboxEnums.ItemType.Message) {
  // Do something.
} else {
  // Do something else.
}
```

####  <a name="location-stringlocationjavascriptapioutlookofficelocation"></a><span data-ttu-id="44017-494">location :String|[Location](/javascript/api/outlook/office.location)</span><span class="sxs-lookup"><span data-stu-id="44017-494">location :String|[Location](/javascript/api/outlook/office.location)</span></span>

<span data-ttu-id="44017-495">获取或设置约会的位置。</span><span class="sxs-lookup"><span data-stu-id="44017-495">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="44017-496">阅读模式</span><span class="sxs-lookup"><span data-stu-id="44017-496">Read mode</span></span>

<span data-ttu-id="44017-497">`location` 属性返回一个包含约会位置的字符串。</span><span class="sxs-lookup"><span data-stu-id="44017-497">The `location` property returns a string that contains the location of the appointment.</span></span>

```javascript
var location = Office.context.mailbox.item.location;
console.log("location: " + location);
```

##### <a name="compose-mode"></a><span data-ttu-id="44017-498">撰写模式</span><span class="sxs-lookup"><span data-stu-id="44017-498">Compose mode</span></span>

<span data-ttu-id="44017-499">`location` 属性返回一个 `Location` 对象，该对象提供用于获取和设置约会位置的方法。</span><span class="sxs-lookup"><span data-stu-id="44017-499">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

```javascript
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="44017-500">类型</span><span class="sxs-lookup"><span data-stu-id="44017-500">Type</span></span>

*   <span data-ttu-id="44017-501">String | [Location](/javascript/api/outlook/office.location)</span><span class="sxs-lookup"><span data-stu-id="44017-501">String | [Location](/javascript/api/outlook/office.location)</span></span>

##### <a name="requirements"></a><span data-ttu-id="44017-502">Requirements</span><span class="sxs-lookup"><span data-stu-id="44017-502">Requirements</span></span>

|<span data-ttu-id="44017-503">要求</span><span class="sxs-lookup"><span data-stu-id="44017-503">Requirement</span></span>|<span data-ttu-id="44017-504">值</span><span class="sxs-lookup"><span data-stu-id="44017-504">Value</span></span>|
|---|---|
|[<span data-ttu-id="44017-505">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="44017-505">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="44017-506">1.0</span><span class="sxs-lookup"><span data-stu-id="44017-506">1.0</span></span>|
|[<span data-ttu-id="44017-507">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="44017-507">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="44017-508">ReadItem</span><span class="sxs-lookup"><span data-stu-id="44017-508">ReadItem</span></span>|
|[<span data-ttu-id="44017-509">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="44017-509">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="44017-510">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="44017-510">Compose or Read</span></span>|

#### <a name="normalizedsubject-string"></a><span data-ttu-id="44017-511">normalizedSubject :String</span><span class="sxs-lookup"><span data-stu-id="44017-511">normalizedSubject :String</span></span>

<span data-ttu-id="44017-p120">获取删除了所有前缀（包括 `RE:` 和 `FWD:`）的项目主题。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="44017-p120">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="44017-p121">normalizedSubject 属性获取包含由电子邮件程序添加的任何标准前缀（如 `RE:` 和 `FW:`）的项目主题。若要获取包含完整前缀的项目主题，请使用 [`subject`](#subject-stringsubjectjavascriptapioutlookofficesubject) 属性。</span><span class="sxs-lookup"><span data-stu-id="44017-p121">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubjectjavascriptapioutlookofficesubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="44017-516">Type</span><span class="sxs-lookup"><span data-stu-id="44017-516">Type</span></span>

*   <span data-ttu-id="44017-517">String</span><span class="sxs-lookup"><span data-stu-id="44017-517">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="44017-518">要求</span><span class="sxs-lookup"><span data-stu-id="44017-518">Requirements</span></span>

|<span data-ttu-id="44017-519">要求</span><span class="sxs-lookup"><span data-stu-id="44017-519">Requirement</span></span>|<span data-ttu-id="44017-520">值</span><span class="sxs-lookup"><span data-stu-id="44017-520">Value</span></span>|
|---|---|
|[<span data-ttu-id="44017-521">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="44017-521">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="44017-522">1.0</span><span class="sxs-lookup"><span data-stu-id="44017-522">1.0</span></span>|
|[<span data-ttu-id="44017-523">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="44017-523">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="44017-524">ReadItem</span><span class="sxs-lookup"><span data-stu-id="44017-524">ReadItem</span></span>|
|[<span data-ttu-id="44017-525">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="44017-525">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="44017-526">阅读</span><span class="sxs-lookup"><span data-stu-id="44017-526">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="44017-527">示例</span><span class="sxs-lookup"><span data-stu-id="44017-527">Example</span></span>

```javascript
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
console.log("Normalized subject: " + normalizedSubject);
```

####  <a name="notificationmessages-notificationmessagesjavascriptapioutlookofficenotificationmessages"></a><span data-ttu-id="44017-528">notificationMessages :[NotificationMessages](/javascript/api/outlook/office.notificationmessages)</span><span class="sxs-lookup"><span data-stu-id="44017-528">notificationMessages :[NotificationMessages](/javascript/api/outlook/office.notificationmessages)</span></span>

<span data-ttu-id="44017-529">获取项目的通知邮件。</span><span class="sxs-lookup"><span data-stu-id="44017-529">Gets the notification messages for an item.</span></span>

##### <a name="type"></a><span data-ttu-id="44017-530">类型</span><span class="sxs-lookup"><span data-stu-id="44017-530">Type</span></span>

*   [<span data-ttu-id="44017-531">NotificationMessages</span><span class="sxs-lookup"><span data-stu-id="44017-531">NotificationMessages</span></span>](/javascript/api/outlook/office.notificationmessages)

##### <a name="requirements"></a><span data-ttu-id="44017-532">Requirements</span><span class="sxs-lookup"><span data-stu-id="44017-532">Requirements</span></span>

|<span data-ttu-id="44017-533">要求</span><span class="sxs-lookup"><span data-stu-id="44017-533">Requirement</span></span>|<span data-ttu-id="44017-534">值</span><span class="sxs-lookup"><span data-stu-id="44017-534">Value</span></span>|
|---|---|
|[<span data-ttu-id="44017-535">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="44017-535">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="44017-536">1.3</span><span class="sxs-lookup"><span data-stu-id="44017-536">1.3</span></span>|
|[<span data-ttu-id="44017-537">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="44017-537">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="44017-538">ReadItem</span><span class="sxs-lookup"><span data-stu-id="44017-538">ReadItem</span></span>|
|[<span data-ttu-id="44017-539">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="44017-539">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="44017-540">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="44017-540">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="44017-541">示例</span><span class="sxs-lookup"><span data-stu-id="44017-541">Example</span></span>

```javascript
// Get all notifications.
Office.context.mailbox.item.notificationMessages.getAllAsync(
  function (asyncResult) {
    console.log(JSON.stringify(asyncResult));
  }
);
```

####  <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="44017-542">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="44017-542">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="44017-543">提供对事件的可选与会者的访问权限。</span><span class="sxs-lookup"><span data-stu-id="44017-543">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="44017-544">对象的类型和访问级别取决于当前项目的模式。</span><span class="sxs-lookup"><span data-stu-id="44017-544">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="44017-545">读取模式</span><span class="sxs-lookup"><span data-stu-id="44017-545">Read mode</span></span>

<span data-ttu-id="44017-546">`optionalAttendees` 属性返回一个数组，其中包含每个可选与会者的 `EmailAddressDetails` 对象。</span><span class="sxs-lookup"><span data-stu-id="44017-546">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span>

```javascript
var optionalAttendees = Office.context.mailbox.item.optionalAttendees;
console.log("Optional attendees: " + JSON.stringify(optionalAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="44017-547">撰写模式</span><span class="sxs-lookup"><span data-stu-id="44017-547">Compose mode</span></span>

<span data-ttu-id="44017-548">`optionalAttendees` 属性返回一个 `Recipients` 对象，该对象提供用于获取或更新可选与会者的方法。</span><span class="sxs-lookup"><span data-stu-id="44017-548">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span>

```javascript
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="44017-549">类型</span><span class="sxs-lookup"><span data-stu-id="44017-549">Type</span></span>

*   <span data-ttu-id="44017-550">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="44017-550">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="44017-551">Requirements</span><span class="sxs-lookup"><span data-stu-id="44017-551">Requirements</span></span>

|<span data-ttu-id="44017-552">要求</span><span class="sxs-lookup"><span data-stu-id="44017-552">Requirement</span></span>|<span data-ttu-id="44017-553">值</span><span class="sxs-lookup"><span data-stu-id="44017-553">Value</span></span>|
|---|---|
|[<span data-ttu-id="44017-554">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="44017-554">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="44017-555">1.0</span><span class="sxs-lookup"><span data-stu-id="44017-555">1.0</span></span>|
|[<span data-ttu-id="44017-556">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="44017-556">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="44017-557">ReadItem</span><span class="sxs-lookup"><span data-stu-id="44017-557">ReadItem</span></span>|
|[<span data-ttu-id="44017-558">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="44017-558">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="44017-559">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="44017-559">Compose or Read</span></span>|

#### <a name="organizer-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsorganizerjavascriptapioutlookofficeorganizer"></a><span data-ttu-id="44017-560">organizer :[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[Organizer](/javascript/api/outlook/office.organizer)</span><span class="sxs-lookup"><span data-stu-id="44017-560">organizer :[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[Organizer](/javascript/api/outlook/office.organizer)</span></span>

<span data-ttu-id="44017-561">获取指定会议的组织者的电子邮件地址。</span><span class="sxs-lookup"><span data-stu-id="44017-561">Gets the email address of the organizer for a specified meeting.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="44017-562">阅读模式</span><span class="sxs-lookup"><span data-stu-id="44017-562">Read mode</span></span>

<span data-ttu-id="44017-563">`organizer` 属性返回 [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) 对象，它表示会议组织者。</span><span class="sxs-lookup"><span data-stu-id="44017-563">The `organizer` property returns an [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) object that represents the meeting organizer.</span></span>

```javascript
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
console.log("Organizer: " + organizerName + " (" + organizerAddress + ")");
```

##### <a name="compose-mode"></a><span data-ttu-id="44017-564">撰写模式</span><span class="sxs-lookup"><span data-stu-id="44017-564">Compose mode</span></span>

<span data-ttu-id="44017-565">`organizer` 属性返回 [Organizer](/javascript/api/outlook/office.organizer) 对象，该对象提供获取组织者值的方法。</span><span class="sxs-lookup"><span data-stu-id="44017-565">The `organizer` property returns an [Organizer](/javascript/api/outlook/office.organizer) object that provides a method to get the organizer value.</span></span>

```javascript
Office.context.mailbox.item.organizer.getAsync(
  function(asyncResult) {
    console.log(JSON.stringify(asyncResult));
  }
);
```

##### <a name="type"></a><span data-ttu-id="44017-566">类型</span><span class="sxs-lookup"><span data-stu-id="44017-566">Type</span></span>

*   <span data-ttu-id="44017-567">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [Organizer](/javascript/api/outlook/office.organizer)</span><span class="sxs-lookup"><span data-stu-id="44017-567">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [Organizer](/javascript/api/outlook/office.organizer)</span></span>

##### <a name="requirements"></a><span data-ttu-id="44017-568">要求</span><span class="sxs-lookup"><span data-stu-id="44017-568">Requirements</span></span>

|<span data-ttu-id="44017-569">要求</span><span class="sxs-lookup"><span data-stu-id="44017-569">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="44017-570">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="44017-570">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="44017-571">1.0</span><span class="sxs-lookup"><span data-stu-id="44017-571">1.0</span></span>|<span data-ttu-id="44017-572">1.7</span><span class="sxs-lookup"><span data-stu-id="44017-572">1.7</span></span>|
|[<span data-ttu-id="44017-573">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="44017-573">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="44017-574">ReadItem</span><span class="sxs-lookup"><span data-stu-id="44017-574">ReadItem</span></span>|<span data-ttu-id="44017-575">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="44017-575">ReadWriteItem</span></span>|
|[<span data-ttu-id="44017-576">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="44017-576">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="44017-577">阅读</span><span class="sxs-lookup"><span data-stu-id="44017-577">Read</span></span>|<span data-ttu-id="44017-578">撰写</span><span class="sxs-lookup"><span data-stu-id="44017-578">Compose</span></span>|

#### <a name="nullable-recurrence-recurrencejavascriptapioutlookofficerecurrence"></a><span data-ttu-id="44017-579">(nullable) recurrence :[Recurrence](/javascript/api/outlook/office.recurrence)</span><span class="sxs-lookup"><span data-stu-id="44017-579">(nullable) recurrence :[Recurrence](/javascript/api/outlook/office.recurrence)</span></span>

<span data-ttu-id="44017-580">获取或设置约会的定期模式。</span><span class="sxs-lookup"><span data-stu-id="44017-580">Gets or sets the recurrence pattern of an appointment.</span></span> <span data-ttu-id="44017-581">获取或设置会议请求的定期模式。</span><span class="sxs-lookup"><span data-stu-id="44017-581">Gets the recurrence pattern of a meeting request.</span></span> <span data-ttu-id="44017-582">阅读撰写约会项目的模式。</span><span class="sxs-lookup"><span data-stu-id="44017-582">Read and compose modes for appointment items.</span></span> <span data-ttu-id="44017-583">会议请求项目的阅读模式。</span><span class="sxs-lookup"><span data-stu-id="44017-583">Read mode for meeting request items.</span></span>

<span data-ttu-id="44017-584">如果项目是一个系列或系列中的一个实例，则 `recurrence` 属性将返回定期约会的 [recurrence](/javascript/api/outlook/office.recurrence) 对象或会议请求。</span><span class="sxs-lookup"><span data-stu-id="44017-584">The `recurrence` property returns a [recurrence](/javascript/api/outlook/office.recurrence) object for recurring appointments or meetings requests if an item is a series or an instance in a series.</span></span> <span data-ttu-id="44017-585">针对单个约会和单个约会的会议请求返回 `null`。</span><span class="sxs-lookup"><span data-stu-id="44017-585">`null` is returned for single appointments and meeting requests of single appointments.</span></span> <span data-ttu-id="44017-586">针对非会议请求的邮件返回 `undefined`。</span><span class="sxs-lookup"><span data-stu-id="44017-586">`undefined` is returned for messages that are not meeting requests.</span></span>

> <span data-ttu-id="44017-587">注意：会议请求的 `itemClass` 值为 IPM.Schedule.Meeting.Request。</span><span class="sxs-lookup"><span data-stu-id="44017-587">Note: Meeting requests have an `itemClass` value of IPM.Schedule.Meeting.Request.</span></span>

> <span data-ttu-id="44017-588">注意：如果 recurrence 对象为 `null`，则这表示对象是单个约会或单个约会的会议请求，而不是系列的一部分。</span><span class="sxs-lookup"><span data-stu-id="44017-588">Note: If the recurrence object is `null`, this indicates that the object is a single appointment or a meeting request of a single appointment and NOT a part of a series.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="44017-589">阅读模式</span><span class="sxs-lookup"><span data-stu-id="44017-589">Read mode</span></span>

<span data-ttu-id="44017-590">该`recurrence`属性返回一个代表约会定期的[定期](/javascript/api/outlook/office.recurrence)对象。</span><span class="sxs-lookup"><span data-stu-id="44017-590">The `recurrence` property returns a [Recurrence](/javascript/api/outlook/office.recurrence) object that represents the appointment recurrence.</span></span> <span data-ttu-id="44017-591">此功能适用于约会和会议请求。</span><span class="sxs-lookup"><span data-stu-id="44017-591">This is available for appointments and meeting requests.</span></span>

```javascript
var recurrence = Office.context.mailbox.item.recurrence;
console.log("Recurrence: " + JSON.stringify(recurrence));
```

##### <a name="compose-mode"></a><span data-ttu-id="44017-592">撰写模式</span><span class="sxs-lookup"><span data-stu-id="44017-592">Compose mode</span></span>

<span data-ttu-id="44017-593">该`recurrence`属性返回一个[定期](/javascript/api/outlook/office.recurrence)对象, 该对象提供用于管理约会周期的方法。</span><span class="sxs-lookup"><span data-stu-id="44017-593">The `recurrence` property returns a [Recurrence](/javascript/api/outlook/office.recurrence) object that provides methods to manage the appointment recurrence.</span></span> <span data-ttu-id="44017-594">这可用于约会。</span><span class="sxs-lookup"><span data-stu-id="44017-594">This is available for appointments.</span></span>

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

##### <a name="type"></a><span data-ttu-id="44017-595">类型</span><span class="sxs-lookup"><span data-stu-id="44017-595">Type</span></span>

* [<span data-ttu-id="44017-596">Recurrence</span><span class="sxs-lookup"><span data-stu-id="44017-596">Recurrence</span></span>](/javascript/api/outlook/office.recurrence)

|<span data-ttu-id="44017-597">要求</span><span class="sxs-lookup"><span data-stu-id="44017-597">Requirement</span></span>|<span data-ttu-id="44017-598">值</span><span class="sxs-lookup"><span data-stu-id="44017-598">Value</span></span>|
|---|---|
|[<span data-ttu-id="44017-599">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="44017-599">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="44017-600">1.7</span><span class="sxs-lookup"><span data-stu-id="44017-600">1.7</span></span>|
|[<span data-ttu-id="44017-601">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="44017-601">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="44017-602">ReadItem</span><span class="sxs-lookup"><span data-stu-id="44017-602">ReadItem</span></span>|
|[<span data-ttu-id="44017-603">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="44017-603">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="44017-604">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="44017-604">Compose or Read</span></span>|

####  <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="44017-605">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="44017-605">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="44017-606">提供对事件的必需与会者的访问权限。</span><span class="sxs-lookup"><span data-stu-id="44017-606">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="44017-607">对象的类型和访问级别取决于当前项目的模式。</span><span class="sxs-lookup"><span data-stu-id="44017-607">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="44017-608">读取模式</span><span class="sxs-lookup"><span data-stu-id="44017-608">Read mode</span></span>

<span data-ttu-id="44017-609">`requiredAttendees` 属性返回一个数组，其中包含每个必需与会者的 `EmailAddressDetails` 对象。</span><span class="sxs-lookup"><span data-stu-id="44017-609">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span>

```javascript
var requiredAttendees = Office.context.mailbox.item.requiredAttendees;
console.log("Required attendees: " + JSON.stringify(requiredAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="44017-610">撰写模式</span><span class="sxs-lookup"><span data-stu-id="44017-610">Compose mode</span></span>

<span data-ttu-id="44017-611">`requiredAttendees` 属性返回一个 `Recipients` 对象，该对象提供用于获取或更新必需与会者的方法。</span><span class="sxs-lookup"><span data-stu-id="44017-611">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span>

```javascript
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
  console.log(JSON.stringify(arrayOfRequiredAttendeesRecipients));
}
```

##### <a name="type"></a><span data-ttu-id="44017-612">类型</span><span class="sxs-lookup"><span data-stu-id="44017-612">Type</span></span>

*   <span data-ttu-id="44017-613">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="44017-613">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="44017-614">Requirements</span><span class="sxs-lookup"><span data-stu-id="44017-614">Requirements</span></span>

|<span data-ttu-id="44017-615">要求</span><span class="sxs-lookup"><span data-stu-id="44017-615">Requirement</span></span>|<span data-ttu-id="44017-616">值</span><span class="sxs-lookup"><span data-stu-id="44017-616">Value</span></span>|
|---|---|
|[<span data-ttu-id="44017-617">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="44017-617">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="44017-618">1.0</span><span class="sxs-lookup"><span data-stu-id="44017-618">1.0</span></span>|
|[<span data-ttu-id="44017-619">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="44017-619">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="44017-620">ReadItem</span><span class="sxs-lookup"><span data-stu-id="44017-620">ReadItem</span></span>|
|[<span data-ttu-id="44017-621">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="44017-621">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="44017-622">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="44017-622">Compose or Read</span></span>|

#### <a name="sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetails"></a><span data-ttu-id="44017-623">sender :[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="44017-623">sender :[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)</span></span>

<span data-ttu-id="44017-p128">获取电子邮件发件人的电子邮件地址。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="44017-p128">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="44017-p129">[`from`](#from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsfromjavascriptapioutlookofficefrom) 和 `sender` 属性表示同一个人，邮件由代理人发送的除外。在此情况下，`from` 属性表示代理程序，而 sender 属性表示代理人。</span><span class="sxs-lookup"><span data-stu-id="44017-p129">The [`from`](#from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsfromjavascriptapioutlookofficefrom) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="44017-628">`sender` 属性中 `EmailAddressDetails` 对象的 `recipientType` 属性为 `undefined`。</span><span class="sxs-lookup"><span data-stu-id="44017-628">The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="44017-629">类型</span><span class="sxs-lookup"><span data-stu-id="44017-629">Type</span></span>

*   [<span data-ttu-id="44017-630">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="44017-630">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="44017-631">Requirements</span><span class="sxs-lookup"><span data-stu-id="44017-631">Requirements</span></span>

|<span data-ttu-id="44017-632">要求</span><span class="sxs-lookup"><span data-stu-id="44017-632">Requirement</span></span>|<span data-ttu-id="44017-633">值</span><span class="sxs-lookup"><span data-stu-id="44017-633">Value</span></span>|
|---|---|
|[<span data-ttu-id="44017-634">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="44017-634">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="44017-635">1.0</span><span class="sxs-lookup"><span data-stu-id="44017-635">1.0</span></span>|
|[<span data-ttu-id="44017-636">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="44017-636">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="44017-637">ReadItem</span><span class="sxs-lookup"><span data-stu-id="44017-637">ReadItem</span></span>|
|[<span data-ttu-id="44017-638">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="44017-638">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="44017-639">阅读</span><span class="sxs-lookup"><span data-stu-id="44017-639">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="44017-640">示例</span><span class="sxs-lookup"><span data-stu-id="44017-640">Example</span></span>

```javascript
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
console.log("Sender: " + senderName + " (" + senderAddress + ")");
```

#### <a name="nullable-seriesid-string"></a><span data-ttu-id="44017-641">(nullable) seriesId :String</span><span class="sxs-lookup"><span data-stu-id="44017-641">(nullable) seriesId :String</span></span>

<span data-ttu-id="44017-642">获取实例所属的系列的 ID。</span><span class="sxs-lookup"><span data-stu-id="44017-642">Gets the id of the series that an instance belongs to.</span></span>

<span data-ttu-id="44017-643">在 OWA 和 Outlook 中，`seriesId` 返回此项目所属的父（系列）项目的 Exchange Web 服务 (EWS) ID。</span><span class="sxs-lookup"><span data-stu-id="44017-643">In OWA and Outlook, the `seriesId` returns the Exchange Web Services (EWS) ID of the parent (series) item that this item belongs to.</span></span> <span data-ttu-id="44017-644">但是，在 iOS 和 Android 中，`seriesId` 返回父项目的其余部分 ID。</span><span class="sxs-lookup"><span data-stu-id="44017-644">However, in iOS and Android, the `seriesId` returns the REST ID of the parent item.</span></span>

> [!NOTE]
> <span data-ttu-id="44017-645">`seriesId` 属性返回的标识符与 Exchange Web 服务项目标识符相同。</span><span class="sxs-lookup"><span data-stu-id="44017-645">The identifier returned by the `seriesId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="44017-646">`seriesId` 属性与 Outlook REST API 使用的 Outlook ID 不同。</span><span class="sxs-lookup"><span data-stu-id="44017-646">The `seriesId` property is not identical to the Outlook IDs used by the Outlook REST API.</span></span> <span data-ttu-id="44017-647">使用此值进行 REST API 调用前，应使用 [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string) 对它进行转换。</span><span class="sxs-lookup"><span data-stu-id="44017-647">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="44017-648">有关详细信息，请参阅[从 Outlook 加载项使用 Outlook REST API](https://docs.microsoft.com/outlook/add-ins/use-rest-api)。</span><span class="sxs-lookup"><span data-stu-id="44017-648">For more details, see [Use the Outlook REST APIs from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/use-rest-api).</span></span>

<span data-ttu-id="44017-649">`seriesId` 属性对于没有父项目（如单个约会、系列项目或会议请求）的项目返回 `null`，对于非会议请求的任何其他项目，返回 `undefined`。</span><span class="sxs-lookup"><span data-stu-id="44017-649">The `seriesId` property returns `null` for items that do not have parent items such as single appointments, series items, or meeting requests and returns `undefined` for any other items that are not meeting requests.</span></span>

##### <a name="type"></a><span data-ttu-id="44017-650">Type</span><span class="sxs-lookup"><span data-stu-id="44017-650">Type</span></span>

* <span data-ttu-id="44017-651">String</span><span class="sxs-lookup"><span data-stu-id="44017-651">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="44017-652">要求</span><span class="sxs-lookup"><span data-stu-id="44017-652">Requirements</span></span>

|<span data-ttu-id="44017-653">要求</span><span class="sxs-lookup"><span data-stu-id="44017-653">Requirement</span></span>|<span data-ttu-id="44017-654">值</span><span class="sxs-lookup"><span data-stu-id="44017-654">Value</span></span>|
|---|---|
|[<span data-ttu-id="44017-655">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="44017-655">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="44017-656">1.7</span><span class="sxs-lookup"><span data-stu-id="44017-656">1.7</span></span>|
|[<span data-ttu-id="44017-657">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="44017-657">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="44017-658">ReadItem</span><span class="sxs-lookup"><span data-stu-id="44017-658">ReadItem</span></span>|
|[<span data-ttu-id="44017-659">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="44017-659">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="44017-660">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="44017-660">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="44017-661">示例</span><span class="sxs-lookup"><span data-stu-id="44017-661">Example</span></span>

```javascript
var seriesId = Office.context.mailbox.item.seriesId;

// The seriesId property returns null for items that do
// not have parent items (such as single appointments,
// series items, or meeting requests) and returns
// undefined for messages that are not meeting requests.
var isSeriesInstance = (seriesId != null);
console.log("SeriesId is " + seriesId + " and isSeriesInstance is " + isSeriesInstance);
```

####  <a name="start-datetimejavascriptapioutlookofficetime"></a><span data-ttu-id="44017-662">start :Date|[Time](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="44017-662">start :Date|[Time](/javascript/api/outlook/office.time)</span></span>

<span data-ttu-id="44017-663">获取或设置约会开始的日期和时间。</span><span class="sxs-lookup"><span data-stu-id="44017-663">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="44017-p132">将 `start` 属性表示为协调世界时 (UTC) 的日期和时间值。可使用 [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttime) 方法将该值转换为客户端的本地日期和时间。</span><span class="sxs-lookup"><span data-stu-id="44017-p132">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="44017-666">阅读模式</span><span class="sxs-lookup"><span data-stu-id="44017-666">Read mode</span></span>

<span data-ttu-id="44017-667">`start` 属性返回 `Date` 对象。</span><span class="sxs-lookup"><span data-stu-id="44017-667">The `start` property returns a `Date` object.</span></span>

```javascript
var start = Office.context.mailbox.item.start;
console.log("Appointment start: " + JSON.stringify(start));
```

##### <a name="compose-mode"></a><span data-ttu-id="44017-668">撰写模式</span><span class="sxs-lookup"><span data-stu-id="44017-668">Compose mode</span></span>

<span data-ttu-id="44017-669">`start` 属性返回 `Time` 对象。</span><span class="sxs-lookup"><span data-stu-id="44017-669">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="44017-670">使用 [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) 方法设置开始时间时，应使用 [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) 方法将客户端的本地时间转换为服务器的 UTC。</span><span class="sxs-lookup"><span data-stu-id="44017-670">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="44017-671">以下示例通过使用 `Time` 对象的 [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) 方法，设置撰写模式下约会的开始时间。</span><span class="sxs-lookup"><span data-stu-id="44017-671">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="44017-672">类型</span><span class="sxs-lookup"><span data-stu-id="44017-672">Type</span></span>

*   <span data-ttu-id="44017-673">Date | [Time](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="44017-673">Date | [Time](/javascript/api/outlook/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="44017-674">Requirements</span><span class="sxs-lookup"><span data-stu-id="44017-674">Requirements</span></span>

|<span data-ttu-id="44017-675">要求</span><span class="sxs-lookup"><span data-stu-id="44017-675">Requirement</span></span>|<span data-ttu-id="44017-676">值</span><span class="sxs-lookup"><span data-stu-id="44017-676">Value</span></span>|
|---|---|
|[<span data-ttu-id="44017-677">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="44017-677">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="44017-678">1.0</span><span class="sxs-lookup"><span data-stu-id="44017-678">1.0</span></span>|
|[<span data-ttu-id="44017-679">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="44017-679">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="44017-680">ReadItem</span><span class="sxs-lookup"><span data-stu-id="44017-680">ReadItem</span></span>|
|[<span data-ttu-id="44017-681">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="44017-681">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="44017-682">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="44017-682">Compose or Read</span></span>|

####  <a name="subject-stringsubjectjavascriptapioutlookofficesubject"></a><span data-ttu-id="44017-683">subject :String|[Subject](/javascript/api/outlook/office.subject)</span><span class="sxs-lookup"><span data-stu-id="44017-683">subject :String|[Subject](/javascript/api/outlook/office.subject)</span></span>

<span data-ttu-id="44017-684">获取或设置显示在项目的主题字段中的说明。</span><span class="sxs-lookup"><span data-stu-id="44017-684">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="44017-685">`subject` 属性获取或设置由电子邮件服务器发送项目时的整个主题。</span><span class="sxs-lookup"><span data-stu-id="44017-685">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="44017-686">阅读模式</span><span class="sxs-lookup"><span data-stu-id="44017-686">Read mode</span></span>

<span data-ttu-id="44017-p133">`subject` 属性返回一个字符串。使用 [`normalizedSubject`](#normalizedsubject-string) 属性获取不带任何前导前缀（如 `RE:` 和 `FW:`）的主题。</span><span class="sxs-lookup"><span data-stu-id="44017-p133">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

<span data-ttu-id="44017-689">以下 JavaScript 代码示例显示了如何访问 Outlook 中当前项目的 `subject` 属性。</span><span class="sxs-lookup"><span data-stu-id="44017-689">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

```javascript
var subject = Office.context.mailbox.item.subject;
console.log(subject);
```

##### <a name="compose-mode"></a><span data-ttu-id="44017-690">撰写模式</span><span class="sxs-lookup"><span data-stu-id="44017-690">Compose mode</span></span>
<span data-ttu-id="44017-691">`subject` 属性返回一个 `Subject` 对象，该对象提供用于获取和设置主题的方法。</span><span class="sxs-lookup"><span data-stu-id="44017-691">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```javascript
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
  console.log(subject);
}
```

##### <a name="type"></a><span data-ttu-id="44017-692">类型</span><span class="sxs-lookup"><span data-stu-id="44017-692">Type</span></span>

*   <span data-ttu-id="44017-693">String | [Subject](/javascript/api/outlook/office.subject)</span><span class="sxs-lookup"><span data-stu-id="44017-693">String | [Subject](/javascript/api/outlook/office.subject)</span></span>

##### <a name="requirements"></a><span data-ttu-id="44017-694">要求</span><span class="sxs-lookup"><span data-stu-id="44017-694">Requirements</span></span>

|<span data-ttu-id="44017-695">要求</span><span class="sxs-lookup"><span data-stu-id="44017-695">Requirement</span></span>|<span data-ttu-id="44017-696">值</span><span class="sxs-lookup"><span data-stu-id="44017-696">Value</span></span>|
|---|---|
|[<span data-ttu-id="44017-697">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="44017-697">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="44017-698">1.0</span><span class="sxs-lookup"><span data-stu-id="44017-698">1.0</span></span>|
|[<span data-ttu-id="44017-699">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="44017-699">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="44017-700">ReadItem</span><span class="sxs-lookup"><span data-stu-id="44017-700">ReadItem</span></span>|
|[<span data-ttu-id="44017-701">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="44017-701">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="44017-702">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="44017-702">Compose or Read</span></span>|

####  <a name="to-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="44017-703">to :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="44017-703">to :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="44017-704">提供对邮件的“**收件人**”行上的收件人的访问权限。</span><span class="sxs-lookup"><span data-stu-id="44017-704">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="44017-705">对象的类型和访问级别取决于当前项目的模式。</span><span class="sxs-lookup"><span data-stu-id="44017-705">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="44017-706">阅读模式</span><span class="sxs-lookup"><span data-stu-id="44017-706">Read mode</span></span>

<span data-ttu-id="44017-p135">`to` 属性返回包含邮件的**收件人**行上所列的每个收件人的 `EmailAddressDetails` 对象的数组。集合上限为 100 个成员。</span><span class="sxs-lookup"><span data-stu-id="44017-p135">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.</span></span>

```javascript
console.log(JSON.stringify(Office.context.mailbox.item.to));
```

##### <a name="compose-mode"></a><span data-ttu-id="44017-709">撰写模式</span><span class="sxs-lookup"><span data-stu-id="44017-709">Compose mode</span></span>

<span data-ttu-id="44017-710">`to` 属性返回一个 `Recipients` 对象，该对象提供用于获取或更新邮件的“**收件人**”行上收件人的方法。</span><span class="sxs-lookup"><span data-stu-id="44017-710">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span>

```javascript
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="44017-711">类型</span><span class="sxs-lookup"><span data-stu-id="44017-711">Type</span></span>

*   <span data-ttu-id="44017-712">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="44017-712">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="44017-713">要求</span><span class="sxs-lookup"><span data-stu-id="44017-713">Requirements</span></span>

|<span data-ttu-id="44017-714">要求</span><span class="sxs-lookup"><span data-stu-id="44017-714">Requirement</span></span>|<span data-ttu-id="44017-715">值</span><span class="sxs-lookup"><span data-stu-id="44017-715">Value</span></span>|
|---|---|
|[<span data-ttu-id="44017-716">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="44017-716">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="44017-717">1.0</span><span class="sxs-lookup"><span data-stu-id="44017-717">1.0</span></span>|
|[<span data-ttu-id="44017-718">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="44017-718">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="44017-719">ReadItem</span><span class="sxs-lookup"><span data-stu-id="44017-719">ReadItem</span></span>|
|[<span data-ttu-id="44017-720">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="44017-720">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="44017-721">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="44017-721">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="44017-722">方法</span><span class="sxs-lookup"><span data-stu-id="44017-722">Methods</span></span>

####  <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="44017-723">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="44017-723">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="44017-724">将文件作为附件添加到邮件或约会。</span><span class="sxs-lookup"><span data-stu-id="44017-724">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="44017-725">`addFileAttachmentAsync` 方法在指定的 URI 上载文件并将其附加到撰写窗体中的项目。</span><span class="sxs-lookup"><span data-stu-id="44017-725">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="44017-726">你随后可以将该标识符与 [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) 方法一同使用，以删除同一个会话中的附件。</span><span class="sxs-lookup"><span data-stu-id="44017-726">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="44017-727">Parameters</span><span class="sxs-lookup"><span data-stu-id="44017-727">Parameters</span></span>
|<span data-ttu-id="44017-728">名称</span><span class="sxs-lookup"><span data-stu-id="44017-728">Name</span></span>|<span data-ttu-id="44017-729">类型</span><span class="sxs-lookup"><span data-stu-id="44017-729">Type</span></span>|<span data-ttu-id="44017-730">属性</span><span class="sxs-lookup"><span data-stu-id="44017-730">Attributes</span></span>|<span data-ttu-id="44017-731">说明</span><span class="sxs-lookup"><span data-stu-id="44017-731">Description</span></span>|
|---|---|---|---|
|`uri`|<span data-ttu-id="44017-732">String</span><span class="sxs-lookup"><span data-stu-id="44017-732">String</span></span>||<span data-ttu-id="44017-p136">提供附加到邮件或约会的文件的位置的 URI。最大长度为 2048 个字符。</span><span class="sxs-lookup"><span data-stu-id="44017-p136">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`|<span data-ttu-id="44017-735">字符串</span><span class="sxs-lookup"><span data-stu-id="44017-735">String</span></span>||<span data-ttu-id="44017-p137">在附件上载过程中显示的附件名称。最大长度为 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="44017-p137">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="44017-738">Object</span><span class="sxs-lookup"><span data-stu-id="44017-738">Object</span></span>|<span data-ttu-id="44017-739">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="44017-739">&lt;optional&gt;</span></span>|<span data-ttu-id="44017-740">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="44017-740">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="44017-741">对象</span><span class="sxs-lookup"><span data-stu-id="44017-741">Object</span></span>|<span data-ttu-id="44017-742">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="44017-742">&lt;optional&gt;</span></span>|<span data-ttu-id="44017-743">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="44017-743">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.isInline`|<span data-ttu-id="44017-744">布尔值</span><span class="sxs-lookup"><span data-stu-id="44017-744">Boolean</span></span>|<span data-ttu-id="44017-745">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="44017-745">&lt;optional&gt;</span></span>|<span data-ttu-id="44017-746">如果为 `true`，则表示附件将在邮件正文中内联显示，并且不应显示在附件列表中。</span><span class="sxs-lookup"><span data-stu-id="44017-746">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`callback`|<span data-ttu-id="44017-747">函数</span><span class="sxs-lookup"><span data-stu-id="44017-747">function</span></span>|<span data-ttu-id="44017-748">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="44017-748">&lt;optional&gt;</span></span>|<span data-ttu-id="44017-749">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="44017-749">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="44017-750">如果成功，附件标识符将在 `asyncResult.value` 属性中提供。</span><span class="sxs-lookup"><span data-stu-id="44017-750">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="44017-751">如果上传附件失败，`asyncResult` 对象将包含一个提供错误说明的 `Error` 对象。</span><span class="sxs-lookup"><span data-stu-id="44017-751">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="44017-752">错误</span><span class="sxs-lookup"><span data-stu-id="44017-752">Errors</span></span>

|<span data-ttu-id="44017-753">错误代码</span><span class="sxs-lookup"><span data-stu-id="44017-753">Error code</span></span>|<span data-ttu-id="44017-754">说明</span><span class="sxs-lookup"><span data-stu-id="44017-754">Description</span></span>|
|------------|-------------|
|`AttachmentSizeExceeded`|<span data-ttu-id="44017-755">附件大小超过了允许的大小。</span><span class="sxs-lookup"><span data-stu-id="44017-755">The attachment is larger than allowed.</span></span>|
|`FileTypeNotSupported`|<span data-ttu-id="44017-756">该附件的扩展名不是允许的扩展名。</span><span class="sxs-lookup"><span data-stu-id="44017-756">The attachment has an extension that is not allowed.</span></span>|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="44017-757">邮件或约会具有的附件过多。</span><span class="sxs-lookup"><span data-stu-id="44017-757">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="44017-758">要求</span><span class="sxs-lookup"><span data-stu-id="44017-758">Requirements</span></span>

|<span data-ttu-id="44017-759">要求</span><span class="sxs-lookup"><span data-stu-id="44017-759">Requirement</span></span>|<span data-ttu-id="44017-760">值</span><span class="sxs-lookup"><span data-stu-id="44017-760">Value</span></span>|
|---|---|
|[<span data-ttu-id="44017-761">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="44017-761">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="44017-762">1.1</span><span class="sxs-lookup"><span data-stu-id="44017-762">1.1</span></span>|
|[<span data-ttu-id="44017-763">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="44017-763">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="44017-764">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="44017-764">ReadWriteItem</span></span>|
|[<span data-ttu-id="44017-765">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="44017-765">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="44017-766">撰写</span><span class="sxs-lookup"><span data-stu-id="44017-766">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="44017-767">示例</span><span class="sxs-lookup"><span data-stu-id="44017-767">Examples</span></span>

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

<span data-ttu-id="44017-768">以下示例将图像文件添加为内联附件，并在邮件正文中引用该附件。</span><span class="sxs-lookup"><span data-stu-id="44017-768">The following example adds an image file as an inline attachment and references the attachment in the message body.</span></span>

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

#### <a name="addfileattachmentfrombase64asyncbase64file-attachmentname-options-callback"></a><span data-ttu-id="44017-769">addFileAttachmentFromBase64Async(base64File, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="44017-769">addFileAttachmentFromBase64Async(base64File, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="44017-770">将 base64 编码中的文件作为附件添加到消息或约会。</span><span class="sxs-lookup"><span data-stu-id="44017-770">Adds a file from the base64 encoding to a message or appointment as an attachment.</span></span>

<span data-ttu-id="44017-771">`addFileAttachmentFromBase64Async` 方法从 base64 编码上传文件，并将其附加到撰写表单中的项目。</span><span class="sxs-lookup"><span data-stu-id="44017-771">The `addFileAttachmentFromBase64Async` method uploads the file from the base64 encoding and attaches it to the item in the compose form.</span></span> <span data-ttu-id="44017-772">此方法返回 AsyncResult.value 对象中的附件标识符。</span><span class="sxs-lookup"><span data-stu-id="44017-772">This method returns the attachment identifier in the AsyncResult.value object.</span></span>

<span data-ttu-id="44017-773">随后可以将该标识符与 [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) 方法一同使用，以删除同一个会话中的附件。</span><span class="sxs-lookup"><span data-stu-id="44017-773">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="44017-774">Parameters</span><span class="sxs-lookup"><span data-stu-id="44017-774">Parameters</span></span>
|<span data-ttu-id="44017-775">名称</span><span class="sxs-lookup"><span data-stu-id="44017-775">Name</span></span>|<span data-ttu-id="44017-776">类型</span><span class="sxs-lookup"><span data-stu-id="44017-776">Type</span></span>|<span data-ttu-id="44017-777">属性</span><span class="sxs-lookup"><span data-stu-id="44017-777">Attributes</span></span>|<span data-ttu-id="44017-778">说明</span><span class="sxs-lookup"><span data-stu-id="44017-778">Description</span></span>|
|---|---|---|---|
|`base64File`|<span data-ttu-id="44017-779">String</span><span class="sxs-lookup"><span data-stu-id="44017-779">String</span></span>||<span data-ttu-id="44017-780">要添加到电子邮件或事件的图像或文件的 base64 编码内容。</span><span class="sxs-lookup"><span data-stu-id="44017-780">The base64 encoded content of an image or file to be added to an email or event.</span></span>|
|`attachmentName`|<span data-ttu-id="44017-781">字符串</span><span class="sxs-lookup"><span data-stu-id="44017-781">String</span></span>||<span data-ttu-id="44017-p139">在附件上载过程中显示的附件名称。最大长度为 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="44017-p139">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="44017-784">Object</span><span class="sxs-lookup"><span data-stu-id="44017-784">Object</span></span>|<span data-ttu-id="44017-785">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="44017-785">&lt;optional&gt;</span></span>|<span data-ttu-id="44017-786">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="44017-786">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="44017-787">对象</span><span class="sxs-lookup"><span data-stu-id="44017-787">Object</span></span>|<span data-ttu-id="44017-788">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="44017-788">&lt;optional&gt;</span></span>|<span data-ttu-id="44017-789">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="44017-789">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.isInline`|<span data-ttu-id="44017-790">布尔值</span><span class="sxs-lookup"><span data-stu-id="44017-790">Boolean</span></span>|<span data-ttu-id="44017-791">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="44017-791">&lt;optional&gt;</span></span>|<span data-ttu-id="44017-792">如果为 `true`，则表示附件将在邮件正文中内联显示，并且不应显示在附件列表中。</span><span class="sxs-lookup"><span data-stu-id="44017-792">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`callback`|<span data-ttu-id="44017-793">函数</span><span class="sxs-lookup"><span data-stu-id="44017-793">function</span></span>|<span data-ttu-id="44017-794">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="44017-794">&lt;optional&gt;</span></span>|<span data-ttu-id="44017-795">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="44017-795">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="44017-796">如果成功，附件标识符将在 `asyncResult.value` 属性中提供。</span><span class="sxs-lookup"><span data-stu-id="44017-796">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="44017-797">如果上传附件失败，`asyncResult` 对象将包含一个提供错误说明的 `Error` 对象。</span><span class="sxs-lookup"><span data-stu-id="44017-797">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="44017-798">错误</span><span class="sxs-lookup"><span data-stu-id="44017-798">Errors</span></span>

|<span data-ttu-id="44017-799">错误代码</span><span class="sxs-lookup"><span data-stu-id="44017-799">Error code</span></span>|<span data-ttu-id="44017-800">说明</span><span class="sxs-lookup"><span data-stu-id="44017-800">Description</span></span>|
|------------|-------------|
|`AttachmentSizeExceeded`|<span data-ttu-id="44017-801">附件大小超过了允许的大小。</span><span class="sxs-lookup"><span data-stu-id="44017-801">The attachment is larger than allowed.</span></span>|
|`FileTypeNotSupported`|<span data-ttu-id="44017-802">该附件的扩展名不是允许的扩展名。</span><span class="sxs-lookup"><span data-stu-id="44017-802">The attachment has an extension that is not allowed.</span></span>|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="44017-803">邮件或约会具有的附件过多。</span><span class="sxs-lookup"><span data-stu-id="44017-803">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="44017-804">Requirements</span><span class="sxs-lookup"><span data-stu-id="44017-804">Requirements</span></span>

|<span data-ttu-id="44017-805">要求</span><span class="sxs-lookup"><span data-stu-id="44017-805">Requirement</span></span>|<span data-ttu-id="44017-806">值</span><span class="sxs-lookup"><span data-stu-id="44017-806">Value</span></span>|
|---|---|
|[<span data-ttu-id="44017-807">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="44017-807">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="44017-808">预览</span><span class="sxs-lookup"><span data-stu-id="44017-808">Preview</span></span>|
|[<span data-ttu-id="44017-809">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="44017-809">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="44017-810">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="44017-810">ReadWriteItem</span></span>|
|[<span data-ttu-id="44017-811">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="44017-811">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="44017-812">撰写</span><span class="sxs-lookup"><span data-stu-id="44017-812">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="44017-813">示例</span><span class="sxs-lookup"><span data-stu-id="44017-813">Examples</span></span>

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

####  <a name="addhandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="44017-814">addHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="44017-814">addHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="44017-815">添加支持事件的事件处理程序。</span><span class="sxs-lookup"><span data-stu-id="44017-815">Adds an event handler for a supported event.</span></span>

<span data-ttu-id="44017-816">当前，支持的事件类型是 `Office.EventType.AttachmentsChanged`、`Office.EventType.AppointmentTimeChanged`、`Office.EventType.RecipientsChanged` 和 `Office.EventType.RecurrenceChanged`</span><span class="sxs-lookup"><span data-stu-id="44017-816">Currently the supported event types are `Office.EventType.AttachmentsChanged`, `Office.EventType.AppointmentTimeChanged`, `Office.EventType.RecipientsChanged`, and `Office.EventType.RecurrenceChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="44017-817">Parameters</span><span class="sxs-lookup"><span data-stu-id="44017-817">Parameters</span></span>

| <span data-ttu-id="44017-818">名称</span><span class="sxs-lookup"><span data-stu-id="44017-818">Name</span></span> | <span data-ttu-id="44017-819">类型</span><span class="sxs-lookup"><span data-stu-id="44017-819">Type</span></span> | <span data-ttu-id="44017-820">属性</span><span class="sxs-lookup"><span data-stu-id="44017-820">Attributes</span></span> | <span data-ttu-id="44017-821">说明</span><span class="sxs-lookup"><span data-stu-id="44017-821">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="44017-822">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="44017-822">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="44017-823">应调用处理程序的事件。</span><span class="sxs-lookup"><span data-stu-id="44017-823">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="44017-824">函数</span><span class="sxs-lookup"><span data-stu-id="44017-824">Function</span></span> || <span data-ttu-id="44017-p140">用于处理事件的函数。此函数必须接受一个参数，即对象文本。参数上的 `type` 属性将匹配传递给 `addHandlerAsync` 的 `eventType` 参数。</span><span class="sxs-lookup"><span data-stu-id="44017-p140">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `addHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="44017-828">Object</span><span class="sxs-lookup"><span data-stu-id="44017-828">Object</span></span> | <span data-ttu-id="44017-829">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="44017-829">&lt;optional&gt;</span></span> | <span data-ttu-id="44017-830">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="44017-830">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="44017-831">对象</span><span class="sxs-lookup"><span data-stu-id="44017-831">Object</span></span> | <span data-ttu-id="44017-832">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="44017-832">&lt;optional&gt;</span></span> | <span data-ttu-id="44017-833">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="44017-833">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="44017-834">函数</span><span class="sxs-lookup"><span data-stu-id="44017-834">function</span></span>| <span data-ttu-id="44017-835">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="44017-835">&lt;optional&gt;</span></span>|<span data-ttu-id="44017-836">方法完成后，使用单个参数 `callback`（一个 [`asyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `AsyncResult` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="44017-836">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="44017-837">Requirements</span><span class="sxs-lookup"><span data-stu-id="44017-837">Requirements</span></span>

|<span data-ttu-id="44017-838">要求</span><span class="sxs-lookup"><span data-stu-id="44017-838">Requirement</span></span>| <span data-ttu-id="44017-839">值</span><span class="sxs-lookup"><span data-stu-id="44017-839">Value</span></span>|
|---|---|
|[<span data-ttu-id="44017-840">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="44017-840">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="44017-841">1.7</span><span class="sxs-lookup"><span data-stu-id="44017-841">1.7</span></span> |
|[<span data-ttu-id="44017-842">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="44017-842">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="44017-843">ReadItem</span><span class="sxs-lookup"><span data-stu-id="44017-843">ReadItem</span></span> |
|[<span data-ttu-id="44017-844">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="44017-844">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="44017-845">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="44017-845">Compose or Read</span></span> |

##### <a name="example"></a><span data-ttu-id="44017-846">示例</span><span class="sxs-lookup"><span data-stu-id="44017-846">Example</span></span>

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

####  <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="44017-847">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="44017-847">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="44017-848">将 Exchange 项目（如邮件）作为附件添加到邮件或约会。</span><span class="sxs-lookup"><span data-stu-id="44017-848">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="44017-p141">`addItemAttachmentAsync` 方法将包含指定 Exchange 标识符的项目附加到撰写窗体中的项目。如果指定一个回调方法，此方法使用 `asyncResult` 参数调用，该参数包含一个附件标识符或代码，指示附加项目过程中出现的任何错误。可以使用 `options` 参数将状态信息传递给回调方法（如果需要）。</span><span class="sxs-lookup"><span data-stu-id="44017-p141">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="44017-852">随后可以将该标识符与 [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) 方法一同使用，以删除同一个会话中的附件。</span><span class="sxs-lookup"><span data-stu-id="44017-852">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="44017-853">如果 Office 加载项在 Outlook Web App 中运行，则 `addItemAttachmentAsync` 方法可以将项目附加到项目（正在编辑的项目除外）中；然而，不支持也不建议这样做。</span><span class="sxs-lookup"><span data-stu-id="44017-853">If your Office Add-in is running in Outlook Web App, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="44017-854">Parameters</span><span class="sxs-lookup"><span data-stu-id="44017-854">Parameters</span></span>

|<span data-ttu-id="44017-855">名称</span><span class="sxs-lookup"><span data-stu-id="44017-855">Name</span></span>|<span data-ttu-id="44017-856">类型</span><span class="sxs-lookup"><span data-stu-id="44017-856">Type</span></span>|<span data-ttu-id="44017-857">属性</span><span class="sxs-lookup"><span data-stu-id="44017-857">Attributes</span></span>|<span data-ttu-id="44017-858">说明</span><span class="sxs-lookup"><span data-stu-id="44017-858">Description</span></span>|
|---|---|---|---|
|`itemId`|<span data-ttu-id="44017-859">String</span><span class="sxs-lookup"><span data-stu-id="44017-859">String</span></span>||<span data-ttu-id="44017-p142">要附加的项目的 Exchange 标识符。最大长度为 100 个字符。</span><span class="sxs-lookup"><span data-stu-id="44017-p142">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`|<span data-ttu-id="44017-862">字符串</span><span class="sxs-lookup"><span data-stu-id="44017-862">String</span></span>||<span data-ttu-id="44017-863">要附加的项目的主题。</span><span class="sxs-lookup"><span data-stu-id="44017-863">The subject of the item to be attached.</span></span> <span data-ttu-id="44017-864">最大长度为 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="44017-864">The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="44017-865">对象</span><span class="sxs-lookup"><span data-stu-id="44017-865">Object</span></span>|<span data-ttu-id="44017-866">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="44017-866">&lt;optional&gt;</span></span>|<span data-ttu-id="44017-867">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="44017-867">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="44017-868">对象</span><span class="sxs-lookup"><span data-stu-id="44017-868">Object</span></span>|<span data-ttu-id="44017-869">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="44017-869">&lt;optional&gt;</span></span>|<span data-ttu-id="44017-870">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="44017-870">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="44017-871">函数</span><span class="sxs-lookup"><span data-stu-id="44017-871">function</span></span>|<span data-ttu-id="44017-872">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="44017-872">&lt;optional&gt;</span></span>|<span data-ttu-id="44017-873">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="44017-873">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="44017-874">如果成功，附件标识符将在 `asyncResult.value` 属性中提供。</span><span class="sxs-lookup"><span data-stu-id="44017-874">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="44017-875">如果添加附件失败，`asyncResult` 对象将包含一个提供错误说明的 `Error` 对象。</span><span class="sxs-lookup"><span data-stu-id="44017-875">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="44017-876">错误</span><span class="sxs-lookup"><span data-stu-id="44017-876">Errors</span></span>

|<span data-ttu-id="44017-877">错误代码</span><span class="sxs-lookup"><span data-stu-id="44017-877">Error code</span></span>|<span data-ttu-id="44017-878">说明</span><span class="sxs-lookup"><span data-stu-id="44017-878">Description</span></span>|
|------------|-------------|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="44017-879">邮件或约会具有的附件过多。</span><span class="sxs-lookup"><span data-stu-id="44017-879">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="44017-880">Requirements</span><span class="sxs-lookup"><span data-stu-id="44017-880">Requirements</span></span>

|<span data-ttu-id="44017-881">要求</span><span class="sxs-lookup"><span data-stu-id="44017-881">Requirement</span></span>|<span data-ttu-id="44017-882">值</span><span class="sxs-lookup"><span data-stu-id="44017-882">Value</span></span>|
|---|---|
|[<span data-ttu-id="44017-883">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="44017-883">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="44017-884">1.1</span><span class="sxs-lookup"><span data-stu-id="44017-884">1.1</span></span>|
|[<span data-ttu-id="44017-885">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="44017-885">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="44017-886">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="44017-886">ReadWriteItem</span></span>|
|[<span data-ttu-id="44017-887">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="44017-887">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="44017-888">撰写</span><span class="sxs-lookup"><span data-stu-id="44017-888">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="44017-889">示例</span><span class="sxs-lookup"><span data-stu-id="44017-889">Example</span></span>

<span data-ttu-id="44017-890">以下示例将现有的 Outlook 项目添加为名为 `My Attachment` 的附件。</span><span class="sxs-lookup"><span data-stu-id="44017-890">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

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

####  <a name="close"></a><span data-ttu-id="44017-891">close()</span><span class="sxs-lookup"><span data-stu-id="44017-891">close()</span></span>

<span data-ttu-id="44017-892">关闭当前正在撰写的项目。</span><span class="sxs-lookup"><span data-stu-id="44017-892">Closes the current item that is being composed.</span></span>

<span data-ttu-id="44017-p144">
            \`close\` 方法的行为取决于要撰写的项目的当前状态。如果项目具有未保存的更改，客户端将提示用户保存、放弃或取消关闭操作。</span><span class="sxs-lookup"><span data-stu-id="44017-p144">The behavior of the `close` method depends on the current state of the item being composed. If the item has unsaved changes, the client prompts the user to save, discard, or cancel the close action.</span></span>

> [!NOTE]
> <span data-ttu-id="44017-895">在 Outlook 网页版中，如果该项目是约会并且之前已使用 `saveAsync` 保存，则即使自上次保存项目后未发生任何更改，也会提示用户保存、放弃或取消。</span><span class="sxs-lookup"><span data-stu-id="44017-895">In Outlook on the web, if the item is an appointment and it has previously been saved using `saveAsync`, the user is prompted to save, discard, or cancel even if no changes have occurred since the item was last saved.</span></span>

<span data-ttu-id="44017-896">在 Outlook 桌面客户端中，如果邮件是内联答复，`close` 方法不起作用。</span><span class="sxs-lookup"><span data-stu-id="44017-896">In the Outlook desktop client, if the message is an inline reply, the `close` method has no effect.</span></span>

##### <a name="requirements"></a><span data-ttu-id="44017-897">要求</span><span class="sxs-lookup"><span data-stu-id="44017-897">Requirements</span></span>

|<span data-ttu-id="44017-898">要求</span><span class="sxs-lookup"><span data-stu-id="44017-898">Requirement</span></span>|<span data-ttu-id="44017-899">值</span><span class="sxs-lookup"><span data-stu-id="44017-899">Value</span></span>|
|---|---|
|[<span data-ttu-id="44017-900">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="44017-900">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="44017-901">1.3</span><span class="sxs-lookup"><span data-stu-id="44017-901">1.3</span></span>|
|[<span data-ttu-id="44017-902">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="44017-902">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="44017-903">受限</span><span class="sxs-lookup"><span data-stu-id="44017-903">Restricted</span></span>|
|[<span data-ttu-id="44017-904">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="44017-904">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="44017-905">撰写</span><span class="sxs-lookup"><span data-stu-id="44017-905">Compose</span></span>|

#### <a name="displayreplyallformformdata-callback"></a><span data-ttu-id="44017-906">displayReplyAllForm (formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="44017-906">displayReplyAllForm(formData, [callback])</span></span>

<span data-ttu-id="44017-907">显示答复窗体，其中包括所选邮件的发件人和所有收件人或所选约会的组织者和所有与会者。</span><span class="sxs-lookup"><span data-stu-id="44017-907">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="44017-908">在 Outlook for iOS 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="44017-908">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="44017-909">在 Outlook Web App 中，答复窗体显示为包含 3 列视图的弹出式窗体以及包含 2 列或 1 列视图的弹出式窗体。</span><span class="sxs-lookup"><span data-stu-id="44017-909">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="44017-910">如果任意字符串参数超出其限制，`displayReplyAllForm` 将引发异常。</span><span class="sxs-lookup"><span data-stu-id="44017-910">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="44017-p145">当在 `formData.attachments` 参数中指定附件时，Outlook 和 Outlook Web App 尝试下载所有附件并将其附加到答复窗体。如果无法添加任何附件，则在窗体 UI 中显示错误。如果这不可能，则不引发错误消息。</span><span class="sxs-lookup"><span data-stu-id="44017-p145">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="44017-914">Parameters</span><span class="sxs-lookup"><span data-stu-id="44017-914">Parameters</span></span>

|<span data-ttu-id="44017-915">名称</span><span class="sxs-lookup"><span data-stu-id="44017-915">Name</span></span>|<span data-ttu-id="44017-916">类型</span><span class="sxs-lookup"><span data-stu-id="44017-916">Type</span></span>|<span data-ttu-id="44017-917">属性</span><span class="sxs-lookup"><span data-stu-id="44017-917">Attributes</span></span>|<span data-ttu-id="44017-918">说明</span><span class="sxs-lookup"><span data-stu-id="44017-918">Description</span></span>|
|---|---|---|---|
|`formData`|<span data-ttu-id="44017-919">字符串 &#124; 对象</span><span class="sxs-lookup"><span data-stu-id="44017-919">String &#124; Object</span></span>||<span data-ttu-id="44017-p146">一个包含文本和 HTML 且表示答复窗体的正文的字符串。字符串限制为 32 KB。</span><span class="sxs-lookup"><span data-stu-id="44017-p146">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="44017-922">**或**</span><span class="sxs-lookup"><span data-stu-id="44017-922">**OR**</span></span><br/><span data-ttu-id="44017-p147">包含正文或附件数据和回调函数的对象。对象定义如下。</span><span class="sxs-lookup"><span data-stu-id="44017-p147">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span>|
|`formData.htmlBody`|<span data-ttu-id="44017-925">String</span><span class="sxs-lookup"><span data-stu-id="44017-925">String</span></span>|<span data-ttu-id="44017-926">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="44017-926">&lt;optional&gt;</span></span>|<span data-ttu-id="44017-p148">一个包含文本和 HTML 且表示答复窗体的正文的字符串。字符串限制为 32 KB。</span><span class="sxs-lookup"><span data-stu-id="44017-p148">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
|`formData.attachments`|<span data-ttu-id="44017-929">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="44017-929">Array.&lt;Object&gt;</span></span>|<span data-ttu-id="44017-930">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="44017-930">&lt;optional&gt;</span></span>|<span data-ttu-id="44017-931">JSON 对象（是文件或项目附件）的数组。</span><span class="sxs-lookup"><span data-stu-id="44017-931">An array of JSON objects that are either file or item attachments.</span></span>|
|`formData.attachments.type`|<span data-ttu-id="44017-932">String</span><span class="sxs-lookup"><span data-stu-id="44017-932">String</span></span>||<span data-ttu-id="44017-p149">指示附件的类型。必须是文件附件的 `file` 或项目附件的 `item`。</span><span class="sxs-lookup"><span data-stu-id="44017-p149">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span>|
|`formData.attachments.name`|<span data-ttu-id="44017-935">字符串</span><span class="sxs-lookup"><span data-stu-id="44017-935">String</span></span>||<span data-ttu-id="44017-936">一个包含附件的名称的字符串，最多包含 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="44017-936">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
|`formData.attachments.url`|<span data-ttu-id="44017-937">String</span><span class="sxs-lookup"><span data-stu-id="44017-937">String</span></span>||<span data-ttu-id="44017-p150">仅在将 `type` 设置为 `file` 时使用。文件的位置的 URI。</span><span class="sxs-lookup"><span data-stu-id="44017-p150">Only used if `type` is set to `file`. The URI of the location for the file.</span></span>|
|`formData.attachments.isInline`|<span data-ttu-id="44017-940">Boolean</span><span class="sxs-lookup"><span data-stu-id="44017-940">Boolean</span></span>||<span data-ttu-id="44017-p151">仅在将 `type` 设置为 `file` 时使用。如果为 `true`，则表示附件将在邮件正文中内联显示，并且不应显示在附件列表中。</span><span class="sxs-lookup"><span data-stu-id="44017-p151">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`formData.attachments.itemId`|<span data-ttu-id="44017-943">String</span><span class="sxs-lookup"><span data-stu-id="44017-943">String</span></span>||<span data-ttu-id="44017-p152">仅在将 `type` 设置为 `item` 时使用。附件的 EWS 项目 ID。字符串最长为 100 个字符。</span><span class="sxs-lookup"><span data-stu-id="44017-p152">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span>|
|`callback`|<span data-ttu-id="44017-947">函数</span><span class="sxs-lookup"><span data-stu-id="44017-947">function</span></span>|<span data-ttu-id="44017-948">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="44017-948">&lt;optional&gt;</span></span>|<span data-ttu-id="44017-949">方法完成后，使用单个参数 `asyncResult`（一个 [AsyncResult](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="44017-949">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="44017-950">要求</span><span class="sxs-lookup"><span data-stu-id="44017-950">Requirements</span></span>

|<span data-ttu-id="44017-951">要求</span><span class="sxs-lookup"><span data-stu-id="44017-951">Requirement</span></span>|<span data-ttu-id="44017-952">值</span><span class="sxs-lookup"><span data-stu-id="44017-952">Value</span></span>|
|---|---|
|[<span data-ttu-id="44017-953">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="44017-953">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="44017-954">1.0</span><span class="sxs-lookup"><span data-stu-id="44017-954">1.0</span></span>|
|[<span data-ttu-id="44017-955">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="44017-955">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="44017-956">ReadItem</span><span class="sxs-lookup"><span data-stu-id="44017-956">ReadItem</span></span>|
|[<span data-ttu-id="44017-957">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="44017-957">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="44017-958">阅读</span><span class="sxs-lookup"><span data-stu-id="44017-958">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="44017-959">示例</span><span class="sxs-lookup"><span data-stu-id="44017-959">Examples</span></span>

<span data-ttu-id="44017-960">以下代码将一个字符串传递到 `displayReplyAllForm` 函数。</span><span class="sxs-lookup"><span data-stu-id="44017-960">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="44017-961">使用空白正文答复。</span><span class="sxs-lookup"><span data-stu-id="44017-961">Reply with an empty body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="44017-962">仅使用正文答复。</span><span class="sxs-lookup"><span data-stu-id="44017-962">Reply with just a body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="44017-963">使用正文和文件附件答复。</span><span class="sxs-lookup"><span data-stu-id="44017-963">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="44017-964">使用正文和项目附件答复。</span><span class="sxs-lookup"><span data-stu-id="44017-964">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="44017-965">使用正文、文件附件、项目附件和回调答复。</span><span class="sxs-lookup"><span data-stu-id="44017-965">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="displayreplyformformdata-callback"></a><span data-ttu-id="44017-966">displayReplyForm (formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="44017-966">displayReplyForm(formData, [callback])</span></span>

<span data-ttu-id="44017-967">显示答复窗体，其中仅包括所选邮件的发件人或所选约会的组织者。</span><span class="sxs-lookup"><span data-stu-id="44017-967">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="44017-968">在 Outlook for iOS 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="44017-968">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="44017-969">在 Outlook Web App 中，答复窗体显示为包含 3 列视图的弹出式窗体以及包含 2 列或 1 列视图的弹出式窗体。</span><span class="sxs-lookup"><span data-stu-id="44017-969">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="44017-970">如果任意字符串参数超出其限制，`displayReplyForm` 将引发异常。</span><span class="sxs-lookup"><span data-stu-id="44017-970">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="44017-p153">当在 `formData.attachments` 参数中指定附件时，Outlook 和 Outlook Web App 尝试下载所有附件并将其附加到答复窗体。如果无法添加任何附件，则在窗体 UI 中显示错误。如果这不可能，则不引发错误消息。</span><span class="sxs-lookup"><span data-stu-id="44017-p153">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="44017-974">Parameters</span><span class="sxs-lookup"><span data-stu-id="44017-974">Parameters</span></span>

|<span data-ttu-id="44017-975">名称</span><span class="sxs-lookup"><span data-stu-id="44017-975">Name</span></span>|<span data-ttu-id="44017-976">类型</span><span class="sxs-lookup"><span data-stu-id="44017-976">Type</span></span>|<span data-ttu-id="44017-977">属性</span><span class="sxs-lookup"><span data-stu-id="44017-977">Attributes</span></span>|<span data-ttu-id="44017-978">说明</span><span class="sxs-lookup"><span data-stu-id="44017-978">Description</span></span>|
|---|---|---|---|
|`formData`|<span data-ttu-id="44017-979">字符串 &#124; 对象</span><span class="sxs-lookup"><span data-stu-id="44017-979">String &#124; Object</span></span>||<span data-ttu-id="44017-p154">一个包含文本和 HTML 且表示答复窗体的正文的字符串。字符串限制为 32 KB。</span><span class="sxs-lookup"><span data-stu-id="44017-p154">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="44017-982">**或**</span><span class="sxs-lookup"><span data-stu-id="44017-982">**OR**</span></span><br/><span data-ttu-id="44017-p155">包含正文或附件数据和回调函数的对象。对象定义如下。</span><span class="sxs-lookup"><span data-stu-id="44017-p155">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span>|
|`formData.htmlBody`|<span data-ttu-id="44017-985">String</span><span class="sxs-lookup"><span data-stu-id="44017-985">String</span></span>|<span data-ttu-id="44017-986">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="44017-986">&lt;optional&gt;</span></span>|<span data-ttu-id="44017-p156">一个包含文本和 HTML 且表示答复窗体的正文的字符串。字符串限制为 32 KB。</span><span class="sxs-lookup"><span data-stu-id="44017-p156">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
|`formData.attachments`|<span data-ttu-id="44017-989">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="44017-989">Array.&lt;Object&gt;</span></span>|<span data-ttu-id="44017-990">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="44017-990">&lt;optional&gt;</span></span>|<span data-ttu-id="44017-991">JSON 对象（是文件或项目附件）的数组。</span><span class="sxs-lookup"><span data-stu-id="44017-991">An array of JSON objects that are either file or item attachments.</span></span>|
|`formData.attachments.type`|<span data-ttu-id="44017-992">String</span><span class="sxs-lookup"><span data-stu-id="44017-992">String</span></span>||<span data-ttu-id="44017-p157">指示附件的类型。必须是文件附件的 `file` 或项目附件的 `item`。</span><span class="sxs-lookup"><span data-stu-id="44017-p157">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span>|
|`formData.attachments.name`|<span data-ttu-id="44017-995">字符串</span><span class="sxs-lookup"><span data-stu-id="44017-995">String</span></span>||<span data-ttu-id="44017-996">一个包含附件的名称的字符串，最多包含 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="44017-996">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
|`formData.attachments.url`|<span data-ttu-id="44017-997">String</span><span class="sxs-lookup"><span data-stu-id="44017-997">String</span></span>||<span data-ttu-id="44017-p158">仅在将 `type` 设置为 `file` 时使用。文件的位置的 URI。</span><span class="sxs-lookup"><span data-stu-id="44017-p158">Only used if `type` is set to `file`. The URI of the location for the file.</span></span>|
|`formData.attachments.isInline`|<span data-ttu-id="44017-1000">Boolean</span><span class="sxs-lookup"><span data-stu-id="44017-1000">Boolean</span></span>||<span data-ttu-id="44017-p159">仅在将 `type` 设置为 `file` 时使用。如果为 `true`，则表示附件将在邮件正文中内联显示，并且不应显示在附件列表中。</span><span class="sxs-lookup"><span data-stu-id="44017-p159">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`formData.attachments.itemId`|<span data-ttu-id="44017-1003">String</span><span class="sxs-lookup"><span data-stu-id="44017-1003">String</span></span>||<span data-ttu-id="44017-p160">仅在将 `type` 设置为 `item` 时使用。附件的 EWS 项目 ID。字符串最长为 100 个字符。</span><span class="sxs-lookup"><span data-stu-id="44017-p160">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span>|
|`callback`|<span data-ttu-id="44017-1007">函数</span><span class="sxs-lookup"><span data-stu-id="44017-1007">function</span></span>|<span data-ttu-id="44017-1008">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="44017-1008">&lt;optional&gt;</span></span>|<span data-ttu-id="44017-1009">方法完成后，使用单个参数 `asyncResult`（一个 [AsyncResult](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="44017-1009">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="44017-1010">要求</span><span class="sxs-lookup"><span data-stu-id="44017-1010">Requirements</span></span>

|<span data-ttu-id="44017-1011">要求</span><span class="sxs-lookup"><span data-stu-id="44017-1011">Requirement</span></span>|<span data-ttu-id="44017-1012">值</span><span class="sxs-lookup"><span data-stu-id="44017-1012">Value</span></span>|
|---|---|
|[<span data-ttu-id="44017-1013">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="44017-1013">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="44017-1014">1.0</span><span class="sxs-lookup"><span data-stu-id="44017-1014">1.0</span></span>|
|[<span data-ttu-id="44017-1015">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="44017-1015">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="44017-1016">ReadItem</span><span class="sxs-lookup"><span data-stu-id="44017-1016">ReadItem</span></span>|
|[<span data-ttu-id="44017-1017">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="44017-1017">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="44017-1018">阅读</span><span class="sxs-lookup"><span data-stu-id="44017-1018">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="44017-1019">示例</span><span class="sxs-lookup"><span data-stu-id="44017-1019">Examples</span></span>

<span data-ttu-id="44017-1020">以下代码将一个字符串传递到 `displayReplyForm` 函数。</span><span class="sxs-lookup"><span data-stu-id="44017-1020">The following code passes a string to the `displayReplyForm` function.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="44017-1021">使用空白正文答复。</span><span class="sxs-lookup"><span data-stu-id="44017-1021">Reply with an empty body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="44017-1022">仅使用正文答复。</span><span class="sxs-lookup"><span data-stu-id="44017-1022">Reply with just a body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="44017-1023">使用正文和文件附件答复。</span><span class="sxs-lookup"><span data-stu-id="44017-1023">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="44017-1024">使用正文和项目附件答复。</span><span class="sxs-lookup"><span data-stu-id="44017-1024">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="44017-1025">使用正文、文件附件、项目附件和回调答复。</span><span class="sxs-lookup"><span data-stu-id="44017-1025">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="getattachmentcontentasyncattachmentid-options-callback--attachmentcontentjavascriptapioutlookofficeattachmentcontent"></a><span data-ttu-id="44017-1026">getAttachmentContentAsync (attachmentId, [options], [callback]) → [AttachmentContent](/javascript/api/outlook/office.attachmentcontent)</span><span class="sxs-lookup"><span data-stu-id="44017-1026">getAttachmentContentAsync(attachmentId, [options], [callback]) → [AttachmentContent](/javascript/api/outlook/office.attachmentcontent)</span></span>

<span data-ttu-id="44017-1027">从消息或约会中获取指定的附件，并将其作为 `AttachmentContent` 对象返回。</span><span class="sxs-lookup"><span data-stu-id="44017-1027">Gets the specified attachment from a message or appointment and returns it as an `AttachmentContent` object.</span></span>

<span data-ttu-id="44017-1028">`getAttachmentContentAsync` 方法获取项目中带指定标识符的附件。</span><span class="sxs-lookup"><span data-stu-id="44017-1028">The `getAttachmentContentAsync` method gets the attachment with the specified identifier from the item.</span></span> <span data-ttu-id="44017-1029">作为最佳做法，应使用标识符检索同一会话中的附件，在该会话中，使用 `getAttachmentsAsync` 或 `item.attachments` 调用检索附件 ID。</span><span class="sxs-lookup"><span data-stu-id="44017-1029">As a best practice, you should use the identifier to retrieve an attachment in the same session that the attachmentIds were retrieved with the `getAttachmentsAsync` or `item.attachments` call.</span></span> <span data-ttu-id="44017-1030">在 Outlook Web App 和适用于设备的 OWA 中，附件标识符只在同一个会话中才有效。</span><span class="sxs-lookup"><span data-stu-id="44017-1030">In Outlook Web App and OWA for Devices, the attachment identifier is valid only within the same session.</span></span> <span data-ttu-id="44017-1031">当用户关闭应用，或者如果用户开始在内嵌窗体中撰写，则随后弹出的窗体中继续在单独的窗口撰写时，会话即结束。</span><span class="sxs-lookup"><span data-stu-id="44017-1031">A session is over when the user closes the app, or if the user starts composing an inline form then subsequently pops out the form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="44017-1032">Parameters</span><span class="sxs-lookup"><span data-stu-id="44017-1032">Parameters</span></span>

|<span data-ttu-id="44017-1033">名称</span><span class="sxs-lookup"><span data-stu-id="44017-1033">Name</span></span>|<span data-ttu-id="44017-1034">类型</span><span class="sxs-lookup"><span data-stu-id="44017-1034">Type</span></span>|<span data-ttu-id="44017-1035">属性</span><span class="sxs-lookup"><span data-stu-id="44017-1035">Attributes</span></span>|<span data-ttu-id="44017-1036">说明</span><span class="sxs-lookup"><span data-stu-id="44017-1036">Description</span></span>|
|---|---|---|---|
|`attachmentId`|<span data-ttu-id="44017-1037">字符串</span><span class="sxs-lookup"><span data-stu-id="44017-1037">String</span></span>||<span data-ttu-id="44017-1038">要获取的附件的标识符。</span><span class="sxs-lookup"><span data-stu-id="44017-1038">The identifier of the attachment you want to get.</span></span>|
|`options`|<span data-ttu-id="44017-1039">对象</span><span class="sxs-lookup"><span data-stu-id="44017-1039">Object</span></span>|<span data-ttu-id="44017-1040">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="44017-1040">&lt;optional&gt;</span></span>|<span data-ttu-id="44017-1041">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="44017-1041">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="44017-1042">对象</span><span class="sxs-lookup"><span data-stu-id="44017-1042">Object</span></span>|<span data-ttu-id="44017-1043">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="44017-1043">&lt;optional&gt;</span></span>|<span data-ttu-id="44017-1044">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="44017-1044">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="44017-1045">函数</span><span class="sxs-lookup"><span data-stu-id="44017-1045">function</span></span>|<span data-ttu-id="44017-1046">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="44017-1046">&lt;optional&gt;</span></span>|<span data-ttu-id="44017-1047">方法完成后，使用单个参数 `asyncResult`（一个 [AsyncResult](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="44017-1047">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="44017-1048">要求</span><span class="sxs-lookup"><span data-stu-id="44017-1048">Requirements</span></span>

|<span data-ttu-id="44017-1049">要求</span><span class="sxs-lookup"><span data-stu-id="44017-1049">Requirement</span></span>|<span data-ttu-id="44017-1050">值</span><span class="sxs-lookup"><span data-stu-id="44017-1050">Value</span></span>|
|---|---|
|[<span data-ttu-id="44017-1051">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="44017-1051">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="44017-1052">预览</span><span class="sxs-lookup"><span data-stu-id="44017-1052">Preview</span></span>|
|[<span data-ttu-id="44017-1053">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="44017-1053">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="44017-1054">ReadItem</span><span class="sxs-lookup"><span data-stu-id="44017-1054">ReadItem</span></span>|
|[<span data-ttu-id="44017-1055">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="44017-1055">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="44017-1056">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="44017-1056">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="44017-1057">返回：</span><span class="sxs-lookup"><span data-stu-id="44017-1057">Returns:</span></span>

<span data-ttu-id="44017-1058">类型：[AttachmentContent](/javascript/api/outlook/office.attachmentcontent)</span><span class="sxs-lookup"><span data-stu-id="44017-1058">Type: [AttachmentContent](/javascript/api/outlook/office.attachmentcontent)</span></span>

##### <a name="example"></a><span data-ttu-id="44017-1059">示例</span><span class="sxs-lookup"><span data-stu-id="44017-1059">Example</span></span>

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

#### <a name="getattachmentsasyncoptions-callback--arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetails"></a><span data-ttu-id="44017-1060">getAttachmentsAsync ([options], [callback]) → Array. <[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="44017-1060">getAttachmentsAsync([options], [callback]) → Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span></span>

<span data-ttu-id="44017-1061">获取项目的附件作为数组。</span><span class="sxs-lookup"><span data-stu-id="44017-1061">Gets the item's attachments as an array.</span></span> <span data-ttu-id="44017-1062">仅限撰写模式。</span><span class="sxs-lookup"><span data-stu-id="44017-1062">Compose mode only.</span></span>

##### <a name="parameters"></a><span data-ttu-id="44017-1063">Parameters</span><span class="sxs-lookup"><span data-stu-id="44017-1063">Parameters</span></span>

|<span data-ttu-id="44017-1064">名称</span><span class="sxs-lookup"><span data-stu-id="44017-1064">Name</span></span>|<span data-ttu-id="44017-1065">类型</span><span class="sxs-lookup"><span data-stu-id="44017-1065">Type</span></span>|<span data-ttu-id="44017-1066">属性</span><span class="sxs-lookup"><span data-stu-id="44017-1066">Attributes</span></span>|<span data-ttu-id="44017-1067">说明</span><span class="sxs-lookup"><span data-stu-id="44017-1067">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="44017-1068">Object</span><span class="sxs-lookup"><span data-stu-id="44017-1068">Object</span></span>|<span data-ttu-id="44017-1069">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="44017-1069">&lt;optional&gt;</span></span>|<span data-ttu-id="44017-1070">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="44017-1070">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="44017-1071">对象</span><span class="sxs-lookup"><span data-stu-id="44017-1071">Object</span></span>|<span data-ttu-id="44017-1072">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="44017-1072">&lt;optional&gt;</span></span>|<span data-ttu-id="44017-1073">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="44017-1073">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="44017-1074">函数</span><span class="sxs-lookup"><span data-stu-id="44017-1074">function</span></span>|<span data-ttu-id="44017-1075">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="44017-1075">&lt;optional&gt;</span></span>|<span data-ttu-id="44017-1076">方法完成后，使用单个参数 `asyncResult`（一个 [AsyncResult](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="44017-1076">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="44017-1077">Requirements</span><span class="sxs-lookup"><span data-stu-id="44017-1077">Requirements</span></span>

|<span data-ttu-id="44017-1078">要求</span><span class="sxs-lookup"><span data-stu-id="44017-1078">Requirement</span></span>|<span data-ttu-id="44017-1079">值</span><span class="sxs-lookup"><span data-stu-id="44017-1079">Value</span></span>|
|---|---|
|[<span data-ttu-id="44017-1080">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="44017-1080">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="44017-1081">预览</span><span class="sxs-lookup"><span data-stu-id="44017-1081">Preview</span></span>|
|[<span data-ttu-id="44017-1082">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="44017-1082">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="44017-1083">ReadItem</span><span class="sxs-lookup"><span data-stu-id="44017-1083">ReadItem</span></span>|
|[<span data-ttu-id="44017-1084">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="44017-1084">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="44017-1085">撰写</span><span class="sxs-lookup"><span data-stu-id="44017-1085">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="44017-1086">返回：</span><span class="sxs-lookup"><span data-stu-id="44017-1086">Returns:</span></span>

<span data-ttu-id="44017-1087">类型：Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="44017-1087">Type: Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span></span>

##### <a name="example"></a><span data-ttu-id="44017-1088">示例</span><span class="sxs-lookup"><span data-stu-id="44017-1088">Example</span></span>

<span data-ttu-id="44017-1089">以下示例使用当前项目上所有附件的详细信息构成 HTML 字符串。</span><span class="sxs-lookup"><span data-stu-id="44017-1089">The following example builds an HTML string with details of all attachments on the current item.</span></span>

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

#### <a name="getentities--entitiesjavascriptapioutlookofficeentities"></a><span data-ttu-id="44017-1090">getEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="44017-1090">getEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span></span>

<span data-ttu-id="44017-1091">获取在所选项目的正文中找到的实体。</span><span class="sxs-lookup"><span data-stu-id="44017-1091">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="44017-1092">在 Outlook for iOS 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="44017-1092">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="44017-1093">要求</span><span class="sxs-lookup"><span data-stu-id="44017-1093">Requirements</span></span>

|<span data-ttu-id="44017-1094">要求</span><span class="sxs-lookup"><span data-stu-id="44017-1094">Requirement</span></span>|<span data-ttu-id="44017-1095">值</span><span class="sxs-lookup"><span data-stu-id="44017-1095">Value</span></span>|
|---|---|
|[<span data-ttu-id="44017-1096">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="44017-1096">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="44017-1097">1.0</span><span class="sxs-lookup"><span data-stu-id="44017-1097">1.0</span></span>|
|[<span data-ttu-id="44017-1098">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="44017-1098">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="44017-1099">ReadItem</span><span class="sxs-lookup"><span data-stu-id="44017-1099">ReadItem</span></span>|
|[<span data-ttu-id="44017-1100">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="44017-1100">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="44017-1101">阅读</span><span class="sxs-lookup"><span data-stu-id="44017-1101">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="44017-1102">返回：</span><span class="sxs-lookup"><span data-stu-id="44017-1102">Returns:</span></span>

<span data-ttu-id="44017-1103">类型：[Entities](/javascript/api/outlook/office.entities)</span><span class="sxs-lookup"><span data-stu-id="44017-1103">Type: [Entities](/javascript/api/outlook/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="44017-1104">示例</span><span class="sxs-lookup"><span data-stu-id="44017-1104">Example</span></span>

<span data-ttu-id="44017-1105">以下示例访问当前项目的正文中的联系人实体。</span><span class="sxs-lookup"><span data-stu-id="44017-1105">The following example accesses the contacts entities in the current item's body.</span></span>

```javascript
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlookofficecontactmeetingsuggestionjavascriptapioutlookofficemeetingsuggestionphonenumberjavascriptapioutlookofficephonenumbertasksuggestionjavascriptapioutlookofficetasksuggestion"></a><span data-ttu-id="44017-1106">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="44017-1106">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span></span>

<span data-ttu-id="44017-1107">获取所选项目的正文中找到的指定实体类型的所有实体的数组。</span><span class="sxs-lookup"><span data-stu-id="44017-1107">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="44017-1108">在 Outlook for iOS 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="44017-1108">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="44017-1109">Parameters</span><span class="sxs-lookup"><span data-stu-id="44017-1109">Parameters</span></span>

|<span data-ttu-id="44017-1110">名称</span><span class="sxs-lookup"><span data-stu-id="44017-1110">Name</span></span>|<span data-ttu-id="44017-1111">说明</span><span class="sxs-lookup"><span data-stu-id="44017-1111">Type</span></span>|<span data-ttu-id="44017-1112">说明</span><span class="sxs-lookup"><span data-stu-id="44017-1112">Description</span></span>|
|---|---|---|
|`entityType`|[<span data-ttu-id="44017-1113">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="44017-1113">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook/office.mailboxenums.entitytype)|<span data-ttu-id="44017-1114">EntityType 枚举值之一。</span><span class="sxs-lookup"><span data-stu-id="44017-1114">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="44017-1115">Requirements</span><span class="sxs-lookup"><span data-stu-id="44017-1115">Requirements</span></span>

|<span data-ttu-id="44017-1116">要求</span><span class="sxs-lookup"><span data-stu-id="44017-1116">Requirement</span></span>|<span data-ttu-id="44017-1117">值</span><span class="sxs-lookup"><span data-stu-id="44017-1117">Value</span></span>|
|---|---|
|[<span data-ttu-id="44017-1118">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="44017-1118">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="44017-1119">1.0</span><span class="sxs-lookup"><span data-stu-id="44017-1119">1.0</span></span>|
|[<span data-ttu-id="44017-1120">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="44017-1120">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="44017-1121">受限</span><span class="sxs-lookup"><span data-stu-id="44017-1121">Restricted</span></span>|
|[<span data-ttu-id="44017-1122">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="44017-1122">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="44017-1123">阅读</span><span class="sxs-lookup"><span data-stu-id="44017-1123">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="44017-1124">返回：</span><span class="sxs-lookup"><span data-stu-id="44017-1124">Returns:</span></span>

<span data-ttu-id="44017-1125">如果在 `entityType` 中传递的值不是 `EntityType` 枚举的有效成员，该方法返回 null。</span><span class="sxs-lookup"><span data-stu-id="44017-1125">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="44017-1126">如果指定类型的任何实体都不存在于该项目的正文中，该方法将返回空数组。</span><span class="sxs-lookup"><span data-stu-id="44017-1126">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="44017-1127">否则，返回的数组中对象的类型取决于 `entityType` 参数中请求实体的类型。</span><span class="sxs-lookup"><span data-stu-id="44017-1127">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="44017-1128">当使用此方法的最低权限级别**受限**时，某些实体类型需要 **ReadItem** 才能进行访问，如下表中所指定。</span><span class="sxs-lookup"><span data-stu-id="44017-1128">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

|<span data-ttu-id="44017-1129">`entityType` 的值</span><span class="sxs-lookup"><span data-stu-id="44017-1129">Value of `entityType`</span></span>|<span data-ttu-id="44017-1130">返回的数组中对象的类型</span><span class="sxs-lookup"><span data-stu-id="44017-1130">Type of objects in returned array</span></span>|<span data-ttu-id="44017-1131">所需权限级别</span><span class="sxs-lookup"><span data-stu-id="44017-1131">Required Permission Level</span></span>|
|---|---|---|
|`Address`|<span data-ttu-id="44017-1132">字符串</span><span class="sxs-lookup"><span data-stu-id="44017-1132">String</span></span>|<span data-ttu-id="44017-1133">**受限**</span><span class="sxs-lookup"><span data-stu-id="44017-1133">**Restricted**</span></span>|
|`Contact`|<span data-ttu-id="44017-1134">Contact</span><span class="sxs-lookup"><span data-stu-id="44017-1134">Contact</span></span>|<span data-ttu-id="44017-1135">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="44017-1135">**ReadItem**</span></span>|
|`EmailAddress`|<span data-ttu-id="44017-1136">String</span><span class="sxs-lookup"><span data-stu-id="44017-1136">String</span></span>|<span data-ttu-id="44017-1137">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="44017-1137">**ReadItem**</span></span>|
|`MeetingSuggestion`|<span data-ttu-id="44017-1138">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="44017-1138">MeetingSuggestion</span></span>|<span data-ttu-id="44017-1139">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="44017-1139">**ReadItem**</span></span>|
|`PhoneNumber`|<span data-ttu-id="44017-1140">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="44017-1140">PhoneNumber</span></span>|<span data-ttu-id="44017-1141">**受限**</span><span class="sxs-lookup"><span data-stu-id="44017-1141">**Restricted**</span></span>|
|`TaskSuggestion`|<span data-ttu-id="44017-1142">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="44017-1142">TaskSuggestion</span></span>|<span data-ttu-id="44017-1143">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="44017-1143">**ReadItem**</span></span>|
|`URL`|<span data-ttu-id="44017-1144">String</span><span class="sxs-lookup"><span data-stu-id="44017-1144">String</span></span>|<span data-ttu-id="44017-1145">**受限**</span><span class="sxs-lookup"><span data-stu-id="44017-1145">**Restricted**</span></span>|

<span data-ttu-id="44017-1146">类型：Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="44017-1146">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span></span>

##### <a name="example"></a><span data-ttu-id="44017-1147">示例</span><span class="sxs-lookup"><span data-stu-id="44017-1147">Example</span></span>

<span data-ttu-id="44017-1148">以下示例显示了如何访问表示当前项目的正文中的邮政地址的字符串数组。</span><span class="sxs-lookup"><span data-stu-id="44017-1148">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlookofficecontactmeetingsuggestionjavascriptapioutlookofficemeetingsuggestionphonenumberjavascriptapioutlookofficephonenumbertasksuggestionjavascriptapioutlookofficetasksuggestion"></a><span data-ttu-id="44017-1149">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="44017-1149">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span></span>

<span data-ttu-id="44017-1150">返回传递清单 XML 文件中定义的命名筛选器的所选项目中的已知实体。</span><span class="sxs-lookup"><span data-stu-id="44017-1150">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="44017-1151">在 Outlook for iOS 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="44017-1151">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="44017-1152">`getFilteredEntitiesByName` 方法返回匹配在具有指定 `FilterName` 元素值的清单 XML 文件中的 [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) 规则元素中定义的正则表达式的实体。</span><span class="sxs-lookup"><span data-stu-id="44017-1152">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="44017-1153">Parameters</span><span class="sxs-lookup"><span data-stu-id="44017-1153">Parameters</span></span>

|<span data-ttu-id="44017-1154">名称</span><span class="sxs-lookup"><span data-stu-id="44017-1154">Name</span></span>|<span data-ttu-id="44017-1155">说明</span><span class="sxs-lookup"><span data-stu-id="44017-1155">Type</span></span>|<span data-ttu-id="44017-1156">说明</span><span class="sxs-lookup"><span data-stu-id="44017-1156">Description</span></span>|
|---|---|---|
|`name`|<span data-ttu-id="44017-1157">字符串</span><span class="sxs-lookup"><span data-stu-id="44017-1157">String</span></span>|<span data-ttu-id="44017-1158">定义筛选器匹配的 `ItemHasKnownEntity` 规则元素的名称。</span><span class="sxs-lookup"><span data-stu-id="44017-1158">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="44017-1159">Requirements</span><span class="sxs-lookup"><span data-stu-id="44017-1159">Requirements</span></span>

|<span data-ttu-id="44017-1160">要求</span><span class="sxs-lookup"><span data-stu-id="44017-1160">Requirement</span></span>|<span data-ttu-id="44017-1161">值</span><span class="sxs-lookup"><span data-stu-id="44017-1161">Value</span></span>|
|---|---|
|[<span data-ttu-id="44017-1162">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="44017-1162">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="44017-1163">1.0</span><span class="sxs-lookup"><span data-stu-id="44017-1163">1.0</span></span>|
|[<span data-ttu-id="44017-1164">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="44017-1164">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="44017-1165">ReadItem</span><span class="sxs-lookup"><span data-stu-id="44017-1165">ReadItem</span></span>|
|[<span data-ttu-id="44017-1166">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="44017-1166">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="44017-1167">阅读</span><span class="sxs-lookup"><span data-stu-id="44017-1167">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="44017-1168">返回：</span><span class="sxs-lookup"><span data-stu-id="44017-1168">Returns:</span></span>

<span data-ttu-id="44017-p164">如果具有匹配 `name` 参数的 `FilterName` 元素值的清单中没有任何 `ItemHasKnownEntity` 元素，则该方法返回 `null`。如果 `name` 参数匹配清单中的 `ItemHasKnownEntity` 元素，但在匹配的当前项目中没有实体，则该方法返回一个空数组。</span><span class="sxs-lookup"><span data-stu-id="44017-p164">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="44017-1171">类型：Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="44017-1171">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span></span>

#### <a name="getinitializationcontextasyncoptions-callback"></a><span data-ttu-id="44017-1172">getInitializationContextAsync([options], [callback])</span><span class="sxs-lookup"><span data-stu-id="44017-1172">getInitializationContextAsync([options], [callback])</span></span>

<span data-ttu-id="44017-1173">当[可操作邮件激活](https://docs.microsoft.com/outlook/actionable-messages/invoke-add-in-from-actionable-message)加载项时，获取传递的初始化数据。</span><span class="sxs-lookup"><span data-stu-id="44017-1173">Gets initialization data passed when the add-in is [activated by an actionable message](https://docs.microsoft.com/outlook/actionable-messages/invoke-add-in-from-actionable-message).</span></span>

> [!NOTE]
> <span data-ttu-id="44017-1174">仅 Outlook 2016 for Windows 或更高版本（高于 16.0.8413.1000 的即点即用版本）和适用于 Office 365 的 Outlook 网页版支持此方法。</span><span class="sxs-lookup"><span data-stu-id="44017-1174">This method is only supported by Outlook 2016 or later for Windows (Click-to-Run versions later than 16.0.8413.1000) and Outlook on the web for Office 365.</span></span>

##### <a name="parameters"></a><span data-ttu-id="44017-1175">Parameters</span><span class="sxs-lookup"><span data-stu-id="44017-1175">Parameters</span></span>
|<span data-ttu-id="44017-1176">名称</span><span class="sxs-lookup"><span data-stu-id="44017-1176">Name</span></span>|<span data-ttu-id="44017-1177">类型</span><span class="sxs-lookup"><span data-stu-id="44017-1177">Type</span></span>|<span data-ttu-id="44017-1178">属性</span><span class="sxs-lookup"><span data-stu-id="44017-1178">Attributes</span></span>|<span data-ttu-id="44017-1179">说明</span><span class="sxs-lookup"><span data-stu-id="44017-1179">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="44017-1180">Object</span><span class="sxs-lookup"><span data-stu-id="44017-1180">Object</span></span>|<span data-ttu-id="44017-1181">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="44017-1181">&lt;optional&gt;</span></span>|<span data-ttu-id="44017-1182">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="44017-1182">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="44017-1183">对象</span><span class="sxs-lookup"><span data-stu-id="44017-1183">Object</span></span>|<span data-ttu-id="44017-1184">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="44017-1184">&lt;optional&gt;</span></span>|<span data-ttu-id="44017-1185">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="44017-1185">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="44017-1186">函数</span><span class="sxs-lookup"><span data-stu-id="44017-1186">function</span></span>|<span data-ttu-id="44017-1187">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="44017-1187">&lt;optional&gt;</span></span>|<span data-ttu-id="44017-1188">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="44017-1188">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="44017-1189">成功后，`asyncResult.value` 属性便以字符串形式提供初始化数据。</span><span class="sxs-lookup"><span data-stu-id="44017-1189">On success, the initialization data is provided in the `asyncResult.value` property as a string.</span></span><br/><span data-ttu-id="44017-1190">如果没有初始化上下文，`asyncResult` 对象包含 `Error` 对象，并将它的 `code` 和 `name` 属性分别设置为 `9020` 和 `GenericResponseError`。</span><span class="sxs-lookup"><span data-stu-id="44017-1190">If there is no initialization context, the `asyncResult` object will contain an `Error` object with its `code` property set to `9020` and its `name` property set to `GenericResponseError`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="44017-1191">要求</span><span class="sxs-lookup"><span data-stu-id="44017-1191">Requirements</span></span>

|<span data-ttu-id="44017-1192">要求</span><span class="sxs-lookup"><span data-stu-id="44017-1192">Requirement</span></span>|<span data-ttu-id="44017-1193">值</span><span class="sxs-lookup"><span data-stu-id="44017-1193">Value</span></span>|
|---|---|
|[<span data-ttu-id="44017-1194">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="44017-1194">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="44017-1195">预览</span><span class="sxs-lookup"><span data-stu-id="44017-1195">Preview</span></span>|
|[<span data-ttu-id="44017-1196">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="44017-1196">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="44017-1197">ReadItem</span><span class="sxs-lookup"><span data-stu-id="44017-1197">ReadItem</span></span>|
|[<span data-ttu-id="44017-1198">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="44017-1198">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="44017-1199">阅读</span><span class="sxs-lookup"><span data-stu-id="44017-1199">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="44017-1200">示例</span><span class="sxs-lookup"><span data-stu-id="44017-1200">Example</span></span>

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

#### <a name="getregexmatches--object"></a><span data-ttu-id="44017-1201">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="44017-1201">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="44017-1202">返回所选项目中匹配在清单 XML 文件中定义的正则表达式的字符串值。</span><span class="sxs-lookup"><span data-stu-id="44017-1202">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="44017-1203">在 Outlook for iOS 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="44017-1203">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="44017-p165">`getRegExMatches` 方法返回匹配在清单 XML 文件中的每个 `ItemHasRegularExpressionMatch` 或 `ItemHasKnownEntity` 规则元素中定义的正则表达式的字符串。对于 `ItemHasRegularExpressionMatch` 规则，匹配字符串必须发生在该规则指定的项目的属性中。`PropertyName` 简单类型定义支持的属性。</span><span class="sxs-lookup"><span data-stu-id="44017-p165">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="44017-1207">例如，考虑一个外接程序清单具有以下 `Rule` 元素：</span><span class="sxs-lookup"><span data-stu-id="44017-1207">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="44017-1208">从 `getRegExMatches` 返回的对象应有两个属性：`fruits` 和 `veggies`。</span><span class="sxs-lookup"><span data-stu-id="44017-1208">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="44017-p166">如果在项目的正文属性上指定 `ItemHasRegularExpressionMatch` 规则，则正则表达式应进一步筛选正文，不应尝试返回该项目的整个正文。使用正则表达式（如 `.*`）获取项目的整个正文并不总是返回预期的结果。而是使用 [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) 方法检索整个正文。</span><span class="sxs-lookup"><span data-stu-id="44017-p166">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="44017-1212">Requirements</span><span class="sxs-lookup"><span data-stu-id="44017-1212">Requirements</span></span>

|<span data-ttu-id="44017-1213">要求</span><span class="sxs-lookup"><span data-stu-id="44017-1213">Requirement</span></span>|<span data-ttu-id="44017-1214">值</span><span class="sxs-lookup"><span data-stu-id="44017-1214">Value</span></span>|
|---|---|
|[<span data-ttu-id="44017-1215">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="44017-1215">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="44017-1216">1.0</span><span class="sxs-lookup"><span data-stu-id="44017-1216">1.0</span></span>|
|[<span data-ttu-id="44017-1217">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="44017-1217">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="44017-1218">ReadItem</span><span class="sxs-lookup"><span data-stu-id="44017-1218">ReadItem</span></span>|
|[<span data-ttu-id="44017-1219">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="44017-1219">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="44017-1220">阅读</span><span class="sxs-lookup"><span data-stu-id="44017-1220">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="44017-1221">返回：</span><span class="sxs-lookup"><span data-stu-id="44017-1221">Returns:</span></span>

<span data-ttu-id="44017-p167">一个包含匹配在清单 XML 文件中定义的正则表达式的字符串数组的对象。每个数组的名称等于匹配 `ItemHasRegularExpressionMatch` 规则的 `RegExName` 属性或匹配 `ItemHasKnownEntity` 规则的 `FilterName` 属性的相应值。</span><span class="sxs-lookup"><span data-stu-id="44017-p167">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<dl class="param-type"><span data-ttu-id="44017-1224">

<dt>
类型</dt>


</span><span class="sxs-lookup"><span data-stu-id="44017-1224">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="44017-1225">对象</span><span class="sxs-lookup"><span data-stu-id="44017-1225">Object</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="44017-1226">示例</span><span class="sxs-lookup"><span data-stu-id="44017-1226">Example</span></span>

<span data-ttu-id="44017-1227">以下示例显示了如何访问正则表达式规则元素 `fruits` 和 `veggies` 的匹配项的数组，这些元素在清单中指定。</span><span class="sxs-lookup"><span data-stu-id="44017-1227">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```javascript
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veggies = allMatches.veggies;
```

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="44017-1228">getRegExMatchesByName(name) → (nullable) {Array.<String>}</span><span class="sxs-lookup"><span data-stu-id="44017-1228">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="44017-1229">返回匹配在清单 XML 文件中定义的命名正则表达式的所选项目中的字符串值。</span><span class="sxs-lookup"><span data-stu-id="44017-1229">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="44017-1230">在 Outlook for iOS 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="44017-1230">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="44017-1231">`getRegExMatchesByName` 方法返回匹配在具有指定 `RegExName` 元素值的清单 XML 文件中的 `ItemHasRegularExpressionMatch` 规则元素中定义的正则表达式的字符串。</span><span class="sxs-lookup"><span data-stu-id="44017-1231">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="44017-p168">如果在项目的正文属性上指定 `ItemHasRegularExpressionMatch` 规则，则正则表达式应进一步筛选正文，不应尝试返回该项目的整个正文。使用正则表达式（如 `.*`）获取项目的整个正文并不总是返回预期的结果。</span><span class="sxs-lookup"><span data-stu-id="44017-p168">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="44017-1234">Parameters</span><span class="sxs-lookup"><span data-stu-id="44017-1234">Parameters</span></span>

|<span data-ttu-id="44017-1235">名称</span><span class="sxs-lookup"><span data-stu-id="44017-1235">Name</span></span>|<span data-ttu-id="44017-1236">说明</span><span class="sxs-lookup"><span data-stu-id="44017-1236">Type</span></span>|<span data-ttu-id="44017-1237">说明</span><span class="sxs-lookup"><span data-stu-id="44017-1237">Description</span></span>|
|---|---|---|
|`name`|<span data-ttu-id="44017-1238">字符串</span><span class="sxs-lookup"><span data-stu-id="44017-1238">String</span></span>|<span data-ttu-id="44017-1239">定义筛选器匹配的 `ItemHasRegularExpressionMatch` 规则元素的名称。</span><span class="sxs-lookup"><span data-stu-id="44017-1239">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="44017-1240">Requirements</span><span class="sxs-lookup"><span data-stu-id="44017-1240">Requirements</span></span>

|<span data-ttu-id="44017-1241">要求</span><span class="sxs-lookup"><span data-stu-id="44017-1241">Requirement</span></span>|<span data-ttu-id="44017-1242">值</span><span class="sxs-lookup"><span data-stu-id="44017-1242">Value</span></span>|
|---|---|
|[<span data-ttu-id="44017-1243">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="44017-1243">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="44017-1244">1.0</span><span class="sxs-lookup"><span data-stu-id="44017-1244">1.0</span></span>|
|[<span data-ttu-id="44017-1245">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="44017-1245">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="44017-1246">ReadItem</span><span class="sxs-lookup"><span data-stu-id="44017-1246">ReadItem</span></span>|
|[<span data-ttu-id="44017-1247">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="44017-1247">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="44017-1248">阅读</span><span class="sxs-lookup"><span data-stu-id="44017-1248">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="44017-1249">返回：</span><span class="sxs-lookup"><span data-stu-id="44017-1249">Returns:</span></span>

<span data-ttu-id="44017-1250">一个包含匹配在清单 XML 文件中定义的正则表达式的字符串的数组。</span><span class="sxs-lookup"><span data-stu-id="44017-1250">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<dl class="param-type"><span data-ttu-id="44017-1251">

<dt>类型</dt>

</span><span class="sxs-lookup"><span data-stu-id="44017-1251">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="44017-1252">Array.< String ></span><span class="sxs-lookup"><span data-stu-id="44017-1252">Array.< String ></span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="44017-1253">示例</span><span class="sxs-lookup"><span data-stu-id="44017-1253">Example</span></span>

```javascript
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

####  <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="44017-1254">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="44017-1254">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="44017-1255">以异步方式返回邮件的主题或正文中选定的数据。</span><span class="sxs-lookup"><span data-stu-id="44017-1255">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="44017-p169">如果没有选定内容，但光标位于正文或主题中，此方法将会为所选数据返回 null。如果选定的是字段，而不是正文或主题，则此方法返回 `InvalidSelection` 错误。</span><span class="sxs-lookup"><span data-stu-id="44017-p169">If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="44017-1258">Parameters</span><span class="sxs-lookup"><span data-stu-id="44017-1258">Parameters</span></span>

|<span data-ttu-id="44017-1259">名称</span><span class="sxs-lookup"><span data-stu-id="44017-1259">Name</span></span>|<span data-ttu-id="44017-1260">类型</span><span class="sxs-lookup"><span data-stu-id="44017-1260">Type</span></span>|<span data-ttu-id="44017-1261">属性</span><span class="sxs-lookup"><span data-stu-id="44017-1261">Attributes</span></span>|<span data-ttu-id="44017-1262">说明</span><span class="sxs-lookup"><span data-stu-id="44017-1262">Description</span></span>|
|---|---|---|---|
|`coercionType`|[<span data-ttu-id="44017-1263">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="44017-1263">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="44017-p170">请求数据的格式。如果为文本，则此方法返回纯文本作为字符串，删除任何显示的 HTML 标记。如果为 HTML，则此方法返回所选文本，不论是纯文本还是 HTML。</span><span class="sxs-lookup"><span data-stu-id="44017-p170">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`|<span data-ttu-id="44017-1267">对象</span><span class="sxs-lookup"><span data-stu-id="44017-1267">Object</span></span>|<span data-ttu-id="44017-1268">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="44017-1268">&lt;optional&gt;</span></span>|<span data-ttu-id="44017-1269">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="44017-1269">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="44017-1270">对象</span><span class="sxs-lookup"><span data-stu-id="44017-1270">Object</span></span>|<span data-ttu-id="44017-1271">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="44017-1271">&lt;optional&gt;</span></span>|<span data-ttu-id="44017-1272">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="44017-1272">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="44017-1273">函数</span><span class="sxs-lookup"><span data-stu-id="44017-1273">function</span></span>||<span data-ttu-id="44017-1274">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="44017-1274">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="44017-1275">若要从回调方法访问所选数据，请调用 `asyncResult.value.data`。</span><span class="sxs-lookup"><span data-stu-id="44017-1275">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="44017-1276">若要访问选定内容的源属性，请调用 `asyncResult.value.sourceProperty`，这将为 `body` 或 `subject`。</span><span class="sxs-lookup"><span data-stu-id="44017-1276">To access the source property that the selection comes from, call `asyncResult.value.sourceProperty`, which will be either `body` or `subject`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="44017-1277">Requirements</span><span class="sxs-lookup"><span data-stu-id="44017-1277">Requirements</span></span>

|<span data-ttu-id="44017-1278">要求</span><span class="sxs-lookup"><span data-stu-id="44017-1278">Requirement</span></span>|<span data-ttu-id="44017-1279">值</span><span class="sxs-lookup"><span data-stu-id="44017-1279">Value</span></span>|
|---|---|
|[<span data-ttu-id="44017-1280">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="44017-1280">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="44017-1281">1.2</span><span class="sxs-lookup"><span data-stu-id="44017-1281">1.2</span></span>|
|[<span data-ttu-id="44017-1282">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="44017-1282">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="44017-1283">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="44017-1283">ReadWriteItem</span></span>|
|[<span data-ttu-id="44017-1284">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="44017-1284">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="44017-1285">撰写</span><span class="sxs-lookup"><span data-stu-id="44017-1285">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="44017-1286">返回：</span><span class="sxs-lookup"><span data-stu-id="44017-1286">Returns:</span></span>

<span data-ttu-id="44017-1287">作为字符串的所选数据的格式由 `coercionType` 确定。</span><span class="sxs-lookup"><span data-stu-id="44017-1287">The selected data as a string with format determined by `coercionType`.</span></span>

<dl class="param-type"><span data-ttu-id="44017-1288">

<dt>
类型</dt>


</span><span class="sxs-lookup"><span data-stu-id="44017-1288">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="44017-1289">字符串</span><span class="sxs-lookup"><span data-stu-id="44017-1289">String</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="44017-1290">示例</span><span class="sxs-lookup"><span data-stu-id="44017-1290">Example</span></span>

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

#### <a name="getselectedentities--entitiesjavascriptapioutlookofficeentities"></a><span data-ttu-id="44017-1291">getSelectedEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="44017-1291">getSelectedEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span></span>

<span data-ttu-id="44017-p172">获取在用户已选择的突出显示匹配项中找到的实体。突出显示匹配项适用于[上下文加载项](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins)。</span><span class="sxs-lookup"><span data-stu-id="44017-p172">Gets the entities found in a highlighted match a user has selected. Highlighted matches apply to [contextual add-ins](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="44017-1294">在 Outlook for iOS 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="44017-1294">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="44017-1295">Requirements</span><span class="sxs-lookup"><span data-stu-id="44017-1295">Requirements</span></span>

|<span data-ttu-id="44017-1296">要求</span><span class="sxs-lookup"><span data-stu-id="44017-1296">Requirement</span></span>|<span data-ttu-id="44017-1297">值</span><span class="sxs-lookup"><span data-stu-id="44017-1297">Value</span></span>|
|---|---|
|[<span data-ttu-id="44017-1298">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="44017-1298">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="44017-1299">1.6</span><span class="sxs-lookup"><span data-stu-id="44017-1299">1.6</span></span>|
|[<span data-ttu-id="44017-1300">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="44017-1300">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="44017-1301">ReadItem</span><span class="sxs-lookup"><span data-stu-id="44017-1301">ReadItem</span></span>|
|[<span data-ttu-id="44017-1302">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="44017-1302">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="44017-1303">阅读</span><span class="sxs-lookup"><span data-stu-id="44017-1303">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="44017-1304">返回：</span><span class="sxs-lookup"><span data-stu-id="44017-1304">Returns:</span></span>

<span data-ttu-id="44017-1305">类型：[Entities](/javascript/api/outlook/office.entities)</span><span class="sxs-lookup"><span data-stu-id="44017-1305">Type: [Entities](/javascript/api/outlook/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="44017-1306">示例</span><span class="sxs-lookup"><span data-stu-id="44017-1306">Example</span></span>

<span data-ttu-id="44017-1307">以下示例访问用户选择的突出显示匹配项中的地址实体。</span><span class="sxs-lookup"><span data-stu-id="44017-1307">The following example accesses the addresses entities in the highlighted match selected by the user.</span></span>

```javascript
var contacts = Office.context.mailbox.item.getSelectedEntities().addresses;
```

#### <a name="getselectedregexmatches--object"></a><span data-ttu-id="44017-1308">getSelectedRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="44017-1308">getSelectedRegExMatches() → {Object}</span></span>

<span data-ttu-id="44017-p173">返回突出显示匹配项中匹配在清单 XML 文件中定义的正则表达式的字符串值。突出显示匹配项适用于[上下文外接程序](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins)。</span><span class="sxs-lookup"><span data-stu-id="44017-p173">Returns string values in a highlighted match that match the regular expressions defined in the manifest XML file. Highlighted matches apply to [contextual add-ins](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="44017-1311">在 Outlook for iOS 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="44017-1311">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="44017-p174">`getSelectedRegExMatches` 方法返回匹配在清单 XML 文件中的每个 `ItemHasRegularExpressionMatch` 或 `ItemHasKnownEntity` 规则元素中定义的正则表达式的字符串。对于 `ItemHasRegularExpressionMatch` 规则，匹配字符串必须发生在该规则指定的项目的属性中。`PropertyName` 简单类型定义支持的属性。</span><span class="sxs-lookup"><span data-stu-id="44017-p174">The `getSelectedRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="44017-1315">例如，考虑一个外接程序清单具有以下 `Rule` 元素：</span><span class="sxs-lookup"><span data-stu-id="44017-1315">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="44017-1316">从 `getRegExMatches` 返回的对象应有两个属性：`fruits` 和 `veggies`。</span><span class="sxs-lookup"><span data-stu-id="44017-1316">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="44017-p175">如果在项目的正文属性上指定 `ItemHasRegularExpressionMatch` 规则，则正则表达式应进一步筛选正文，不应尝试返回该项目的整个正文。使用正则表达式（如 `.*`）获取项目的整个正文并不总是返回预期的结果。而是使用 [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) 方法检索整个正文。</span><span class="sxs-lookup"><span data-stu-id="44017-p175">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="44017-1320">Requirements</span><span class="sxs-lookup"><span data-stu-id="44017-1320">Requirements</span></span>

|<span data-ttu-id="44017-1321">要求</span><span class="sxs-lookup"><span data-stu-id="44017-1321">Requirement</span></span>|<span data-ttu-id="44017-1322">值</span><span class="sxs-lookup"><span data-stu-id="44017-1322">Value</span></span>|
|---|---|
|[<span data-ttu-id="44017-1323">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="44017-1323">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="44017-1324">1.6</span><span class="sxs-lookup"><span data-stu-id="44017-1324">1.6</span></span>|
|[<span data-ttu-id="44017-1325">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="44017-1325">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="44017-1326">ReadItem</span><span class="sxs-lookup"><span data-stu-id="44017-1326">ReadItem</span></span>|
|[<span data-ttu-id="44017-1327">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="44017-1327">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="44017-1328">阅读</span><span class="sxs-lookup"><span data-stu-id="44017-1328">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="44017-1329">返回：</span><span class="sxs-lookup"><span data-stu-id="44017-1329">Returns:</span></span>

<span data-ttu-id="44017-p176">一个包含匹配在清单 XML 文件中定义的正则表达式的字符串数组的对象。每个数组的名称等于匹配 `ItemHasRegularExpressionMatch` 规则的 `RegExName` 属性或匹配 `ItemHasKnownEntity` 规则的 `FilterName` 属性的相应值。</span><span class="sxs-lookup"><span data-stu-id="44017-p176">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

##### <a name="example"></a><span data-ttu-id="44017-1332">示例</span><span class="sxs-lookup"><span data-stu-id="44017-1332">Example</span></span>

<span data-ttu-id="44017-1333">以下示例显示了如何访问正则表达式规则元素 `fruits` 和 `veggies` 的匹配项的数组，这些元素在清单中指定。</span><span class="sxs-lookup"><span data-stu-id="44017-1333">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```javascript
var selectedMatches = Office.context.mailbox.item.getSelectedRegExMatches();
var fruits = selectedMatches.fruits;
var veggies = selectedMatches.veggies;
```

#### <a name="getsharedpropertiesasyncoptions-callback"></a><span data-ttu-id="44017-1334">getSharedPropertiesAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="44017-1334">getSharedPropertiesAsync([options], callback)</span></span>

<span data-ttu-id="44017-1335">获取共享文件夹、日历或邮箱中所选约会或邮件的属性。</span><span class="sxs-lookup"><span data-stu-id="44017-1335">Gets the properties of the selected appointment or message in a shared folder, calendar, or mailbox.</span></span>

##### <a name="parameters"></a><span data-ttu-id="44017-1336">Parameters</span><span class="sxs-lookup"><span data-stu-id="44017-1336">Parameters</span></span>

|<span data-ttu-id="44017-1337">名称</span><span class="sxs-lookup"><span data-stu-id="44017-1337">Name</span></span>|<span data-ttu-id="44017-1338">类型</span><span class="sxs-lookup"><span data-stu-id="44017-1338">Type</span></span>|<span data-ttu-id="44017-1339">属性</span><span class="sxs-lookup"><span data-stu-id="44017-1339">Attributes</span></span>|<span data-ttu-id="44017-1340">说明</span><span class="sxs-lookup"><span data-stu-id="44017-1340">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="44017-1341">Object</span><span class="sxs-lookup"><span data-stu-id="44017-1341">Object</span></span>|<span data-ttu-id="44017-1342">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="44017-1342">&lt;optional&gt;</span></span>|<span data-ttu-id="44017-1343">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="44017-1343">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="44017-1344">对象</span><span class="sxs-lookup"><span data-stu-id="44017-1344">Object</span></span>|<span data-ttu-id="44017-1345">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="44017-1345">&lt;optional&gt;</span></span>|<span data-ttu-id="44017-1346">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="44017-1346">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="44017-1347">函数</span><span class="sxs-lookup"><span data-stu-id="44017-1347">function</span></span>||<span data-ttu-id="44017-1348">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="44017-1348">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="44017-1349">共享属性作为 `asyncResult.value` 属性中的 [`SharedProperties`](/javascript/api/outlook/office.sharedproperties) 对象提供。</span><span class="sxs-lookup"><span data-stu-id="44017-1349">The shared properties are provided as a [`SharedProperties`](/javascript/api/outlook/office.sharedproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="44017-1350">此对象可用于获取项目的共享属性。</span><span class="sxs-lookup"><span data-stu-id="44017-1350">This object can be used to get the item's shared properties.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="44017-1351">要求</span><span class="sxs-lookup"><span data-stu-id="44017-1351">Requirements</span></span>

|<span data-ttu-id="44017-1352">要求</span><span class="sxs-lookup"><span data-stu-id="44017-1352">Requirement</span></span>|<span data-ttu-id="44017-1353">值</span><span class="sxs-lookup"><span data-stu-id="44017-1353">Value</span></span>|
|---|---|
|[<span data-ttu-id="44017-1354">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="44017-1354">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="44017-1355">预览</span><span class="sxs-lookup"><span data-stu-id="44017-1355">Preview</span></span>|
|[<span data-ttu-id="44017-1356">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="44017-1356">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="44017-1357">ReadItem</span><span class="sxs-lookup"><span data-stu-id="44017-1357">ReadItem</span></span>|
|[<span data-ttu-id="44017-1358">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="44017-1358">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="44017-1359">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="44017-1359">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="44017-1360">示例</span><span class="sxs-lookup"><span data-stu-id="44017-1360">Example</span></span>

```javascript
Office.context.mailbox.item.getSharedPropertiesAsync(callback);

function callback (asyncResult) {
  var context = asyncResult.context;
  var sharedProperties = asyncResult.value;
}
```

####  <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="44017-1361">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="44017-1361">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="44017-1362">异步加载所选项目上此外接程序的自定义属性。</span><span class="sxs-lookup"><span data-stu-id="44017-1362">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="44017-p178">自定义属性基于每个应用、每个项目存储为键/值对。此方法在回调中返回 `CustomProperties` 对象，该回调提供访问特定于当前项目和当前外接程序的自定义属性的方法。自定义属性未在项目上加密，因此这不应用作安全存储。</span><span class="sxs-lookup"><span data-stu-id="44017-p178">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="44017-1366">Parameters</span><span class="sxs-lookup"><span data-stu-id="44017-1366">Parameters</span></span>

|<span data-ttu-id="44017-1367">名称</span><span class="sxs-lookup"><span data-stu-id="44017-1367">Name</span></span>|<span data-ttu-id="44017-1368">类型</span><span class="sxs-lookup"><span data-stu-id="44017-1368">Type</span></span>|<span data-ttu-id="44017-1369">属性</span><span class="sxs-lookup"><span data-stu-id="44017-1369">Attributes</span></span>|<span data-ttu-id="44017-1370">说明</span><span class="sxs-lookup"><span data-stu-id="44017-1370">Description</span></span>|
|---|---|---|---|
|`callback`|<span data-ttu-id="44017-1371">函数</span><span class="sxs-lookup"><span data-stu-id="44017-1371">function</span></span>||<span data-ttu-id="44017-1372">方法完成后，使用单个参数 `callback`（一个 [`asyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `AsyncResult` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="44017-1372">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="44017-1373">自定义属性作为 `asyncResult.value` 属性中的 [`CustomProperties`](/javascript/api/outlook/office.customproperties) 对象提供。</span><span class="sxs-lookup"><span data-stu-id="44017-1373">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook/office.customproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="44017-1374">此对象可用于获取、设置以及从项目中删除自定义属性，并将自定义属性集的更改重新保存到服务器。</span><span class="sxs-lookup"><span data-stu-id="44017-1374">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`|<span data-ttu-id="44017-1375">对象</span><span class="sxs-lookup"><span data-stu-id="44017-1375">Object</span></span>|<span data-ttu-id="44017-1376">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="44017-1376">&lt;optional&gt;</span></span>|<span data-ttu-id="44017-1377">开发人员可以提供他们想要在回调函数中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="44017-1377">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="44017-1378">此对象可以通过回调函数中的 `asyncResult.asyncContext` 属性进行访问。</span><span class="sxs-lookup"><span data-stu-id="44017-1378">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="44017-1379">要求</span><span class="sxs-lookup"><span data-stu-id="44017-1379">Requirements</span></span>

|<span data-ttu-id="44017-1380">要求</span><span class="sxs-lookup"><span data-stu-id="44017-1380">Requirement</span></span>|<span data-ttu-id="44017-1381">值</span><span class="sxs-lookup"><span data-stu-id="44017-1381">Value</span></span>|
|---|---|
|[<span data-ttu-id="44017-1382">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="44017-1382">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="44017-1383">1.0</span><span class="sxs-lookup"><span data-stu-id="44017-1383">1.0</span></span>|
|[<span data-ttu-id="44017-1384">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="44017-1384">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="44017-1385">ReadItem</span><span class="sxs-lookup"><span data-stu-id="44017-1385">ReadItem</span></span>|
|[<span data-ttu-id="44017-1386">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="44017-1386">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="44017-1387">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="44017-1387">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="44017-1388">示例</span><span class="sxs-lookup"><span data-stu-id="44017-1388">Example</span></span>

<span data-ttu-id="44017-p181">以下代码示例显示了如何使用 `loadCustomPropertiesAsync` 方法异步加载特定于当前项目的自定义属性。该示例还显示了如何使用 `CustomProperties.saveAsync` 方法将这些属性重新保存到服务器。加载自定义属性后，该代码示例将使用 `CustomProperties.get` 方法读取自定义属性 `myProp`，使用 `CustomProperties.set` 方法写入自定义属性 `otherProp`，最后调用 `saveAsync` 方法保存这些自定义属性。</span><span class="sxs-lookup"><span data-stu-id="44017-p181">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

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

####  <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="44017-1392">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="44017-1392">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="44017-1393">将附件从邮件或约会中删除。</span><span class="sxs-lookup"><span data-stu-id="44017-1393">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="44017-1394">`removeAttachmentAsync` 方法删除项目中带指定标识符的附件。</span><span class="sxs-lookup"><span data-stu-id="44017-1394">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item.</span></span> <span data-ttu-id="44017-1395">最佳做法是，仅当同一个邮件应用程序在同一会话中添加了一个附件时，你才应使用该附件标识符来删除该附件。</span><span class="sxs-lookup"><span data-stu-id="44017-1395">As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session.</span></span> <span data-ttu-id="44017-1396">在 Outlook Web App 和适用于设备的 OWA 中，附件标识符只在同一个会话中才有效。</span><span class="sxs-lookup"><span data-stu-id="44017-1396">In Outlook Web App and OWA for Devices, the attachment identifier is valid only within the same session.</span></span> <span data-ttu-id="44017-1397">当用户关闭应用，或者如果用户开始在内嵌窗体中撰写，则随后弹出的窗体中继续在单独的窗口撰写时，会话即结束。</span><span class="sxs-lookup"><span data-stu-id="44017-1397">A session is over when the user closes the app, or if the user starts composing an inline form then subsequently pops out the form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="44017-1398">Parameters</span><span class="sxs-lookup"><span data-stu-id="44017-1398">Parameters</span></span>

|<span data-ttu-id="44017-1399">名称</span><span class="sxs-lookup"><span data-stu-id="44017-1399">Name</span></span>|<span data-ttu-id="44017-1400">类型</span><span class="sxs-lookup"><span data-stu-id="44017-1400">Type</span></span>|<span data-ttu-id="44017-1401">属性</span><span class="sxs-lookup"><span data-stu-id="44017-1401">Attributes</span></span>|<span data-ttu-id="44017-1402">说明</span><span class="sxs-lookup"><span data-stu-id="44017-1402">Description</span></span>|
|---|---|---|---|
|`attachmentId`|<span data-ttu-id="44017-1403">字符串</span><span class="sxs-lookup"><span data-stu-id="44017-1403">String</span></span>||<span data-ttu-id="44017-1404">要删除的附件的标识符。</span><span class="sxs-lookup"><span data-stu-id="44017-1404">The identifier of the attachment to remove.</span></span>|
|`options`|<span data-ttu-id="44017-1405">对象</span><span class="sxs-lookup"><span data-stu-id="44017-1405">Object</span></span>|<span data-ttu-id="44017-1406">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="44017-1406">&lt;optional&gt;</span></span>|<span data-ttu-id="44017-1407">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="44017-1407">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="44017-1408">对象</span><span class="sxs-lookup"><span data-stu-id="44017-1408">Object</span></span>|<span data-ttu-id="44017-1409">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="44017-1409">&lt;optional&gt;</span></span>|<span data-ttu-id="44017-1410">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="44017-1410">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="44017-1411">函数</span><span class="sxs-lookup"><span data-stu-id="44017-1411">function</span></span>|<span data-ttu-id="44017-1412">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="44017-1412">&lt;optional&gt;</span></span>|<span data-ttu-id="44017-1413">方法完成后，使用单个参数 `callback`（一个 [`asyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `AsyncResult` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="44017-1413">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="44017-1414">如果删除附件失败，`asyncResult.error` 属性将包含一个说明失败原因的错误代码。</span><span class="sxs-lookup"><span data-stu-id="44017-1414">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="44017-1415">错误</span><span class="sxs-lookup"><span data-stu-id="44017-1415">Errors</span></span>

|<span data-ttu-id="44017-1416">错误代码</span><span class="sxs-lookup"><span data-stu-id="44017-1416">Error code</span></span>|<span data-ttu-id="44017-1417">说明</span><span class="sxs-lookup"><span data-stu-id="44017-1417">Description</span></span>|
|------------|-------------|
|`InvalidAttachmentId`|<span data-ttu-id="44017-1418">附件标识符不存在。</span><span class="sxs-lookup"><span data-stu-id="44017-1418">The attachment identifier does not exist.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="44017-1419">Requirements</span><span class="sxs-lookup"><span data-stu-id="44017-1419">Requirements</span></span>

|<span data-ttu-id="44017-1420">要求</span><span class="sxs-lookup"><span data-stu-id="44017-1420">Requirement</span></span>|<span data-ttu-id="44017-1421">值</span><span class="sxs-lookup"><span data-stu-id="44017-1421">Value</span></span>|
|---|---|
|[<span data-ttu-id="44017-1422">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="44017-1422">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="44017-1423">1.1</span><span class="sxs-lookup"><span data-stu-id="44017-1423">1.1</span></span>|
|[<span data-ttu-id="44017-1424">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="44017-1424">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="44017-1425">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="44017-1425">ReadWriteItem</span></span>|
|[<span data-ttu-id="44017-1426">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="44017-1426">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="44017-1427">撰写</span><span class="sxs-lookup"><span data-stu-id="44017-1427">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="44017-1428">示例</span><span class="sxs-lookup"><span data-stu-id="44017-1428">Example</span></span>

<span data-ttu-id="44017-1429">以下代码删除包含标识符 '0' 的附件。</span><span class="sxs-lookup"><span data-stu-id="44017-1429">The following code removes an attachment with an identifier of '0'.</span></span>

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

####  <a name="removehandlerasynceventtype-options-callback"></a><span data-ttu-id="44017-1430">removeHandlerAsync(eventType, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="44017-1430">removeHandlerAsync(eventType, [options], [callback])</span></span>

<span data-ttu-id="44017-1431">删除受支持事件类型的事件处理程序。</span><span class="sxs-lookup"><span data-stu-id="44017-1431">Removes the event handlers for a supported event type.</span></span>

<span data-ttu-id="44017-1432">当前，支持的事件类型是 `Office.EventType.AttachmentsChanged`、`Office.EventType.AppointmentTimeChanged`、`Office.EventType.RecipientsChanged` 和 `Office.EventType.RecurrenceChanged`</span><span class="sxs-lookup"><span data-stu-id="44017-1432">Currently the supported event types are `Office.EventType.AttachmentsChanged`, `Office.EventType.AppointmentTimeChanged`, `Office.EventType.RecipientsChanged`, and `Office.EventType.RecurrenceChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="44017-1433">Parameters</span><span class="sxs-lookup"><span data-stu-id="44017-1433">Parameters</span></span>

| <span data-ttu-id="44017-1434">名称</span><span class="sxs-lookup"><span data-stu-id="44017-1434">Name</span></span> | <span data-ttu-id="44017-1435">类型</span><span class="sxs-lookup"><span data-stu-id="44017-1435">Type</span></span> | <span data-ttu-id="44017-1436">属性</span><span class="sxs-lookup"><span data-stu-id="44017-1436">Attributes</span></span> | <span data-ttu-id="44017-1437">说明</span><span class="sxs-lookup"><span data-stu-id="44017-1437">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="44017-1438">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="44017-1438">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="44017-1439">应撤销处理程序的事件。</span><span class="sxs-lookup"><span data-stu-id="44017-1439">The event that should revoke the handler.</span></span> |
| `options` | <span data-ttu-id="44017-1440">对象</span><span class="sxs-lookup"><span data-stu-id="44017-1440">Object</span></span> | <span data-ttu-id="44017-1441">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="44017-1441">&lt;optional&gt;</span></span> | <span data-ttu-id="44017-1442">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="44017-1442">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="44017-1443">对象</span><span class="sxs-lookup"><span data-stu-id="44017-1443">Object</span></span> | <span data-ttu-id="44017-1444">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="44017-1444">&lt;optional&gt;</span></span> | <span data-ttu-id="44017-1445">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="44017-1445">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="44017-1446">函数</span><span class="sxs-lookup"><span data-stu-id="44017-1446">function</span></span>| <span data-ttu-id="44017-1447">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="44017-1447">&lt;optional&gt;</span></span>|<span data-ttu-id="44017-1448">方法完成后，使用单个参数 `callback`（一个 [`asyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `AsyncResult` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="44017-1448">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="44017-1449">Requirements</span><span class="sxs-lookup"><span data-stu-id="44017-1449">Requirements</span></span>

|<span data-ttu-id="44017-1450">要求</span><span class="sxs-lookup"><span data-stu-id="44017-1450">Requirement</span></span>| <span data-ttu-id="44017-1451">值</span><span class="sxs-lookup"><span data-stu-id="44017-1451">Value</span></span>|
|---|---|
|[<span data-ttu-id="44017-1452">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="44017-1452">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="44017-1453">1.7</span><span class="sxs-lookup"><span data-stu-id="44017-1453">1.7</span></span> |
|[<span data-ttu-id="44017-1454">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="44017-1454">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="44017-1455">ReadItem</span><span class="sxs-lookup"><span data-stu-id="44017-1455">ReadItem</span></span> |
|[<span data-ttu-id="44017-1456">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="44017-1456">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="44017-1457">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="44017-1457">Compose or Read</span></span> |

####  <a name="saveasyncoptions-callback"></a><span data-ttu-id="44017-1458">saveAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="44017-1458">saveAsync([options], callback)</span></span>

<span data-ttu-id="44017-1459">异步保存项目。</span><span class="sxs-lookup"><span data-stu-id="44017-1459">Asynchronously saves an item.</span></span>

<span data-ttu-id="44017-p183">调用时，此方法将当前邮件保存为草稿，并通过回调方法返回项目 ID。在 Outlook Web App 或 Outlook 联机模式下，该项目被保存到服务器中。在 Outlook 缓存模式下，该项目被保存到本地缓存中。</span><span class="sxs-lookup"><span data-stu-id="44017-p183">When invoked, this method saves the current message as a draft and returns the item id via the callback method. In Outlook Web App or Outlook in online mode, the item is saved to the server. In Outlook in cached mode, the item is saved to the local cache.</span></span>

> [!NOTE]
> <span data-ttu-id="44017-1463">如果加载项在撰写模式下对某个项目调用 `saveAsync` 来获得 `itemId`，以便与 EWS 或 REST API 一同使用，请注意，当 Outlook 处于高速缓存模式时，可能需要一段时间项目才能真正同步到服务器。</span><span class="sxs-lookup"><span data-stu-id="44017-1463">If your add-in calls `saveAsync` on an item in compose mode in order to get an `itemId` to use with EWS or the REST API, be aware that when Outlook is in cached mode, it may take some time before the item is actually synced to the server.</span></span> <span data-ttu-id="44017-1464">在项目同步前，使用 `itemId` 将返回一个错误。</span><span class="sxs-lookup"><span data-stu-id="44017-1464">Until the item is synced, using the `itemId` will return an error.</span></span>

<span data-ttu-id="44017-p185">由于约会没有草稿状态，如果以撰写模式在约会中调用 `saveAsync`，则该项将被保存为用户日历中的正常约会。对于之前未保存过的新约会，则不会发送邀请。保存现有约会将向添加或删除的与会者发送更新。</span><span class="sxs-lookup"><span data-stu-id="44017-p185">Since appointments have no draft state, if `saveAsync` is called on an appointment in compose mode, the item will be saved as a normal appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent. Saving an existing appointment will send an update to added or removed attendees.</span></span>

> [!NOTE]
> <span data-ttu-id="44017-1468">以下客户端在撰写模式下对约会上的 `saveAsync` 具有不同的行为：</span><span class="sxs-lookup"><span data-stu-id="44017-1468">The following clients have different behavior for `saveAsync` on appointments in compose mode:</span></span>
>
> - <span data-ttu-id="44017-1469">Mac Outlook 不支持在撰写模式下对会议执行 `saveAsync` 操作。</span><span class="sxs-lookup"><span data-stu-id="44017-1469">Mac Outlook does not support `saveAsync` on a meeting in compose mode.</span></span> <span data-ttu-id="44017-1470">对 Mac Outlook 中的会议调用 `saveAsync` 将会返回一个错误。</span><span class="sxs-lookup"><span data-stu-id="44017-1470">Calling `saveAsync` on a meeting in Mac Outlook will return an error.</span></span>
> - <span data-ttu-id="44017-1471">在撰写模式下的约会上调用 `saveAsync` 时，Outlook 网页版始终发送邀请或更新。</span><span class="sxs-lookup"><span data-stu-id="44017-1471">Outlook on the web always sends an invitation or update when `saveAsync` is called on an appointment in compose mode.</span></span>

##### <a name="parameters"></a><span data-ttu-id="44017-1472">Parameters</span><span class="sxs-lookup"><span data-stu-id="44017-1472">Parameters</span></span>

|<span data-ttu-id="44017-1473">名称</span><span class="sxs-lookup"><span data-stu-id="44017-1473">Name</span></span>|<span data-ttu-id="44017-1474">类型</span><span class="sxs-lookup"><span data-stu-id="44017-1474">Type</span></span>|<span data-ttu-id="44017-1475">属性</span><span class="sxs-lookup"><span data-stu-id="44017-1475">Attributes</span></span>|<span data-ttu-id="44017-1476">说明</span><span class="sxs-lookup"><span data-stu-id="44017-1476">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="44017-1477">Object</span><span class="sxs-lookup"><span data-stu-id="44017-1477">Object</span></span>|<span data-ttu-id="44017-1478">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="44017-1478">&lt;optional&gt;</span></span>|<span data-ttu-id="44017-1479">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="44017-1479">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="44017-1480">对象</span><span class="sxs-lookup"><span data-stu-id="44017-1480">Object</span></span>|<span data-ttu-id="44017-1481">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="44017-1481">&lt;optional&gt;</span></span>|<span data-ttu-id="44017-1482">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="44017-1482">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="44017-1483">函数</span><span class="sxs-lookup"><span data-stu-id="44017-1483">function</span></span>||<span data-ttu-id="44017-1484">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="44017-1484">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="44017-1485">如果成功，该项目标识符将在 `asyncResult.value` 属性中提供。</span><span class="sxs-lookup"><span data-stu-id="44017-1485">On success, the item identifier is provided in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="44017-1486">要求</span><span class="sxs-lookup"><span data-stu-id="44017-1486">Requirements</span></span>

|<span data-ttu-id="44017-1487">要求</span><span class="sxs-lookup"><span data-stu-id="44017-1487">Requirement</span></span>|<span data-ttu-id="44017-1488">值</span><span class="sxs-lookup"><span data-stu-id="44017-1488">Value</span></span>|
|---|---|
|[<span data-ttu-id="44017-1489">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="44017-1489">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="44017-1490">1.3</span><span class="sxs-lookup"><span data-stu-id="44017-1490">1.3</span></span>|
|[<span data-ttu-id="44017-1491">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="44017-1491">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="44017-1492">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="44017-1492">ReadWriteItem</span></span>|
|[<span data-ttu-id="44017-1493">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="44017-1493">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="44017-1494">撰写</span><span class="sxs-lookup"><span data-stu-id="44017-1494">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="44017-1495">示例</span><span class="sxs-lookup"><span data-stu-id="44017-1495">Examples</span></span>

```javascript
Office.context.mailbox.item.saveAsync(
  function callback(result) {
    // Process the result.
  });
```

<span data-ttu-id="44017-p187">下面是传递给回调函数的 `result` 参数的示例。`value` 属性包含的项目的项目 ID。</span><span class="sxs-lookup"><span data-stu-id="44017-p187">The following is an example of the `result` parameter passed to the callback function. The `value` property contains the item ID of the item.</span></span>

```json
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

####  <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="44017-1498">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="44017-1498">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="44017-1499">以异步方式将数据插入到邮件的正文或主题中。</span><span class="sxs-lookup"><span data-stu-id="44017-1499">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="44017-p188">`setSelectedDataAsync` 方法将指定字符串插入到项目主题或正文的光标位置，或者，如果在编辑器中已选择文本，则该方法将替换选择的文本。如果光标不在正文或主题字段中，则返回错误。插入之后，光标会位于插入内容的末尾。</span><span class="sxs-lookup"><span data-stu-id="44017-p188">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="44017-1503">Parameters</span><span class="sxs-lookup"><span data-stu-id="44017-1503">Parameters</span></span>

|<span data-ttu-id="44017-1504">名称</span><span class="sxs-lookup"><span data-stu-id="44017-1504">Name</span></span>|<span data-ttu-id="44017-1505">类型</span><span class="sxs-lookup"><span data-stu-id="44017-1505">Type</span></span>|<span data-ttu-id="44017-1506">属性</span><span class="sxs-lookup"><span data-stu-id="44017-1506">Attributes</span></span>|<span data-ttu-id="44017-1507">说明</span><span class="sxs-lookup"><span data-stu-id="44017-1507">Description</span></span>|
|---|---|---|---|
|`data`|<span data-ttu-id="44017-1508">字符串</span><span class="sxs-lookup"><span data-stu-id="44017-1508">String</span></span>||<span data-ttu-id="44017-p189">要插入的数据。数据不得超过 1,000,000 个字符。如果传入的数据超过 1,000,000 个字符，则会引发 `ArgumentOutOfRange` 异常。</span><span class="sxs-lookup"><span data-stu-id="44017-p189">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`|<span data-ttu-id="44017-1512">Object</span><span class="sxs-lookup"><span data-stu-id="44017-1512">Object</span></span>|<span data-ttu-id="44017-1513">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="44017-1513">&lt;optional&gt;</span></span>|<span data-ttu-id="44017-1514">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="44017-1514">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="44017-1515">对象</span><span class="sxs-lookup"><span data-stu-id="44017-1515">Object</span></span>|<span data-ttu-id="44017-1516">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="44017-1516">&lt;optional&gt;</span></span>|<span data-ttu-id="44017-1517">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="44017-1517">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`|[<span data-ttu-id="44017-1518">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="44017-1518">Office.CoercionType</span></span>](office.md#coerciontype-string)|<span data-ttu-id="44017-1519">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="44017-1519">&lt;optional&gt;</span></span>|<span data-ttu-id="44017-p190">如果为 `text`，则在 Outlook Web App 和 Outlook 中应用当前样式。如果该字段是 HTML 编辑器，则仅插入文本数据，即使数据为 HTML。</span><span class="sxs-lookup"><span data-stu-id="44017-p190">If `text`, the current style is applied in Outlook Web App and Outlook. If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="44017-p191">如果 `html` 和该字段支持 HTML（主题不支持），则在 Outlook Web App 中应用当前样式，而在 Outlook 中应用默认样式。如果该字段是文本字段，则返回 `InvalidDataFormat` 错误。</span><span class="sxs-lookup"><span data-stu-id="44017-p191">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook Web App and the default style is applied in Outlook. If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="44017-1524">如果未设置 `coercionType`，则结果取决于该字段：如果该字段是 HTML，则使用 HTML；如果该字段是文本，则使用纯文本。</span><span class="sxs-lookup"><span data-stu-id="44017-1524">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`|<span data-ttu-id="44017-1525">function</span><span class="sxs-lookup"><span data-stu-id="44017-1525">function</span></span>||<span data-ttu-id="44017-1526">方法完成后，使用单个参数 `callback`（一个 [`asyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `AsyncResult` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="44017-1526">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="44017-1527">Requirements</span><span class="sxs-lookup"><span data-stu-id="44017-1527">Requirements</span></span>

|<span data-ttu-id="44017-1528">要求</span><span class="sxs-lookup"><span data-stu-id="44017-1528">Requirement</span></span>|<span data-ttu-id="44017-1529">值</span><span class="sxs-lookup"><span data-stu-id="44017-1529">Value</span></span>|
|---|---|
|[<span data-ttu-id="44017-1530">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="44017-1530">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="44017-1531">1.2</span><span class="sxs-lookup"><span data-stu-id="44017-1531">1.2</span></span>|
|[<span data-ttu-id="44017-1532">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="44017-1532">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="44017-1533">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="44017-1533">ReadWriteItem</span></span>|
|[<span data-ttu-id="44017-1534">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="44017-1534">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="44017-1535">撰写</span><span class="sxs-lookup"><span data-stu-id="44017-1535">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="44017-1536">示例</span><span class="sxs-lookup"><span data-stu-id="44017-1536">Example</span></span>

```javascript
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```
