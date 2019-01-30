---
title: Office.context.mailbox.item-预览要求集
description: ''
ms.date: 01/16/2019
localization_priority: Normal
ms.openlocfilehash: b4b2ec9c735270d9b1bfca3d1c24ef6b0f1ca1cb
ms.sourcegitcommit: d1aa7201820176ed986b9f00bb9c88e055906c77
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 01/23/2019
ms.locfileid: "29389597"
---
# <a name="item"></a><span data-ttu-id="96cbb-102">item</span><span class="sxs-lookup"><span data-stu-id="96cbb-102">item</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmditem"></a><span data-ttu-id="96cbb-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span><span class="sxs-lookup"><span data-stu-id="96cbb-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span></span>

<span data-ttu-id="96cbb-p101">`item` 命名空间用于访问当前选定的邮件、会议请求或约会。可以通过使用 [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtype) 属性确定 `item` 的类型。</span><span class="sxs-lookup"><span data-stu-id="96cbb-p101">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="96cbb-106">Requirements</span><span class="sxs-lookup"><span data-stu-id="96cbb-106">Requirements</span></span>

|<span data-ttu-id="96cbb-107">要求</span><span class="sxs-lookup"><span data-stu-id="96cbb-107">Requirement</span></span>|<span data-ttu-id="96cbb-108">值</span><span class="sxs-lookup"><span data-stu-id="96cbb-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="96cbb-109">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="96cbb-109">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="96cbb-110">1.0</span><span class="sxs-lookup"><span data-stu-id="96cbb-110">1.0</span></span>|
|[<span data-ttu-id="96cbb-111">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="96cbb-111">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="96cbb-112">受限</span><span class="sxs-lookup"><span data-stu-id="96cbb-112">Restricted</span></span>|
|[<span data-ttu-id="96cbb-113">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="96cbb-113">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="96cbb-114">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="96cbb-114">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="96cbb-115">成员和方法</span><span class="sxs-lookup"><span data-stu-id="96cbb-115">Members and methods</span></span>

| <span data-ttu-id="96cbb-116">成员</span><span class="sxs-lookup"><span data-stu-id="96cbb-116">Member</span></span> | <span data-ttu-id="96cbb-117">类型</span><span class="sxs-lookup"><span data-stu-id="96cbb-117">Type</span></span> |
|--------|------|
| [<span data-ttu-id="96cbb-118">attachments</span><span class="sxs-lookup"><span data-stu-id="96cbb-118">attachments</span></span>](#attachments-arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetails) | <span data-ttu-id="96cbb-119">成员</span><span class="sxs-lookup"><span data-stu-id="96cbb-119">Member</span></span> |
| [<span data-ttu-id="96cbb-120">bcc</span><span class="sxs-lookup"><span data-stu-id="96cbb-120">bcc</span></span>](#bcc-recipientsjavascriptapioutlookofficerecipients) | <span data-ttu-id="96cbb-121">成员</span><span class="sxs-lookup"><span data-stu-id="96cbb-121">Member</span></span> |
| [<span data-ttu-id="96cbb-122">body</span><span class="sxs-lookup"><span data-stu-id="96cbb-122">body</span></span>](#body-bodyjavascriptapioutlookofficebody) | <span data-ttu-id="96cbb-123">成员</span><span class="sxs-lookup"><span data-stu-id="96cbb-123">Member</span></span> |
| [<span data-ttu-id="96cbb-124">cc</span><span class="sxs-lookup"><span data-stu-id="96cbb-124">cc</span></span>](#cc-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients) | <span data-ttu-id="96cbb-125">成员</span><span class="sxs-lookup"><span data-stu-id="96cbb-125">Member</span></span> |
| [<span data-ttu-id="96cbb-126">conversationId</span><span class="sxs-lookup"><span data-stu-id="96cbb-126">conversationId</span></span>](#nullable-conversationid-string) | <span data-ttu-id="96cbb-127">成员</span><span class="sxs-lookup"><span data-stu-id="96cbb-127">Member</span></span> |
| [<span data-ttu-id="96cbb-128">dateTimeCreated</span><span class="sxs-lookup"><span data-stu-id="96cbb-128">dateTimeCreated</span></span>](#datetimecreated-date) | <span data-ttu-id="96cbb-129">成员</span><span class="sxs-lookup"><span data-stu-id="96cbb-129">Member</span></span> |
| [<span data-ttu-id="96cbb-130">dateTimeModified</span><span class="sxs-lookup"><span data-stu-id="96cbb-130">dateTimeModified</span></span>](#datetimemodified-date) | <span data-ttu-id="96cbb-131">成员</span><span class="sxs-lookup"><span data-stu-id="96cbb-131">Member</span></span> |
| [<span data-ttu-id="96cbb-132">end</span><span class="sxs-lookup"><span data-stu-id="96cbb-132">end</span></span>](#end-datetimejavascriptapioutlookofficetime) | <span data-ttu-id="96cbb-133">成员</span><span class="sxs-lookup"><span data-stu-id="96cbb-133">Member</span></span> |
| [<span data-ttu-id="96cbb-134">from</span><span class="sxs-lookup"><span data-stu-id="96cbb-134">from</span></span>](#from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsfromjavascriptapioutlookofficefrom) | <span data-ttu-id="96cbb-135">Member</span><span class="sxs-lookup"><span data-stu-id="96cbb-135">Member</span></span> |
| [<span data-ttu-id="96cbb-136">internetHeaders</span><span class="sxs-lookup"><span data-stu-id="96cbb-136">internetHeaders</span></span>](#internetheaders-internetheadersjavascriptapioutlookofficeinternetheaders) | <span data-ttu-id="96cbb-137">Member</span><span class="sxs-lookup"><span data-stu-id="96cbb-137">Member</span></span> |
| [<span data-ttu-id="96cbb-138">internetMessageId</span><span class="sxs-lookup"><span data-stu-id="96cbb-138">internetMessageId</span></span>](#internetmessageid-string) | <span data-ttu-id="96cbb-139">成员</span><span class="sxs-lookup"><span data-stu-id="96cbb-139">Member</span></span> |
| [<span data-ttu-id="96cbb-140">itemClass</span><span class="sxs-lookup"><span data-stu-id="96cbb-140">itemClass</span></span>](#itemclass-string) | <span data-ttu-id="96cbb-141">成员</span><span class="sxs-lookup"><span data-stu-id="96cbb-141">Member</span></span> |
| [<span data-ttu-id="96cbb-142">itemId</span><span class="sxs-lookup"><span data-stu-id="96cbb-142">itemId</span></span>](#nullable-itemid-string) | <span data-ttu-id="96cbb-143">成员</span><span class="sxs-lookup"><span data-stu-id="96cbb-143">Member</span></span> |
| [<span data-ttu-id="96cbb-144">itemType</span><span class="sxs-lookup"><span data-stu-id="96cbb-144">itemType</span></span>](#itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtype) | <span data-ttu-id="96cbb-145">成员</span><span class="sxs-lookup"><span data-stu-id="96cbb-145">Member</span></span> |
| [<span data-ttu-id="96cbb-146">location</span><span class="sxs-lookup"><span data-stu-id="96cbb-146">location</span></span>](#location-stringlocationjavascriptapioutlookofficelocation) | <span data-ttu-id="96cbb-147">成员</span><span class="sxs-lookup"><span data-stu-id="96cbb-147">Member</span></span> |
| [<span data-ttu-id="96cbb-148">normalizedSubject</span><span class="sxs-lookup"><span data-stu-id="96cbb-148">normalizedSubject</span></span>](#normalizedsubject-string) | <span data-ttu-id="96cbb-149">成员</span><span class="sxs-lookup"><span data-stu-id="96cbb-149">Member</span></span> |
| [<span data-ttu-id="96cbb-150">notificationMessages</span><span class="sxs-lookup"><span data-stu-id="96cbb-150">notificationMessages</span></span>](#notificationmessages-notificationmessagesjavascriptapioutlookofficenotificationmessages) | <span data-ttu-id="96cbb-151">成员</span><span class="sxs-lookup"><span data-stu-id="96cbb-151">Member</span></span> |
| [<span data-ttu-id="96cbb-152">optionalAttendees</span><span class="sxs-lookup"><span data-stu-id="96cbb-152">optionalAttendees</span></span>](#optionalattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients) | <span data-ttu-id="96cbb-153">成员</span><span class="sxs-lookup"><span data-stu-id="96cbb-153">Member</span></span> |
| [<span data-ttu-id="96cbb-154">organizer</span><span class="sxs-lookup"><span data-stu-id="96cbb-154">organizer</span></span>](#organizer-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsorganizerjavascriptapioutlookofficeorganizer) | <span data-ttu-id="96cbb-155">Member</span><span class="sxs-lookup"><span data-stu-id="96cbb-155">Member</span></span> |
| [<span data-ttu-id="96cbb-156">recurrence</span><span class="sxs-lookup"><span data-stu-id="96cbb-156">recurrence</span></span>](#nullable-recurrence-recurrencejavascriptapioutlookofficerecurrence) | <span data-ttu-id="96cbb-157">Member</span><span class="sxs-lookup"><span data-stu-id="96cbb-157">Member</span></span> |
| [<span data-ttu-id="96cbb-158">requiredAttendees</span><span class="sxs-lookup"><span data-stu-id="96cbb-158">requiredAttendees</span></span>](#requiredattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients) | <span data-ttu-id="96cbb-159">成员</span><span class="sxs-lookup"><span data-stu-id="96cbb-159">Member</span></span> |
| [<span data-ttu-id="96cbb-160">sender</span><span class="sxs-lookup"><span data-stu-id="96cbb-160">sender</span></span>](#sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetails) | <span data-ttu-id="96cbb-161">Member</span><span class="sxs-lookup"><span data-stu-id="96cbb-161">Member</span></span> |
| [<span data-ttu-id="96cbb-162">seriesId</span><span class="sxs-lookup"><span data-stu-id="96cbb-162">seriesId</span></span>](#nullable-seriesid-string) | <span data-ttu-id="96cbb-163">Member</span><span class="sxs-lookup"><span data-stu-id="96cbb-163">Member</span></span> |
| [<span data-ttu-id="96cbb-164">start</span><span class="sxs-lookup"><span data-stu-id="96cbb-164">start</span></span>](#start-datetimejavascriptapioutlookofficetime) | <span data-ttu-id="96cbb-165">成员</span><span class="sxs-lookup"><span data-stu-id="96cbb-165">Member</span></span> |
| [<span data-ttu-id="96cbb-166">subject</span><span class="sxs-lookup"><span data-stu-id="96cbb-166">subject</span></span>](#subject-stringsubjectjavascriptapioutlookofficesubject) | <span data-ttu-id="96cbb-167">成员</span><span class="sxs-lookup"><span data-stu-id="96cbb-167">Member</span></span> |
| [<span data-ttu-id="96cbb-168">to</span><span class="sxs-lookup"><span data-stu-id="96cbb-168">to</span></span>](#to-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients) | <span data-ttu-id="96cbb-169">成员</span><span class="sxs-lookup"><span data-stu-id="96cbb-169">Member</span></span> |
| [<span data-ttu-id="96cbb-170">addFileAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="96cbb-170">addFileAttachmentAsync</span></span>](#addfileattachmentasyncuri-attachmentname-options-callback) | <span data-ttu-id="96cbb-171">方法</span><span class="sxs-lookup"><span data-stu-id="96cbb-171">Method</span></span> |
| [<span data-ttu-id="96cbb-172">addFileAttachmentFromBase64Async</span><span class="sxs-lookup"><span data-stu-id="96cbb-172">addFileAttachmentFromBase64Async</span></span>](#addfileattachmentfrombase64asyncbase64file-attachmentname-options-callback) | <span data-ttu-id="96cbb-173">方法</span><span class="sxs-lookup"><span data-stu-id="96cbb-173">Method</span></span> |
| [<span data-ttu-id="96cbb-174">addHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="96cbb-174">addHandlerAsync</span></span>](#addhandlerasynceventtype-handler-options-callback) | <span data-ttu-id="96cbb-175">方法</span><span class="sxs-lookup"><span data-stu-id="96cbb-175">Method</span></span> |
| [<span data-ttu-id="96cbb-176">addItemAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="96cbb-176">addItemAttachmentAsync</span></span>](#additemattachmentasyncitemid-attachmentname-options-callback) | <span data-ttu-id="96cbb-177">方法</span><span class="sxs-lookup"><span data-stu-id="96cbb-177">Method</span></span> |
| [<span data-ttu-id="96cbb-178">close</span><span class="sxs-lookup"><span data-stu-id="96cbb-178">close</span></span>](#close) | <span data-ttu-id="96cbb-179">方法</span><span class="sxs-lookup"><span data-stu-id="96cbb-179">Method</span></span> |
| [<span data-ttu-id="96cbb-180">displayReplyAllForm</span><span class="sxs-lookup"><span data-stu-id="96cbb-180">displayReplyAllForm</span></span>](#displayreplyallformformdata) | <span data-ttu-id="96cbb-181">方法</span><span class="sxs-lookup"><span data-stu-id="96cbb-181">Method</span></span> |
| [<span data-ttu-id="96cbb-182">displayReplyForm</span><span class="sxs-lookup"><span data-stu-id="96cbb-182">displayReplyForm</span></span>](#displayreplyformformdata) | <span data-ttu-id="96cbb-183">方法</span><span class="sxs-lookup"><span data-stu-id="96cbb-183">Method</span></span> |
| [<span data-ttu-id="96cbb-184">getAttachmentContentAsync</span><span class="sxs-lookup"><span data-stu-id="96cbb-184">getAttachmentContentAsync</span></span>](#getattachmentcontentasyncattachmentid-options-callback--attachmentcontentjavascriptapioutlookofficeattachmentcontent) | <span data-ttu-id="96cbb-185">方法</span><span class="sxs-lookup"><span data-stu-id="96cbb-185">Method</span></span> |
| [<span data-ttu-id="96cbb-186">getAttachmentsAsync</span><span class="sxs-lookup"><span data-stu-id="96cbb-186">getAttachmentsAsync</span></span>](#getattachmentsasyncoptions-callback--arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetails) | <span data-ttu-id="96cbb-187">方法</span><span class="sxs-lookup"><span data-stu-id="96cbb-187">Method</span></span> |
| [<span data-ttu-id="96cbb-188">getEntities</span><span class="sxs-lookup"><span data-stu-id="96cbb-188">getEntities</span></span>](#getentities--entitiesjavascriptapioutlookofficeentities) | <span data-ttu-id="96cbb-189">方法</span><span class="sxs-lookup"><span data-stu-id="96cbb-189">Method</span></span> |
| [<span data-ttu-id="96cbb-190">getEntitiesByType</span><span class="sxs-lookup"><span data-stu-id="96cbb-190">getEntitiesByType</span></span>](#getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlookofficecontactmeetingsuggestionjavascriptapioutlookofficemeetingsuggestionphonenumberjavascriptapioutlookofficephonenumbertasksuggestionjavascriptapioutlookofficetasksuggestion) | <span data-ttu-id="96cbb-191">方法</span><span class="sxs-lookup"><span data-stu-id="96cbb-191">Method</span></span> |
| [<span data-ttu-id="96cbb-192">getFilteredEntitiesByName</span><span class="sxs-lookup"><span data-stu-id="96cbb-192">getFilteredEntitiesByName</span></span>](#getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlookofficecontactmeetingsuggestionjavascriptapioutlookofficemeetingsuggestionphonenumberjavascriptapioutlookofficephonenumbertasksuggestionjavascriptapioutlookofficetasksuggestion) | <span data-ttu-id="96cbb-193">方法</span><span class="sxs-lookup"><span data-stu-id="96cbb-193">Method</span></span> |
| [<span data-ttu-id="96cbb-194">getInitializationContextAsync</span><span class="sxs-lookup"><span data-stu-id="96cbb-194">getInitializationContextAsync</span></span>](#getinitializationcontextasyncoptions-callback) | <span data-ttu-id="96cbb-195">方法</span><span class="sxs-lookup"><span data-stu-id="96cbb-195">Method</span></span> |
| [<span data-ttu-id="96cbb-196">getRegExMatches</span><span class="sxs-lookup"><span data-stu-id="96cbb-196">getRegExMatches</span></span>](#getregexmatches--object) | <span data-ttu-id="96cbb-197">方法</span><span class="sxs-lookup"><span data-stu-id="96cbb-197">Method</span></span> |
| [<span data-ttu-id="96cbb-198">getRegExMatchesByName</span><span class="sxs-lookup"><span data-stu-id="96cbb-198">getRegExMatchesByName</span></span>](#getregexmatchesbynamename--nullable-array-string-) | <span data-ttu-id="96cbb-199">方法</span><span class="sxs-lookup"><span data-stu-id="96cbb-199">Method</span></span> |
| [<span data-ttu-id="96cbb-200">getSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="96cbb-200">getSelectedDataAsync</span></span>](#getselecteddataasynccoerciontype-options-callback--string) | <span data-ttu-id="96cbb-201">方法</span><span class="sxs-lookup"><span data-stu-id="96cbb-201">Method</span></span> |
| [<span data-ttu-id="96cbb-202">getSelectedEntities</span><span class="sxs-lookup"><span data-stu-id="96cbb-202">getSelectedEntities</span></span>](#getselectedentities--entitiesjavascriptapioutlookofficeentities) | <span data-ttu-id="96cbb-203">方法</span><span class="sxs-lookup"><span data-stu-id="96cbb-203">Method</span></span> |
| [<span data-ttu-id="96cbb-204">getSelectedRegExMatches</span><span class="sxs-lookup"><span data-stu-id="96cbb-204">getSelectedRegExMatches</span></span>](#getselectedregexmatches--object) | <span data-ttu-id="96cbb-205">方法</span><span class="sxs-lookup"><span data-stu-id="96cbb-205">Method</span></span> |
| [<span data-ttu-id="96cbb-206">getSharedPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="96cbb-206">getSharedPropertiesAsync</span></span>](#getsharedpropertiesasyncoptions-callback) | <span data-ttu-id="96cbb-207">方法</span><span class="sxs-lookup"><span data-stu-id="96cbb-207">Method</span></span> |
| [<span data-ttu-id="96cbb-208">loadCustomPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="96cbb-208">loadCustomPropertiesAsync</span></span>](#loadcustompropertiesasynccallback-usercontext) | <span data-ttu-id="96cbb-209">方法</span><span class="sxs-lookup"><span data-stu-id="96cbb-209">Method</span></span> |
| [<span data-ttu-id="96cbb-210">removeAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="96cbb-210">removeAttachmentAsync</span></span>](#removeattachmentasyncattachmentid-options-callback) | <span data-ttu-id="96cbb-211">方法</span><span class="sxs-lookup"><span data-stu-id="96cbb-211">Method</span></span> |
| [<span data-ttu-id="96cbb-212">removeHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="96cbb-212">removeHandlerAsync</span></span>](#removehandlerasynceventtype-options-callback) | <span data-ttu-id="96cbb-213">方法</span><span class="sxs-lookup"><span data-stu-id="96cbb-213">Method</span></span> |
| [<span data-ttu-id="96cbb-214">saveAsync</span><span class="sxs-lookup"><span data-stu-id="96cbb-214">saveAsync</span></span>](#saveasyncoptions-callback) | <span data-ttu-id="96cbb-215">方法</span><span class="sxs-lookup"><span data-stu-id="96cbb-215">Method</span></span> |
| [<span data-ttu-id="96cbb-216">setSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="96cbb-216">setSelectedDataAsync</span></span>](#setselecteddataasyncdata-options-callback) | <span data-ttu-id="96cbb-217">方法</span><span class="sxs-lookup"><span data-stu-id="96cbb-217">Method</span></span> |

### <a name="example"></a><span data-ttu-id="96cbb-218">示例</span><span class="sxs-lookup"><span data-stu-id="96cbb-218">Example</span></span>

<span data-ttu-id="96cbb-219">以下 JavaScript 代码示例显示了如何访问 Outlook 中当前项目的 `subject` 属性。</span><span class="sxs-lookup"><span data-stu-id="96cbb-219">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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
}
```

### <a name="members"></a><span data-ttu-id="96cbb-220">成员</span><span class="sxs-lookup"><span data-stu-id="96cbb-220">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetails"></a><span data-ttu-id="96cbb-221">attachments :Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="96cbb-221">attachments :Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span></span>

<span data-ttu-id="96cbb-222">获取项目的附件作为数组。</span><span class="sxs-lookup"><span data-stu-id="96cbb-222">Gets the item's attachments as an array.</span></span> <span data-ttu-id="96cbb-223">仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="96cbb-223">Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="96cbb-224">由于潜在的安全问题，某些类型的文件会受到 Outlook 阻止，并且不会返回。</span><span class="sxs-lookup"><span data-stu-id="96cbb-224">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="96cbb-225">如需了解更多信息，请参阅 [Outlook 中阻止的附件](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519)。</span><span class="sxs-lookup"><span data-stu-id="96cbb-225">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="96cbb-226">类型：</span><span class="sxs-lookup"><span data-stu-id="96cbb-226">Type:</span></span>

*   <span data-ttu-id="96cbb-227">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="96cbb-227">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span></span>

##### <a name="requirements"></a><span data-ttu-id="96cbb-228">要求</span><span class="sxs-lookup"><span data-stu-id="96cbb-228">Requirements</span></span>

|<span data-ttu-id="96cbb-229">要求</span><span class="sxs-lookup"><span data-stu-id="96cbb-229">Requirement</span></span>|<span data-ttu-id="96cbb-230">值</span><span class="sxs-lookup"><span data-stu-id="96cbb-230">Value</span></span>|
|---|---|
|[<span data-ttu-id="96cbb-231">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="96cbb-231">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="96cbb-232">1.0</span><span class="sxs-lookup"><span data-stu-id="96cbb-232">1.0</span></span>|
|[<span data-ttu-id="96cbb-233">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="96cbb-233">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="96cbb-234">ReadItem</span><span class="sxs-lookup"><span data-stu-id="96cbb-234">ReadItem</span></span>|
|[<span data-ttu-id="96cbb-235">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="96cbb-235">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="96cbb-236">阅读</span><span class="sxs-lookup"><span data-stu-id="96cbb-236">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="96cbb-237">示例</span><span class="sxs-lookup"><span data-stu-id="96cbb-237">Example</span></span>

<span data-ttu-id="96cbb-238">以下代码使用当前项目上所有附件的详细信息构成 HTML 字符串。</span><span class="sxs-lookup"><span data-stu-id="96cbb-238">The following code builds an HTML string with details of all attachments on the current item.</span></span>

```javascript
var _Item = Office.context.mailbox.item;
var outputString = "";

if (_Item.attachments.length > 0) {
  for (i = 0 ; i < _Item.attachments.length ; i++) {
    var _att = _Item.attachments[i];
    outputString += "<BR>" + i + ". Name: ";
    outputString += _att.name;
    outputString += "<BR>ID: " + _att.id;
    outputString += "<BR>contentType: " + _att.contentType;
    outputString += "<BR>size: " + _att.size;
    outputString += "<BR>attachmentType: " + _att.attachmentType;
    outputString += "<BR>isInline: " + _att.isInline;
  }
}

// Do something with outputString
```

####  <a name="bcc-recipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="96cbb-239">bcc :[Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="96cbb-239">bcc :[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="96cbb-240">获取对象，该对象提供用于获取或更新邮件的密件抄送 (Bcc) 行上的收件人的方法。</span><span class="sxs-lookup"><span data-stu-id="96cbb-240">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="96cbb-241">仅限撰写模式。</span><span class="sxs-lookup"><span data-stu-id="96cbb-241">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="96cbb-242">类型：</span><span class="sxs-lookup"><span data-stu-id="96cbb-242">Type:</span></span>

*   [<span data-ttu-id="96cbb-243">收件人</span><span class="sxs-lookup"><span data-stu-id="96cbb-243">Recipients</span></span>](/javascript/api/outlook/office.recipients)

##### <a name="requirements"></a><span data-ttu-id="96cbb-244">要求</span><span class="sxs-lookup"><span data-stu-id="96cbb-244">Requirements</span></span>

|<span data-ttu-id="96cbb-245">要求</span><span class="sxs-lookup"><span data-stu-id="96cbb-245">Requirement</span></span>|<span data-ttu-id="96cbb-246">值</span><span class="sxs-lookup"><span data-stu-id="96cbb-246">Value</span></span>|
|---|---|
|[<span data-ttu-id="96cbb-247">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="96cbb-247">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="96cbb-248">1.1</span><span class="sxs-lookup"><span data-stu-id="96cbb-248">1.1</span></span>|
|[<span data-ttu-id="96cbb-249">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="96cbb-249">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="96cbb-250">ReadItem</span><span class="sxs-lookup"><span data-stu-id="96cbb-250">ReadItem</span></span>|
|[<span data-ttu-id="96cbb-251">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="96cbb-251">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="96cbb-252">撰写</span><span class="sxs-lookup"><span data-stu-id="96cbb-252">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="96cbb-253">示例</span><span class="sxs-lookup"><span data-stu-id="96cbb-253">Example</span></span>

```javascript
Office.context.mailbox.item.bcc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.bcc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.bcc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfBccRecipients = asyncResult.value;
}
```

####  <a name="body-bodyjavascriptapioutlookofficebody"></a><span data-ttu-id="96cbb-254">body :[Body](/javascript/api/outlook/office.body)</span><span class="sxs-lookup"><span data-stu-id="96cbb-254">body :[Body](/javascript/api/outlook/office.body)</span></span>

<span data-ttu-id="96cbb-255">获取一个提供用于处理项目正文的方法的对象。</span><span class="sxs-lookup"><span data-stu-id="96cbb-255">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="96cbb-256">类型：</span><span class="sxs-lookup"><span data-stu-id="96cbb-256">Type:</span></span>

*   [<span data-ttu-id="96cbb-257">Body</span><span class="sxs-lookup"><span data-stu-id="96cbb-257">Body</span></span>](/javascript/api/outlook/office.body)

##### <a name="requirements"></a><span data-ttu-id="96cbb-258">要求</span><span class="sxs-lookup"><span data-stu-id="96cbb-258">Requirements</span></span>

|<span data-ttu-id="96cbb-259">要求</span><span class="sxs-lookup"><span data-stu-id="96cbb-259">Requirement</span></span>|<span data-ttu-id="96cbb-260">值</span><span class="sxs-lookup"><span data-stu-id="96cbb-260">Value</span></span>|
|---|---|
|[<span data-ttu-id="96cbb-261">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="96cbb-261">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="96cbb-262">1.1</span><span class="sxs-lookup"><span data-stu-id="96cbb-262">1.1</span></span>|
|[<span data-ttu-id="96cbb-263">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="96cbb-263">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="96cbb-264">ReadItem</span><span class="sxs-lookup"><span data-stu-id="96cbb-264">ReadItem</span></span>|
|[<span data-ttu-id="96cbb-265">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="96cbb-265">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="96cbb-266">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="96cbb-266">Compose or read</span></span>|

####  <a name="cc-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="96cbb-267">cc :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="96cbb-267">cc :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="96cbb-268">提供对邮件的抄送 (Cc) 收件人的访问权限。</span><span class="sxs-lookup"><span data-stu-id="96cbb-268">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="96cbb-269">对象的类型和访问级别取决于当前项目的模式。</span><span class="sxs-lookup"><span data-stu-id="96cbb-269">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="96cbb-270">阅读模式</span><span class="sxs-lookup"><span data-stu-id="96cbb-270">Read mode</span></span>

<span data-ttu-id="96cbb-p106">`cc` 属性返回包含邮件的**抄送**行上所列的每个收件人的 `EmailAddressDetails` 对象的数组。集合上限为 100 个成员。</span><span class="sxs-lookup"><span data-stu-id="96cbb-p106">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="96cbb-273">撰写模式</span><span class="sxs-lookup"><span data-stu-id="96cbb-273">Compose mode</span></span>

<span data-ttu-id="96cbb-274">`cc` 属性返回一个 `Recipients` 对象，该对象提供用于获取或更新邮件的“**抄送**”行上收件人的方法。</span><span class="sxs-lookup"><span data-stu-id="96cbb-274">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="96cbb-275">类型：</span><span class="sxs-lookup"><span data-stu-id="96cbb-275">Type:</span></span>

*   <span data-ttu-id="96cbb-276">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="96cbb-276">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="96cbb-277">要求</span><span class="sxs-lookup"><span data-stu-id="96cbb-277">Requirements</span></span>

|<span data-ttu-id="96cbb-278">要求</span><span class="sxs-lookup"><span data-stu-id="96cbb-278">Requirement</span></span>|<span data-ttu-id="96cbb-279">值</span><span class="sxs-lookup"><span data-stu-id="96cbb-279">Value</span></span>|
|---|---|
|[<span data-ttu-id="96cbb-280">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="96cbb-280">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="96cbb-281">1.0</span><span class="sxs-lookup"><span data-stu-id="96cbb-281">1.0</span></span>|
|[<span data-ttu-id="96cbb-282">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="96cbb-282">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="96cbb-283">ReadItem</span><span class="sxs-lookup"><span data-stu-id="96cbb-283">ReadItem</span></span>|
|[<span data-ttu-id="96cbb-284">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="96cbb-284">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="96cbb-285">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="96cbb-285">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="96cbb-286">示例</span><span class="sxs-lookup"><span data-stu-id="96cbb-286">Example</span></span>

```javascript
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

####  <a name="nullable-conversationid-string"></a><span data-ttu-id="96cbb-287">(nullable) conversationId :String</span><span class="sxs-lookup"><span data-stu-id="96cbb-287">(nullable) conversationId :String</span></span>

<span data-ttu-id="96cbb-288">获取包含特定消息的电子邮件会话的标识符。</span><span class="sxs-lookup"><span data-stu-id="96cbb-288">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="96cbb-p107">如果在阅读窗体或撰写窗体的回复中激活邮件应用程序，则此属性可以获得一个整数值。如果用户随后更改了回复邮件的主题（若发送回复），则该邮件的对话 ID 将改变且之前获取的值将不适用。</span><span class="sxs-lookup"><span data-stu-id="96cbb-p107">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="96cbb-p108">对于撰写窗体的新项目，此属性获得一个 null 值。如果用户设置一个主题并保存该项目，`conversationId` 属性将返回一个值。</span><span class="sxs-lookup"><span data-stu-id="96cbb-p108">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="96cbb-293">类型:</span><span class="sxs-lookup"><span data-stu-id="96cbb-293">Type:</span></span>

*   <span data-ttu-id="96cbb-294">String</span><span class="sxs-lookup"><span data-stu-id="96cbb-294">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="96cbb-295">要求</span><span class="sxs-lookup"><span data-stu-id="96cbb-295">Requirements</span></span>

|<span data-ttu-id="96cbb-296">要求</span><span class="sxs-lookup"><span data-stu-id="96cbb-296">Requirement</span></span>|<span data-ttu-id="96cbb-297">值</span><span class="sxs-lookup"><span data-stu-id="96cbb-297">Value</span></span>|
|---|---|
|[<span data-ttu-id="96cbb-298">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="96cbb-298">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="96cbb-299">1.0</span><span class="sxs-lookup"><span data-stu-id="96cbb-299">1.0</span></span>|
|[<span data-ttu-id="96cbb-300">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="96cbb-300">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="96cbb-301">ReadItem</span><span class="sxs-lookup"><span data-stu-id="96cbb-301">ReadItem</span></span>|
|[<span data-ttu-id="96cbb-302">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="96cbb-302">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="96cbb-303">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="96cbb-303">Compose or read</span></span>|

#### <a name="datetimecreated-date"></a><span data-ttu-id="96cbb-304">dateTimeCreated :Date</span><span class="sxs-lookup"><span data-stu-id="96cbb-304">dateTimeCreated :Date</span></span>

<span data-ttu-id="96cbb-p109">获取项目创建的日期和时间。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="96cbb-p109">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="96cbb-307">类型：</span><span class="sxs-lookup"><span data-stu-id="96cbb-307">Type:</span></span>

*   <span data-ttu-id="96cbb-308">日期</span><span class="sxs-lookup"><span data-stu-id="96cbb-308">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="96cbb-309">要求</span><span class="sxs-lookup"><span data-stu-id="96cbb-309">Requirements</span></span>

|<span data-ttu-id="96cbb-310">要求</span><span class="sxs-lookup"><span data-stu-id="96cbb-310">Requirement</span></span>|<span data-ttu-id="96cbb-311">值</span><span class="sxs-lookup"><span data-stu-id="96cbb-311">Value</span></span>|
|---|---|
|[<span data-ttu-id="96cbb-312">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="96cbb-312">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="96cbb-313">1.0</span><span class="sxs-lookup"><span data-stu-id="96cbb-313">1.0</span></span>|
|[<span data-ttu-id="96cbb-314">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="96cbb-314">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="96cbb-315">ReadItem</span><span class="sxs-lookup"><span data-stu-id="96cbb-315">ReadItem</span></span>|
|[<span data-ttu-id="96cbb-316">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="96cbb-316">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="96cbb-317">阅读</span><span class="sxs-lookup"><span data-stu-id="96cbb-317">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="96cbb-318">示例</span><span class="sxs-lookup"><span data-stu-id="96cbb-318">Example</span></span>

```javascript
var created = Office.context.mailbox.item.dateTimeCreated;
```

#### <a name="datetimemodified-date"></a><span data-ttu-id="96cbb-319">dateTimeModified :Date</span><span class="sxs-lookup"><span data-stu-id="96cbb-319">dateTimeModified :Date</span></span>

<span data-ttu-id="96cbb-p110">获取项目最近一次修改的日期和时间。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="96cbb-p110">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="96cbb-322">在 Outlook for iOS 或 Outlook for Android 中不支持此成员。</span><span class="sxs-lookup"><span data-stu-id="96cbb-322">This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="type"></a><span data-ttu-id="96cbb-323">类型:</span><span class="sxs-lookup"><span data-stu-id="96cbb-323">Type:</span></span>

*   <span data-ttu-id="96cbb-324">日期</span><span class="sxs-lookup"><span data-stu-id="96cbb-324">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="96cbb-325">要求</span><span class="sxs-lookup"><span data-stu-id="96cbb-325">Requirements</span></span>

|<span data-ttu-id="96cbb-326">要求</span><span class="sxs-lookup"><span data-stu-id="96cbb-326">Requirement</span></span>|<span data-ttu-id="96cbb-327">值</span><span class="sxs-lookup"><span data-stu-id="96cbb-327">Value</span></span>|
|---|---|
|[<span data-ttu-id="96cbb-328">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="96cbb-328">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="96cbb-329">1.0</span><span class="sxs-lookup"><span data-stu-id="96cbb-329">1.0</span></span>|
|[<span data-ttu-id="96cbb-330">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="96cbb-330">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="96cbb-331">ReadItem</span><span class="sxs-lookup"><span data-stu-id="96cbb-331">ReadItem</span></span>|
|[<span data-ttu-id="96cbb-332">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="96cbb-332">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="96cbb-333">阅读</span><span class="sxs-lookup"><span data-stu-id="96cbb-333">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="96cbb-334">示例</span><span class="sxs-lookup"><span data-stu-id="96cbb-334">Example</span></span>

```javascript
var modified = Office.context.mailbox.item.dateTimeModified;
```

####  <a name="end-datetimejavascriptapioutlookofficetime"></a><span data-ttu-id="96cbb-335">end :Date|[Time](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="96cbb-335">end :Date|[Time](/javascript/api/outlook/office.time)</span></span>

<span data-ttu-id="96cbb-336">获取或设置约会结束的日期和时间。</span><span class="sxs-lookup"><span data-stu-id="96cbb-336">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="96cbb-p111">将 `end` 属性表示为协调世界时 (UTC) 的日期和时间值。可使用 [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttime) 方法将 end 属性值转换为客户端的本地日期和时间。</span><span class="sxs-lookup"><span data-stu-id="96cbb-p111">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="96cbb-339">阅读模式</span><span class="sxs-lookup"><span data-stu-id="96cbb-339">Read mode</span></span>

<span data-ttu-id="96cbb-340">`end` 属性返回 `Date` 对象。</span><span class="sxs-lookup"><span data-stu-id="96cbb-340">The `end` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="96cbb-341">撰写模式</span><span class="sxs-lookup"><span data-stu-id="96cbb-341">Compose mode</span></span>

<span data-ttu-id="96cbb-342">`end` 属性返回 `Time` 对象。</span><span class="sxs-lookup"><span data-stu-id="96cbb-342">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="96cbb-343">使用 [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) 方法设置结束时间时，应使用 [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) 方法将客户端的本地时间转换为服务器的 UTC。</span><span class="sxs-lookup"><span data-stu-id="96cbb-343">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="96cbb-344">类型：</span><span class="sxs-lookup"><span data-stu-id="96cbb-344">Type:</span></span>

*   <span data-ttu-id="96cbb-345">Date | [Time](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="96cbb-345">Date | [Time](/javascript/api/outlook/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="96cbb-346">要求</span><span class="sxs-lookup"><span data-stu-id="96cbb-346">Requirements</span></span>

|<span data-ttu-id="96cbb-347">要求</span><span class="sxs-lookup"><span data-stu-id="96cbb-347">Requirement</span></span>|<span data-ttu-id="96cbb-348">值</span><span class="sxs-lookup"><span data-stu-id="96cbb-348">Value</span></span>|
|---|---|
|[<span data-ttu-id="96cbb-349">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="96cbb-349">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="96cbb-350">1.0</span><span class="sxs-lookup"><span data-stu-id="96cbb-350">1.0</span></span>|
|[<span data-ttu-id="96cbb-351">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="96cbb-351">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="96cbb-352">ReadItem</span><span class="sxs-lookup"><span data-stu-id="96cbb-352">ReadItem</span></span>|
|[<span data-ttu-id="96cbb-353">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="96cbb-353">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="96cbb-354">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="96cbb-354">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="96cbb-355">示例</span><span class="sxs-lookup"><span data-stu-id="96cbb-355">Example</span></span>

<span data-ttu-id="96cbb-356">以下示例通过使用 `Time` 对象的 [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) 方法，设置撰写模式下约会的结束时间。</span><span class="sxs-lookup"><span data-stu-id="96cbb-356">The following example sets the end time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

```javascript
var endTime = new Date("3/14/2015");
var options = {
  // Pass information that can be used
  // in the callback
  asyncContext: {verb:"Set"}
}
Office.context.mailbox.item.end.setAsync(endTime, options, function(result) {
  if (result.error) {
    console.debug(result.error);
  } else {
    // Access the asyncContext that was passed to the setAsync function
    console.debug("End Time " + result.asyncContext.verb);
  }
});
```

#### <a name="from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsfromjavascriptapioutlookofficefrom"></a><span data-ttu-id="96cbb-357">from :[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[From](/javascript/api/outlook/office.from)</span><span class="sxs-lookup"><span data-stu-id="96cbb-357">from :[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[From](/javascript/api/outlook/office.from)</span></span>

<span data-ttu-id="96cbb-358">获取邮件发件人的电子邮件地址。</span><span class="sxs-lookup"><span data-stu-id="96cbb-358">Gets the email address of the sender of a message.</span></span>

<span data-ttu-id="96cbb-p112">`from` 和 [`sender`](#sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetails) 属性表示同一个人，邮件由代理人发送的除外。在此情况下，`from` 属性表示代理程序，而 sender 属性表示代理人。</span><span class="sxs-lookup"><span data-stu-id="96cbb-p112">The `from` and [`sender`](#sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="96cbb-361">`from` 属性中 `EmailAddressDetails` 对象的 `recipientType` 属性为 `undefined`。</span><span class="sxs-lookup"><span data-stu-id="96cbb-361">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="96cbb-362">阅读模式</span><span class="sxs-lookup"><span data-stu-id="96cbb-362">Read mode</span></span>

<span data-ttu-id="96cbb-363">`from` 属性返回一个 `EmailAddressDetails` 对象。</span><span class="sxs-lookup"><span data-stu-id="96cbb-363">The `from` property returns an `EmailAddressDetails` object.</span></span>

```javascript
var subject = Office.context.mailbox.item.from;
```

##### <a name="compose-mode"></a><span data-ttu-id="96cbb-364">撰写模式</span><span class="sxs-lookup"><span data-stu-id="96cbb-364">Compose mode</span></span>

<span data-ttu-id="96cbb-365">`from` 属性返回一个 `From` 对象，该对象提供从值中进行获取的方法。</span><span class="sxs-lookup"><span data-stu-id="96cbb-365">The `from` property returns a `From` object that provides a method to get the from value.</span></span>

```javascript
Office.context.mailbox.item.from.getAsync(callback);

function callback(asyncResult) {
  var from = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="96cbb-366">类型：</span><span class="sxs-lookup"><span data-stu-id="96cbb-366">Type:</span></span>

*   <span data-ttu-id="96cbb-367">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [From](/javascript/api/outlook/office.from)</span><span class="sxs-lookup"><span data-stu-id="96cbb-367">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [From](/javascript/api/outlook/office.from)</span></span>

##### <a name="requirements"></a><span data-ttu-id="96cbb-368">要求</span><span class="sxs-lookup"><span data-stu-id="96cbb-368">Requirements</span></span>

|<span data-ttu-id="96cbb-369">要求</span><span class="sxs-lookup"><span data-stu-id="96cbb-369">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="96cbb-370">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="96cbb-370">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="96cbb-371">1.0</span><span class="sxs-lookup"><span data-stu-id="96cbb-371">1.0</span></span>|<span data-ttu-id="96cbb-372">1.7</span><span class="sxs-lookup"><span data-stu-id="96cbb-372">1.7</span></span>|
|[<span data-ttu-id="96cbb-373">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="96cbb-373">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="96cbb-374">ReadItem</span><span class="sxs-lookup"><span data-stu-id="96cbb-374">ReadItem</span></span>|<span data-ttu-id="96cbb-375">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="96cbb-375">ReadWriteItem</span></span>|
|[<span data-ttu-id="96cbb-376">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="96cbb-376">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="96cbb-377">阅读</span><span class="sxs-lookup"><span data-stu-id="96cbb-377">Read</span></span>|<span data-ttu-id="96cbb-378">撰写</span><span class="sxs-lookup"><span data-stu-id="96cbb-378">Compose</span></span>|

#### <a name="internetheaders-internetheadersjavascriptapioutlookofficeinternetheaders"></a><span data-ttu-id="96cbb-379">internetHeaders :[InternetHeaders](/javascript/api/outlook/office.internetheaders)</span><span class="sxs-lookup"><span data-stu-id="96cbb-379">internetHeaders :[InternetHeaders](/javascript/api/outlook/office.internetheaders)</span></span>

<span data-ttu-id="96cbb-380">获取或设置消息的 Internet 标头。</span><span class="sxs-lookup"><span data-stu-id="96cbb-380">Gets or sets the internet headers of a message.</span></span>

##### <a name="type"></a><span data-ttu-id="96cbb-381">类型：</span><span class="sxs-lookup"><span data-stu-id="96cbb-381">Type:</span></span>

*   [<span data-ttu-id="96cbb-382">InternetHeaders</span><span class="sxs-lookup"><span data-stu-id="96cbb-382">InternetHeaders</span></span>](/javascript/api/outlook/office.internetheaders)

##### <a name="requirements"></a><span data-ttu-id="96cbb-383">要求</span><span class="sxs-lookup"><span data-stu-id="96cbb-383">Requirements</span></span>

|<span data-ttu-id="96cbb-384">要求</span><span class="sxs-lookup"><span data-stu-id="96cbb-384">Requirement</span></span>|<span data-ttu-id="96cbb-385">值</span><span class="sxs-lookup"><span data-stu-id="96cbb-385">Value</span></span>|
|---|---|
|[<span data-ttu-id="96cbb-386">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="96cbb-386">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="96cbb-387">预览</span><span class="sxs-lookup"><span data-stu-id="96cbb-387">Preview</span></span>|
|[<span data-ttu-id="96cbb-388">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="96cbb-388">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="96cbb-389">ReadItem</span><span class="sxs-lookup"><span data-stu-id="96cbb-389">ReadItem</span></span>|
|[<span data-ttu-id="96cbb-390">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="96cbb-390">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="96cbb-391">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="96cbb-391">Compose or read</span></span>|

#### <a name="internetmessageid-string"></a><span data-ttu-id="96cbb-392">internetMessageId :String</span><span class="sxs-lookup"><span data-stu-id="96cbb-392">internetMessageId :String</span></span>

<span data-ttu-id="96cbb-p113">获取电子邮件的 Internet 消息标识符。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="96cbb-p113">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="96cbb-395">类型：</span><span class="sxs-lookup"><span data-stu-id="96cbb-395">Type:</span></span>

*   <span data-ttu-id="96cbb-396">String</span><span class="sxs-lookup"><span data-stu-id="96cbb-396">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="96cbb-397">要求</span><span class="sxs-lookup"><span data-stu-id="96cbb-397">Requirements</span></span>

|<span data-ttu-id="96cbb-398">要求</span><span class="sxs-lookup"><span data-stu-id="96cbb-398">Requirement</span></span>|<span data-ttu-id="96cbb-399">值</span><span class="sxs-lookup"><span data-stu-id="96cbb-399">Value</span></span>|
|---|---|
|[<span data-ttu-id="96cbb-400">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="96cbb-400">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="96cbb-401">1.0</span><span class="sxs-lookup"><span data-stu-id="96cbb-401">1.0</span></span>|
|[<span data-ttu-id="96cbb-402">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="96cbb-402">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="96cbb-403">ReadItem</span><span class="sxs-lookup"><span data-stu-id="96cbb-403">ReadItem</span></span>|
|[<span data-ttu-id="96cbb-404">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="96cbb-404">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="96cbb-405">阅读</span><span class="sxs-lookup"><span data-stu-id="96cbb-405">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="96cbb-406">示例</span><span class="sxs-lookup"><span data-stu-id="96cbb-406">Example</span></span>

```javascript
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

#### <a name="itemclass-string"></a><span data-ttu-id="96cbb-407">itemClass :String</span><span class="sxs-lookup"><span data-stu-id="96cbb-407">itemClass :String</span></span>

<span data-ttu-id="96cbb-p114">获取选定项目的 Exchange Web 服务项目类。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="96cbb-p114">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="96cbb-p115">`itemClass` 属性指定所选项目的邮件类别。以下是邮件或约会项目的默认邮件类别。</span><span class="sxs-lookup"><span data-stu-id="96cbb-p115">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

|<span data-ttu-id="96cbb-412">类型</span><span class="sxs-lookup"><span data-stu-id="96cbb-412">Type</span></span>|<span data-ttu-id="96cbb-413">说明</span><span class="sxs-lookup"><span data-stu-id="96cbb-413">Description</span></span>|<span data-ttu-id="96cbb-414">项目类</span><span class="sxs-lookup"><span data-stu-id="96cbb-414">item class</span></span>|
|---|---|---|
|<span data-ttu-id="96cbb-415">约会项目</span><span class="sxs-lookup"><span data-stu-id="96cbb-415">Appointment items</span></span>|<span data-ttu-id="96cbb-416">这些是项目类 `IPM.Appointment` 或 `IPM.Appointment.Occurence` 的日历项目。</span><span class="sxs-lookup"><span data-stu-id="96cbb-416">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurence`.</span></span>|`IPM.Appointment`<br />`IPM.Appointment.Occurence`|
|<span data-ttu-id="96cbb-417">邮件项目</span><span class="sxs-lookup"><span data-stu-id="96cbb-417">Message items</span></span>|<span data-ttu-id="96cbb-418">这些项目包括具有默认邮件类别 `IPM.Note` 的电子邮件，以及使用 `IPM.Schedule.Meeting` 作为基础邮件类别的会议请求、响应和取消。</span><span class="sxs-lookup"><span data-stu-id="96cbb-418">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span>|`IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled`|

<span data-ttu-id="96cbb-419">你可以创建用于扩展默认邮件类别的自定义邮件类别，例如，自定义约会邮件类别 `IPM.Appointment.Contoso`。</span><span class="sxs-lookup"><span data-stu-id="96cbb-419">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="96cbb-420">类型：</span><span class="sxs-lookup"><span data-stu-id="96cbb-420">Type:</span></span>

*   <span data-ttu-id="96cbb-421">String</span><span class="sxs-lookup"><span data-stu-id="96cbb-421">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="96cbb-422">要求</span><span class="sxs-lookup"><span data-stu-id="96cbb-422">Requirements</span></span>

|<span data-ttu-id="96cbb-423">要求</span><span class="sxs-lookup"><span data-stu-id="96cbb-423">Requirement</span></span>|<span data-ttu-id="96cbb-424">值</span><span class="sxs-lookup"><span data-stu-id="96cbb-424">Value</span></span>|
|---|---|
|[<span data-ttu-id="96cbb-425">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="96cbb-425">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="96cbb-426">1.0</span><span class="sxs-lookup"><span data-stu-id="96cbb-426">1.0</span></span>|
|[<span data-ttu-id="96cbb-427">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="96cbb-427">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="96cbb-428">ReadItem</span><span class="sxs-lookup"><span data-stu-id="96cbb-428">ReadItem</span></span>|
|[<span data-ttu-id="96cbb-429">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="96cbb-429">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="96cbb-430">阅读</span><span class="sxs-lookup"><span data-stu-id="96cbb-430">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="96cbb-431">示例</span><span class="sxs-lookup"><span data-stu-id="96cbb-431">Example</span></span>

```javascript
var itemClass = Office.context.mailbox.item.itemClass;
```

#### <a name="nullable-itemid-string"></a><span data-ttu-id="96cbb-432">(nullable) itemId :String</span><span class="sxs-lookup"><span data-stu-id="96cbb-432">(nullable) itemId :String</span></span>

<span data-ttu-id="96cbb-p116">获取当前项目的 Exchange Web 服务项目标识符。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="96cbb-p116">Gets the Exchange Web Services item identifier for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="96cbb-435">`itemId` 属性返回的标识符与 Exchange Web 服务项目标识符相同。</span><span class="sxs-lookup"><span data-stu-id="96cbb-435">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="96cbb-436">`itemId` 属性与 Outlook 条目 ID 或 Outlook REST API 使用的 ID 不同。</span><span class="sxs-lookup"><span data-stu-id="96cbb-436">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="96cbb-437">使用此值进行 REST API 调用前，应使用 [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string) 对它进行转换。</span><span class="sxs-lookup"><span data-stu-id="96cbb-437">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="96cbb-438">有关详细信息，请参阅[从 Outlook 加载项使用 Outlook REST API](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id)。</span><span class="sxs-lookup"><span data-stu-id="96cbb-438">For more details, see [Use the Outlook REST APIs from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

<span data-ttu-id="96cbb-p118">`itemId` 属性在撰写模式下不可用。如果需要项目标识符，[`saveAsync`](#saveasyncoptions-callback) 方法可用于将项目保存到存储，这将在回调函数的 [`AsyncResult.value`](/javascript/api/office/office.asyncresult) 参数中返回项目标识符。</span><span class="sxs-lookup"><span data-stu-id="96cbb-p118">The `itemId` property is not available in compose mode. If an item identifier is required, the [`saveAsync`](#saveasyncoptions-callback) method can be used to save the item to the store, which will return the item identifier in the [`AsyncResult.value`](/javascript/api/office/office.asyncresult) parameter in the callback function.</span></span>

##### <a name="type"></a><span data-ttu-id="96cbb-441">类型:</span><span class="sxs-lookup"><span data-stu-id="96cbb-441">Type:</span></span>

*   <span data-ttu-id="96cbb-442">String</span><span class="sxs-lookup"><span data-stu-id="96cbb-442">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="96cbb-443">要求</span><span class="sxs-lookup"><span data-stu-id="96cbb-443">Requirements</span></span>

|<span data-ttu-id="96cbb-444">要求</span><span class="sxs-lookup"><span data-stu-id="96cbb-444">Requirement</span></span>|<span data-ttu-id="96cbb-445">值</span><span class="sxs-lookup"><span data-stu-id="96cbb-445">Value</span></span>|
|---|---|
|[<span data-ttu-id="96cbb-446">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="96cbb-446">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="96cbb-447">1.0</span><span class="sxs-lookup"><span data-stu-id="96cbb-447">1.0</span></span>|
|[<span data-ttu-id="96cbb-448">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="96cbb-448">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="96cbb-449">ReadItem</span><span class="sxs-lookup"><span data-stu-id="96cbb-449">ReadItem</span></span>|
|[<span data-ttu-id="96cbb-450">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="96cbb-450">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="96cbb-451">阅读</span><span class="sxs-lookup"><span data-stu-id="96cbb-451">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="96cbb-452">示例</span><span class="sxs-lookup"><span data-stu-id="96cbb-452">Example</span></span>

<span data-ttu-id="96cbb-p119">以下代码检查项目标识符是否存在。如果 `itemId` 属性返回 `null` 或 `undefined`，则将项目保存到存储，并从异步结果中获取项目标识符。</span><span class="sxs-lookup"><span data-stu-id="96cbb-p119">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

```javascript
var itemId = Office.context.mailbox.item.itemId;
if (itemId === null || itemId == undefined) {
  Office.context.mailbox.item.saveAsync(function(result){
    itemId = result.value;
  });
}
```

####  <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtype"></a><span data-ttu-id="96cbb-455">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype)</span><span class="sxs-lookup"><span data-stu-id="96cbb-455">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype)</span></span>

<span data-ttu-id="96cbb-456">获取实例表示的项的类型。</span><span class="sxs-lookup"><span data-stu-id="96cbb-456">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="96cbb-457">`itemType` 属性返回其中一个 `ItemType` 枚举值，指示 `item` 对象实例是邮件还是约会。</span><span class="sxs-lookup"><span data-stu-id="96cbb-457">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="96cbb-458">类型：</span><span class="sxs-lookup"><span data-stu-id="96cbb-458">Type:</span></span>

*   [<span data-ttu-id="96cbb-459">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="96cbb-459">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook/office.mailboxenums.itemtype)

##### <a name="requirements"></a><span data-ttu-id="96cbb-460">要求</span><span class="sxs-lookup"><span data-stu-id="96cbb-460">Requirements</span></span>

|<span data-ttu-id="96cbb-461">要求</span><span class="sxs-lookup"><span data-stu-id="96cbb-461">Requirement</span></span>|<span data-ttu-id="96cbb-462">值</span><span class="sxs-lookup"><span data-stu-id="96cbb-462">Value</span></span>|
|---|---|
|[<span data-ttu-id="96cbb-463">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="96cbb-463">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="96cbb-464">1.0</span><span class="sxs-lookup"><span data-stu-id="96cbb-464">1.0</span></span>|
|[<span data-ttu-id="96cbb-465">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="96cbb-465">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="96cbb-466">ReadItem</span><span class="sxs-lookup"><span data-stu-id="96cbb-466">ReadItem</span></span>|
|[<span data-ttu-id="96cbb-467">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="96cbb-467">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="96cbb-468">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="96cbb-468">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="96cbb-469">示例</span><span class="sxs-lookup"><span data-stu-id="96cbb-469">Example</span></span>

```javascript
if (Office.context.mailbox.item.itemType == Office.MailboxEnums.ItemType.Message)
  // do something
else
  // do something else
```

####  <a name="location-stringlocationjavascriptapioutlookofficelocation"></a><span data-ttu-id="96cbb-470">location :String|[Location](/javascript/api/outlook/office.location)</span><span class="sxs-lookup"><span data-stu-id="96cbb-470">location :String|[Location](/javascript/api/outlook/office.location)</span></span>

<span data-ttu-id="96cbb-471">获取或设置约会的位置。</span><span class="sxs-lookup"><span data-stu-id="96cbb-471">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="96cbb-472">阅读模式</span><span class="sxs-lookup"><span data-stu-id="96cbb-472">Read mode</span></span>

<span data-ttu-id="96cbb-473">`location` 属性返回一个包含约会位置的字符串。</span><span class="sxs-lookup"><span data-stu-id="96cbb-473">The `location` property returns a string that contains the location of the appointment.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="96cbb-474">撰写模式</span><span class="sxs-lookup"><span data-stu-id="96cbb-474">Compose mode</span></span>

<span data-ttu-id="96cbb-475">`location` 属性返回一个 `Location` 对象，该对象提供用于获取和设置约会位置的方法。</span><span class="sxs-lookup"><span data-stu-id="96cbb-475">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="96cbb-476">类型：</span><span class="sxs-lookup"><span data-stu-id="96cbb-476">Type:</span></span>

*   <span data-ttu-id="96cbb-477">String | [Location](/javascript/api/outlook/office.location)</span><span class="sxs-lookup"><span data-stu-id="96cbb-477">String | [Location](/javascript/api/outlook/office.location)</span></span>

##### <a name="requirements"></a><span data-ttu-id="96cbb-478">要求</span><span class="sxs-lookup"><span data-stu-id="96cbb-478">Requirements</span></span>

|<span data-ttu-id="96cbb-479">要求</span><span class="sxs-lookup"><span data-stu-id="96cbb-479">Requirement</span></span>|<span data-ttu-id="96cbb-480">值</span><span class="sxs-lookup"><span data-stu-id="96cbb-480">Value</span></span>|
|---|---|
|[<span data-ttu-id="96cbb-481">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="96cbb-481">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="96cbb-482">1.0</span><span class="sxs-lookup"><span data-stu-id="96cbb-482">1.0</span></span>|
|[<span data-ttu-id="96cbb-483">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="96cbb-483">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="96cbb-484">ReadItem</span><span class="sxs-lookup"><span data-stu-id="96cbb-484">ReadItem</span></span>|
|[<span data-ttu-id="96cbb-485">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="96cbb-485">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="96cbb-486">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="96cbb-486">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="96cbb-487">示例</span><span class="sxs-lookup"><span data-stu-id="96cbb-487">Example</span></span>

```javascript
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

#### <a name="normalizedsubject-string"></a><span data-ttu-id="96cbb-488">normalizedSubject :String</span><span class="sxs-lookup"><span data-stu-id="96cbb-488">normalizedSubject :String</span></span>

<span data-ttu-id="96cbb-p120">获取删除了所有前缀（包括 `RE:` 和 `FWD:`）的项目主题。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="96cbb-p120">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="96cbb-p121">normalizedSubject 属性获取包含由电子邮件程序添加的任何标准前缀（如 `RE:` 和 `FW:`）的项目主题。若要获取包含完整前缀的项目主题，请使用 [`subject`](#subject-stringsubjectjavascriptapioutlookofficesubject) 属性。</span><span class="sxs-lookup"><span data-stu-id="96cbb-p121">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubjectjavascriptapioutlookofficesubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="96cbb-493">类型:</span><span class="sxs-lookup"><span data-stu-id="96cbb-493">Type:</span></span>

*   <span data-ttu-id="96cbb-494">String</span><span class="sxs-lookup"><span data-stu-id="96cbb-494">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="96cbb-495">要求</span><span class="sxs-lookup"><span data-stu-id="96cbb-495">Requirements</span></span>

|<span data-ttu-id="96cbb-496">要求</span><span class="sxs-lookup"><span data-stu-id="96cbb-496">Requirement</span></span>|<span data-ttu-id="96cbb-497">值</span><span class="sxs-lookup"><span data-stu-id="96cbb-497">Value</span></span>|
|---|---|
|[<span data-ttu-id="96cbb-498">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="96cbb-498">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="96cbb-499">1.0</span><span class="sxs-lookup"><span data-stu-id="96cbb-499">1.0</span></span>|
|[<span data-ttu-id="96cbb-500">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="96cbb-500">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="96cbb-501">ReadItem</span><span class="sxs-lookup"><span data-stu-id="96cbb-501">ReadItem</span></span>|
|[<span data-ttu-id="96cbb-502">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="96cbb-502">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="96cbb-503">阅读</span><span class="sxs-lookup"><span data-stu-id="96cbb-503">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="96cbb-504">示例</span><span class="sxs-lookup"><span data-stu-id="96cbb-504">Example</span></span>

```javascript
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
```

####  <a name="notificationmessages-notificationmessagesjavascriptapioutlookofficenotificationmessages"></a><span data-ttu-id="96cbb-505">notificationMessages :[NotificationMessages](/javascript/api/outlook/office.notificationmessages)</span><span class="sxs-lookup"><span data-stu-id="96cbb-505">notificationMessages :[NotificationMessages](/javascript/api/outlook/office.notificationmessages)</span></span>

<span data-ttu-id="96cbb-506">获取项目的通知邮件。</span><span class="sxs-lookup"><span data-stu-id="96cbb-506">Gets the notification messages for an item.</span></span>

##### <a name="type"></a><span data-ttu-id="96cbb-507">类型：</span><span class="sxs-lookup"><span data-stu-id="96cbb-507">Type:</span></span>

*   [<span data-ttu-id="96cbb-508">NotificationMessages</span><span class="sxs-lookup"><span data-stu-id="96cbb-508">NotificationMessages</span></span>](/javascript/api/outlook/office.notificationmessages)

##### <a name="requirements"></a><span data-ttu-id="96cbb-509">要求</span><span class="sxs-lookup"><span data-stu-id="96cbb-509">Requirements</span></span>

|<span data-ttu-id="96cbb-510">要求</span><span class="sxs-lookup"><span data-stu-id="96cbb-510">Requirement</span></span>|<span data-ttu-id="96cbb-511">值</span><span class="sxs-lookup"><span data-stu-id="96cbb-511">Value</span></span>|
|---|---|
|[<span data-ttu-id="96cbb-512">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="96cbb-512">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="96cbb-513">1.3</span><span class="sxs-lookup"><span data-stu-id="96cbb-513">1.3</span></span>|
|[<span data-ttu-id="96cbb-514">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="96cbb-514">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="96cbb-515">ReadItem</span><span class="sxs-lookup"><span data-stu-id="96cbb-515">ReadItem</span></span>|
|[<span data-ttu-id="96cbb-516">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="96cbb-516">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="96cbb-517">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="96cbb-517">Compose or read</span></span>|

####  <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="96cbb-518">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="96cbb-518">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="96cbb-519">提供对事件的可选与会者的访问权限。</span><span class="sxs-lookup"><span data-stu-id="96cbb-519">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="96cbb-520">对象的类型和访问级别取决于当前项目的模式。</span><span class="sxs-lookup"><span data-stu-id="96cbb-520">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="96cbb-521">阅读模式</span><span class="sxs-lookup"><span data-stu-id="96cbb-521">Read mode</span></span>

<span data-ttu-id="96cbb-522">`optionalAttendees` 属性返回一个数组，其中包含每个可选与会者的 `EmailAddressDetails` 对象。</span><span class="sxs-lookup"><span data-stu-id="96cbb-522">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="96cbb-523">撰写模式</span><span class="sxs-lookup"><span data-stu-id="96cbb-523">Compose mode</span></span>

<span data-ttu-id="96cbb-524">`optionalAttendees` 属性返回一个 `Recipients` 对象，该对象提供用于获取或更新可选与会者的方法。</span><span class="sxs-lookup"><span data-stu-id="96cbb-524">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="96cbb-525">类型：</span><span class="sxs-lookup"><span data-stu-id="96cbb-525">Type:</span></span>

*   <span data-ttu-id="96cbb-526">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="96cbb-526">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="96cbb-527">要求</span><span class="sxs-lookup"><span data-stu-id="96cbb-527">Requirements</span></span>

|<span data-ttu-id="96cbb-528">要求</span><span class="sxs-lookup"><span data-stu-id="96cbb-528">Requirement</span></span>|<span data-ttu-id="96cbb-529">值</span><span class="sxs-lookup"><span data-stu-id="96cbb-529">Value</span></span>|
|---|---|
|[<span data-ttu-id="96cbb-530">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="96cbb-530">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="96cbb-531">1.0</span><span class="sxs-lookup"><span data-stu-id="96cbb-531">1.0</span></span>|
|[<span data-ttu-id="96cbb-532">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="96cbb-532">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="96cbb-533">ReadItem</span><span class="sxs-lookup"><span data-stu-id="96cbb-533">ReadItem</span></span>|
|[<span data-ttu-id="96cbb-534">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="96cbb-534">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="96cbb-535">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="96cbb-535">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="96cbb-536">示例</span><span class="sxs-lookup"><span data-stu-id="96cbb-536">Example</span></span>

```javascript
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

#### <a name="organizer-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsorganizerjavascriptapioutlookofficeorganizer"></a><span data-ttu-id="96cbb-537">organizer :[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[Organizer](/javascript/api/outlook/office.organizer)</span><span class="sxs-lookup"><span data-stu-id="96cbb-537">organizer :[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[Organizer](/javascript/api/outlook/office.organizer)</span></span>

<span data-ttu-id="96cbb-538">获取指定会议的组织者的电子邮件地址。</span><span class="sxs-lookup"><span data-stu-id="96cbb-538">Gets the email address of the organizer for a specified meeting.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="96cbb-539">阅读模式</span><span class="sxs-lookup"><span data-stu-id="96cbb-539">Read mode</span></span>

<span data-ttu-id="96cbb-540">`organizer` 属性返回 [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) 对象，它表示会议组织者。</span><span class="sxs-lookup"><span data-stu-id="96cbb-540">The `organizer` property returns an [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) object that represents the meeting organizer.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="96cbb-541">撰写模式</span><span class="sxs-lookup"><span data-stu-id="96cbb-541">Compose mode</span></span>

<span data-ttu-id="96cbb-542">`organizer` 属性返回 [Organizer](/javascript/api/outlook/office.organizer) 对象，该对象提供获取组织者值的方法。</span><span class="sxs-lookup"><span data-stu-id="96cbb-542">The `organizer` property returns an [Organizer](/javascript/api/outlook/office.organizer) object that provides a method to get the organizer value.</span></span>

##### <a name="type"></a><span data-ttu-id="96cbb-543">类型：</span><span class="sxs-lookup"><span data-stu-id="96cbb-543">Type:</span></span>

*   <span data-ttu-id="96cbb-544">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [Organizer](/javascript/api/outlook/office.organizer)</span><span class="sxs-lookup"><span data-stu-id="96cbb-544">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [Organizer](/javascript/api/outlook/office.organizer)</span></span>

##### <a name="requirements"></a><span data-ttu-id="96cbb-545">要求</span><span class="sxs-lookup"><span data-stu-id="96cbb-545">Requirements</span></span>

|<span data-ttu-id="96cbb-546">要求</span><span class="sxs-lookup"><span data-stu-id="96cbb-546">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="96cbb-547">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="96cbb-547">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="96cbb-548">1.0</span><span class="sxs-lookup"><span data-stu-id="96cbb-548">1.0</span></span>|<span data-ttu-id="96cbb-549">1.7</span><span class="sxs-lookup"><span data-stu-id="96cbb-549">1.7</span></span>|
|[<span data-ttu-id="96cbb-550">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="96cbb-550">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="96cbb-551">ReadItem</span><span class="sxs-lookup"><span data-stu-id="96cbb-551">ReadItem</span></span>|<span data-ttu-id="96cbb-552">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="96cbb-552">ReadWriteItem</span></span>|
|[<span data-ttu-id="96cbb-553">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="96cbb-553">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="96cbb-554">阅读</span><span class="sxs-lookup"><span data-stu-id="96cbb-554">Read</span></span>|<span data-ttu-id="96cbb-555">撰写</span><span class="sxs-lookup"><span data-stu-id="96cbb-555">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="96cbb-556">示例</span><span class="sxs-lookup"><span data-stu-id="96cbb-556">Example</span></span>

```javascript
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
```

#### <a name="nullable-recurrence-recurrencejavascriptapioutlookofficerecurrence"></a><span data-ttu-id="96cbb-557">(nullable) recurrence :[Recurrence](/javascript/api/outlook/office.recurrence)</span><span class="sxs-lookup"><span data-stu-id="96cbb-557">(nullable) recurrence :[Recurrence](/javascript/api/outlook/office.recurrence)</span></span>

<span data-ttu-id="96cbb-558">获取或设置约会的定期模式。</span><span class="sxs-lookup"><span data-stu-id="96cbb-558">Gets or sets the recurrence pattern of an appointment.</span></span> <span data-ttu-id="96cbb-559">获取或设置会议请求的定期模式。</span><span class="sxs-lookup"><span data-stu-id="96cbb-559">Gets the recurrence pattern of a meeting request.</span></span> <span data-ttu-id="96cbb-560">阅读撰写约会项目的模式。</span><span class="sxs-lookup"><span data-stu-id="96cbb-560">Read and compose modes for appointment items.</span></span> <span data-ttu-id="96cbb-561">会议请求项目的阅读模式。</span><span class="sxs-lookup"><span data-stu-id="96cbb-561">Read mode for meeting request items.</span></span>

<span data-ttu-id="96cbb-562">如果项目是一个系列或系列中的一个实例，则 `recurrence` 属性将返回定期约会的 [recurrence](/javascript/api/outlook/office.recurrence) 对象或会议请求。</span><span class="sxs-lookup"><span data-stu-id="96cbb-562">The `recurrence` property returns a [recurrence](/javascript/api/outlook/office.recurrence) object for recurring appointments or meetings requests if an item is a series or an instance in a series.</span></span> <span data-ttu-id="96cbb-563">针对单个约会和单个约会的会议请求返回 `null`。</span><span class="sxs-lookup"><span data-stu-id="96cbb-563">`null` is returned for single appointments and meeting requests of single appointments.</span></span> <span data-ttu-id="96cbb-564">针对非会议请求的邮件返回 `undefined`。</span><span class="sxs-lookup"><span data-stu-id="96cbb-564">`undefined` is returned for messages that are not meeting requests.</span></span>

> <span data-ttu-id="96cbb-565">注意：会议请求的 `itemClass` 值为 IPM.Schedule.Meeting.Request。</span><span class="sxs-lookup"><span data-stu-id="96cbb-565">Note: Meeting requests have an `itemClass` value of IPM.Schedule.Meeting.Request.</span></span>

> <span data-ttu-id="96cbb-566">注意：如果 recurrence 对象为 `null`，则这表示对象是单个约会或单个约会的会议请求，而不是系列的一部分。</span><span class="sxs-lookup"><span data-stu-id="96cbb-566">Note: If the recurrence object is `null`, this indicates that the object is a single appointment or a meeting request of a single appointment and NOT a part of a series.</span></span>

##### <a name="type"></a><span data-ttu-id="96cbb-567">类型：</span><span class="sxs-lookup"><span data-stu-id="96cbb-567">Type:</span></span>

* [<span data-ttu-id="96cbb-568">Recurrence</span><span class="sxs-lookup"><span data-stu-id="96cbb-568">Recurrence</span></span>](/javascript/api/outlook/office.recurrence)

|<span data-ttu-id="96cbb-569">要求</span><span class="sxs-lookup"><span data-stu-id="96cbb-569">Requirement</span></span>|<span data-ttu-id="96cbb-570">值</span><span class="sxs-lookup"><span data-stu-id="96cbb-570">Value</span></span>|
|---|---|
|[<span data-ttu-id="96cbb-571">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="96cbb-571">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="96cbb-572">1.7</span><span class="sxs-lookup"><span data-stu-id="96cbb-572">1.7</span></span>|
|[<span data-ttu-id="96cbb-573">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="96cbb-573">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="96cbb-574">ReadItem</span><span class="sxs-lookup"><span data-stu-id="96cbb-574">ReadItem</span></span>|
|[<span data-ttu-id="96cbb-575">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="96cbb-575">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="96cbb-576">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="96cbb-576">Compose or read</span></span>|

####  <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="96cbb-577">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="96cbb-577">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="96cbb-578">提供对事件的必需与会者的访问权限。</span><span class="sxs-lookup"><span data-stu-id="96cbb-578">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="96cbb-579">对象的类型和访问级别取决于当前项目的模式。</span><span class="sxs-lookup"><span data-stu-id="96cbb-579">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="96cbb-580">阅读模式</span><span class="sxs-lookup"><span data-stu-id="96cbb-580">Read mode</span></span>

<span data-ttu-id="96cbb-581">`requiredAttendees` 属性返回一个数组，其中包含每个必需与会者的 `EmailAddressDetails` 对象。</span><span class="sxs-lookup"><span data-stu-id="96cbb-581">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="96cbb-582">撰写模式</span><span class="sxs-lookup"><span data-stu-id="96cbb-582">Compose mode</span></span>

<span data-ttu-id="96cbb-583">`requiredAttendees` 属性返回一个 `Recipients` 对象，该对象提供用于获取或更新必需与会者的方法。</span><span class="sxs-lookup"><span data-stu-id="96cbb-583">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="96cbb-584">类型：</span><span class="sxs-lookup"><span data-stu-id="96cbb-584">Type:</span></span>

*   <span data-ttu-id="96cbb-585">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="96cbb-585">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="96cbb-586">要求</span><span class="sxs-lookup"><span data-stu-id="96cbb-586">Requirements</span></span>

|<span data-ttu-id="96cbb-587">要求</span><span class="sxs-lookup"><span data-stu-id="96cbb-587">Requirement</span></span>|<span data-ttu-id="96cbb-588">值</span><span class="sxs-lookup"><span data-stu-id="96cbb-588">Value</span></span>|
|---|---|
|[<span data-ttu-id="96cbb-589">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="96cbb-589">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="96cbb-590">1.0</span><span class="sxs-lookup"><span data-stu-id="96cbb-590">1.0</span></span>|
|[<span data-ttu-id="96cbb-591">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="96cbb-591">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="96cbb-592">ReadItem</span><span class="sxs-lookup"><span data-stu-id="96cbb-592">ReadItem</span></span>|
|[<span data-ttu-id="96cbb-593">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="96cbb-593">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="96cbb-594">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="96cbb-594">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="96cbb-595">示例</span><span class="sxs-lookup"><span data-stu-id="96cbb-595">Example</span></span>

```javascript
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
}
```

#### <a name="sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetails"></a><span data-ttu-id="96cbb-596">sender :[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="96cbb-596">sender :[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)</span></span>

<span data-ttu-id="96cbb-p126">获取电子邮件发件人的电子邮件地址。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="96cbb-p126">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="96cbb-p127">[`from`](#from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsfromjavascriptapioutlookofficefrom) 和 `sender` 属性表示同一个人，邮件由代理人发送的除外。在此情况下，`from` 属性表示代理程序，而 sender 属性表示代理人。</span><span class="sxs-lookup"><span data-stu-id="96cbb-p127">The [`from`](#from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsfromjavascriptapioutlookofficefrom) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="96cbb-601">`sender` 属性中 `EmailAddressDetails` 对象的 `recipientType` 属性为 `undefined`。</span><span class="sxs-lookup"><span data-stu-id="96cbb-601">The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="96cbb-602">类型：</span><span class="sxs-lookup"><span data-stu-id="96cbb-602">Type:</span></span>

*   [<span data-ttu-id="96cbb-603">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="96cbb-603">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="96cbb-604">要求</span><span class="sxs-lookup"><span data-stu-id="96cbb-604">Requirements</span></span>

|<span data-ttu-id="96cbb-605">要求</span><span class="sxs-lookup"><span data-stu-id="96cbb-605">Requirement</span></span>|<span data-ttu-id="96cbb-606">值</span><span class="sxs-lookup"><span data-stu-id="96cbb-606">Value</span></span>|
|---|---|
|[<span data-ttu-id="96cbb-607">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="96cbb-607">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="96cbb-608">1.0</span><span class="sxs-lookup"><span data-stu-id="96cbb-608">1.0</span></span>|
|[<span data-ttu-id="96cbb-609">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="96cbb-609">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="96cbb-610">ReadItem</span><span class="sxs-lookup"><span data-stu-id="96cbb-610">ReadItem</span></span>|
|[<span data-ttu-id="96cbb-611">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="96cbb-611">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="96cbb-612">阅读</span><span class="sxs-lookup"><span data-stu-id="96cbb-612">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="96cbb-613">示例</span><span class="sxs-lookup"><span data-stu-id="96cbb-613">Example</span></span>

```javascript
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
```

#### <a name="nullable-seriesid-string"></a><span data-ttu-id="96cbb-614">(nullable) seriesId :String</span><span class="sxs-lookup"><span data-stu-id="96cbb-614">(nullable) seriesId :String</span></span>

<span data-ttu-id="96cbb-615">获取实例所属的系列的 ID。</span><span class="sxs-lookup"><span data-stu-id="96cbb-615">Gets the id of the series that an instance belongs to.</span></span>

<span data-ttu-id="96cbb-616">在 OWA 和 Outlook 中，`seriesId` 返回此项目所属的父（系列）项目的 Exchange Web 服务 (EWS) ID。</span><span class="sxs-lookup"><span data-stu-id="96cbb-616">In OWA and Outlook, the `seriesId` returns the Exchange Web Services (EWS) ID of the parent (series) item that this item belongs to.</span></span> <span data-ttu-id="96cbb-617">但是，在 iOS 和 Android 中，`seriesId` 返回父项目的其余部分 ID。</span><span class="sxs-lookup"><span data-stu-id="96cbb-617">However, in iOS and Android, the `seriesId` returns the REST ID of the parent item.</span></span>

> [!NOTE]
> <span data-ttu-id="96cbb-618">`seriesId` 属性返回的标识符与 Exchange Web 服务项目标识符相同。</span><span class="sxs-lookup"><span data-stu-id="96cbb-618">The identifier returned by the `seriesId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="96cbb-619">`seriesId` 属性与 Outlook REST API 使用的 Outlook ID 不同。</span><span class="sxs-lookup"><span data-stu-id="96cbb-619">The `seriesId` property is not identical to the Outlook IDs used by the Outlook REST API.</span></span> <span data-ttu-id="96cbb-620">使用此值进行 REST API 调用前，应使用 [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string) 对它进行转换。</span><span class="sxs-lookup"><span data-stu-id="96cbb-620">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="96cbb-621">有关详细信息，请参阅[从 Outlook 加载项使用 Outlook REST API](https://docs.microsoft.com/outlook/add-ins/use-rest-api)。</span><span class="sxs-lookup"><span data-stu-id="96cbb-621">For more details, see [Use the Outlook REST APIs from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/use-rest-api).</span></span>

<span data-ttu-id="96cbb-622">`seriesId` 属性对于没有父项目（如单个约会、系列项目或会议请求）的项目返回 `null`，对于非会议请求的任何其他项目，返回 `undefined`。</span><span class="sxs-lookup"><span data-stu-id="96cbb-622">The `seriesId` property returns `null` for items that do not have parent items such as single appointments, series items, or meeting requests and returns `undefined` for any other items that are not meeting requests.</span></span>

##### <a name="type"></a><span data-ttu-id="96cbb-623">类型：</span><span class="sxs-lookup"><span data-stu-id="96cbb-623">Type:</span></span>

* <span data-ttu-id="96cbb-624">String</span><span class="sxs-lookup"><span data-stu-id="96cbb-624">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="96cbb-625">要求</span><span class="sxs-lookup"><span data-stu-id="96cbb-625">Requirements</span></span>

|<span data-ttu-id="96cbb-626">要求</span><span class="sxs-lookup"><span data-stu-id="96cbb-626">Requirement</span></span>|<span data-ttu-id="96cbb-627">值</span><span class="sxs-lookup"><span data-stu-id="96cbb-627">Value</span></span>|
|---|---|
|[<span data-ttu-id="96cbb-628">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="96cbb-628">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="96cbb-629">1.7</span><span class="sxs-lookup"><span data-stu-id="96cbb-629">1.7</span></span>|
|[<span data-ttu-id="96cbb-630">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="96cbb-630">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="96cbb-631">ReadItem</span><span class="sxs-lookup"><span data-stu-id="96cbb-631">ReadItem</span></span>|
|[<span data-ttu-id="96cbb-632">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="96cbb-632">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="96cbb-633">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="96cbb-633">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="96cbb-634">示例</span><span class="sxs-lookup"><span data-stu-id="96cbb-634">Example</span></span>

```javascript
var seriesId = Office.context.mailbox.item.seriesId;
var isSeries = (seriesId == null);
```

####  <a name="start-datetimejavascriptapioutlookofficetime"></a><span data-ttu-id="96cbb-635">start :Date|[Time](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="96cbb-635">start :Date|[Time](/javascript/api/outlook/office.time)</span></span>

<span data-ttu-id="96cbb-636">获取或设置约会开始的日期和时间。</span><span class="sxs-lookup"><span data-stu-id="96cbb-636">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="96cbb-p130">将 `start` 属性表示为协调世界时 (UTC) 的日期和时间值。可使用 [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttime) 方法将该值转换为客户端的本地日期和时间。</span><span class="sxs-lookup"><span data-stu-id="96cbb-p130">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="96cbb-639">阅读模式</span><span class="sxs-lookup"><span data-stu-id="96cbb-639">Read mode</span></span>

<span data-ttu-id="96cbb-640">`start` 属性返回 `Date` 对象。</span><span class="sxs-lookup"><span data-stu-id="96cbb-640">The `start` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="96cbb-641">撰写模式</span><span class="sxs-lookup"><span data-stu-id="96cbb-641">Compose mode</span></span>

<span data-ttu-id="96cbb-642">`start` 属性返回 `Time` 对象。</span><span class="sxs-lookup"><span data-stu-id="96cbb-642">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="96cbb-643">使用 [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) 方法设置开始时间时，应使用 [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) 方法将客户端的本地时间转换为服务器的 UTC。</span><span class="sxs-lookup"><span data-stu-id="96cbb-643">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="96cbb-644">类型：</span><span class="sxs-lookup"><span data-stu-id="96cbb-644">Type:</span></span>

*   <span data-ttu-id="96cbb-645">Date | [Time](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="96cbb-645">Date | [Time](/javascript/api/outlook/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="96cbb-646">要求</span><span class="sxs-lookup"><span data-stu-id="96cbb-646">Requirements</span></span>

|<span data-ttu-id="96cbb-647">要求</span><span class="sxs-lookup"><span data-stu-id="96cbb-647">Requirement</span></span>|<span data-ttu-id="96cbb-648">值</span><span class="sxs-lookup"><span data-stu-id="96cbb-648">Value</span></span>|
|---|---|
|[<span data-ttu-id="96cbb-649">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="96cbb-649">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="96cbb-650">1.0</span><span class="sxs-lookup"><span data-stu-id="96cbb-650">1.0</span></span>|
|[<span data-ttu-id="96cbb-651">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="96cbb-651">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="96cbb-652">ReadItem</span><span class="sxs-lookup"><span data-stu-id="96cbb-652">ReadItem</span></span>|
|[<span data-ttu-id="96cbb-653">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="96cbb-653">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="96cbb-654">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="96cbb-654">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="96cbb-655">示例</span><span class="sxs-lookup"><span data-stu-id="96cbb-655">Example</span></span>

<span data-ttu-id="96cbb-656">以下示例通过使用 `Time` 对象的 [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) 方法，设置撰写模式下约会的开始时间。</span><span class="sxs-lookup"><span data-stu-id="96cbb-656">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

```javascript
var startTime = new Date("3/14/2015");
var options = {
  // Pass information that can be used
  // in the callback
  asyncContext: {verb:"Set"}
}
Office.context.mailbox.item.start.setAsync(startTime, options, function(result) {
  if (result.error) {
    console.debug(result.error);
  } else {
    // Access the asyncContext that was passed to the setAsync function
    console.debug("Start Time " + result.asyncContext.verb);
  }
});
```

####  <a name="subject-stringsubjectjavascriptapioutlookofficesubject"></a><span data-ttu-id="96cbb-657">subject :String|[Subject](/javascript/api/outlook/office.subject)</span><span class="sxs-lookup"><span data-stu-id="96cbb-657">subject :String|[Subject](/javascript/api/outlook/office.subject)</span></span>

<span data-ttu-id="96cbb-658">获取或设置显示在项目的主题字段中的说明。</span><span class="sxs-lookup"><span data-stu-id="96cbb-658">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="96cbb-659">`subject` 属性获取或设置由电子邮件服务器发送项目时的整个主题。</span><span class="sxs-lookup"><span data-stu-id="96cbb-659">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="96cbb-660">阅读模式</span><span class="sxs-lookup"><span data-stu-id="96cbb-660">Read mode</span></span>

<span data-ttu-id="96cbb-p131">`subject` 属性返回一个字符串。使用 [`normalizedSubject`](#normalizedsubject-string) 属性获取不带任何前导前缀（如 `RE:` 和 `FW:`）的主题。</span><span class="sxs-lookup"><span data-stu-id="96cbb-p131">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

```javascript
var subject = Office.context.mailbox.item.subject;
```

##### <a name="compose-mode"></a><span data-ttu-id="96cbb-663">撰写模式</span><span class="sxs-lookup"><span data-stu-id="96cbb-663">Compose mode</span></span>

<span data-ttu-id="96cbb-664">`subject` 属性返回一个 `Subject` 对象，该对象提供用于获取和设置主题的方法。</span><span class="sxs-lookup"><span data-stu-id="96cbb-664">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```javascript
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="96cbb-665">类型：</span><span class="sxs-lookup"><span data-stu-id="96cbb-665">Type:</span></span>

*   <span data-ttu-id="96cbb-666">String | [Subject](/javascript/api/outlook/office.subject)</span><span class="sxs-lookup"><span data-stu-id="96cbb-666">String | [Subject](/javascript/api/outlook/office.subject)</span></span>

##### <a name="requirements"></a><span data-ttu-id="96cbb-667">要求</span><span class="sxs-lookup"><span data-stu-id="96cbb-667">Requirements</span></span>

|<span data-ttu-id="96cbb-668">要求</span><span class="sxs-lookup"><span data-stu-id="96cbb-668">Requirement</span></span>|<span data-ttu-id="96cbb-669">值</span><span class="sxs-lookup"><span data-stu-id="96cbb-669">Value</span></span>|
|---|---|
|[<span data-ttu-id="96cbb-670">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="96cbb-670">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="96cbb-671">1.0</span><span class="sxs-lookup"><span data-stu-id="96cbb-671">1.0</span></span>|
|[<span data-ttu-id="96cbb-672">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="96cbb-672">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="96cbb-673">ReadItem</span><span class="sxs-lookup"><span data-stu-id="96cbb-673">ReadItem</span></span>|
|[<span data-ttu-id="96cbb-674">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="96cbb-674">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="96cbb-675">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="96cbb-675">Compose or read</span></span>|

####  <a name="to-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="96cbb-676">to :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="96cbb-676">to :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="96cbb-677">提供对邮件的“**收件人**”行上的收件人的访问权限。</span><span class="sxs-lookup"><span data-stu-id="96cbb-677">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="96cbb-678">对象的类型和访问级别取决于当前项目的模式。</span><span class="sxs-lookup"><span data-stu-id="96cbb-678">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="96cbb-679">阅读模式</span><span class="sxs-lookup"><span data-stu-id="96cbb-679">Read mode</span></span>

<span data-ttu-id="96cbb-p133">`to` 属性返回包含邮件的**收件人**行上所列的每个收件人的 `EmailAddressDetails` 对象的数组。集合上限为 100 个成员。</span><span class="sxs-lookup"><span data-stu-id="96cbb-p133">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="96cbb-682">撰写模式</span><span class="sxs-lookup"><span data-stu-id="96cbb-682">Compose mode</span></span>

<span data-ttu-id="96cbb-683">`to` 属性返回一个 `Recipients` 对象，该对象提供用于获取或更新邮件的“**收件人**”行上收件人的方法。</span><span class="sxs-lookup"><span data-stu-id="96cbb-683">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="96cbb-684">类型：</span><span class="sxs-lookup"><span data-stu-id="96cbb-684">Type:</span></span>

*   <span data-ttu-id="96cbb-685">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="96cbb-685">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="96cbb-686">要求</span><span class="sxs-lookup"><span data-stu-id="96cbb-686">Requirements</span></span>

|<span data-ttu-id="96cbb-687">要求</span><span class="sxs-lookup"><span data-stu-id="96cbb-687">Requirement</span></span>|<span data-ttu-id="96cbb-688">值</span><span class="sxs-lookup"><span data-stu-id="96cbb-688">Value</span></span>|
|---|---|
|[<span data-ttu-id="96cbb-689">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="96cbb-689">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="96cbb-690">1.0</span><span class="sxs-lookup"><span data-stu-id="96cbb-690">1.0</span></span>|
|[<span data-ttu-id="96cbb-691">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="96cbb-691">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="96cbb-692">ReadItem</span><span class="sxs-lookup"><span data-stu-id="96cbb-692">ReadItem</span></span>|
|[<span data-ttu-id="96cbb-693">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="96cbb-693">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="96cbb-694">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="96cbb-694">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="96cbb-695">示例</span><span class="sxs-lookup"><span data-stu-id="96cbb-695">Example</span></span>

```javascript
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

### <a name="methods"></a><span data-ttu-id="96cbb-696">方法</span><span class="sxs-lookup"><span data-stu-id="96cbb-696">Methods</span></span>

####  <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="96cbb-697">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="96cbb-697">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="96cbb-698">将文件作为附件添加到邮件或约会。</span><span class="sxs-lookup"><span data-stu-id="96cbb-698">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="96cbb-699">`addFileAttachmentAsync` 方法在指定的 URI 上载文件并将其附加到撰写窗体中的项目。</span><span class="sxs-lookup"><span data-stu-id="96cbb-699">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="96cbb-700">你随后可以将该标识符与 [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) 方法一同使用，以删除同一个会话中的附件。</span><span class="sxs-lookup"><span data-stu-id="96cbb-700">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="96cbb-701">参数：</span><span class="sxs-lookup"><span data-stu-id="96cbb-701">Parameters:</span></span>
|<span data-ttu-id="96cbb-702">名称</span><span class="sxs-lookup"><span data-stu-id="96cbb-702">Name</span></span>|<span data-ttu-id="96cbb-703">类型</span><span class="sxs-lookup"><span data-stu-id="96cbb-703">Type</span></span>|<span data-ttu-id="96cbb-704">属性</span><span class="sxs-lookup"><span data-stu-id="96cbb-704">Attributes</span></span>|<span data-ttu-id="96cbb-705">说明</span><span class="sxs-lookup"><span data-stu-id="96cbb-705">Description</span></span>|
|---|---|---|---|
|`uri`|<span data-ttu-id="96cbb-706">字符串</span><span class="sxs-lookup"><span data-stu-id="96cbb-706">String</span></span>||<span data-ttu-id="96cbb-p134">提供附加到邮件或约会的文件的位置的 URI。最大长度为 2048 个字符。</span><span class="sxs-lookup"><span data-stu-id="96cbb-p134">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`|<span data-ttu-id="96cbb-709">字符串</span><span class="sxs-lookup"><span data-stu-id="96cbb-709">String</span></span>||<span data-ttu-id="96cbb-p135">在附件上载过程中显示的附件名称。最大长度为 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="96cbb-p135">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="96cbb-712">对象</span><span class="sxs-lookup"><span data-stu-id="96cbb-712">Object</span></span>|<span data-ttu-id="96cbb-713">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="96cbb-713">&lt;optional&gt;</span></span>|<span data-ttu-id="96cbb-714">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="96cbb-714">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="96cbb-715">对象</span><span class="sxs-lookup"><span data-stu-id="96cbb-715">Object</span></span>|<span data-ttu-id="96cbb-716">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="96cbb-716">&lt;optional&gt;</span></span>|<span data-ttu-id="96cbb-717">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="96cbb-717">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.isInline`|<span data-ttu-id="96cbb-718">布尔值</span><span class="sxs-lookup"><span data-stu-id="96cbb-718">Boolean</span></span>|<span data-ttu-id="96cbb-719">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="96cbb-719">&lt;optional&gt;</span></span>|<span data-ttu-id="96cbb-720">如果为 `true`，则表示附件将在邮件正文中内联显示，并且不应显示在附件列表中。</span><span class="sxs-lookup"><span data-stu-id="96cbb-720">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`callback`|<span data-ttu-id="96cbb-721">函数</span><span class="sxs-lookup"><span data-stu-id="96cbb-721">function</span></span>|<span data-ttu-id="96cbb-722">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="96cbb-722">&lt;optional&gt;</span></span>|<span data-ttu-id="96cbb-723">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="96cbb-723">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="96cbb-724">如果成功，附件标识符将在 `asyncResult.value` 属性中提供。</span><span class="sxs-lookup"><span data-stu-id="96cbb-724">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="96cbb-725">如果上传附件失败，`asyncResult` 对象将包含一个提供错误说明的 `Error` 对象。</span><span class="sxs-lookup"><span data-stu-id="96cbb-725">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="96cbb-726">错误</span><span class="sxs-lookup"><span data-stu-id="96cbb-726">Errors</span></span>

|<span data-ttu-id="96cbb-727">错误代码</span><span class="sxs-lookup"><span data-stu-id="96cbb-727">Error code</span></span>|<span data-ttu-id="96cbb-728">说明</span><span class="sxs-lookup"><span data-stu-id="96cbb-728">Description</span></span>|
|------------|-------------|
|`AttachmentSizeExceeded`|<span data-ttu-id="96cbb-729">附件大小超过了允许的大小。</span><span class="sxs-lookup"><span data-stu-id="96cbb-729">The attachment is larger than allowed.</span></span>|
|`FileTypeNotSupported`|<span data-ttu-id="96cbb-730">该附件的扩展名不是允许的扩展名。</span><span class="sxs-lookup"><span data-stu-id="96cbb-730">The attachment has an extension that is not allowed.</span></span>|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="96cbb-731">邮件或约会具有的附件过多。</span><span class="sxs-lookup"><span data-stu-id="96cbb-731">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="96cbb-732">要求</span><span class="sxs-lookup"><span data-stu-id="96cbb-732">Requirements</span></span>

|<span data-ttu-id="96cbb-733">要求</span><span class="sxs-lookup"><span data-stu-id="96cbb-733">Requirement</span></span>|<span data-ttu-id="96cbb-734">值</span><span class="sxs-lookup"><span data-stu-id="96cbb-734">Value</span></span>|
|---|---|
|[<span data-ttu-id="96cbb-735">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="96cbb-735">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="96cbb-736">1.1</span><span class="sxs-lookup"><span data-stu-id="96cbb-736">1.1</span></span>|
|[<span data-ttu-id="96cbb-737">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="96cbb-737">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="96cbb-738">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="96cbb-738">ReadWriteItem</span></span>|
|[<span data-ttu-id="96cbb-739">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="96cbb-739">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="96cbb-740">撰写</span><span class="sxs-lookup"><span data-stu-id="96cbb-740">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="96cbb-741">示例</span><span class="sxs-lookup"><span data-stu-id="96cbb-741">Examples</span></span>

```js
function callback(result) {
  if (result.error) {
    showMessage(result.error);
  } else {
    showMessage("Attachment added");
  }
}

function addAttachment() {
  // The values in asyncContext can be accessed in the callback
  var options = { 'asyncContext': { var1: 1, var2: 2 } };

  var attachmentURL = "https://contoso.com/rtm/icon.png";
  Office.context.mailbox.item.addFileAttachmentAsync(attachmentURL, attachmentURL, options, callback);
}
```

<span data-ttu-id="96cbb-742">以下示例将图像文件添加为内联附件，并在邮件正文中引用该附件。</span><span class="sxs-lookup"><span data-stu-id="96cbb-742">The following example adds an image file as an inline attachment and references the attachment in the message body.</span></span>

```js
Office.context.mailbox.item.addFileAttachmentAsync
(
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
        
      }
    );
  }
);
```

#### <a name="addfileattachmentfrombase64asyncbase64file-attachmentname-options-callback"></a><span data-ttu-id="96cbb-743">addFileAttachmentFromBase64Async(base64File, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="96cbb-743">addFileAttachmentFromBase64Async(base64File, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="96cbb-744">将 base64 编码中的文件作为附件添加到消息或约会。</span><span class="sxs-lookup"><span data-stu-id="96cbb-744">Adds a file from the base64 encoding to a message or appointment as an attachment.</span></span>

<span data-ttu-id="96cbb-745">`addFileAttachmentFromBase64Async` 方法从 base64 编码上传文件，并将其附加到撰写表单中的项目。</span><span class="sxs-lookup"><span data-stu-id="96cbb-745">The `addFileAttachmentFromBase64Async` method uploads the file from the base64 encoding and attaches it to the item in the compose form.</span></span> <span data-ttu-id="96cbb-746">此方法返回 AsyncResult.value 对象中的附件标识符。</span><span class="sxs-lookup"><span data-stu-id="96cbb-746">This method returns the attachment identifier in the AsyncResult.value object.</span></span>

<span data-ttu-id="96cbb-747">随后可以将该标识符与 [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) 方法一同使用，以删除同一个会话中的附件。</span><span class="sxs-lookup"><span data-stu-id="96cbb-747">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="96cbb-748">参数：</span><span class="sxs-lookup"><span data-stu-id="96cbb-748">Parameters:</span></span>
|<span data-ttu-id="96cbb-749">名称</span><span class="sxs-lookup"><span data-stu-id="96cbb-749">Name</span></span>|<span data-ttu-id="96cbb-750">类型</span><span class="sxs-lookup"><span data-stu-id="96cbb-750">Type</span></span>|<span data-ttu-id="96cbb-751">属性</span><span class="sxs-lookup"><span data-stu-id="96cbb-751">Attributes</span></span>|<span data-ttu-id="96cbb-752">说明</span><span class="sxs-lookup"><span data-stu-id="96cbb-752">Description</span></span>|
|---|---|---|---|
|`base64File`|<span data-ttu-id="96cbb-753">字符串</span><span class="sxs-lookup"><span data-stu-id="96cbb-753">String</span></span>||<span data-ttu-id="96cbb-754">要添加到电子邮件或事件的图像或文件的 base64 编码内容。</span><span class="sxs-lookup"><span data-stu-id="96cbb-754">The base64 encoded content of an image or file to be added to an email or event.</span></span>|
|`attachmentName`|<span data-ttu-id="96cbb-755">字符串</span><span class="sxs-lookup"><span data-stu-id="96cbb-755">String</span></span>||<span data-ttu-id="96cbb-p137">在附件上载过程中显示的附件名称。最大长度为 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="96cbb-p137">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="96cbb-758">对象</span><span class="sxs-lookup"><span data-stu-id="96cbb-758">Object</span></span>|<span data-ttu-id="96cbb-759">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="96cbb-759">&lt;optional&gt;</span></span>|<span data-ttu-id="96cbb-760">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="96cbb-760">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="96cbb-761">对象</span><span class="sxs-lookup"><span data-stu-id="96cbb-761">Object</span></span>|<span data-ttu-id="96cbb-762">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="96cbb-762">&lt;optional&gt;</span></span>|<span data-ttu-id="96cbb-763">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="96cbb-763">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.isInline`|<span data-ttu-id="96cbb-764">布尔值</span><span class="sxs-lookup"><span data-stu-id="96cbb-764">Boolean</span></span>|<span data-ttu-id="96cbb-765">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="96cbb-765">&lt;optional&gt;</span></span>|<span data-ttu-id="96cbb-766">如果为 `true`，则表示附件将在邮件正文中内联显示，并且不应显示在附件列表中。</span><span class="sxs-lookup"><span data-stu-id="96cbb-766">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`callback`|<span data-ttu-id="96cbb-767">函数</span><span class="sxs-lookup"><span data-stu-id="96cbb-767">function</span></span>|<span data-ttu-id="96cbb-768">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="96cbb-768">&lt;optional&gt;</span></span>|<span data-ttu-id="96cbb-769">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="96cbb-769">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="96cbb-770">如果成功，附件标识符将在 `asyncResult.value` 属性中提供。</span><span class="sxs-lookup"><span data-stu-id="96cbb-770">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="96cbb-771">如果上传附件失败，`asyncResult` 对象将包含一个提供错误说明的 `Error` 对象。</span><span class="sxs-lookup"><span data-stu-id="96cbb-771">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="96cbb-772">错误</span><span class="sxs-lookup"><span data-stu-id="96cbb-772">Errors</span></span>

|<span data-ttu-id="96cbb-773">错误代码</span><span class="sxs-lookup"><span data-stu-id="96cbb-773">Error code</span></span>|<span data-ttu-id="96cbb-774">说明</span><span class="sxs-lookup"><span data-stu-id="96cbb-774">Description</span></span>|
|------------|-------------|
|`AttachmentSizeExceeded`|<span data-ttu-id="96cbb-775">附件大小超过了允许的大小。</span><span class="sxs-lookup"><span data-stu-id="96cbb-775">The attachment is larger than allowed.</span></span>|
|`FileTypeNotSupported`|<span data-ttu-id="96cbb-776">该附件的扩展名不是允许的扩展名。</span><span class="sxs-lookup"><span data-stu-id="96cbb-776">The attachment has an extension that is not allowed.</span></span>|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="96cbb-777">邮件或约会具有的附件过多。</span><span class="sxs-lookup"><span data-stu-id="96cbb-777">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="96cbb-778">要求</span><span class="sxs-lookup"><span data-stu-id="96cbb-778">Requirements</span></span>

|<span data-ttu-id="96cbb-779">要求</span><span class="sxs-lookup"><span data-stu-id="96cbb-779">Requirement</span></span>|<span data-ttu-id="96cbb-780">值</span><span class="sxs-lookup"><span data-stu-id="96cbb-780">Value</span></span>|
|---|---|
|[<span data-ttu-id="96cbb-781">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="96cbb-781">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="96cbb-782">预览</span><span class="sxs-lookup"><span data-stu-id="96cbb-782">Preview</span></span>|
|[<span data-ttu-id="96cbb-783">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="96cbb-783">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="96cbb-784">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="96cbb-784">ReadWriteItem</span></span>|
|[<span data-ttu-id="96cbb-785">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="96cbb-785">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="96cbb-786">撰写</span><span class="sxs-lookup"><span data-stu-id="96cbb-786">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="96cbb-787">示例</span><span class="sxs-lookup"><span data-stu-id="96cbb-787">Examples</span></span>

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
      }
    );
  }
);
```

####  <a name="addhandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="96cbb-788">addHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="96cbb-788">addHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="96cbb-789">添加支持事件的事件处理程序。</span><span class="sxs-lookup"><span data-stu-id="96cbb-789">Adds an event handler for a supported event.</span></span>

<span data-ttu-id="96cbb-790">当前，支持的事件类型是 `Office.EventType.AttachmentsChanged`、`Office.EventType.AppointmentTimeChanged`、`Office.EventType.RecipientsChanged` 和 `Office.EventType.RecurrenceChanged`</span><span class="sxs-lookup"><span data-stu-id="96cbb-790">Currently the supported event types are `Office.EventType.AttachmentsChanged`, `Office.EventType.AppointmentTimeChanged`, `Office.EventType.RecipientsChanged`, and `Office.EventType.RecurrenceChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="96cbb-791">参数：</span><span class="sxs-lookup"><span data-stu-id="96cbb-791">Parameters:</span></span>

| <span data-ttu-id="96cbb-792">名称</span><span class="sxs-lookup"><span data-stu-id="96cbb-792">Name</span></span> | <span data-ttu-id="96cbb-793">类型</span><span class="sxs-lookup"><span data-stu-id="96cbb-793">Type</span></span> | <span data-ttu-id="96cbb-794">属性</span><span class="sxs-lookup"><span data-stu-id="96cbb-794">Attributes</span></span> | <span data-ttu-id="96cbb-795">说明</span><span class="sxs-lookup"><span data-stu-id="96cbb-795">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="96cbb-796">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="96cbb-796">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="96cbb-797">应调用处理程序的事件。</span><span class="sxs-lookup"><span data-stu-id="96cbb-797">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="96cbb-798">函数</span><span class="sxs-lookup"><span data-stu-id="96cbb-798">Function</span></span> || <span data-ttu-id="96cbb-p138">用于处理事件的函数。此函数必须接受一个参数，即对象文本。参数上的 `type` 属性将匹配传递给 `addHandlerAsync` 的 `eventType` 参数。</span><span class="sxs-lookup"><span data-stu-id="96cbb-p138">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `addHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="96cbb-802">Object</span><span class="sxs-lookup"><span data-stu-id="96cbb-802">Object</span></span> | <span data-ttu-id="96cbb-803">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="96cbb-803">&lt;optional&gt;</span></span> | <span data-ttu-id="96cbb-804">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="96cbb-804">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="96cbb-805">对象</span><span class="sxs-lookup"><span data-stu-id="96cbb-805">Object</span></span> | <span data-ttu-id="96cbb-806">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="96cbb-806">&lt;optional&gt;</span></span> | <span data-ttu-id="96cbb-807">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="96cbb-807">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="96cbb-808">函数</span><span class="sxs-lookup"><span data-stu-id="96cbb-808">function</span></span>| <span data-ttu-id="96cbb-809">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="96cbb-809">&lt;optional&gt;</span></span>|<span data-ttu-id="96cbb-810">方法完成后，使用单个参数 `callback`（一个 [`asyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `AsyncResult` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="96cbb-810">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="96cbb-811">Requirements</span><span class="sxs-lookup"><span data-stu-id="96cbb-811">Requirements</span></span>

|<span data-ttu-id="96cbb-812">要求</span><span class="sxs-lookup"><span data-stu-id="96cbb-812">Requirement</span></span>| <span data-ttu-id="96cbb-813">值</span><span class="sxs-lookup"><span data-stu-id="96cbb-813">Value</span></span>|
|---|---|
|[<span data-ttu-id="96cbb-814">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="96cbb-814">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="96cbb-815">1.7</span><span class="sxs-lookup"><span data-stu-id="96cbb-815">1.7</span></span> |
|[<span data-ttu-id="96cbb-816">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="96cbb-816">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="96cbb-817">ReadItem</span><span class="sxs-lookup"><span data-stu-id="96cbb-817">ReadItem</span></span> |
|[<span data-ttu-id="96cbb-818">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="96cbb-818">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="96cbb-819">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="96cbb-819">Compose or read</span></span> |

####  <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="96cbb-820">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="96cbb-820">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="96cbb-821">将 Exchange 项目（如邮件）作为附件添加到邮件或约会。</span><span class="sxs-lookup"><span data-stu-id="96cbb-821">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="96cbb-p139">`addItemAttachmentAsync` 方法将包含指定 Exchange 标识符的项目附加到撰写窗体中的项目。如果指定一个回调方法，此方法使用 `asyncResult` 参数调用，该参数包含一个附件标识符或代码，指示附加项目过程中出现的任何错误。可以使用 `options` 参数将状态信息传递给回调方法（如果需要）。</span><span class="sxs-lookup"><span data-stu-id="96cbb-p139">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="96cbb-825">随后可以将该标识符与 [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) 方法一同使用，以删除同一个会话中的附件。</span><span class="sxs-lookup"><span data-stu-id="96cbb-825">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="96cbb-826">如果 Office 加载项在 Outlook Web App 中运行，则 `addItemAttachmentAsync` 方法可以将项目附加到项目（正在编辑的项目除外）中；然而，不支持也不建议这样做。</span><span class="sxs-lookup"><span data-stu-id="96cbb-826">If your Office Add-in is running in Outlook Web App, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="96cbb-827">参数：</span><span class="sxs-lookup"><span data-stu-id="96cbb-827">Parameters:</span></span>

|<span data-ttu-id="96cbb-828">名称</span><span class="sxs-lookup"><span data-stu-id="96cbb-828">Name</span></span>|<span data-ttu-id="96cbb-829">类型</span><span class="sxs-lookup"><span data-stu-id="96cbb-829">Type</span></span>|<span data-ttu-id="96cbb-830">属性</span><span class="sxs-lookup"><span data-stu-id="96cbb-830">Attributes</span></span>|<span data-ttu-id="96cbb-831">说明</span><span class="sxs-lookup"><span data-stu-id="96cbb-831">Description</span></span>|
|---|---|---|---|
|`itemId`|<span data-ttu-id="96cbb-832">字符串</span><span class="sxs-lookup"><span data-stu-id="96cbb-832">String</span></span>||<span data-ttu-id="96cbb-p140">要附加的项目的 Exchange 标识符。最大长度为 100 个字符。</span><span class="sxs-lookup"><span data-stu-id="96cbb-p140">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`|<span data-ttu-id="96cbb-835">字符串</span><span class="sxs-lookup"><span data-stu-id="96cbb-835">String</span></span>||<span data-ttu-id="96cbb-p141">要附加的项目的主题。最大长度为 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="96cbb-p141">The sujbect of the item to be attached. The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="96cbb-838">对象</span><span class="sxs-lookup"><span data-stu-id="96cbb-838">Object</span></span>|<span data-ttu-id="96cbb-839">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="96cbb-839">&lt;optional&gt;</span></span>|<span data-ttu-id="96cbb-840">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="96cbb-840">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="96cbb-841">对象</span><span class="sxs-lookup"><span data-stu-id="96cbb-841">Object</span></span>|<span data-ttu-id="96cbb-842">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="96cbb-842">&lt;optional&gt;</span></span>|<span data-ttu-id="96cbb-843">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="96cbb-843">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="96cbb-844">函数</span><span class="sxs-lookup"><span data-stu-id="96cbb-844">function</span></span>|<span data-ttu-id="96cbb-845">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="96cbb-845">&lt;optional&gt;</span></span>|<span data-ttu-id="96cbb-846">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="96cbb-846">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="96cbb-847">如果成功，附件标识符将在 `asyncResult.value` 属性中提供。</span><span class="sxs-lookup"><span data-stu-id="96cbb-847">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="96cbb-848">如果添加附件失败，`asyncResult` 对象将包含一个提供错误说明的 `Error` 对象。</span><span class="sxs-lookup"><span data-stu-id="96cbb-848">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="96cbb-849">错误</span><span class="sxs-lookup"><span data-stu-id="96cbb-849">Errors</span></span>

|<span data-ttu-id="96cbb-850">错误代码</span><span class="sxs-lookup"><span data-stu-id="96cbb-850">Error code</span></span>|<span data-ttu-id="96cbb-851">说明</span><span class="sxs-lookup"><span data-stu-id="96cbb-851">Description</span></span>|
|------------|-------------|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="96cbb-852">邮件或约会具有的附件过多。</span><span class="sxs-lookup"><span data-stu-id="96cbb-852">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="96cbb-853">要求</span><span class="sxs-lookup"><span data-stu-id="96cbb-853">Requirements</span></span>

|<span data-ttu-id="96cbb-854">要求</span><span class="sxs-lookup"><span data-stu-id="96cbb-854">Requirement</span></span>|<span data-ttu-id="96cbb-855">值</span><span class="sxs-lookup"><span data-stu-id="96cbb-855">Value</span></span>|
|---|---|
|[<span data-ttu-id="96cbb-856">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="96cbb-856">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="96cbb-857">1.1</span><span class="sxs-lookup"><span data-stu-id="96cbb-857">1.1</span></span>|
|[<span data-ttu-id="96cbb-858">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="96cbb-858">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="96cbb-859">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="96cbb-859">ReadWriteItem</span></span>|
|[<span data-ttu-id="96cbb-860">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="96cbb-860">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="96cbb-861">撰写</span><span class="sxs-lookup"><span data-stu-id="96cbb-861">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="96cbb-862">示例</span><span class="sxs-lookup"><span data-stu-id="96cbb-862">Example</span></span>

<span data-ttu-id="96cbb-863">以下示例将现有的 Outlook 项目添加为名为 `My Attachment` 的附件。</span><span class="sxs-lookup"><span data-stu-id="96cbb-863">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

```javascript
function callback(result) {
  if (result.error) {
    showMessage(result.error);
  } else {
    showMessage("Attachment added");
  }
}

function addAttachment() {
  // EWS ID of item to attach
  // (Shortened for readability)
  var itemId = "AAMkADI1...AAA=";

  // The values in asyncContext can be accessed in the callback
  var options = { 'asyncContext': { var1: 1, var2: 2 } };

  Office.context.mailbox.item.addItemAttachmentAsync(itemId, "My Attachment", options, callback);
}
```

####  <a name="close"></a><span data-ttu-id="96cbb-864">close()</span><span class="sxs-lookup"><span data-stu-id="96cbb-864">close()</span></span>

<span data-ttu-id="96cbb-865">关闭当前正在撰写的项目。</span><span class="sxs-lookup"><span data-stu-id="96cbb-865">Closes the current item that is being composed.</span></span>

<span data-ttu-id="96cbb-p142">
            \`close\` 方法的行为取决于要撰写的项目的当前状态。如果项目具有未保存的更改，客户端将提示用户保存、放弃或取消关闭操作。</span><span class="sxs-lookup"><span data-stu-id="96cbb-p142">The behavior of the `close` method depends on the current state of the item being composed. If the item has unsaved changes, the client prompts the user to save, discard, or cancel the close action.</span></span>

> [!NOTE]
> <span data-ttu-id="96cbb-868">在 Outlook 网页版中，如果该项目是约会并且之前已使用 `saveAsync` 保存，则即使自上次保存项目后未发生任何更改，也会提示用户保存、放弃或取消。</span><span class="sxs-lookup"><span data-stu-id="96cbb-868">In Outlook on the web, if the item is an appointment and it has previously been saved using `saveAsync`, the user is prompted to save, discard, or cancel even if no changes have occurred since the item was last saved.</span></span>

<span data-ttu-id="96cbb-869">在 Outlook 桌面客户端中，如果邮件是内联答复，`close` 方法不起作用。</span><span class="sxs-lookup"><span data-stu-id="96cbb-869">In the Outlook desktop client, if the message is an inline reply, the `close` method has no effect.</span></span>

##### <a name="requirements"></a><span data-ttu-id="96cbb-870">要求</span><span class="sxs-lookup"><span data-stu-id="96cbb-870">Requirements</span></span>

|<span data-ttu-id="96cbb-871">要求</span><span class="sxs-lookup"><span data-stu-id="96cbb-871">Requirement</span></span>|<span data-ttu-id="96cbb-872">值</span><span class="sxs-lookup"><span data-stu-id="96cbb-872">Value</span></span>|
|---|---|
|[<span data-ttu-id="96cbb-873">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="96cbb-873">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="96cbb-874">1.3</span><span class="sxs-lookup"><span data-stu-id="96cbb-874">1.3</span></span>|
|[<span data-ttu-id="96cbb-875">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="96cbb-875">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="96cbb-876">受限</span><span class="sxs-lookup"><span data-stu-id="96cbb-876">Restricted</span></span>|
|[<span data-ttu-id="96cbb-877">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="96cbb-877">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="96cbb-878">撰写</span><span class="sxs-lookup"><span data-stu-id="96cbb-878">Compose</span></span>|

#### <a name="displayreplyallformformdata"></a><span data-ttu-id="96cbb-879">displayReplyAllForm(formData)</span><span class="sxs-lookup"><span data-stu-id="96cbb-879">displayReplyAllForm(formData)</span></span>

<span data-ttu-id="96cbb-880">显示答复窗体，其中包括所选邮件的发件人和所有收件人或所选约会的组织者和所有与会者。</span><span class="sxs-lookup"><span data-stu-id="96cbb-880">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="96cbb-881">在 Outlook for iOS 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="96cbb-881">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="96cbb-882">在 Outlook Web App 中，答复窗体显示为包含 3 列视图的弹出式窗体以及包含 2 列或 1 列视图的弹出式窗体。</span><span class="sxs-lookup"><span data-stu-id="96cbb-882">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="96cbb-883">如果任意字符串参数超出其限制，`displayReplyAllForm` 将引发异常。</span><span class="sxs-lookup"><span data-stu-id="96cbb-883">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="96cbb-p143">当在 `formData.attachments` 参数中指定附件时，Outlook 和 Outlook Web App 尝试下载所有附件并将其附加到答复窗体。如果无法添加任何附件，则在窗体 UI 中显示错误。如果这不可能，则不引发错误消息。</span><span class="sxs-lookup"><span data-stu-id="96cbb-p143">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="96cbb-887">参数：</span><span class="sxs-lookup"><span data-stu-id="96cbb-887">Parameters:</span></span>

|<span data-ttu-id="96cbb-888">名称</span><span class="sxs-lookup"><span data-stu-id="96cbb-888">Name</span></span>|<span data-ttu-id="96cbb-889">类型</span><span class="sxs-lookup"><span data-stu-id="96cbb-889">Type</span></span>|<span data-ttu-id="96cbb-890">属性</span><span class="sxs-lookup"><span data-stu-id="96cbb-890">Attributes</span></span>|<span data-ttu-id="96cbb-891">说明</span><span class="sxs-lookup"><span data-stu-id="96cbb-891">Description</span></span>|
|---|---|---|---|
|`formData`|<span data-ttu-id="96cbb-892">字符串 &#124; 对象</span><span class="sxs-lookup"><span data-stu-id="96cbb-892">String &#124; Object</span></span>||<span data-ttu-id="96cbb-p144">一个包含文本和 HTML 且表示答复窗体的正文的字符串。字符串限制为 32 KB。</span><span class="sxs-lookup"><span data-stu-id="96cbb-p144">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="96cbb-895">**OR**</span><span class="sxs-lookup"><span data-stu-id="96cbb-895">**OR**</span></span><br/><span data-ttu-id="96cbb-p145">包含正文或附件数据和回调函数的对象。对象定义如下。</span><span class="sxs-lookup"><span data-stu-id="96cbb-p145">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span>|
|`formData.htmlBody`|<span data-ttu-id="96cbb-898">字符串</span><span class="sxs-lookup"><span data-stu-id="96cbb-898">String</span></span>|<span data-ttu-id="96cbb-899">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="96cbb-899">&lt;optional&gt;</span></span>|<span data-ttu-id="96cbb-p146">一个包含文本和 HTML 且表示答复窗体的正文的字符串。字符串限制为 32 KB。</span><span class="sxs-lookup"><span data-stu-id="96cbb-p146">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
|`formData.attachments`|<span data-ttu-id="96cbb-902">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="96cbb-902">Array.&lt;Object&gt;</span></span>|<span data-ttu-id="96cbb-903">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="96cbb-903">&lt;optional&gt;</span></span>|<span data-ttu-id="96cbb-904">JSON 对象（是文件或项目附件）的数组。</span><span class="sxs-lookup"><span data-stu-id="96cbb-904">An array of JSON objects that are either file or item attachments.</span></span>|
|`formData.attachments.type`|<span data-ttu-id="96cbb-905">字符串</span><span class="sxs-lookup"><span data-stu-id="96cbb-905">String</span></span>||<span data-ttu-id="96cbb-p147">指示附件的类型。必须是文件附件的 `file` 或项目附件的 `item`。</span><span class="sxs-lookup"><span data-stu-id="96cbb-p147">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span>|
|`formData.attachments.name`|<span data-ttu-id="96cbb-908">字符串</span><span class="sxs-lookup"><span data-stu-id="96cbb-908">String</span></span>||<span data-ttu-id="96cbb-909">一个包含附件的名称的字符串，最多包含 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="96cbb-909">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
|`formData.attachments.url`|<span data-ttu-id="96cbb-910">String</span><span class="sxs-lookup"><span data-stu-id="96cbb-910">String</span></span>||<span data-ttu-id="96cbb-p148">仅在将 `type` 设置为 `file` 时使用。文件的位置的 URI。</span><span class="sxs-lookup"><span data-stu-id="96cbb-p148">Only used if `type` is set to `file`. The URI of the location for the file.</span></span>|
|`formData.attachments.isInline`|<span data-ttu-id="96cbb-913">Boolean</span><span class="sxs-lookup"><span data-stu-id="96cbb-913">Boolean</span></span>||<span data-ttu-id="96cbb-p149">仅在将 `type` 设置为 `file` 时使用。如果为 `true`，则表示附件将在邮件正文中内联显示，并且不应显示在附件列表中。</span><span class="sxs-lookup"><span data-stu-id="96cbb-p149">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`formData.attachments.itemId`|<span data-ttu-id="96cbb-916">String</span><span class="sxs-lookup"><span data-stu-id="96cbb-916">String</span></span>||<span data-ttu-id="96cbb-p150">仅在将 `type` 设置为 `item` 时使用。附件的 EWS 项目 ID。字符串最长为 100 个字符。</span><span class="sxs-lookup"><span data-stu-id="96cbb-p150">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span>|
|`callback`|<span data-ttu-id="96cbb-920">函数</span><span class="sxs-lookup"><span data-stu-id="96cbb-920">function</span></span>|<span data-ttu-id="96cbb-921">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="96cbb-921">&lt;optional&gt;</span></span>|<span data-ttu-id="96cbb-922">方法完成后，使用单个参数 `asyncResult`（一个 [AsyncResult](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="96cbb-922">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="96cbb-923">要求</span><span class="sxs-lookup"><span data-stu-id="96cbb-923">Requirements</span></span>

|<span data-ttu-id="96cbb-924">要求</span><span class="sxs-lookup"><span data-stu-id="96cbb-924">Requirement</span></span>|<span data-ttu-id="96cbb-925">值</span><span class="sxs-lookup"><span data-stu-id="96cbb-925">Value</span></span>|
|---|---|
|[<span data-ttu-id="96cbb-926">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="96cbb-926">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="96cbb-927">1.0</span><span class="sxs-lookup"><span data-stu-id="96cbb-927">1.0</span></span>|
|[<span data-ttu-id="96cbb-928">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="96cbb-928">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="96cbb-929">ReadItem</span><span class="sxs-lookup"><span data-stu-id="96cbb-929">ReadItem</span></span>|
|[<span data-ttu-id="96cbb-930">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="96cbb-930">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="96cbb-931">阅读</span><span class="sxs-lookup"><span data-stu-id="96cbb-931">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="96cbb-932">示例</span><span class="sxs-lookup"><span data-stu-id="96cbb-932">Examples</span></span>

<span data-ttu-id="96cbb-933">以下代码将一个字符串传递到 `displayReplyAllForm` 函数。</span><span class="sxs-lookup"><span data-stu-id="96cbb-933">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="96cbb-934">使用空白正文答复。</span><span class="sxs-lookup"><span data-stu-id="96cbb-934">Reply with an empty body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="96cbb-935">仅使用正文答复。</span><span class="sxs-lookup"><span data-stu-id="96cbb-935">Reply with just a body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="96cbb-936">使用正文和文件附件答复。</span><span class="sxs-lookup"><span data-stu-id="96cbb-936">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="96cbb-937">使用正文和项目附件答复。</span><span class="sxs-lookup"><span data-stu-id="96cbb-937">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="96cbb-938">使用正文、文件附件、项目附件和回调答复。</span><span class="sxs-lookup"><span data-stu-id="96cbb-938">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="displayreplyformformdata"></a><span data-ttu-id="96cbb-939">displayReplyForm(formData)</span><span class="sxs-lookup"><span data-stu-id="96cbb-939">displayReplyForm(formData)</span></span>

<span data-ttu-id="96cbb-940">显示答复窗体，其中仅包括所选邮件的发件人或所选约会的组织者。</span><span class="sxs-lookup"><span data-stu-id="96cbb-940">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="96cbb-941">在 Outlook for iOS 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="96cbb-941">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="96cbb-942">在 Outlook Web App 中，答复窗体显示为包含 3 列视图的弹出式窗体以及包含 2 列或 1 列视图的弹出式窗体。</span><span class="sxs-lookup"><span data-stu-id="96cbb-942">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="96cbb-943">如果任意字符串参数超出其限制，`displayReplyForm` 将引发异常。</span><span class="sxs-lookup"><span data-stu-id="96cbb-943">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="96cbb-p151">当在 `formData.attachments` 参数中指定附件时，Outlook 和 Outlook Web App 尝试下载所有附件并将其附加到答复窗体。如果无法添加任何附件，则在窗体 UI 中显示错误。如果这不可能，则不引发错误消息。</span><span class="sxs-lookup"><span data-stu-id="96cbb-p151">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="96cbb-947">参数：</span><span class="sxs-lookup"><span data-stu-id="96cbb-947">Parameters:</span></span>

|<span data-ttu-id="96cbb-948">名称</span><span class="sxs-lookup"><span data-stu-id="96cbb-948">Name</span></span>|<span data-ttu-id="96cbb-949">类型</span><span class="sxs-lookup"><span data-stu-id="96cbb-949">Type</span></span>|<span data-ttu-id="96cbb-950">属性</span><span class="sxs-lookup"><span data-stu-id="96cbb-950">Attributes</span></span>|<span data-ttu-id="96cbb-951">说明</span><span class="sxs-lookup"><span data-stu-id="96cbb-951">Description</span></span>|
|---|---|---|---|
|`formData`|<span data-ttu-id="96cbb-952">字符串 &#124; 对象</span><span class="sxs-lookup"><span data-stu-id="96cbb-952">String &#124; Object</span></span>||<span data-ttu-id="96cbb-p152">一个包含文本和 HTML 且表示答复窗体的正文的字符串。字符串限制为 32 KB。</span><span class="sxs-lookup"><span data-stu-id="96cbb-p152">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="96cbb-955">**OR**</span><span class="sxs-lookup"><span data-stu-id="96cbb-955">**OR**</span></span><br/><span data-ttu-id="96cbb-p153">包含正文或附件数据和回调函数的对象。对象定义如下。</span><span class="sxs-lookup"><span data-stu-id="96cbb-p153">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span>|
|`formData.htmlBody`|<span data-ttu-id="96cbb-958">字符串</span><span class="sxs-lookup"><span data-stu-id="96cbb-958">String</span></span>|<span data-ttu-id="96cbb-959">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="96cbb-959">&lt;optional&gt;</span></span>|<span data-ttu-id="96cbb-p154">一个包含文本和 HTML 且表示答复窗体的正文的字符串。字符串限制为 32 KB。</span><span class="sxs-lookup"><span data-stu-id="96cbb-p154">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
|`formData.attachments`|<span data-ttu-id="96cbb-962">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="96cbb-962">Array.&lt;Object&gt;</span></span>|<span data-ttu-id="96cbb-963">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="96cbb-963">&lt;optional&gt;</span></span>|<span data-ttu-id="96cbb-964">JSON 对象（是文件或项目附件）的数组。</span><span class="sxs-lookup"><span data-stu-id="96cbb-964">An array of JSON objects that are either file or item attachments.</span></span>|
|`formData.attachments.type`|<span data-ttu-id="96cbb-965">字符串</span><span class="sxs-lookup"><span data-stu-id="96cbb-965">String</span></span>||<span data-ttu-id="96cbb-p155">指示附件的类型。必须是文件附件的 `file` 或项目附件的 `item`。</span><span class="sxs-lookup"><span data-stu-id="96cbb-p155">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span>|
|`formData.attachments.name`|<span data-ttu-id="96cbb-968">字符串</span><span class="sxs-lookup"><span data-stu-id="96cbb-968">String</span></span>||<span data-ttu-id="96cbb-969">一个包含附件的名称的字符串，最多包含 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="96cbb-969">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
|`formData.attachments.url`|<span data-ttu-id="96cbb-970">String</span><span class="sxs-lookup"><span data-stu-id="96cbb-970">String</span></span>||<span data-ttu-id="96cbb-p156">仅在将 `type` 设置为 `file` 时使用。文件的位置的 URI。</span><span class="sxs-lookup"><span data-stu-id="96cbb-p156">Only used if `type` is set to `file`. The URI of the location for the file.</span></span>|
|`formData.attachments.isInline`|<span data-ttu-id="96cbb-973">Boolean</span><span class="sxs-lookup"><span data-stu-id="96cbb-973">Boolean</span></span>||<span data-ttu-id="96cbb-p157">仅在将 `type` 设置为 `file` 时使用。如果为 `true`，则表示附件将在邮件正文中内联显示，并且不应显示在附件列表中。</span><span class="sxs-lookup"><span data-stu-id="96cbb-p157">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`formData.attachments.itemId`|<span data-ttu-id="96cbb-976">String</span><span class="sxs-lookup"><span data-stu-id="96cbb-976">String</span></span>||<span data-ttu-id="96cbb-p158">仅在将 `type` 设置为 `item` 时使用。附件的 EWS 项目 ID。字符串最长为 100 个字符。</span><span class="sxs-lookup"><span data-stu-id="96cbb-p158">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span>|
|`callback`|<span data-ttu-id="96cbb-980">函数</span><span class="sxs-lookup"><span data-stu-id="96cbb-980">function</span></span>|<span data-ttu-id="96cbb-981">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="96cbb-981">&lt;optional&gt;</span></span>|<span data-ttu-id="96cbb-982">方法完成后，使用单个参数 `asyncResult`（一个 [AsyncResult](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="96cbb-982">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="96cbb-983">要求</span><span class="sxs-lookup"><span data-stu-id="96cbb-983">Requirements</span></span>

|<span data-ttu-id="96cbb-984">要求</span><span class="sxs-lookup"><span data-stu-id="96cbb-984">Requirement</span></span>|<span data-ttu-id="96cbb-985">值</span><span class="sxs-lookup"><span data-stu-id="96cbb-985">Value</span></span>|
|---|---|
|[<span data-ttu-id="96cbb-986">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="96cbb-986">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="96cbb-987">1.0</span><span class="sxs-lookup"><span data-stu-id="96cbb-987">1.0</span></span>|
|[<span data-ttu-id="96cbb-988">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="96cbb-988">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="96cbb-989">ReadItem</span><span class="sxs-lookup"><span data-stu-id="96cbb-989">ReadItem</span></span>|
|[<span data-ttu-id="96cbb-990">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="96cbb-990">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="96cbb-991">阅读</span><span class="sxs-lookup"><span data-stu-id="96cbb-991">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="96cbb-992">示例</span><span class="sxs-lookup"><span data-stu-id="96cbb-992">Examples</span></span>

<span data-ttu-id="96cbb-993">以下代码将一个字符串传递到 `displayReplyForm` 函数。</span><span class="sxs-lookup"><span data-stu-id="96cbb-993">The following code passes a string to the `displayReplyForm` function.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="96cbb-994">使用空白正文答复。</span><span class="sxs-lookup"><span data-stu-id="96cbb-994">Reply with an empty body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="96cbb-995">仅使用正文答复。</span><span class="sxs-lookup"><span data-stu-id="96cbb-995">Reply with just a body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="96cbb-996">使用正文和文件附件答复。</span><span class="sxs-lookup"><span data-stu-id="96cbb-996">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="96cbb-997">使用正文和项目附件答复。</span><span class="sxs-lookup"><span data-stu-id="96cbb-997">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="96cbb-998">使用正文、文件附件、项目附件和回调答复。</span><span class="sxs-lookup"><span data-stu-id="96cbb-998">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="getattachmentcontentasyncattachmentid-options-callback--attachmentcontentjavascriptapioutlookofficeattachmentcontent"></a><span data-ttu-id="96cbb-999">getAttachmentContentAsync(attachmentId, [options], callback) → [AttachmentContent](/javascript/api/outlook/office.attachmentcontent)</span><span class="sxs-lookup"><span data-stu-id="96cbb-999">getAttachmentContentAsync(attachmentId, [options], callback) → [AttachmentContent](/javascript/api/outlook/office.attachmentcontent)</span></span>

<span data-ttu-id="96cbb-1000">从消息或约会中获取指定的附件，并将其作为 `AttachmentContent` 对象返回。</span><span class="sxs-lookup"><span data-stu-id="96cbb-1000">Gets the specified attachment from a message or appointment and returns it as an `AttachmentContent` object.</span></span>

<span data-ttu-id="96cbb-1001">`getAttachmentContentAsync` 方法获取项目中带指定标识符的附件。</span><span class="sxs-lookup"><span data-stu-id="96cbb-1001">The `getAttachmentContentAsync` method gets the attachment with the specified identifier from the item.</span></span> <span data-ttu-id="96cbb-1002">作为最佳做法，应使用标识符检索同一会话中的附件，在该会话中，使用 `getAttachmentsAsync` 或 `item.attachments` 调用检索附件 ID。</span><span class="sxs-lookup"><span data-stu-id="96cbb-1002">As a best practice, you should use the identifier to retrieve an attachment in the same session that the attachmentIds were retrieved with the `getAttachmentsAsync` or `item.attachments` call.</span></span> <span data-ttu-id="96cbb-1003">在 Outlook Web App 和适用于设备的 OWA 中，附件标识符只在同一个会话中才有效。</span><span class="sxs-lookup"><span data-stu-id="96cbb-1003">In Outlook Web App and OWA for Devices, the attachment identifier is valid only within the same session.</span></span> <span data-ttu-id="96cbb-1004">当用户关闭应用，或者如果用户开始在内嵌窗体中撰写，则随后弹出的窗体中继续在单独的窗口撰写时，会话即结束。</span><span class="sxs-lookup"><span data-stu-id="96cbb-1004">A session is over when the user closes the app, or if the user starts composing an inline form then subsequently pops out the form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="96cbb-1005">参数：</span><span class="sxs-lookup"><span data-stu-id="96cbb-1005">Parameters:</span></span>

|<span data-ttu-id="96cbb-1006">名称</span><span class="sxs-lookup"><span data-stu-id="96cbb-1006">Name</span></span>|<span data-ttu-id="96cbb-1007">类型</span><span class="sxs-lookup"><span data-stu-id="96cbb-1007">Type</span></span>|<span data-ttu-id="96cbb-1008">属性</span><span class="sxs-lookup"><span data-stu-id="96cbb-1008">Attributes</span></span>|<span data-ttu-id="96cbb-1009">说明</span><span class="sxs-lookup"><span data-stu-id="96cbb-1009">Description</span></span>|
|---|---|---|---|
|`attachmentId`|<span data-ttu-id="96cbb-1010">字符串</span><span class="sxs-lookup"><span data-stu-id="96cbb-1010">String</span></span>||<span data-ttu-id="96cbb-1011">要获取的附件的标识符。</span><span class="sxs-lookup"><span data-stu-id="96cbb-1011">The identifier of the attachment you want to get.</span></span>|
|`options`|<span data-ttu-id="96cbb-1012">对象</span><span class="sxs-lookup"><span data-stu-id="96cbb-1012">Object</span></span>|<span data-ttu-id="96cbb-1013">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="96cbb-1013">&lt;optional&gt;</span></span>|<span data-ttu-id="96cbb-1014">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="96cbb-1014">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="96cbb-1015">对象</span><span class="sxs-lookup"><span data-stu-id="96cbb-1015">Object</span></span>|<span data-ttu-id="96cbb-1016">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="96cbb-1016">&lt;optional&gt;</span></span>|<span data-ttu-id="96cbb-1017">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="96cbb-1017">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="96cbb-1018">函数</span><span class="sxs-lookup"><span data-stu-id="96cbb-1018">function</span></span>|<span data-ttu-id="96cbb-1019">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="96cbb-1019">&lt;optional&gt;</span></span>|<span data-ttu-id="96cbb-1020">方法完成后，使用单个参数 `asyncResult`（一个 [AsyncResult](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="96cbb-1020">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="96cbb-1021">要求</span><span class="sxs-lookup"><span data-stu-id="96cbb-1021">Requirements</span></span>

|<span data-ttu-id="96cbb-1022">要求</span><span class="sxs-lookup"><span data-stu-id="96cbb-1022">Requirement</span></span>|<span data-ttu-id="96cbb-1023">值</span><span class="sxs-lookup"><span data-stu-id="96cbb-1023">Value</span></span>|
|---|---|
|[<span data-ttu-id="96cbb-1024">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="96cbb-1024">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="96cbb-1025">预览</span><span class="sxs-lookup"><span data-stu-id="96cbb-1025">Preview</span></span>|
|[<span data-ttu-id="96cbb-1026">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="96cbb-1026">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="96cbb-1027">ReadItem</span><span class="sxs-lookup"><span data-stu-id="96cbb-1027">ReadItem</span></span>|
|[<span data-ttu-id="96cbb-1028">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="96cbb-1028">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="96cbb-1029">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="96cbb-1029">Compose or read</span></span>|

##### <a name="returns"></a><span data-ttu-id="96cbb-1030">返回：</span><span class="sxs-lookup"><span data-stu-id="96cbb-1030">Returns:</span></span>

<span data-ttu-id="96cbb-1031">类型：[AttachmentContent](/javascript/api/outlook/office.attachmentcontent)</span><span class="sxs-lookup"><span data-stu-id="96cbb-1031">Type: [AttachmentContent](/javascript/api/outlook/office.attachmentcontent)</span></span>

##### <a name="example"></a><span data-ttu-id="96cbb-1032">示例</span><span class="sxs-lookup"><span data-stu-id="96cbb-1032">Example</span></span>

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
    // parse string to be a url, an .eml file, a base64-encoded string, or an .icalendar file
    if (result.format == Office.MailboxEnums.AttachmentContentFormat.Base64) {
        // handle file attachment
    }
    else if (result.format == Office.MailboxEnums.AttachmentContentFormat.Eml) {
        // handle item attachment
    }
    else if (result.format == Office.MailboxEnums.AttachmentContentFormat.ICalendar) {
        // handle .icalender attachment
    }
    else {
        // handle cloud attachment  
    }
}
```

#### <a name="getattachmentsasyncoptions-callback--arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetails"></a><span data-ttu-id="96cbb-1033">getAttachmentsAsync([options], callback) → Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="96cbb-1033">getAttachmentsAsync([options], callback) → Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span></span>

<span data-ttu-id="96cbb-1034">获取项目的附件作为数组。</span><span class="sxs-lookup"><span data-stu-id="96cbb-1034">Gets the item's attachments as an array.</span></span> <span data-ttu-id="96cbb-1035">仅限撰写模式。</span><span class="sxs-lookup"><span data-stu-id="96cbb-1035">Compose mode only.</span></span>

##### <a name="parameters"></a><span data-ttu-id="96cbb-1036">参数：</span><span class="sxs-lookup"><span data-stu-id="96cbb-1036">Parameters:</span></span>

|<span data-ttu-id="96cbb-1037">名称</span><span class="sxs-lookup"><span data-stu-id="96cbb-1037">Name</span></span>|<span data-ttu-id="96cbb-1038">类型</span><span class="sxs-lookup"><span data-stu-id="96cbb-1038">Type</span></span>|<span data-ttu-id="96cbb-1039">属性</span><span class="sxs-lookup"><span data-stu-id="96cbb-1039">Attributes</span></span>|<span data-ttu-id="96cbb-1040">说明</span><span class="sxs-lookup"><span data-stu-id="96cbb-1040">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="96cbb-1041">对象</span><span class="sxs-lookup"><span data-stu-id="96cbb-1041">Object</span></span>|<span data-ttu-id="96cbb-1042">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="96cbb-1042">&lt;optional&gt;</span></span>|<span data-ttu-id="96cbb-1043">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="96cbb-1043">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="96cbb-1044">对象</span><span class="sxs-lookup"><span data-stu-id="96cbb-1044">Object</span></span>|<span data-ttu-id="96cbb-1045">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="96cbb-1045">&lt;optional&gt;</span></span>|<span data-ttu-id="96cbb-1046">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="96cbb-1046">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="96cbb-1047">函数</span><span class="sxs-lookup"><span data-stu-id="96cbb-1047">function</span></span>|<span data-ttu-id="96cbb-1048">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="96cbb-1048">&lt;optional&gt;</span></span>|<span data-ttu-id="96cbb-1049">方法完成后，使用单个参数 `asyncResult`（一个 [AsyncResult](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="96cbb-1049">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="96cbb-1050">要求</span><span class="sxs-lookup"><span data-stu-id="96cbb-1050">Requirements</span></span>

|<span data-ttu-id="96cbb-1051">要求</span><span class="sxs-lookup"><span data-stu-id="96cbb-1051">Requirement</span></span>|<span data-ttu-id="96cbb-1052">值</span><span class="sxs-lookup"><span data-stu-id="96cbb-1052">Value</span></span>|
|---|---|
|[<span data-ttu-id="96cbb-1053">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="96cbb-1053">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="96cbb-1054">预览</span><span class="sxs-lookup"><span data-stu-id="96cbb-1054">Preview</span></span>|
|[<span data-ttu-id="96cbb-1055">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="96cbb-1055">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="96cbb-1056">ReadItem</span><span class="sxs-lookup"><span data-stu-id="96cbb-1056">ReadItem</span></span>|
|[<span data-ttu-id="96cbb-1057">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="96cbb-1057">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="96cbb-1058">撰写</span><span class="sxs-lookup"><span data-stu-id="96cbb-1058">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="96cbb-1059">返回：</span><span class="sxs-lookup"><span data-stu-id="96cbb-1059">Returns:</span></span>

<span data-ttu-id="96cbb-1060">类型：Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="96cbb-1060">Type: Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span></span>

##### <a name="example"></a><span data-ttu-id="96cbb-1061">示例</span><span class="sxs-lookup"><span data-stu-id="96cbb-1061">Example</span></span>

<span data-ttu-id="96cbb-1062">以下示例使用当前项目上所有附件的详细信息构成 HTML 字符串。</span><span class="sxs-lookup"><span data-stu-id="96cbb-1062">The following example builds an HTML string with details of all attachments on the current item.</span></span>

```javascript
var item = Office.context.mailbox.item;
var outputString = "";
item.getAttachmentsAsync(callback);  
function callback(result) {
    if (result.value.length > 0) {
        for (i = 0 ; i < result.value.length ; i++) {
            var _att = result.value [i];
            outputString += "<BR>" + i + ". Name: ";
            outputString += _att.name;
            outputString += "<BR>ID: " + _att.id;
            outputString += "<BR>contentType: " + _att.contentType;
            outputString += "<BR>size: " + _att.size;
            outputString += "<BR>attachmentType: " + _att.attachmentType;
            outputString += "<BR>isInline: " + _att.isInline;
        }
    }
}
```

#### <a name="getentities--entitiesjavascriptapioutlookofficeentities"></a><span data-ttu-id="96cbb-1063">getEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="96cbb-1063">getEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span></span>

<span data-ttu-id="96cbb-1064">获取在所选项目的正文中找到的实体。</span><span class="sxs-lookup"><span data-stu-id="96cbb-1064">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="96cbb-1065">在 Outlook for iOS 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="96cbb-1065">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="96cbb-1066">要求</span><span class="sxs-lookup"><span data-stu-id="96cbb-1066">Requirements</span></span>

|<span data-ttu-id="96cbb-1067">要求</span><span class="sxs-lookup"><span data-stu-id="96cbb-1067">Requirement</span></span>|<span data-ttu-id="96cbb-1068">值</span><span class="sxs-lookup"><span data-stu-id="96cbb-1068">Value</span></span>|
|---|---|
|[<span data-ttu-id="96cbb-1069">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="96cbb-1069">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="96cbb-1070">1.0</span><span class="sxs-lookup"><span data-stu-id="96cbb-1070">1.0</span></span>|
|[<span data-ttu-id="96cbb-1071">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="96cbb-1071">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="96cbb-1072">ReadItem</span><span class="sxs-lookup"><span data-stu-id="96cbb-1072">ReadItem</span></span>|
|[<span data-ttu-id="96cbb-1073">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="96cbb-1073">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="96cbb-1074">阅读</span><span class="sxs-lookup"><span data-stu-id="96cbb-1074">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="96cbb-1075">返回：</span><span class="sxs-lookup"><span data-stu-id="96cbb-1075">Returns:</span></span>

<span data-ttu-id="96cbb-1076">类型：[Entities](/javascript/api/outlook/office.entities)</span><span class="sxs-lookup"><span data-stu-id="96cbb-1076">Type: [Entities](/javascript/api/outlook/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="96cbb-1077">示例</span><span class="sxs-lookup"><span data-stu-id="96cbb-1077">Example</span></span>

<span data-ttu-id="96cbb-1078">以下示例访问当前项目的正文中的联系人实体。</span><span class="sxs-lookup"><span data-stu-id="96cbb-1078">The following example accesses the contacts entities in the current item's body.</span></span>

```javascript
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlookofficecontactmeetingsuggestionjavascriptapioutlookofficemeetingsuggestionphonenumberjavascriptapioutlookofficephonenumbertasksuggestionjavascriptapioutlookofficetasksuggestion"></a><span data-ttu-id="96cbb-1079">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="96cbb-1079">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span></span>

<span data-ttu-id="96cbb-1080">获取所选项目的正文中找到的指定实体类型的所有实体的数组。</span><span class="sxs-lookup"><span data-stu-id="96cbb-1080">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="96cbb-1081">在 Outlook for iOS 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="96cbb-1081">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="96cbb-1082">参数：</span><span class="sxs-lookup"><span data-stu-id="96cbb-1082">Parameters:</span></span>

|<span data-ttu-id="96cbb-1083">名称</span><span class="sxs-lookup"><span data-stu-id="96cbb-1083">Name</span></span>|<span data-ttu-id="96cbb-1084">类型</span><span class="sxs-lookup"><span data-stu-id="96cbb-1084">Type</span></span>|<span data-ttu-id="96cbb-1085">说明</span><span class="sxs-lookup"><span data-stu-id="96cbb-1085">Description</span></span>|
|---|---|---|
|`entityType`|[<span data-ttu-id="96cbb-1086">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="96cbb-1086">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook/office.mailboxenums.entitytype)|<span data-ttu-id="96cbb-1087">EntityType 枚举值之一。</span><span class="sxs-lookup"><span data-stu-id="96cbb-1087">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="96cbb-1088">要求</span><span class="sxs-lookup"><span data-stu-id="96cbb-1088">Requirements</span></span>

|<span data-ttu-id="96cbb-1089">要求</span><span class="sxs-lookup"><span data-stu-id="96cbb-1089">Requirement</span></span>|<span data-ttu-id="96cbb-1090">值</span><span class="sxs-lookup"><span data-stu-id="96cbb-1090">Value</span></span>|
|---|---|
|[<span data-ttu-id="96cbb-1091">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="96cbb-1091">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="96cbb-1092">1.0</span><span class="sxs-lookup"><span data-stu-id="96cbb-1092">1.0</span></span>|
|[<span data-ttu-id="96cbb-1093">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="96cbb-1093">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="96cbb-1094">受限</span><span class="sxs-lookup"><span data-stu-id="96cbb-1094">Restricted</span></span>|
|[<span data-ttu-id="96cbb-1095">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="96cbb-1095">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="96cbb-1096">阅读</span><span class="sxs-lookup"><span data-stu-id="96cbb-1096">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="96cbb-1097">返回：</span><span class="sxs-lookup"><span data-stu-id="96cbb-1097">Returns:</span></span>

<span data-ttu-id="96cbb-1098">如果在 `entityType` 中传递的值不是 `EntityType` 枚举的有效成员，该方法返回 null。</span><span class="sxs-lookup"><span data-stu-id="96cbb-1098">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="96cbb-1099">如果指定类型的任何实体都不存在于该项目的正文中，该方法将返回空数组。</span><span class="sxs-lookup"><span data-stu-id="96cbb-1099">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="96cbb-1100">否则，返回的数组中对象的类型取决于 `entityType` 参数中请求实体的类型。</span><span class="sxs-lookup"><span data-stu-id="96cbb-1100">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="96cbb-1101">当使用此方法的最低权限级别**受限**时，某些实体类型需要 **ReadItem** 才能进行访问，如下表中所指定。</span><span class="sxs-lookup"><span data-stu-id="96cbb-1101">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

|<span data-ttu-id="96cbb-1102">`entityType` 的值</span><span class="sxs-lookup"><span data-stu-id="96cbb-1102">Value of `entityType`</span></span>|<span data-ttu-id="96cbb-1103">返回的数组中对象的类型</span><span class="sxs-lookup"><span data-stu-id="96cbb-1103">Type of objects in returned array</span></span>|<span data-ttu-id="96cbb-1104">所需权限级别</span><span class="sxs-lookup"><span data-stu-id="96cbb-1104">Required Permission Level</span></span>|
|---|---|---|
|`Address`|<span data-ttu-id="96cbb-1105">String</span><span class="sxs-lookup"><span data-stu-id="96cbb-1105">String</span></span>|<span data-ttu-id="96cbb-1106">**受限**</span><span class="sxs-lookup"><span data-stu-id="96cbb-1106">**Restricted**</span></span>|
|`Contact`|<span data-ttu-id="96cbb-1107">Contact</span><span class="sxs-lookup"><span data-stu-id="96cbb-1107">Contact</span></span>|<span data-ttu-id="96cbb-1108">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="96cbb-1108">**ReadItem**</span></span>|
|`EmailAddress`|<span data-ttu-id="96cbb-1109">String</span><span class="sxs-lookup"><span data-stu-id="96cbb-1109">String</span></span>|<span data-ttu-id="96cbb-1110">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="96cbb-1110">**ReadItem**</span></span>|
|`MeetingSuggestion`|<span data-ttu-id="96cbb-1111">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="96cbb-1111">MeetingSuggestion</span></span>|<span data-ttu-id="96cbb-1112">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="96cbb-1112">**ReadItem**</span></span>|
|`PhoneNumber`|<span data-ttu-id="96cbb-1113">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="96cbb-1113">PhoneNumber</span></span>|<span data-ttu-id="96cbb-1114">**受限**</span><span class="sxs-lookup"><span data-stu-id="96cbb-1114">**Restricted**</span></span>|
|`TaskSuggestion`|<span data-ttu-id="96cbb-1115">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="96cbb-1115">TaskSuggestion</span></span>|<span data-ttu-id="96cbb-1116">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="96cbb-1116">**ReadItem**</span></span>|
|`URL`|<span data-ttu-id="96cbb-1117">String</span><span class="sxs-lookup"><span data-stu-id="96cbb-1117">String</span></span>|<span data-ttu-id="96cbb-1118">**受限**</span><span class="sxs-lookup"><span data-stu-id="96cbb-1118">**Restricted**</span></span>|

<span data-ttu-id="96cbb-1119">类型：Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="96cbb-1119">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span></span>

##### <a name="example"></a><span data-ttu-id="96cbb-1120">示例</span><span class="sxs-lookup"><span data-stu-id="96cbb-1120">Example</span></span>

<span data-ttu-id="96cbb-1121">以下示例显示了如何访问表示当前项目的正文中的邮政地址的字符串数组。</span><span class="sxs-lookup"><span data-stu-id="96cbb-1121">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlookofficecontactmeetingsuggestionjavascriptapioutlookofficemeetingsuggestionphonenumberjavascriptapioutlookofficephonenumbertasksuggestionjavascriptapioutlookofficetasksuggestion"></a><span data-ttu-id="96cbb-1122">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="96cbb-1122">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span></span>

<span data-ttu-id="96cbb-1123">返回传递清单 XML 文件中定义的命名筛选器的所选项目中的已知实体。</span><span class="sxs-lookup"><span data-stu-id="96cbb-1123">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="96cbb-1124">在 Outlook for iOS 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="96cbb-1124">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="96cbb-1125">`getFilteredEntitiesByName` 方法返回匹配在具有指定 `FilterName` 元素值的清单 XML 文件中的 [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) 规则元素中定义的正则表达式的实体。</span><span class="sxs-lookup"><span data-stu-id="96cbb-1125">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="96cbb-1126">参数：</span><span class="sxs-lookup"><span data-stu-id="96cbb-1126">Parameters:</span></span>

|<span data-ttu-id="96cbb-1127">名称</span><span class="sxs-lookup"><span data-stu-id="96cbb-1127">Name</span></span>|<span data-ttu-id="96cbb-1128">类型</span><span class="sxs-lookup"><span data-stu-id="96cbb-1128">Type</span></span>|<span data-ttu-id="96cbb-1129">说明</span><span class="sxs-lookup"><span data-stu-id="96cbb-1129">Description</span></span>|
|---|---|---|
|`name`|<span data-ttu-id="96cbb-1130">字符串</span><span class="sxs-lookup"><span data-stu-id="96cbb-1130">String</span></span>|<span data-ttu-id="96cbb-1131">定义筛选器匹配的 `ItemHasKnownEntity` 规则元素的名称。</span><span class="sxs-lookup"><span data-stu-id="96cbb-1131">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="96cbb-1132">要求</span><span class="sxs-lookup"><span data-stu-id="96cbb-1132">Requirements</span></span>

|<span data-ttu-id="96cbb-1133">要求</span><span class="sxs-lookup"><span data-stu-id="96cbb-1133">Requirement</span></span>|<span data-ttu-id="96cbb-1134">值</span><span class="sxs-lookup"><span data-stu-id="96cbb-1134">Value</span></span>|
|---|---|
|[<span data-ttu-id="96cbb-1135">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="96cbb-1135">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="96cbb-1136">1.0</span><span class="sxs-lookup"><span data-stu-id="96cbb-1136">1.0</span></span>|
|[<span data-ttu-id="96cbb-1137">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="96cbb-1137">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="96cbb-1138">ReadItem</span><span class="sxs-lookup"><span data-stu-id="96cbb-1138">ReadItem</span></span>|
|[<span data-ttu-id="96cbb-1139">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="96cbb-1139">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="96cbb-1140">阅读</span><span class="sxs-lookup"><span data-stu-id="96cbb-1140">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="96cbb-1141">返回：</span><span class="sxs-lookup"><span data-stu-id="96cbb-1141">Returns:</span></span>

<span data-ttu-id="96cbb-p162">如果具有匹配 `name` 参数的 `FilterName` 元素值的清单中没有任何 `ItemHasKnownEntity` 元素，则该方法返回 `null`。如果 `name` 参数匹配清单中的 `ItemHasKnownEntity` 元素，但在匹配的当前项目中没有实体，则该方法返回一个空数组。</span><span class="sxs-lookup"><span data-stu-id="96cbb-p162">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="96cbb-1144">类型：Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="96cbb-1144">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span></span>

#### <a name="getinitializationcontextasyncoptions-callback"></a><span data-ttu-id="96cbb-1145">getInitializationContextAsync([options], [callback])</span><span class="sxs-lookup"><span data-stu-id="96cbb-1145">getInitializationContextAsync([options], [callback])</span></span>

<span data-ttu-id="96cbb-1146">当[可操作邮件激活](https://docs.microsoft.com/outlook/actionable-messages/invoke-add-in-from-actionable-message)加载项时，获取传递的初始化数据。</span><span class="sxs-lookup"><span data-stu-id="96cbb-1146">Gets initialization data passed when the add-in is [activated by an actionable message](https://docs.microsoft.com/outlook/actionable-messages/invoke-add-in-from-actionable-message).</span></span>

> [!NOTE]
> <span data-ttu-id="96cbb-1147">仅 Outlook 2016 for Windows 或更高版本（高于 16.0.8413.1000 的即点即用版本）和适用于 Office 365 的 Outlook 网页版支持此方法。</span><span class="sxs-lookup"><span data-stu-id="96cbb-1147">This method is only supported by Outlook 2016 or later for Windows (Click-to-Run versions later than 16.0.8413.1000) and Outlook on the web for Office 365.</span></span>

##### <a name="parameters"></a><span data-ttu-id="96cbb-1148">参数：</span><span class="sxs-lookup"><span data-stu-id="96cbb-1148">Parameters:</span></span>
|<span data-ttu-id="96cbb-1149">名称</span><span class="sxs-lookup"><span data-stu-id="96cbb-1149">Name</span></span>|<span data-ttu-id="96cbb-1150">类型</span><span class="sxs-lookup"><span data-stu-id="96cbb-1150">Type</span></span>|<span data-ttu-id="96cbb-1151">属性</span><span class="sxs-lookup"><span data-stu-id="96cbb-1151">Attributes</span></span>|<span data-ttu-id="96cbb-1152">说明</span><span class="sxs-lookup"><span data-stu-id="96cbb-1152">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="96cbb-1153">对象</span><span class="sxs-lookup"><span data-stu-id="96cbb-1153">Object</span></span>|<span data-ttu-id="96cbb-1154">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="96cbb-1154">&lt;optional&gt;</span></span>|<span data-ttu-id="96cbb-1155">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="96cbb-1155">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="96cbb-1156">对象</span><span class="sxs-lookup"><span data-stu-id="96cbb-1156">Object</span></span>|<span data-ttu-id="96cbb-1157">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="96cbb-1157">&lt;optional&gt;</span></span>|<span data-ttu-id="96cbb-1158">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="96cbb-1158">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="96cbb-1159">函数</span><span class="sxs-lookup"><span data-stu-id="96cbb-1159">function</span></span>|<span data-ttu-id="96cbb-1160">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="96cbb-1160">&lt;optional&gt;</span></span>|<span data-ttu-id="96cbb-1161">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="96cbb-1161">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="96cbb-1162">成功后，`asyncResult.value` 属性便以字符串形式提供初始化数据。</span><span class="sxs-lookup"><span data-stu-id="96cbb-1162">On success, the initialization data is provided in the `asyncResult.value` property as a string.</span></span><br/><span data-ttu-id="96cbb-1163">如果没有初始化上下文，`asyncResult` 对象包含 `Error` 对象，并将它的 `code` 和 `name` 属性分别设置为 `9020` 和 `GenericResponseError`。</span><span class="sxs-lookup"><span data-stu-id="96cbb-1163">If there is no initialization context, the `asyncResult` object will contain an `Error` object with its `code` property set to `9020` and its `name` property set to `GenericResponseError`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="96cbb-1164">要求</span><span class="sxs-lookup"><span data-stu-id="96cbb-1164">Requirements</span></span>

|<span data-ttu-id="96cbb-1165">要求</span><span class="sxs-lookup"><span data-stu-id="96cbb-1165">Requirement</span></span>|<span data-ttu-id="96cbb-1166">值</span><span class="sxs-lookup"><span data-stu-id="96cbb-1166">Value</span></span>|
|---|---|
|[<span data-ttu-id="96cbb-1167">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="96cbb-1167">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="96cbb-1168">预览</span><span class="sxs-lookup"><span data-stu-id="96cbb-1168">Preview</span></span>|
|[<span data-ttu-id="96cbb-1169">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="96cbb-1169">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="96cbb-1170">ReadItem</span><span class="sxs-lookup"><span data-stu-id="96cbb-1170">ReadItem</span></span>|
|[<span data-ttu-id="96cbb-1171">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="96cbb-1171">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="96cbb-1172">阅读</span><span class="sxs-lookup"><span data-stu-id="96cbb-1172">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="96cbb-1173">示例</span><span class="sxs-lookup"><span data-stu-id="96cbb-1173">Example</span></span>

```javascript
// Get the initialization context (if present)
Office.context.mailbox.item.getInitializationContextAsync(
  function(asyncResult) {
    if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
      if (asyncResult.value != null && asyncResult.value.length > 0) {
        // The value is a string, parse to an object
        var context = JSON.parse(asyncResult.value);
        // Do something with context
      } else {
        // Empty context, treat as no context
      }
    } else {
      if (asyncResult.error.code == 9020) {
        // GenericResponseError returned when there is
        // no context
        // Treat as no context
      } else {
        // Handle the error
      }
    }
  }
);
```

#### <a name="getregexmatches--object"></a><span data-ttu-id="96cbb-1174">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="96cbb-1174">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="96cbb-1175">返回所选项目中匹配在清单 XML 文件中定义的正则表达式的字符串值。</span><span class="sxs-lookup"><span data-stu-id="96cbb-1175">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="96cbb-1176">在 Outlook for iOS 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="96cbb-1176">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="96cbb-p163">`getRegExMatches` 方法返回匹配在清单 XML 文件中的每个 `ItemHasRegularExpressionMatch` 或 `ItemHasKnownEntity` 规则元素中定义的正则表达式的字符串。对于 `ItemHasRegularExpressionMatch` 规则，匹配字符串必须发生在该规则指定的项目的属性中。`PropertyName` 简单类型定义支持的属性。</span><span class="sxs-lookup"><span data-stu-id="96cbb-p163">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="96cbb-1180">例如，考虑一个外接程序清单具有以下 `Rule` 元素：</span><span class="sxs-lookup"><span data-stu-id="96cbb-1180">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="96cbb-1181">从 `getRegExMatches` 返回的对象应有两个属性：`fruits` 和 `veggies`。</span><span class="sxs-lookup"><span data-stu-id="96cbb-1181">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="96cbb-p164">如果在项目的正文属性上指定 `ItemHasRegularExpressionMatch` 规则，则正则表达式应进一步筛选正文，不应尝试返回该项目的整个正文。使用正则表达式（如 `.*`）获取项目的整个正文并不总是返回预期的结果。而是使用 [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) 方法检索整个正文。</span><span class="sxs-lookup"><span data-stu-id="96cbb-p164">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="96cbb-1185">Requirements</span><span class="sxs-lookup"><span data-stu-id="96cbb-1185">Requirements</span></span>

|<span data-ttu-id="96cbb-1186">要求</span><span class="sxs-lookup"><span data-stu-id="96cbb-1186">Requirement</span></span>|<span data-ttu-id="96cbb-1187">值</span><span class="sxs-lookup"><span data-stu-id="96cbb-1187">Value</span></span>|
|---|---|
|[<span data-ttu-id="96cbb-1188">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="96cbb-1188">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="96cbb-1189">1.0</span><span class="sxs-lookup"><span data-stu-id="96cbb-1189">1.0</span></span>|
|[<span data-ttu-id="96cbb-1190">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="96cbb-1190">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="96cbb-1191">ReadItem</span><span class="sxs-lookup"><span data-stu-id="96cbb-1191">ReadItem</span></span>|
|[<span data-ttu-id="96cbb-1192">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="96cbb-1192">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="96cbb-1193">阅读</span><span class="sxs-lookup"><span data-stu-id="96cbb-1193">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="96cbb-1194">返回：</span><span class="sxs-lookup"><span data-stu-id="96cbb-1194">Returns:</span></span>

<span data-ttu-id="96cbb-p165">一个包含匹配在清单 XML 文件中定义的正则表达式的字符串数组的对象。每个数组的名称等于匹配 `ItemHasRegularExpressionMatch` 规则的 `RegExName` 属性或匹配 `ItemHasKnownEntity` 规则的 `FilterName` 属性的相应值。</span><span class="sxs-lookup"><span data-stu-id="96cbb-p165">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<dl class="param-type"><span data-ttu-id="96cbb-1197">

<dt>
类型</dt>


</span><span class="sxs-lookup"><span data-stu-id="96cbb-1197">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="96cbb-1198">对象</span><span class="sxs-lookup"><span data-stu-id="96cbb-1198">Object</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="96cbb-1199">示例</span><span class="sxs-lookup"><span data-stu-id="96cbb-1199">Example</span></span>

<span data-ttu-id="96cbb-1200">以下示例显示了如何访问正则表达式规则元素 `fruits` 和 `veggies` 的匹配项的数组，这些元素在清单中指定。</span><span class="sxs-lookup"><span data-stu-id="96cbb-1200">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```javascript
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veges = allMatches.veggies;
```

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="96cbb-1201">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span><span class="sxs-lookup"><span data-stu-id="96cbb-1201">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="96cbb-1202">返回匹配在清单 XML 文件中定义的命名正则表达式的所选项目中的字符串值。</span><span class="sxs-lookup"><span data-stu-id="96cbb-1202">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="96cbb-1203">在 Outlook for iOS 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="96cbb-1203">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="96cbb-1204">`getRegExMatchesByName` 方法返回匹配在具有指定 `RegExName` 元素值的清单 XML 文件中的 `ItemHasRegularExpressionMatch` 规则元素中定义的正则表达式的字符串。</span><span class="sxs-lookup"><span data-stu-id="96cbb-1204">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="96cbb-p166">如果在项目的正文属性上指定 `ItemHasRegularExpressionMatch` 规则，则正则表达式应进一步筛选正文，不应尝试返回该项目的整个正文。使用正则表达式（如 `.*`）获取项目的整个正文并不总是返回预期的结果。</span><span class="sxs-lookup"><span data-stu-id="96cbb-p166">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="96cbb-1207">参数：</span><span class="sxs-lookup"><span data-stu-id="96cbb-1207">Parameters:</span></span>

|<span data-ttu-id="96cbb-1208">名称</span><span class="sxs-lookup"><span data-stu-id="96cbb-1208">Name</span></span>|<span data-ttu-id="96cbb-1209">类型</span><span class="sxs-lookup"><span data-stu-id="96cbb-1209">Type</span></span>|<span data-ttu-id="96cbb-1210">说明</span><span class="sxs-lookup"><span data-stu-id="96cbb-1210">Description</span></span>|
|---|---|---|
|`name`|<span data-ttu-id="96cbb-1211">字符串</span><span class="sxs-lookup"><span data-stu-id="96cbb-1211">String</span></span>|<span data-ttu-id="96cbb-1212">定义筛选器匹配的 `ItemHasRegularExpressionMatch` 规则元素的名称。</span><span class="sxs-lookup"><span data-stu-id="96cbb-1212">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="96cbb-1213">要求</span><span class="sxs-lookup"><span data-stu-id="96cbb-1213">Requirements</span></span>

|<span data-ttu-id="96cbb-1214">要求</span><span class="sxs-lookup"><span data-stu-id="96cbb-1214">Requirement</span></span>|<span data-ttu-id="96cbb-1215">值</span><span class="sxs-lookup"><span data-stu-id="96cbb-1215">Value</span></span>|
|---|---|
|[<span data-ttu-id="96cbb-1216">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="96cbb-1216">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="96cbb-1217">1.0</span><span class="sxs-lookup"><span data-stu-id="96cbb-1217">1.0</span></span>|
|[<span data-ttu-id="96cbb-1218">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="96cbb-1218">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="96cbb-1219">ReadItem</span><span class="sxs-lookup"><span data-stu-id="96cbb-1219">ReadItem</span></span>|
|[<span data-ttu-id="96cbb-1220">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="96cbb-1220">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="96cbb-1221">阅读</span><span class="sxs-lookup"><span data-stu-id="96cbb-1221">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="96cbb-1222">返回：</span><span class="sxs-lookup"><span data-stu-id="96cbb-1222">Returns:</span></span>

<span data-ttu-id="96cbb-1223">一个包含匹配在清单 XML 文件中定义的正则表达式的字符串的数组。</span><span class="sxs-lookup"><span data-stu-id="96cbb-1223">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<dl class="param-type"><span data-ttu-id="96cbb-1224">

<dt>类型</dt>

</span><span class="sxs-lookup"><span data-stu-id="96cbb-1224">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="96cbb-1225">Array.< String ></span><span class="sxs-lookup"><span data-stu-id="96cbb-1225">Array.< String ></span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="96cbb-1226">示例</span><span class="sxs-lookup"><span data-stu-id="96cbb-1226">Example</span></span>

```javascript
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

####  <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="96cbb-1227">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="96cbb-1227">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="96cbb-1228">以异步方式返回邮件的主题或正文中选定的数据。</span><span class="sxs-lookup"><span data-stu-id="96cbb-1228">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="96cbb-p167">如果没有选定内容，但光标位于正文或主题中，此方法将会为所选数据返回 null。如果选定的是字段，而不是正文或主题，则此方法返回 `InvalidSelection` 错误。</span><span class="sxs-lookup"><span data-stu-id="96cbb-p167">If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="96cbb-1231">参数：</span><span class="sxs-lookup"><span data-stu-id="96cbb-1231">Parameters:</span></span>

|<span data-ttu-id="96cbb-1232">名称</span><span class="sxs-lookup"><span data-stu-id="96cbb-1232">Name</span></span>|<span data-ttu-id="96cbb-1233">类型</span><span class="sxs-lookup"><span data-stu-id="96cbb-1233">Type</span></span>|<span data-ttu-id="96cbb-1234">属性</span><span class="sxs-lookup"><span data-stu-id="96cbb-1234">Attributes</span></span>|<span data-ttu-id="96cbb-1235">说明</span><span class="sxs-lookup"><span data-stu-id="96cbb-1235">Description</span></span>|
|---|---|---|---|
|`coercionType`|[<span data-ttu-id="96cbb-1236">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="96cbb-1236">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="96cbb-p168">请求数据的格式。如果为文本，则此方法返回纯文本作为字符串，删除任何显示的 HTML 标记。如果为 HTML，则此方法返回所选文本，不论是纯文本还是 HTML。</span><span class="sxs-lookup"><span data-stu-id="96cbb-p168">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`|<span data-ttu-id="96cbb-1240">对象</span><span class="sxs-lookup"><span data-stu-id="96cbb-1240">Object</span></span>|<span data-ttu-id="96cbb-1241">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="96cbb-1241">&lt;optional&gt;</span></span>|<span data-ttu-id="96cbb-1242">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="96cbb-1242">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="96cbb-1243">对象</span><span class="sxs-lookup"><span data-stu-id="96cbb-1243">Object</span></span>|<span data-ttu-id="96cbb-1244">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="96cbb-1244">&lt;optional&gt;</span></span>|<span data-ttu-id="96cbb-1245">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="96cbb-1245">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="96cbb-1246">函数</span><span class="sxs-lookup"><span data-stu-id="96cbb-1246">function</span></span>||<span data-ttu-id="96cbb-1247">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="96cbb-1247">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="96cbb-1248">若要从回调方法访问所选数据，请调用 `asyncResult.value.data`。</span><span class="sxs-lookup"><span data-stu-id="96cbb-1248">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="96cbb-1249">若要访问选定内容的源属性，请调用 `asyncResult.value.sourceProperty`，这将为 `body` 或 `subject`。</span><span class="sxs-lookup"><span data-stu-id="96cbb-1249">To access the source property that the selection comes from, call `asyncResult.value.sourceProperty`, which will be either `body` or `subject`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="96cbb-1250">要求</span><span class="sxs-lookup"><span data-stu-id="96cbb-1250">Requirements</span></span>

|<span data-ttu-id="96cbb-1251">要求</span><span class="sxs-lookup"><span data-stu-id="96cbb-1251">Requirement</span></span>|<span data-ttu-id="96cbb-1252">值</span><span class="sxs-lookup"><span data-stu-id="96cbb-1252">Value</span></span>|
|---|---|
|[<span data-ttu-id="96cbb-1253">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="96cbb-1253">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="96cbb-1254">1.2</span><span class="sxs-lookup"><span data-stu-id="96cbb-1254">1.2</span></span>|
|[<span data-ttu-id="96cbb-1255">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="96cbb-1255">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="96cbb-1256">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="96cbb-1256">ReadWriteItem</span></span>|
|[<span data-ttu-id="96cbb-1257">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="96cbb-1257">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="96cbb-1258">撰写</span><span class="sxs-lookup"><span data-stu-id="96cbb-1258">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="96cbb-1259">返回：</span><span class="sxs-lookup"><span data-stu-id="96cbb-1259">Returns:</span></span>

<span data-ttu-id="96cbb-1260">作为字符串的所选数据的格式由 `coercionType` 确定。</span><span class="sxs-lookup"><span data-stu-id="96cbb-1260">The selected data as a string with format determined by `coercionType`.</span></span>

<dl class="param-type"><span data-ttu-id="96cbb-1261">

<dt>
类型</dt>


</span><span class="sxs-lookup"><span data-stu-id="96cbb-1261">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="96cbb-1262">字符串</span><span class="sxs-lookup"><span data-stu-id="96cbb-1262">String</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="96cbb-1263">示例</span><span class="sxs-lookup"><span data-stu-id="96cbb-1263">Example</span></span>

```javascript
// getting selected data
Office.initialize = function () {
    Office.context.mailbox.item.getSelectedDataAsync(Office.CoercionType.Text, {}, getCallback);
}

function getCallback(asyncResult) {
    var text = asyncResult.value.data;
    var prop = asyncResult.value.sourceProperty;

    Office.context.mailbox.item.setSelectedDataAsync('Setting ' + prop + ': ' + text, {}, setCallback);
}

function setCallback(asyncResult) {
    // check for errors
}
```

#### <a name="getselectedentities--entitiesjavascriptapioutlookofficeentities"></a><span data-ttu-id="96cbb-1264">getSelectedEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="96cbb-1264">getSelectedEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span></span>

<span data-ttu-id="96cbb-p170">获取在用户已选择的突出显示匹配项中找到的实体。突出显示匹配项适用于[上下文加载项](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins)。</span><span class="sxs-lookup"><span data-stu-id="96cbb-p170">Gets the entities found in a highlighted match a user has selected. Highlighted matches apply to [contextual add-ins](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="96cbb-1267">在 Outlook for iOS 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="96cbb-1267">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="96cbb-1268">要求</span><span class="sxs-lookup"><span data-stu-id="96cbb-1268">Requirements</span></span>

|<span data-ttu-id="96cbb-1269">要求</span><span class="sxs-lookup"><span data-stu-id="96cbb-1269">Requirement</span></span>|<span data-ttu-id="96cbb-1270">值</span><span class="sxs-lookup"><span data-stu-id="96cbb-1270">Value</span></span>|
|---|---|
|[<span data-ttu-id="96cbb-1271">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="96cbb-1271">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="96cbb-1272">1.6</span><span class="sxs-lookup"><span data-stu-id="96cbb-1272">1.6</span></span>|
|[<span data-ttu-id="96cbb-1273">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="96cbb-1273">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="96cbb-1274">ReadItem</span><span class="sxs-lookup"><span data-stu-id="96cbb-1274">ReadItem</span></span>|
|[<span data-ttu-id="96cbb-1275">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="96cbb-1275">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="96cbb-1276">阅读</span><span class="sxs-lookup"><span data-stu-id="96cbb-1276">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="96cbb-1277">返回：</span><span class="sxs-lookup"><span data-stu-id="96cbb-1277">Returns:</span></span>

<span data-ttu-id="96cbb-1278">类型：[Entities](/javascript/api/outlook/office.entities)</span><span class="sxs-lookup"><span data-stu-id="96cbb-1278">Type: [Entities](/javascript/api/outlook/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="96cbb-1279">示例</span><span class="sxs-lookup"><span data-stu-id="96cbb-1279">Example</span></span>

<span data-ttu-id="96cbb-1280">以下示例访问用户选择的突出显示匹配项中的地址实体。</span><span class="sxs-lookup"><span data-stu-id="96cbb-1280">The following example accesses the addresses entities in the highlighted match selected by the user.</span></span>

```javascript
var contacts = Office.context.mailbox.item.getSelectedEntities().addresses;
```

#### <a name="getselectedregexmatches--object"></a><span data-ttu-id="96cbb-1281">getSelectedRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="96cbb-1281">getSelectedRegExMatches() → {Object}</span></span>

<span data-ttu-id="96cbb-p171">返回突出显示匹配项中匹配在清单 XML 文件中定义的正则表达式的字符串值。突出显示匹配项适用于[上下文外接程序](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins)。</span><span class="sxs-lookup"><span data-stu-id="96cbb-p171">Returns string values in a highlighted match that match the regular expressions defined in the manifest XML file. Highlighted matches apply to [contextual add-ins](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="96cbb-1284">在 Outlook for iOS 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="96cbb-1284">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="96cbb-p172">`getSelectedRegExMatches` 方法返回匹配在清单 XML 文件中的每个 `ItemHasRegularExpressionMatch` 或 `ItemHasKnownEntity` 规则元素中定义的正则表达式的字符串。对于 `ItemHasRegularExpressionMatch` 规则，匹配字符串必须发生在该规则指定的项目的属性中。`PropertyName` 简单类型定义支持的属性。</span><span class="sxs-lookup"><span data-stu-id="96cbb-p172">The `getSelectedRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="96cbb-1288">例如，考虑一个外接程序清单具有以下 `Rule` 元素：</span><span class="sxs-lookup"><span data-stu-id="96cbb-1288">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="96cbb-1289">从 `getRegExMatches` 返回的对象应有两个属性：`fruits` 和 `veggies`。</span><span class="sxs-lookup"><span data-stu-id="96cbb-1289">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="96cbb-p173">如果在项目的正文属性上指定 `ItemHasRegularExpressionMatch` 规则，则正则表达式应进一步筛选正文，不应尝试返回该项目的整个正文。使用正则表达式（如 `.*`）获取项目的整个正文并不总是返回预期的结果。而是使用 [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) 方法检索整个正文。</span><span class="sxs-lookup"><span data-stu-id="96cbb-p173">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="96cbb-1293">Requirements</span><span class="sxs-lookup"><span data-stu-id="96cbb-1293">Requirements</span></span>

|<span data-ttu-id="96cbb-1294">要求</span><span class="sxs-lookup"><span data-stu-id="96cbb-1294">Requirement</span></span>|<span data-ttu-id="96cbb-1295">值</span><span class="sxs-lookup"><span data-stu-id="96cbb-1295">Value</span></span>|
|---|---|
|[<span data-ttu-id="96cbb-1296">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="96cbb-1296">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="96cbb-1297">1.6</span><span class="sxs-lookup"><span data-stu-id="96cbb-1297">1.6</span></span>|
|[<span data-ttu-id="96cbb-1298">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="96cbb-1298">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="96cbb-1299">ReadItem</span><span class="sxs-lookup"><span data-stu-id="96cbb-1299">ReadItem</span></span>|
|[<span data-ttu-id="96cbb-1300">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="96cbb-1300">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="96cbb-1301">阅读</span><span class="sxs-lookup"><span data-stu-id="96cbb-1301">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="96cbb-1302">返回：</span><span class="sxs-lookup"><span data-stu-id="96cbb-1302">Returns:</span></span>

<span data-ttu-id="96cbb-p174">一个包含匹配在清单 XML 文件中定义的正则表达式的字符串数组的对象。每个数组的名称等于匹配 `ItemHasRegularExpressionMatch` 规则的 `RegExName` 属性或匹配 `ItemHasKnownEntity` 规则的 `FilterName` 属性的相应值。</span><span class="sxs-lookup"><span data-stu-id="96cbb-p174">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

##### <a name="example"></a><span data-ttu-id="96cbb-1305">示例</span><span class="sxs-lookup"><span data-stu-id="96cbb-1305">Example</span></span>

<span data-ttu-id="96cbb-1306">以下示例显示了如何访问正则表达式规则元素 `fruits` 和 `veggies` 的匹配项的数组，这些元素在清单中指定。</span><span class="sxs-lookup"><span data-stu-id="96cbb-1306">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```javascript
var selectedMatches = Office.context.mailbox.item.getSelectedRegExMatches();
var fruits = selectedMatches.fruits;
var veggies = selectedMatches.veggies;
```

#### <a name="getsharedpropertiesasyncoptions-callback"></a><span data-ttu-id="96cbb-1307">getSharedPropertiesAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="96cbb-1307">getSharedPropertiesAsync([options], callback)</span></span>

<span data-ttu-id="96cbb-1308">获取共享文件夹、日历或邮箱中所选约会或邮件的属性。</span><span class="sxs-lookup"><span data-stu-id="96cbb-1308">Gets the properties of the selected appointment or message in a shared folder, calendar, or mailbox.</span></span>

##### <a name="parameters"></a><span data-ttu-id="96cbb-1309">参数：</span><span class="sxs-lookup"><span data-stu-id="96cbb-1309">Parameters:</span></span>

|<span data-ttu-id="96cbb-1310">名称</span><span class="sxs-lookup"><span data-stu-id="96cbb-1310">Name</span></span>|<span data-ttu-id="96cbb-1311">类型</span><span class="sxs-lookup"><span data-stu-id="96cbb-1311">Type</span></span>|<span data-ttu-id="96cbb-1312">属性</span><span class="sxs-lookup"><span data-stu-id="96cbb-1312">Attributes</span></span>|<span data-ttu-id="96cbb-1313">说明</span><span class="sxs-lookup"><span data-stu-id="96cbb-1313">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="96cbb-1314">对象</span><span class="sxs-lookup"><span data-stu-id="96cbb-1314">Object</span></span>|<span data-ttu-id="96cbb-1315">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="96cbb-1315">&lt;optional&gt;</span></span>|<span data-ttu-id="96cbb-1316">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="96cbb-1316">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="96cbb-1317">对象</span><span class="sxs-lookup"><span data-stu-id="96cbb-1317">Object</span></span>|<span data-ttu-id="96cbb-1318">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="96cbb-1318">&lt;optional&gt;</span></span>|<span data-ttu-id="96cbb-1319">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="96cbb-1319">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="96cbb-1320">函数</span><span class="sxs-lookup"><span data-stu-id="96cbb-1320">function</span></span>||<span data-ttu-id="96cbb-1321">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="96cbb-1321">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="96cbb-1322">共享属性作为 `asyncResult.value` 属性中的 [`SharedProperties`](/javascript/api/outlook/office.sharedproperties) 对象提供。</span><span class="sxs-lookup"><span data-stu-id="96cbb-1322">The shared properties are provided as a [`SharedProperties`](/javascript/api/outlook/office.sharedproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="96cbb-1323">此对象可用于获取项目的共享属性。</span><span class="sxs-lookup"><span data-stu-id="96cbb-1323">This object can be used to get the item's shared properties.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="96cbb-1324">要求</span><span class="sxs-lookup"><span data-stu-id="96cbb-1324">Requirements</span></span>

|<span data-ttu-id="96cbb-1325">要求</span><span class="sxs-lookup"><span data-stu-id="96cbb-1325">Requirement</span></span>|<span data-ttu-id="96cbb-1326">值</span><span class="sxs-lookup"><span data-stu-id="96cbb-1326">Value</span></span>|
|---|---|
|[<span data-ttu-id="96cbb-1327">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="96cbb-1327">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="96cbb-1328">预览</span><span class="sxs-lookup"><span data-stu-id="96cbb-1328">Preview</span></span>|
|[<span data-ttu-id="96cbb-1329">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="96cbb-1329">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="96cbb-1330">ReadItem</span><span class="sxs-lookup"><span data-stu-id="96cbb-1330">ReadItem</span></span>|
|[<span data-ttu-id="96cbb-1331">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="96cbb-1331">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="96cbb-1332">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="96cbb-1332">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="96cbb-1333">示例</span><span class="sxs-lookup"><span data-stu-id="96cbb-1333">Example</span></span>

```js
Office.context.mailbox.item.getSharedPropertiesAsync(callback);
function callback (asyncResult) {
  var context=asyncResult.context;
  var sharedProperties = asyncResult.value;
}
```

####  <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="96cbb-1334">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="96cbb-1334">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="96cbb-1335">异步加载所选项目上此外接程序的自定义属性。</span><span class="sxs-lookup"><span data-stu-id="96cbb-1335">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="96cbb-p176">自定义属性基于每个应用、每个项目存储为键/值对。此方法在回调中返回 `CustomProperties` 对象，该回调提供访问特定于当前项目和当前外接程序的自定义属性的方法。自定义属性未在项目上加密，因此这不应用作安全存储。</span><span class="sxs-lookup"><span data-stu-id="96cbb-p176">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="96cbb-1339">参数：</span><span class="sxs-lookup"><span data-stu-id="96cbb-1339">Parameters:</span></span>

|<span data-ttu-id="96cbb-1340">名称</span><span class="sxs-lookup"><span data-stu-id="96cbb-1340">Name</span></span>|<span data-ttu-id="96cbb-1341">类型</span><span class="sxs-lookup"><span data-stu-id="96cbb-1341">Type</span></span>|<span data-ttu-id="96cbb-1342">属性</span><span class="sxs-lookup"><span data-stu-id="96cbb-1342">Attributes</span></span>|<span data-ttu-id="96cbb-1343">说明</span><span class="sxs-lookup"><span data-stu-id="96cbb-1343">Description</span></span>|
|---|---|---|---|
|`callback`|<span data-ttu-id="96cbb-1344">函数</span><span class="sxs-lookup"><span data-stu-id="96cbb-1344">function</span></span>||<span data-ttu-id="96cbb-1345">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="96cbb-1345">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="96cbb-1346">自定义属性作为 `asyncResult.value` 属性中的 [`CustomProperties`](/javascript/api/outlook/office.customproperties) 对象提供。</span><span class="sxs-lookup"><span data-stu-id="96cbb-1346">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook/office.customproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="96cbb-1347">此对象可用于获取、设置以及从项目中删除自定义属性，并将自定义属性集的更改重新保存到服务器。</span><span class="sxs-lookup"><span data-stu-id="96cbb-1347">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`|<span data-ttu-id="96cbb-1348">对象</span><span class="sxs-lookup"><span data-stu-id="96cbb-1348">Object</span></span>|<span data-ttu-id="96cbb-1349">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="96cbb-1349">&lt;optional&gt;</span></span>|<span data-ttu-id="96cbb-1350">开发人员可以提供他们想要在回调函数中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="96cbb-1350">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="96cbb-1351">此对象可以通过回调函数中的 `asyncResult.asyncContext` 属性进行访问。</span><span class="sxs-lookup"><span data-stu-id="96cbb-1351">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="96cbb-1352">要求</span><span class="sxs-lookup"><span data-stu-id="96cbb-1352">Requirements</span></span>

|<span data-ttu-id="96cbb-1353">要求</span><span class="sxs-lookup"><span data-stu-id="96cbb-1353">Requirement</span></span>|<span data-ttu-id="96cbb-1354">值</span><span class="sxs-lookup"><span data-stu-id="96cbb-1354">Value</span></span>|
|---|---|
|[<span data-ttu-id="96cbb-1355">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="96cbb-1355">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="96cbb-1356">1.0</span><span class="sxs-lookup"><span data-stu-id="96cbb-1356">1.0</span></span>|
|[<span data-ttu-id="96cbb-1357">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="96cbb-1357">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="96cbb-1358">ReadItem</span><span class="sxs-lookup"><span data-stu-id="96cbb-1358">ReadItem</span></span>|
|[<span data-ttu-id="96cbb-1359">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="96cbb-1359">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="96cbb-1360">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="96cbb-1360">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="96cbb-1361">示例</span><span class="sxs-lookup"><span data-stu-id="96cbb-1361">Example</span></span>

<span data-ttu-id="96cbb-p179">以下代码示例显示了如何使用 `loadCustomPropertiesAsync` 方法异步加载特定于当前项目的自定义属性。该示例还显示了如何使用 `CustomProperties.saveAsync` 方法将这些属性重新保存到服务器。加载自定义属性后，该代码示例将使用 `CustomProperties.get` 方法读取自定义属性 `myProp`，使用 `CustomProperties.set` 方法写入自定义属性 `otherProp`，最后调用 `saveAsync` 方法保存这些自定义属性。</span><span class="sxs-lookup"><span data-stu-id="96cbb-p179">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

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

####  <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="96cbb-1365">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="96cbb-1365">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="96cbb-1366">将附件从邮件或约会中删除。</span><span class="sxs-lookup"><span data-stu-id="96cbb-1366">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="96cbb-1367">`removeAttachmentAsync` 方法删除项目中带指定标识符的附件。</span><span class="sxs-lookup"><span data-stu-id="96cbb-1367">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item.</span></span> <span data-ttu-id="96cbb-1368">最佳做法是，仅当同一个邮件应用程序在同一会话中添加了一个附件时，你才应使用该附件标识符来删除该附件。</span><span class="sxs-lookup"><span data-stu-id="96cbb-1368">As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session.</span></span> <span data-ttu-id="96cbb-1369">在 Outlook Web App 和适用于设备的 OWA 中，附件标识符只在同一个会话中才有效。</span><span class="sxs-lookup"><span data-stu-id="96cbb-1369">In Outlook Web App and OWA for Devices, the attachment identifier is valid only within the same session.</span></span> <span data-ttu-id="96cbb-1370">当用户关闭应用，或者如果用户开始在内嵌窗体中撰写，则随后弹出的窗体中继续在单独的窗口撰写时，会话即结束。</span><span class="sxs-lookup"><span data-stu-id="96cbb-1370">A session is over when the user closes the app, or if the user starts composing an inline form then subsequently pops out the form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="96cbb-1371">参数：</span><span class="sxs-lookup"><span data-stu-id="96cbb-1371">Parameters:</span></span>

|<span data-ttu-id="96cbb-1372">名称</span><span class="sxs-lookup"><span data-stu-id="96cbb-1372">Name</span></span>|<span data-ttu-id="96cbb-1373">类型</span><span class="sxs-lookup"><span data-stu-id="96cbb-1373">Type</span></span>|<span data-ttu-id="96cbb-1374">属性</span><span class="sxs-lookup"><span data-stu-id="96cbb-1374">Attributes</span></span>|<span data-ttu-id="96cbb-1375">说明</span><span class="sxs-lookup"><span data-stu-id="96cbb-1375">Description</span></span>|
|---|---|---|---|
|`attachmentId`|<span data-ttu-id="96cbb-1376">字符串</span><span class="sxs-lookup"><span data-stu-id="96cbb-1376">String</span></span>||<span data-ttu-id="96cbb-1377">要删除的附件的标识符。</span><span class="sxs-lookup"><span data-stu-id="96cbb-1377">The identifier of the attachment to remove.</span></span>|
|`options`|<span data-ttu-id="96cbb-1378">对象</span><span class="sxs-lookup"><span data-stu-id="96cbb-1378">Object</span></span>|<span data-ttu-id="96cbb-1379">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="96cbb-1379">&lt;optional&gt;</span></span>|<span data-ttu-id="96cbb-1380">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="96cbb-1380">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="96cbb-1381">对象</span><span class="sxs-lookup"><span data-stu-id="96cbb-1381">Object</span></span>|<span data-ttu-id="96cbb-1382">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="96cbb-1382">&lt;optional&gt;</span></span>|<span data-ttu-id="96cbb-1383">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="96cbb-1383">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="96cbb-1384">函数</span><span class="sxs-lookup"><span data-stu-id="96cbb-1384">function</span></span>|<span data-ttu-id="96cbb-1385">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="96cbb-1385">&lt;optional&gt;</span></span>|<span data-ttu-id="96cbb-1386">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="96cbb-1386">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="96cbb-1387">如果删除附件失败，`asyncResult.error` 属性将包含一个说明失败原因的错误代码。</span><span class="sxs-lookup"><span data-stu-id="96cbb-1387">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="96cbb-1388">错误</span><span class="sxs-lookup"><span data-stu-id="96cbb-1388">Errors</span></span>

|<span data-ttu-id="96cbb-1389">错误代码</span><span class="sxs-lookup"><span data-stu-id="96cbb-1389">Error code</span></span>|<span data-ttu-id="96cbb-1390">说明</span><span class="sxs-lookup"><span data-stu-id="96cbb-1390">Description</span></span>|
|------------|-------------|
|`InvalidAttachmentId`|<span data-ttu-id="96cbb-1391">附件标识符不存在。</span><span class="sxs-lookup"><span data-stu-id="96cbb-1391">The attachment identifier does not exist.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="96cbb-1392">要求</span><span class="sxs-lookup"><span data-stu-id="96cbb-1392">Requirements</span></span>

|<span data-ttu-id="96cbb-1393">要求</span><span class="sxs-lookup"><span data-stu-id="96cbb-1393">Requirement</span></span>|<span data-ttu-id="96cbb-1394">值</span><span class="sxs-lookup"><span data-stu-id="96cbb-1394">Value</span></span>|
|---|---|
|[<span data-ttu-id="96cbb-1395">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="96cbb-1395">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="96cbb-1396">1.1</span><span class="sxs-lookup"><span data-stu-id="96cbb-1396">1.1</span></span>|
|[<span data-ttu-id="96cbb-1397">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="96cbb-1397">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="96cbb-1398">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="96cbb-1398">ReadWriteItem</span></span>|
|[<span data-ttu-id="96cbb-1399">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="96cbb-1399">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="96cbb-1400">撰写</span><span class="sxs-lookup"><span data-stu-id="96cbb-1400">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="96cbb-1401">示例</span><span class="sxs-lookup"><span data-stu-id="96cbb-1401">Example</span></span>

<span data-ttu-id="96cbb-1402">以下代码删除包含标识符 '0' 的附件。</span><span class="sxs-lookup"><span data-stu-id="96cbb-1402">The following code removes an attachment with an identifier of '0'.</span></span>

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

####  <a name="removehandlerasynceventtype-options-callback"></a><span data-ttu-id="96cbb-1403">removeHandlerAsync(eventType, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="96cbb-1403">removeHandlerAsync(eventType, [options], [callback])</span></span>

<span data-ttu-id="96cbb-1404">删除受支持事件类型的事件处理程序。</span><span class="sxs-lookup"><span data-stu-id="96cbb-1404">Removes the event handlers for a supported event type.</span></span>

<span data-ttu-id="96cbb-1405">当前，支持的事件类型是 `Office.EventType.AttachmentsChanged`、`Office.EventType.AppointmentTimeChanged`、`Office.EventType.RecipientsChanged` 和 `Office.EventType.RecurrenceChanged`</span><span class="sxs-lookup"><span data-stu-id="96cbb-1405">Currently the supported event types are `Office.EventType.AttachmentsChanged`, `Office.EventType.AppointmentTimeChanged`, `Office.EventType.RecipientsChanged`, and `Office.EventType.RecurrenceChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="96cbb-1406">参数：</span><span class="sxs-lookup"><span data-stu-id="96cbb-1406">Parameters:</span></span>

| <span data-ttu-id="96cbb-1407">名称</span><span class="sxs-lookup"><span data-stu-id="96cbb-1407">Name</span></span> | <span data-ttu-id="96cbb-1408">类型</span><span class="sxs-lookup"><span data-stu-id="96cbb-1408">Type</span></span> | <span data-ttu-id="96cbb-1409">属性</span><span class="sxs-lookup"><span data-stu-id="96cbb-1409">Attributes</span></span> | <span data-ttu-id="96cbb-1410">说明</span><span class="sxs-lookup"><span data-stu-id="96cbb-1410">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="96cbb-1411">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="96cbb-1411">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="96cbb-1412">应撤销处理程序的事件。</span><span class="sxs-lookup"><span data-stu-id="96cbb-1412">The event that should revoke the handler.</span></span> |
| `options` | <span data-ttu-id="96cbb-1413">对象</span><span class="sxs-lookup"><span data-stu-id="96cbb-1413">Object</span></span> | <span data-ttu-id="96cbb-1414">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="96cbb-1414">&lt;optional&gt;</span></span> | <span data-ttu-id="96cbb-1415">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="96cbb-1415">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="96cbb-1416">对象</span><span class="sxs-lookup"><span data-stu-id="96cbb-1416">Object</span></span> | <span data-ttu-id="96cbb-1417">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="96cbb-1417">&lt;optional&gt;</span></span> | <span data-ttu-id="96cbb-1418">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="96cbb-1418">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="96cbb-1419">函数</span><span class="sxs-lookup"><span data-stu-id="96cbb-1419">function</span></span>| <span data-ttu-id="96cbb-1420">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="96cbb-1420">&lt;optional&gt;</span></span>|<span data-ttu-id="96cbb-1421">方法完成后，使用单个参数 `callback`（一个 [`asyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `AsyncResult` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="96cbb-1421">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="96cbb-1422">Requirements</span><span class="sxs-lookup"><span data-stu-id="96cbb-1422">Requirements</span></span>

|<span data-ttu-id="96cbb-1423">要求</span><span class="sxs-lookup"><span data-stu-id="96cbb-1423">Requirement</span></span>| <span data-ttu-id="96cbb-1424">值</span><span class="sxs-lookup"><span data-stu-id="96cbb-1424">Value</span></span>|
|---|---|
|[<span data-ttu-id="96cbb-1425">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="96cbb-1425">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="96cbb-1426">1.7</span><span class="sxs-lookup"><span data-stu-id="96cbb-1426">1.7</span></span> |
|[<span data-ttu-id="96cbb-1427">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="96cbb-1427">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="96cbb-1428">ReadItem</span><span class="sxs-lookup"><span data-stu-id="96cbb-1428">ReadItem</span></span> |
|[<span data-ttu-id="96cbb-1429">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="96cbb-1429">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="96cbb-1430">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="96cbb-1430">Compose or read</span></span> |

####  <a name="saveasyncoptions-callback"></a><span data-ttu-id="96cbb-1431">saveAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="96cbb-1431">saveAsync([options], callback)</span></span>

<span data-ttu-id="96cbb-1432">异步保存项目。</span><span class="sxs-lookup"><span data-stu-id="96cbb-1432">Asynchronously saves an item.</span></span>

<span data-ttu-id="96cbb-p181">调用时，此方法将当前邮件保存为草稿，并通过回调方法返回项目 ID。在 Outlook Web App 或 Outlook 联机模式下，该项目被保存到服务器中。在 Outlook 缓存模式下，该项目被保存到本地缓存中。</span><span class="sxs-lookup"><span data-stu-id="96cbb-p181">When invoked, this method saves the current message as a draft and returns the item id via the callback method. In Outlook Web App or Outlook in online mode, the item is saved to the server. In Outlook in cached mode, the item is saved to the local cache.</span></span>

> [!NOTE]
> <span data-ttu-id="96cbb-1436">如果加载项在撰写模式下对某个项目调用 `saveAsync` 来获得 `itemId`，以便与 EWS 或 REST API 一同使用，请注意，当 Outlook 处于高速缓存模式时，可能需要一段时间项目才能真正同步到服务器。</span><span class="sxs-lookup"><span data-stu-id="96cbb-1436">If your add-in calls `saveAsync` on an item in compose mode in order to get an `itemId` to use with EWS or the REST API, be aware that when Outlook is in cached mode, it may take some time before the item is actually synced to the server.</span></span> <span data-ttu-id="96cbb-1437">在项目同步前，使用 `itemId` 将返回一个错误。</span><span class="sxs-lookup"><span data-stu-id="96cbb-1437">Until the item is synced, using the `itemId` will return an error.</span></span>

<span data-ttu-id="96cbb-p183">由于约会没有草稿状态，如果以撰写模式在约会中调用 `saveAsync`，则该项将被保存为用户日历中的正常约会。对于之前未保存过的新约会，则不会发送邀请。保存现有约会将向添加或删除的与会者发送更新。</span><span class="sxs-lookup"><span data-stu-id="96cbb-p183">Since appointments have no draft state, if `saveAsync` is called on an appointment in compose mode, the item will be saved as a normal appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent. Saving an existing appointment will send an update to added or removed attendees.</span></span>

> [!NOTE]
> <span data-ttu-id="96cbb-1441">以下客户端在撰写模式下对约会上的 `saveAsync` 具有不同的行为：</span><span class="sxs-lookup"><span data-stu-id="96cbb-1441">The following clients have different behavior for `saveAsync` on appointments in compose mode:</span></span>
>
> - <span data-ttu-id="96cbb-1442">Mac Outlook 不支持在撰写模式下对会议执行 `saveAsync` 操作。</span><span class="sxs-lookup"><span data-stu-id="96cbb-1442">Mac Outlook does not support `saveAsync` on a meeting in compose mode.</span></span> <span data-ttu-id="96cbb-1443">对 Mac Outlook 中的会议调用 `saveAsync` 将会返回一个错误。</span><span class="sxs-lookup"><span data-stu-id="96cbb-1443">Calling `saveAsync` on a meeting in Mac Outlook will return an error.</span></span>
> - <span data-ttu-id="96cbb-1444">在撰写模式下的约会上调用 `saveAsync` 时，Outlook 网页版始终发送邀请或更新。</span><span class="sxs-lookup"><span data-stu-id="96cbb-1444">Outlook on the web always sends an invitation or update when `saveAsync` is called on an appointment in compose mode.</span></span>

##### <a name="parameters"></a><span data-ttu-id="96cbb-1445">参数：</span><span class="sxs-lookup"><span data-stu-id="96cbb-1445">Parameters:</span></span>

|<span data-ttu-id="96cbb-1446">名称</span><span class="sxs-lookup"><span data-stu-id="96cbb-1446">Name</span></span>|<span data-ttu-id="96cbb-1447">类型</span><span class="sxs-lookup"><span data-stu-id="96cbb-1447">Type</span></span>|<span data-ttu-id="96cbb-1448">属性</span><span class="sxs-lookup"><span data-stu-id="96cbb-1448">Attributes</span></span>|<span data-ttu-id="96cbb-1449">说明</span><span class="sxs-lookup"><span data-stu-id="96cbb-1449">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="96cbb-1450">对象</span><span class="sxs-lookup"><span data-stu-id="96cbb-1450">Object</span></span>|<span data-ttu-id="96cbb-1451">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="96cbb-1451">&lt;optional&gt;</span></span>|<span data-ttu-id="96cbb-1452">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="96cbb-1452">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="96cbb-1453">对象</span><span class="sxs-lookup"><span data-stu-id="96cbb-1453">Object</span></span>|<span data-ttu-id="96cbb-1454">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="96cbb-1454">&lt;optional&gt;</span></span>|<span data-ttu-id="96cbb-1455">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="96cbb-1455">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="96cbb-1456">函数</span><span class="sxs-lookup"><span data-stu-id="96cbb-1456">function</span></span>||<span data-ttu-id="96cbb-1457">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="96cbb-1457">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="96cbb-1458">如果成功，该项目标识符将在 `asyncResult.value` 属性中提供。</span><span class="sxs-lookup"><span data-stu-id="96cbb-1458">On success, the item identifier is provided in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="96cbb-1459">要求</span><span class="sxs-lookup"><span data-stu-id="96cbb-1459">Requirements</span></span>

|<span data-ttu-id="96cbb-1460">要求</span><span class="sxs-lookup"><span data-stu-id="96cbb-1460">Requirement</span></span>|<span data-ttu-id="96cbb-1461">值</span><span class="sxs-lookup"><span data-stu-id="96cbb-1461">Value</span></span>|
|---|---|
|[<span data-ttu-id="96cbb-1462">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="96cbb-1462">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="96cbb-1463">1.3</span><span class="sxs-lookup"><span data-stu-id="96cbb-1463">1.3</span></span>|
|[<span data-ttu-id="96cbb-1464">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="96cbb-1464">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="96cbb-1465">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="96cbb-1465">ReadWriteItem</span></span>|
|[<span data-ttu-id="96cbb-1466">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="96cbb-1466">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="96cbb-1467">撰写</span><span class="sxs-lookup"><span data-stu-id="96cbb-1467">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="96cbb-1468">示例</span><span class="sxs-lookup"><span data-stu-id="96cbb-1468">Examples</span></span>

```javascript
Office.context.mailbox.item.saveAsync(
  function callback(result) {
    // Process the result
  });
```

<span data-ttu-id="96cbb-p185">下面是传递给回调函数的 `result` 参数的示例。`value` 属性包含的项目的项目 ID。</span><span class="sxs-lookup"><span data-stu-id="96cbb-p185">The following is an example of the `result` parameter passed to the callback function. The `value` property contains the item ID of the item.</span></span>

```json
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

####  <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="96cbb-1471">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="96cbb-1471">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="96cbb-1472">以异步方式将数据插入到邮件的正文或主题中。</span><span class="sxs-lookup"><span data-stu-id="96cbb-1472">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="96cbb-p186">`setSelectedDataAsync` 方法将指定字符串插入到项目主题或正文的光标位置，或者，如果在编辑器中已选择文本，则该方法将替换选择的文本。如果光标不在正文或主题字段中，则返回错误。插入之后，光标会位于插入内容的末尾。</span><span class="sxs-lookup"><span data-stu-id="96cbb-p186">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="96cbb-1476">参数：</span><span class="sxs-lookup"><span data-stu-id="96cbb-1476">Parameters:</span></span>

|<span data-ttu-id="96cbb-1477">名称</span><span class="sxs-lookup"><span data-stu-id="96cbb-1477">Name</span></span>|<span data-ttu-id="96cbb-1478">类型</span><span class="sxs-lookup"><span data-stu-id="96cbb-1478">Type</span></span>|<span data-ttu-id="96cbb-1479">属性</span><span class="sxs-lookup"><span data-stu-id="96cbb-1479">Attributes</span></span>|<span data-ttu-id="96cbb-1480">说明</span><span class="sxs-lookup"><span data-stu-id="96cbb-1480">Description</span></span>|
|---|---|---|---|
|`data`|<span data-ttu-id="96cbb-1481">字符串</span><span class="sxs-lookup"><span data-stu-id="96cbb-1481">String</span></span>||<span data-ttu-id="96cbb-p187">要插入的数据。数据不得超过 1,000,000 个字符。如果传入的数据超过 1,000,000 个字符，则会引发 `ArgumentOutOfRange` 异常。</span><span class="sxs-lookup"><span data-stu-id="96cbb-p187">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`|<span data-ttu-id="96cbb-1485">Object</span><span class="sxs-lookup"><span data-stu-id="96cbb-1485">Object</span></span>|<span data-ttu-id="96cbb-1486">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="96cbb-1486">&lt;optional&gt;</span></span>|<span data-ttu-id="96cbb-1487">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="96cbb-1487">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="96cbb-1488">对象</span><span class="sxs-lookup"><span data-stu-id="96cbb-1488">Object</span></span>|<span data-ttu-id="96cbb-1489">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="96cbb-1489">&lt;optional&gt;</span></span>|<span data-ttu-id="96cbb-1490">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="96cbb-1490">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`|[<span data-ttu-id="96cbb-1491">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="96cbb-1491">Office.CoercionType</span></span>](office.md#coerciontype-string)|<span data-ttu-id="96cbb-1492">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="96cbb-1492">&lt;optional&gt;</span></span>|<span data-ttu-id="96cbb-p188">如果为 `text`，则在 Outlook Web App 和 Outlook 中应用当前样式。如果该字段是 HTML 编辑器，则仅插入文本数据，即使数据为 HTML。</span><span class="sxs-lookup"><span data-stu-id="96cbb-p188">If `text`, the current style is applied in Outlook Web App and Outlook. If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="96cbb-p189">如果 `html` 和该字段支持 HTML（主题不支持），则在 Outlook Web App 中应用当前样式，而在 Outlook 中应用默认样式。如果该字段是文本字段，则返回 `InvalidDataFormat` 错误。</span><span class="sxs-lookup"><span data-stu-id="96cbb-p189">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook Web App and the default style is applied in Outlook. If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="96cbb-1497">如果未设置 `coercionType`，则结果取决于该字段：如果该字段是 HTML，则使用 HTML；如果该字段是文本，则使用纯文本。</span><span class="sxs-lookup"><span data-stu-id="96cbb-1497">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`|<span data-ttu-id="96cbb-1498">function</span><span class="sxs-lookup"><span data-stu-id="96cbb-1498">function</span></span>||<span data-ttu-id="96cbb-1499">方法完成后，使用单个参数 `callback`（一个 [`asyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `AsyncResult` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="96cbb-1499">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="96cbb-1500">Requirements</span><span class="sxs-lookup"><span data-stu-id="96cbb-1500">Requirements</span></span>

|<span data-ttu-id="96cbb-1501">要求</span><span class="sxs-lookup"><span data-stu-id="96cbb-1501">Requirement</span></span>|<span data-ttu-id="96cbb-1502">值</span><span class="sxs-lookup"><span data-stu-id="96cbb-1502">Value</span></span>|
|---|---|
|[<span data-ttu-id="96cbb-1503">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="96cbb-1503">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="96cbb-1504">1.2</span><span class="sxs-lookup"><span data-stu-id="96cbb-1504">1.2</span></span>|
|[<span data-ttu-id="96cbb-1505">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="96cbb-1505">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="96cbb-1506">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="96cbb-1506">ReadWriteItem</span></span>|
|[<span data-ttu-id="96cbb-1507">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="96cbb-1507">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="96cbb-1508">撰写</span><span class="sxs-lookup"><span data-stu-id="96cbb-1508">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="96cbb-1509">示例</span><span class="sxs-lookup"><span data-stu-id="96cbb-1509">Example</span></span>

```javascript
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```
