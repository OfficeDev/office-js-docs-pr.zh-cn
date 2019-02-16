---
title: "\"context\"-\"邮箱\"。项目-要求集1。7"
description: ''
ms.date: 02/15/2019
localization_priority: Normal
ms.openlocfilehash: 87ac31deaabda5fe1d2ca024972e326cfa168ea7
ms.sourcegitcommit: f26778b596b6b022814c39601485ff676ed4e2fa
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 02/16/2019
ms.locfileid: "30068250"
---
# <a name="item"></a><span data-ttu-id="59995-102">item</span><span class="sxs-lookup"><span data-stu-id="59995-102">item</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmditem"></a><span data-ttu-id="59995-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span><span class="sxs-lookup"><span data-stu-id="59995-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span></span>

<span data-ttu-id="59995-p101">`item` 命名空间用于访问当前选定的邮件、会议请求或约会。可以通过使用 [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlook17officemailboxenumsitemtype) 属性确定 `item` 的类型。</span><span class="sxs-lookup"><span data-stu-id="59995-p101">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlook17officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="59995-106">Requirements</span><span class="sxs-lookup"><span data-stu-id="59995-106">Requirements</span></span>

|<span data-ttu-id="59995-107">要求</span><span class="sxs-lookup"><span data-stu-id="59995-107">Requirement</span></span>|<span data-ttu-id="59995-108">值</span><span class="sxs-lookup"><span data-stu-id="59995-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="59995-109">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="59995-109">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="59995-110">1.0</span><span class="sxs-lookup"><span data-stu-id="59995-110">1.0</span></span>|
|[<span data-ttu-id="59995-111">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="59995-111">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="59995-112">受限</span><span class="sxs-lookup"><span data-stu-id="59995-112">Restricted</span></span>|
|[<span data-ttu-id="59995-113">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="59995-113">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="59995-114">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="59995-114">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="59995-115">成员和方法</span><span class="sxs-lookup"><span data-stu-id="59995-115">Members and methods</span></span>

| <span data-ttu-id="59995-116">成员</span><span class="sxs-lookup"><span data-stu-id="59995-116">Member</span></span> | <span data-ttu-id="59995-117">类型</span><span class="sxs-lookup"><span data-stu-id="59995-117">Type</span></span> |
|--------|------|
| [<span data-ttu-id="59995-118">attachments</span><span class="sxs-lookup"><span data-stu-id="59995-118">attachments</span></span>](#attachments-arrayattachmentdetailsjavascriptapioutlook17officeattachmentdetails) | <span data-ttu-id="59995-119">成员</span><span class="sxs-lookup"><span data-stu-id="59995-119">Member</span></span> |
| [<span data-ttu-id="59995-120">bcc</span><span class="sxs-lookup"><span data-stu-id="59995-120">bcc</span></span>](#bcc-recipientsjavascriptapioutlook17officerecipients) | <span data-ttu-id="59995-121">成员</span><span class="sxs-lookup"><span data-stu-id="59995-121">Member</span></span> |
| [<span data-ttu-id="59995-122">body</span><span class="sxs-lookup"><span data-stu-id="59995-122">body</span></span>](#body-bodyjavascriptapioutlook17officebody) | <span data-ttu-id="59995-123">成员</span><span class="sxs-lookup"><span data-stu-id="59995-123">Member</span></span> |
| [<span data-ttu-id="59995-124">cc</span><span class="sxs-lookup"><span data-stu-id="59995-124">cc</span></span>](#cc-arrayemailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsrecipientsjavascriptapioutlook17officerecipients) | <span data-ttu-id="59995-125">成员</span><span class="sxs-lookup"><span data-stu-id="59995-125">Member</span></span> |
| [<span data-ttu-id="59995-126">conversationId</span><span class="sxs-lookup"><span data-stu-id="59995-126">conversationId</span></span>](#nullable-conversationid-string) | <span data-ttu-id="59995-127">成员</span><span class="sxs-lookup"><span data-stu-id="59995-127">Member</span></span> |
| [<span data-ttu-id="59995-128">dateTimeCreated</span><span class="sxs-lookup"><span data-stu-id="59995-128">dateTimeCreated</span></span>](#datetimecreated-date) | <span data-ttu-id="59995-129">成员</span><span class="sxs-lookup"><span data-stu-id="59995-129">Member</span></span> |
| [<span data-ttu-id="59995-130">dateTimeModified</span><span class="sxs-lookup"><span data-stu-id="59995-130">dateTimeModified</span></span>](#datetimemodified-date) | <span data-ttu-id="59995-131">成员</span><span class="sxs-lookup"><span data-stu-id="59995-131">Member</span></span> |
| [<span data-ttu-id="59995-132">end</span><span class="sxs-lookup"><span data-stu-id="59995-132">end</span></span>](#end-datetimejavascriptapioutlook17officetime) | <span data-ttu-id="59995-133">成员</span><span class="sxs-lookup"><span data-stu-id="59995-133">Member</span></span> |
| [<span data-ttu-id="59995-134">from</span><span class="sxs-lookup"><span data-stu-id="59995-134">from</span></span>](#from-emailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsfromjavascriptapioutlook17officefrom) | <span data-ttu-id="59995-135">成员</span><span class="sxs-lookup"><span data-stu-id="59995-135">Member</span></span> |
| [<span data-ttu-id="59995-136">internetMessageId</span><span class="sxs-lookup"><span data-stu-id="59995-136">internetMessageId</span></span>](#internetmessageid-string) | <span data-ttu-id="59995-137">成员</span><span class="sxs-lookup"><span data-stu-id="59995-137">Member</span></span> |
| [<span data-ttu-id="59995-138">itemClass</span><span class="sxs-lookup"><span data-stu-id="59995-138">itemClass</span></span>](#itemclass-string) | <span data-ttu-id="59995-139">成员</span><span class="sxs-lookup"><span data-stu-id="59995-139">Member</span></span> |
| [<span data-ttu-id="59995-140">itemId</span><span class="sxs-lookup"><span data-stu-id="59995-140">itemId</span></span>](#nullable-itemid-string) | <span data-ttu-id="59995-141">成员</span><span class="sxs-lookup"><span data-stu-id="59995-141">Member</span></span> |
| [<span data-ttu-id="59995-142">itemType</span><span class="sxs-lookup"><span data-stu-id="59995-142">itemType</span></span>](#itemtype-officemailboxenumsitemtypejavascriptapioutlook17officemailboxenumsitemtype) | <span data-ttu-id="59995-143">成员</span><span class="sxs-lookup"><span data-stu-id="59995-143">Member</span></span> |
| [<span data-ttu-id="59995-144">location</span><span class="sxs-lookup"><span data-stu-id="59995-144">location</span></span>](#location-stringlocationjavascriptapioutlook17officelocation) | <span data-ttu-id="59995-145">成员</span><span class="sxs-lookup"><span data-stu-id="59995-145">Member</span></span> |
| [<span data-ttu-id="59995-146">normalizedSubject</span><span class="sxs-lookup"><span data-stu-id="59995-146">normalizedSubject</span></span>](#normalizedsubject-string) | <span data-ttu-id="59995-147">成员</span><span class="sxs-lookup"><span data-stu-id="59995-147">Member</span></span> |
| [<span data-ttu-id="59995-148">notificationMessages</span><span class="sxs-lookup"><span data-stu-id="59995-148">notificationMessages</span></span>](#notificationmessages-notificationmessagesjavascriptapioutlook17officenotificationmessages) | <span data-ttu-id="59995-149">成员</span><span class="sxs-lookup"><span data-stu-id="59995-149">Member</span></span> |
| [<span data-ttu-id="59995-150">optionalAttendees</span><span class="sxs-lookup"><span data-stu-id="59995-150">optionalAttendees</span></span>](#optionalattendees-arrayemailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsrecipientsjavascriptapioutlook17officerecipients) | <span data-ttu-id="59995-151">成员</span><span class="sxs-lookup"><span data-stu-id="59995-151">Member</span></span> |
| [<span data-ttu-id="59995-152">organizer</span><span class="sxs-lookup"><span data-stu-id="59995-152">organizer</span></span>](#organizer-emailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsorganizerjavascriptapioutlook17officeorganizer) | <span data-ttu-id="59995-153">成员</span><span class="sxs-lookup"><span data-stu-id="59995-153">Member</span></span> |
| [<span data-ttu-id="59995-154">recurrence</span><span class="sxs-lookup"><span data-stu-id="59995-154">recurrence</span></span>](#nullable-recurrence-recurrencejavascriptapioutlook17officerecurrence) | <span data-ttu-id="59995-155">Member</span><span class="sxs-lookup"><span data-stu-id="59995-155">Member</span></span> |
| [<span data-ttu-id="59995-156">requiredAttendees</span><span class="sxs-lookup"><span data-stu-id="59995-156">requiredAttendees</span></span>](#requiredattendees-arrayemailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsrecipientsjavascriptapioutlook17officerecipients) | <span data-ttu-id="59995-157">成员</span><span class="sxs-lookup"><span data-stu-id="59995-157">Member</span></span> |
| [<span data-ttu-id="59995-158">sender</span><span class="sxs-lookup"><span data-stu-id="59995-158">sender</span></span>](#sender-emailaddressdetailsjavascriptapioutlook17officeemailaddressdetails) | <span data-ttu-id="59995-159">成员</span><span class="sxs-lookup"><span data-stu-id="59995-159">Member</span></span> |
| [<span data-ttu-id="59995-160">seriesId</span><span class="sxs-lookup"><span data-stu-id="59995-160">seriesId</span></span>](#nullable-seriesid-string) | <span data-ttu-id="59995-161">成员</span><span class="sxs-lookup"><span data-stu-id="59995-161">Member</span></span> |
| [<span data-ttu-id="59995-162">start</span><span class="sxs-lookup"><span data-stu-id="59995-162">start</span></span>](#start-datetimejavascriptapioutlook17officetime) | <span data-ttu-id="59995-163">成员</span><span class="sxs-lookup"><span data-stu-id="59995-163">Member</span></span> |
| [<span data-ttu-id="59995-164">subject</span><span class="sxs-lookup"><span data-stu-id="59995-164">subject</span></span>](#subject-stringsubjectjavascriptapioutlook17officesubject) | <span data-ttu-id="59995-165">成员</span><span class="sxs-lookup"><span data-stu-id="59995-165">Member</span></span> |
| [<span data-ttu-id="59995-166">to</span><span class="sxs-lookup"><span data-stu-id="59995-166">to</span></span>](#to-arrayemailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsrecipientsjavascriptapioutlook17officerecipients) | <span data-ttu-id="59995-167">成员</span><span class="sxs-lookup"><span data-stu-id="59995-167">Member</span></span> |
| [<span data-ttu-id="59995-168">addFileAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="59995-168">addFileAttachmentAsync</span></span>](#addfileattachmentasyncuri-attachmentname-options-callback) | <span data-ttu-id="59995-169">方法</span><span class="sxs-lookup"><span data-stu-id="59995-169">Method</span></span> |
| [<span data-ttu-id="59995-170">addHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="59995-170">addHandlerAsync</span></span>](#addhandlerasynceventtype-handler-options-callback) | <span data-ttu-id="59995-171">方法</span><span class="sxs-lookup"><span data-stu-id="59995-171">Method</span></span> |
| [<span data-ttu-id="59995-172">addItemAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="59995-172">addItemAttachmentAsync</span></span>](#additemattachmentasyncitemid-attachmentname-options-callback) | <span data-ttu-id="59995-173">方法</span><span class="sxs-lookup"><span data-stu-id="59995-173">Method</span></span> |
| [<span data-ttu-id="59995-174">close</span><span class="sxs-lookup"><span data-stu-id="59995-174">close</span></span>](#close) | <span data-ttu-id="59995-175">方法</span><span class="sxs-lookup"><span data-stu-id="59995-175">Method</span></span> |
| [<span data-ttu-id="59995-176">displayReplyAllForm</span><span class="sxs-lookup"><span data-stu-id="59995-176">displayReplyAllForm</span></span>](#displayreplyallformformdata-callback) | <span data-ttu-id="59995-177">方法</span><span class="sxs-lookup"><span data-stu-id="59995-177">Method</span></span> |
| [<span data-ttu-id="59995-178">displayReplyForm</span><span class="sxs-lookup"><span data-stu-id="59995-178">displayReplyForm</span></span>](#displayreplyformformdata-callback) | <span data-ttu-id="59995-179">方法</span><span class="sxs-lookup"><span data-stu-id="59995-179">Method</span></span> |
| [<span data-ttu-id="59995-180">getEntities</span><span class="sxs-lookup"><span data-stu-id="59995-180">getEntities</span></span>](#getentities--entitiesjavascriptapioutlook17officeentities) | <span data-ttu-id="59995-181">方法</span><span class="sxs-lookup"><span data-stu-id="59995-181">Method</span></span> |
| [<span data-ttu-id="59995-182">getEntitiesByType</span><span class="sxs-lookup"><span data-stu-id="59995-182">getEntitiesByType</span></span>](#getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlook17officecontactmeetingsuggestionjavascriptapioutlook17officemeetingsuggestionphonenumberjavascriptapioutlook17officephonenumbertasksuggestionjavascriptapioutlook17officetasksuggestion) | <span data-ttu-id="59995-183">方法</span><span class="sxs-lookup"><span data-stu-id="59995-183">Method</span></span> |
| [<span data-ttu-id="59995-184">getFilteredEntitiesByName</span><span class="sxs-lookup"><span data-stu-id="59995-184">getFilteredEntitiesByName</span></span>](#getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlook17officecontactmeetingsuggestionjavascriptapioutlook17officemeetingsuggestionphonenumberjavascriptapioutlook17officephonenumbertasksuggestionjavascriptapioutlook17officetasksuggestion) | <span data-ttu-id="59995-185">方法</span><span class="sxs-lookup"><span data-stu-id="59995-185">Method</span></span> |
| [<span data-ttu-id="59995-186">getRegExMatches</span><span class="sxs-lookup"><span data-stu-id="59995-186">getRegExMatches</span></span>](#getregexmatches--object) | <span data-ttu-id="59995-187">方法</span><span class="sxs-lookup"><span data-stu-id="59995-187">Method</span></span> |
| [<span data-ttu-id="59995-188">getRegExMatchesByName</span><span class="sxs-lookup"><span data-stu-id="59995-188">getRegExMatchesByName</span></span>](#getregexmatchesbynamename--nullable-array-string-) | <span data-ttu-id="59995-189">方法</span><span class="sxs-lookup"><span data-stu-id="59995-189">Method</span></span> |
| [<span data-ttu-id="59995-190">getSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="59995-190">getSelectedDataAsync</span></span>](#getselecteddataasynccoerciontype-options-callback--string) | <span data-ttu-id="59995-191">方法</span><span class="sxs-lookup"><span data-stu-id="59995-191">Method</span></span> |
| [<span data-ttu-id="59995-192">getSelectedEntities</span><span class="sxs-lookup"><span data-stu-id="59995-192">getSelectedEntities</span></span>](#getselectedentities--entitiesjavascriptapioutlook17officeentities) | <span data-ttu-id="59995-193">方法</span><span class="sxs-lookup"><span data-stu-id="59995-193">Method</span></span> |
| [<span data-ttu-id="59995-194">getSelectedRegExMatches</span><span class="sxs-lookup"><span data-stu-id="59995-194">getSelectedRegExMatches</span></span>](#getselectedregexmatches--object) | <span data-ttu-id="59995-195">方法</span><span class="sxs-lookup"><span data-stu-id="59995-195">Method</span></span> |
| [<span data-ttu-id="59995-196">loadCustomPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="59995-196">loadCustomPropertiesAsync</span></span>](#loadcustompropertiesasynccallback-usercontext) | <span data-ttu-id="59995-197">方法</span><span class="sxs-lookup"><span data-stu-id="59995-197">Method</span></span> |
| [<span data-ttu-id="59995-198">removeAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="59995-198">removeAttachmentAsync</span></span>](#removeattachmentasyncattachmentid-options-callback) | <span data-ttu-id="59995-199">方法</span><span class="sxs-lookup"><span data-stu-id="59995-199">Method</span></span> |
| [<span data-ttu-id="59995-200">removeHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="59995-200">removeHandlerAsync</span></span>](#removehandlerasynceventtype-options-callback) | <span data-ttu-id="59995-201">方法</span><span class="sxs-lookup"><span data-stu-id="59995-201">Method</span></span> |
| [<span data-ttu-id="59995-202">saveAsync</span><span class="sxs-lookup"><span data-stu-id="59995-202">saveAsync</span></span>](#saveasyncoptions-callback) | <span data-ttu-id="59995-203">方法</span><span class="sxs-lookup"><span data-stu-id="59995-203">Method</span></span> |
| [<span data-ttu-id="59995-204">setSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="59995-204">setSelectedDataAsync</span></span>](#setselecteddataasyncdata-options-callback) | <span data-ttu-id="59995-205">方法</span><span class="sxs-lookup"><span data-stu-id="59995-205">Method</span></span> |

### <a name="example"></a><span data-ttu-id="59995-206">示例</span><span class="sxs-lookup"><span data-stu-id="59995-206">Example</span></span>

<span data-ttu-id="59995-207">以下 JavaScript 代码示例显示了如何访问 Outlook 中当前项目的 `subject` 属性。</span><span class="sxs-lookup"><span data-stu-id="59995-207">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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

### <a name="members"></a><span data-ttu-id="59995-208">成员</span><span class="sxs-lookup"><span data-stu-id="59995-208">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlook17officeattachmentdetails"></a><span data-ttu-id="59995-209">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_7/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="59995-209">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_7/office.attachmentdetails)></span></span>

<span data-ttu-id="59995-p102">获取项目的附件数组。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="59995-p102">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="59995-212">由于潜在的安全问题，某些类型的文件会受到 Outlook 阻止，并且不会返回。</span><span class="sxs-lookup"><span data-stu-id="59995-212">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="59995-213">如需了解更多信息，请参阅 [Outlook 中阻止的附件](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519)。</span><span class="sxs-lookup"><span data-stu-id="59995-213">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="59995-214">类型</span><span class="sxs-lookup"><span data-stu-id="59995-214">Type</span></span>

*   <span data-ttu-id="59995-215">Array.<[AttachmentDetails](/javascript/api/outlook_1_7/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="59995-215">Array.<[AttachmentDetails](/javascript/api/outlook_1_7/office.attachmentdetails)></span></span>

##### <a name="requirements"></a><span data-ttu-id="59995-216">Requirements</span><span class="sxs-lookup"><span data-stu-id="59995-216">Requirements</span></span>

|<span data-ttu-id="59995-217">要求</span><span class="sxs-lookup"><span data-stu-id="59995-217">Requirement</span></span>|<span data-ttu-id="59995-218">值</span><span class="sxs-lookup"><span data-stu-id="59995-218">Value</span></span>|
|---|---|
|[<span data-ttu-id="59995-219">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="59995-219">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="59995-220">1.0</span><span class="sxs-lookup"><span data-stu-id="59995-220">1.0</span></span>|
|[<span data-ttu-id="59995-221">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="59995-221">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="59995-222">ReadItem</span><span class="sxs-lookup"><span data-stu-id="59995-222">ReadItem</span></span>|
|[<span data-ttu-id="59995-223">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="59995-223">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="59995-224">阅读</span><span class="sxs-lookup"><span data-stu-id="59995-224">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="59995-225">示例</span><span class="sxs-lookup"><span data-stu-id="59995-225">Example</span></span>

<span data-ttu-id="59995-226">以下代码使用当前项目上所有附件的详细信息构成 HTML 字符串。</span><span class="sxs-lookup"><span data-stu-id="59995-226">The following code builds an HTML string with details of all attachments on the current item.</span></span>

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

####  <a name="bcc-recipientsjavascriptapioutlook17officerecipients"></a><span data-ttu-id="59995-227">bcc :[Recipients](/javascript/api/outlook_1_7/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="59995-227">bcc :[Recipients](/javascript/api/outlook_1_7/office.recipients)</span></span>

<span data-ttu-id="59995-228">获取对象，该对象提供用于获取或更新邮件的密件抄送 (Bcc) 行上的收件人的方法。</span><span class="sxs-lookup"><span data-stu-id="59995-228">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="59995-229">仅限撰写模式。</span><span class="sxs-lookup"><span data-stu-id="59995-229">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="59995-230">类型</span><span class="sxs-lookup"><span data-stu-id="59995-230">Type</span></span>

*   [<span data-ttu-id="59995-231">收件人</span><span class="sxs-lookup"><span data-stu-id="59995-231">Recipients</span></span>](/javascript/api/outlook_1_7/office.recipients)

##### <a name="requirements"></a><span data-ttu-id="59995-232">Requirements</span><span class="sxs-lookup"><span data-stu-id="59995-232">Requirements</span></span>

|<span data-ttu-id="59995-233">要求</span><span class="sxs-lookup"><span data-stu-id="59995-233">Requirement</span></span>|<span data-ttu-id="59995-234">值</span><span class="sxs-lookup"><span data-stu-id="59995-234">Value</span></span>|
|---|---|
|[<span data-ttu-id="59995-235">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="59995-235">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="59995-236">1.1</span><span class="sxs-lookup"><span data-stu-id="59995-236">1.1</span></span>|
|[<span data-ttu-id="59995-237">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="59995-237">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="59995-238">ReadItem</span><span class="sxs-lookup"><span data-stu-id="59995-238">ReadItem</span></span>|
|[<span data-ttu-id="59995-239">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="59995-239">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="59995-240">撰写</span><span class="sxs-lookup"><span data-stu-id="59995-240">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="59995-241">示例</span><span class="sxs-lookup"><span data-stu-id="59995-241">Example</span></span>

```javascript
Office.context.mailbox.item.bcc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.bcc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.bcc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfBccRecipients = asyncResult.value;
}
```

####  <a name="body-bodyjavascriptapioutlook17officebody"></a><span data-ttu-id="59995-242">body :[Body](/javascript/api/outlook_1_7/office.body)</span><span class="sxs-lookup"><span data-stu-id="59995-242">body :[Body](/javascript/api/outlook_1_7/office.body)</span></span>

<span data-ttu-id="59995-243">获取一个提供用于处理项目正文的方法的对象。</span><span class="sxs-lookup"><span data-stu-id="59995-243">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="59995-244">类型</span><span class="sxs-lookup"><span data-stu-id="59995-244">Type</span></span>

*   [<span data-ttu-id="59995-245">Body</span><span class="sxs-lookup"><span data-stu-id="59995-245">Body</span></span>](/javascript/api/outlook_1_7/office.body)

##### <a name="requirements"></a><span data-ttu-id="59995-246">要求</span><span class="sxs-lookup"><span data-stu-id="59995-246">Requirements</span></span>

|<span data-ttu-id="59995-247">要求</span><span class="sxs-lookup"><span data-stu-id="59995-247">Requirement</span></span>|<span data-ttu-id="59995-248">值</span><span class="sxs-lookup"><span data-stu-id="59995-248">Value</span></span>|
|---|---|
|[<span data-ttu-id="59995-249">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="59995-249">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="59995-250">1.1</span><span class="sxs-lookup"><span data-stu-id="59995-250">1.1</span></span>|
|[<span data-ttu-id="59995-251">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="59995-251">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="59995-252">ReadItem</span><span class="sxs-lookup"><span data-stu-id="59995-252">ReadItem</span></span>|
|[<span data-ttu-id="59995-253">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="59995-253">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="59995-254">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="59995-254">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="59995-255">示例</span><span class="sxs-lookup"><span data-stu-id="59995-255">Example</span></span>

<span data-ttu-id="59995-256">本示例获取纯文本格式的邮件正文。</span><span class="sxs-lookup"><span data-stu-id="59995-256">This example gets the body of the message in plain text.</span></span>

```javascript
Office.context.mailbox.item.body.getAsync(
  "text",
  { asyncContext: "This is passed to the callback" },
  function callback(result) {
    // Do something with the result.
  });

```

<span data-ttu-id="59995-257">下面是传递给回调函数的 result 参数的示例。</span><span class="sxs-lookup"><span data-stu-id="59995-257">The following is an example of the result parameter passed to the callback function.</span></span>

```json
{
  "value": "TEXT of whole body (including threads below)",
  "status": "succeeded",
  "asyncContext": "This is passed to the callback"
}
```

####  <a name="cc-arrayemailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsrecipientsjavascriptapioutlook17officerecipients"></a><span data-ttu-id="59995-258">cc :Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_7/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="59995-258">cc :Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_7/office.recipients)</span></span>

<span data-ttu-id="59995-259">提供对邮件的抄送 (Cc) 收件人的访问权限。</span><span class="sxs-lookup"><span data-stu-id="59995-259">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="59995-260">对象的类型和访问级别取决于当前项目的模式。</span><span class="sxs-lookup"><span data-stu-id="59995-260">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="59995-261">阅读模式</span><span class="sxs-lookup"><span data-stu-id="59995-261">Read mode</span></span>

<span data-ttu-id="59995-p106">`cc` 属性返回包含邮件的**抄送**行上所列的每个收件人的 `EmailAddressDetails` 对象的数组。集合上限为 100 个成员。</span><span class="sxs-lookup"><span data-stu-id="59995-p106">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.</span></span>

```javascript
console.log(JSON.stringify(Office.context.mailbox.item.cc));
```

##### <a name="compose-mode"></a><span data-ttu-id="59995-264">撰写模式</span><span class="sxs-lookup"><span data-stu-id="59995-264">Compose mode</span></span>

<span data-ttu-id="59995-265">`cc` 属性返回一个 `Recipients` 对象，该对象提供用于获取或更新邮件的“**抄送**”行上收件人的方法。</span><span class="sxs-lookup"><span data-stu-id="59995-265">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span>

```javascript
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="59995-266">类型</span><span class="sxs-lookup"><span data-stu-id="59995-266">Type</span></span>

*   <span data-ttu-id="59995-267">Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_7/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="59995-267">Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_7/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="59995-268">要求</span><span class="sxs-lookup"><span data-stu-id="59995-268">Requirements</span></span>

|<span data-ttu-id="59995-269">要求</span><span class="sxs-lookup"><span data-stu-id="59995-269">Requirement</span></span>|<span data-ttu-id="59995-270">值</span><span class="sxs-lookup"><span data-stu-id="59995-270">Value</span></span>|
|---|---|
|[<span data-ttu-id="59995-271">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="59995-271">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="59995-272">1.0</span><span class="sxs-lookup"><span data-stu-id="59995-272">1.0</span></span>|
|[<span data-ttu-id="59995-273">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="59995-273">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="59995-274">ReadItem</span><span class="sxs-lookup"><span data-stu-id="59995-274">ReadItem</span></span>|
|[<span data-ttu-id="59995-275">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="59995-275">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="59995-276">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="59995-276">Compose or Read</span></span>|

####  <a name="nullable-conversationid-string"></a><span data-ttu-id="59995-277">(nullable) conversationId :String</span><span class="sxs-lookup"><span data-stu-id="59995-277">(nullable) conversationId :String</span></span>

<span data-ttu-id="59995-278">获取包含特定消息的电子邮件会话的标识符。</span><span class="sxs-lookup"><span data-stu-id="59995-278">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="59995-p107">如果在阅读窗体或撰写窗体的回复中激活邮件应用程序，则此属性可以获得一个整数值。如果用户随后更改了回复邮件的主题（若发送回复），则该邮件的对话 ID 将改变且之前获取的值将不适用。</span><span class="sxs-lookup"><span data-stu-id="59995-p107">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="59995-p108">对于撰写窗体的新项目，此属性获得一个 null 值。如果用户设置一个主题并保存该项目，`conversationId` 属性将返回一个值。</span><span class="sxs-lookup"><span data-stu-id="59995-p108">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="59995-283">Type</span><span class="sxs-lookup"><span data-stu-id="59995-283">Type</span></span>

*   <span data-ttu-id="59995-284">String</span><span class="sxs-lookup"><span data-stu-id="59995-284">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="59995-285">要求</span><span class="sxs-lookup"><span data-stu-id="59995-285">Requirements</span></span>

|<span data-ttu-id="59995-286">要求</span><span class="sxs-lookup"><span data-stu-id="59995-286">Requirement</span></span>|<span data-ttu-id="59995-287">值</span><span class="sxs-lookup"><span data-stu-id="59995-287">Value</span></span>|
|---|---|
|[<span data-ttu-id="59995-288">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="59995-288">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="59995-289">1.0</span><span class="sxs-lookup"><span data-stu-id="59995-289">1.0</span></span>|
|[<span data-ttu-id="59995-290">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="59995-290">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="59995-291">ReadItem</span><span class="sxs-lookup"><span data-stu-id="59995-291">ReadItem</span></span>|
|[<span data-ttu-id="59995-292">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="59995-292">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="59995-293">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="59995-293">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="59995-294">示例</span><span class="sxs-lookup"><span data-stu-id="59995-294">Example</span></span>

```javascript
var conversationId = Office.context.mailbox.item.conversationId;
console.log("conversationId: " + conversationId);
```

#### <a name="datetimecreated-date"></a><span data-ttu-id="59995-295">dateTimeCreated :Date</span><span class="sxs-lookup"><span data-stu-id="59995-295">dateTimeCreated :Date</span></span>

<span data-ttu-id="59995-p109">获取项目创建的日期和时间。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="59995-p109">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="59995-298">类型</span><span class="sxs-lookup"><span data-stu-id="59995-298">Type</span></span>

*   <span data-ttu-id="59995-299">日期</span><span class="sxs-lookup"><span data-stu-id="59995-299">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="59995-300">Requirements</span><span class="sxs-lookup"><span data-stu-id="59995-300">Requirements</span></span>

|<span data-ttu-id="59995-301">要求</span><span class="sxs-lookup"><span data-stu-id="59995-301">Requirement</span></span>|<span data-ttu-id="59995-302">值</span><span class="sxs-lookup"><span data-stu-id="59995-302">Value</span></span>|
|---|---|
|[<span data-ttu-id="59995-303">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="59995-303">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="59995-304">1.0</span><span class="sxs-lookup"><span data-stu-id="59995-304">1.0</span></span>|
|[<span data-ttu-id="59995-305">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="59995-305">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="59995-306">ReadItem</span><span class="sxs-lookup"><span data-stu-id="59995-306">ReadItem</span></span>|
|[<span data-ttu-id="59995-307">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="59995-307">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="59995-308">阅读</span><span class="sxs-lookup"><span data-stu-id="59995-308">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="59995-309">示例</span><span class="sxs-lookup"><span data-stu-id="59995-309">Example</span></span>

```javascript
var dateTimeCreated = Office.context.mailbox.item.dateTimeCreated;
console.log("Date and time created: " + dateTimeCreated);
```

#### <a name="datetimemodified-date"></a><span data-ttu-id="59995-310">dateTimeModified :Date</span><span class="sxs-lookup"><span data-stu-id="59995-310">dateTimeModified :Date</span></span>

<span data-ttu-id="59995-p110">获取项目最近一次修改的日期和时间。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="59995-p110">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="59995-313">在 Outlook for iOS 或 Outlook for Android 中不支持此成员。</span><span class="sxs-lookup"><span data-stu-id="59995-313">This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="type"></a><span data-ttu-id="59995-314">类型</span><span class="sxs-lookup"><span data-stu-id="59995-314">Type</span></span>

*   <span data-ttu-id="59995-315">日期</span><span class="sxs-lookup"><span data-stu-id="59995-315">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="59995-316">Requirements</span><span class="sxs-lookup"><span data-stu-id="59995-316">Requirements</span></span>

|<span data-ttu-id="59995-317">要求</span><span class="sxs-lookup"><span data-stu-id="59995-317">Requirement</span></span>|<span data-ttu-id="59995-318">值</span><span class="sxs-lookup"><span data-stu-id="59995-318">Value</span></span>|
|---|---|
|[<span data-ttu-id="59995-319">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="59995-319">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="59995-320">1.0</span><span class="sxs-lookup"><span data-stu-id="59995-320">1.0</span></span>|
|[<span data-ttu-id="59995-321">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="59995-321">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="59995-322">ReadItem</span><span class="sxs-lookup"><span data-stu-id="59995-322">ReadItem</span></span>|
|[<span data-ttu-id="59995-323">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="59995-323">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="59995-324">阅读</span><span class="sxs-lookup"><span data-stu-id="59995-324">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="59995-325">示例</span><span class="sxs-lookup"><span data-stu-id="59995-325">Example</span></span>

```javascript
var dateTimeModified = Office.context.mailbox.item.dateTimeModified;
console.log("Date and time modified: " + dateTimeModified);
```

####  <a name="end-datetimejavascriptapioutlook17officetime"></a><span data-ttu-id="59995-326">end :Date|[Time](/javascript/api/outlook_1_7/office.time)</span><span class="sxs-lookup"><span data-stu-id="59995-326">end :Date|[Time](/javascript/api/outlook_1_7/office.time)</span></span>

<span data-ttu-id="59995-327">获取或设置约会结束的日期和时间。</span><span class="sxs-lookup"><span data-stu-id="59995-327">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="59995-p111">将 `end` 属性表示为协调世界时 (UTC) 的日期和时间值。可使用 [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook17officelocalclienttime) 方法将 end 属性值转换为客户端的本地日期和时间。</span><span class="sxs-lookup"><span data-stu-id="59995-p111">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook17officelocalclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="59995-330">阅读模式</span><span class="sxs-lookup"><span data-stu-id="59995-330">Read mode</span></span>

<span data-ttu-id="59995-331">`end` 属性返回 `Date` 对象。</span><span class="sxs-lookup"><span data-stu-id="59995-331">The `end` property returns a `Date` object.</span></span>

```javascript
var end = Office.context.mailbox.item.end;
console.log("Appointment end: " + end);
```

##### <a name="compose-mode"></a><span data-ttu-id="59995-332">撰写模式</span><span class="sxs-lookup"><span data-stu-id="59995-332">Compose mode</span></span>

<span data-ttu-id="59995-333">`end` 属性返回 `Time` 对象。</span><span class="sxs-lookup"><span data-stu-id="59995-333">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="59995-334">使用 [`Time.setAsync`](/javascript/api/outlook_1_7/office.time#setasync-datetime--options--callback-) 方法设置结束时间时，应使用 [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) 方法将客户端的本地时间转换为服务器的 UTC。</span><span class="sxs-lookup"><span data-stu-id="59995-334">When you use the [`Time.setAsync`](/javascript/api/outlook_1_7/office.time#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="59995-335">下面的示例使用[`setAsync`](/javascript/api/outlook_1_7/office.time#setasync-datetime--options--callback-) `Time`对象的方法设置约会的结束时间。</span><span class="sxs-lookup"><span data-stu-id="59995-335">The following example sets the end time of an appointment by using the [`setAsync`](/javascript/api/outlook_1_7/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="59995-336">类型</span><span class="sxs-lookup"><span data-stu-id="59995-336">Type</span></span>

*   <span data-ttu-id="59995-337">Date | [Time](/javascript/api/outlook_1_7/office.time)</span><span class="sxs-lookup"><span data-stu-id="59995-337">Date | [Time](/javascript/api/outlook_1_7/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="59995-338">Requirements</span><span class="sxs-lookup"><span data-stu-id="59995-338">Requirements</span></span>

|<span data-ttu-id="59995-339">要求</span><span class="sxs-lookup"><span data-stu-id="59995-339">Requirement</span></span>|<span data-ttu-id="59995-340">值</span><span class="sxs-lookup"><span data-stu-id="59995-340">Value</span></span>|
|---|---|
|[<span data-ttu-id="59995-341">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="59995-341">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="59995-342">1.0</span><span class="sxs-lookup"><span data-stu-id="59995-342">1.0</span></span>|
|[<span data-ttu-id="59995-343">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="59995-343">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="59995-344">ReadItem</span><span class="sxs-lookup"><span data-stu-id="59995-344">ReadItem</span></span>|
|[<span data-ttu-id="59995-345">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="59995-345">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="59995-346">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="59995-346">Compose or Read</span></span>|

#### <a name="from-emailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsfromjavascriptapioutlook17officefrom"></a><span data-ttu-id="59995-347">from :[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)|[From](/javascript/api/outlook_1_7/office.from)</span><span class="sxs-lookup"><span data-stu-id="59995-347">from :[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)|[From](/javascript/api/outlook_1_7/office.from)</span></span>

<span data-ttu-id="59995-348">获取邮件发件人的电子邮件地址。</span><span class="sxs-lookup"><span data-stu-id="59995-348">Gets the email address of the sender of a message.</span></span>

<span data-ttu-id="59995-p112">`from` 和 [`sender`](#sender-emailaddressdetailsjavascriptapioutlook17officeemailaddressdetails) 属性表示同一个人，邮件由代理人发送的除外。在此情况下，`from` 属性表示代理程序，而 sender 属性表示代理人。</span><span class="sxs-lookup"><span data-stu-id="59995-p112">The `from` and [`sender`](#sender-emailaddressdetailsjavascriptapioutlook17officeemailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="59995-351">`from` 属性中 `EmailAddressDetails` 对象的 `recipientType` 属性为 `undefined`。</span><span class="sxs-lookup"><span data-stu-id="59995-351">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="59995-352">阅读模式</span><span class="sxs-lookup"><span data-stu-id="59995-352">Read mode</span></span>

<span data-ttu-id="59995-353">`from` 属性返回一个 `EmailAddressDetails` 对象。</span><span class="sxs-lookup"><span data-stu-id="59995-353">The `from` property returns an `EmailAddressDetails` object.</span></span>

```javascript
var from = Office.context.mailbox.item.from;
console.log("From " + from);
```

##### <a name="compose-mode"></a><span data-ttu-id="59995-354">撰写模式</span><span class="sxs-lookup"><span data-stu-id="59995-354">Compose mode</span></span>

<span data-ttu-id="59995-355">`from` 属性返回一个 `From` 对象，该对象提供从值中进行获取的方法。</span><span class="sxs-lookup"><span data-stu-id="59995-355">The `from` property returns a `From` object that provides a method to get the from value.</span></span>

```javascript
Office.context.mailbox.item.from.getAsync(callback);

function callback(asyncResult) {
  var from = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="59995-356">类型</span><span class="sxs-lookup"><span data-stu-id="59995-356">Type</span></span>

*   <span data-ttu-id="59995-357">[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails) | [From](/javascript/api/outlook_1_7/office.from)</span><span class="sxs-lookup"><span data-stu-id="59995-357">[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails) | [From](/javascript/api/outlook_1_7/office.from)</span></span>

##### <a name="requirements"></a><span data-ttu-id="59995-358">要求</span><span class="sxs-lookup"><span data-stu-id="59995-358">Requirements</span></span>

|<span data-ttu-id="59995-359">要求</span><span class="sxs-lookup"><span data-stu-id="59995-359">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="59995-360">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="59995-360">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="59995-361">1.0</span><span class="sxs-lookup"><span data-stu-id="59995-361">1.0</span></span>|<span data-ttu-id="59995-362">1.7</span><span class="sxs-lookup"><span data-stu-id="59995-362">1.7</span></span>|
|[<span data-ttu-id="59995-363">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="59995-363">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="59995-364">ReadItem</span><span class="sxs-lookup"><span data-stu-id="59995-364">ReadItem</span></span>|<span data-ttu-id="59995-365">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="59995-365">ReadWriteItem</span></span>|
|[<span data-ttu-id="59995-366">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="59995-366">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="59995-367">阅读</span><span class="sxs-lookup"><span data-stu-id="59995-367">Read</span></span>|<span data-ttu-id="59995-368">撰写</span><span class="sxs-lookup"><span data-stu-id="59995-368">Compose</span></span>|

#### <a name="internetmessageid-string"></a><span data-ttu-id="59995-369">internetMessageId :String</span><span class="sxs-lookup"><span data-stu-id="59995-369">internetMessageId :String</span></span>

<span data-ttu-id="59995-p113">获取电子邮件的 Internet 消息标识符。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="59995-p113">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="59995-372">Type</span><span class="sxs-lookup"><span data-stu-id="59995-372">Type</span></span>

*   <span data-ttu-id="59995-373">String</span><span class="sxs-lookup"><span data-stu-id="59995-373">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="59995-374">要求</span><span class="sxs-lookup"><span data-stu-id="59995-374">Requirements</span></span>

|<span data-ttu-id="59995-375">要求</span><span class="sxs-lookup"><span data-stu-id="59995-375">Requirement</span></span>|<span data-ttu-id="59995-376">值</span><span class="sxs-lookup"><span data-stu-id="59995-376">Value</span></span>|
|---|---|
|[<span data-ttu-id="59995-377">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="59995-377">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="59995-378">1.0</span><span class="sxs-lookup"><span data-stu-id="59995-378">1.0</span></span>|
|[<span data-ttu-id="59995-379">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="59995-379">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="59995-380">ReadItem</span><span class="sxs-lookup"><span data-stu-id="59995-380">ReadItem</span></span>|
|[<span data-ttu-id="59995-381">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="59995-381">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="59995-382">阅读</span><span class="sxs-lookup"><span data-stu-id="59995-382">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="59995-383">示例</span><span class="sxs-lookup"><span data-stu-id="59995-383">Example</span></span>

```javascript
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

#### <a name="itemclass-string"></a><span data-ttu-id="59995-384">itemClass :String</span><span class="sxs-lookup"><span data-stu-id="59995-384">itemClass :String</span></span>

<span data-ttu-id="59995-p114">获取选定项目的 Exchange Web 服务项目类。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="59995-p114">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="59995-p115">`itemClass` 属性指定所选项目的邮件类别。以下是邮件或约会项目的默认邮件类别。</span><span class="sxs-lookup"><span data-stu-id="59995-p115">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

|<span data-ttu-id="59995-389">类型</span><span class="sxs-lookup"><span data-stu-id="59995-389">Type</span></span>|<span data-ttu-id="59995-390">说明</span><span class="sxs-lookup"><span data-stu-id="59995-390">Description</span></span>|<span data-ttu-id="59995-391">项目类</span><span class="sxs-lookup"><span data-stu-id="59995-391">item class</span></span>|
|---|---|---|
|<span data-ttu-id="59995-392">约会项目</span><span class="sxs-lookup"><span data-stu-id="59995-392">Appointment items</span></span>|<span data-ttu-id="59995-393">这些是项目类 `IPM.Appointment` 或 `IPM.Appointment.Occurrence` 的日历项目。</span><span class="sxs-lookup"><span data-stu-id="59995-393">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurrence`.</span></span>|`IPM.Appointment`<br />`IPM.Appointment.Occurrence`|
|<span data-ttu-id="59995-394">邮件项目</span><span class="sxs-lookup"><span data-stu-id="59995-394">Message items</span></span>|<span data-ttu-id="59995-395">这些项目包括具有默认邮件类别 `IPM.Note` 的电子邮件，以及使用 `IPM.Schedule.Meeting` 作为基础邮件类别的会议请求、响应和取消。</span><span class="sxs-lookup"><span data-stu-id="59995-395">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span>|`IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled`|

<span data-ttu-id="59995-396">你可以创建用于扩展默认邮件类别的自定义邮件类别，例如，自定义约会邮件类别 `IPM.Appointment.Contoso`。</span><span class="sxs-lookup"><span data-stu-id="59995-396">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="59995-397">Type</span><span class="sxs-lookup"><span data-stu-id="59995-397">Type</span></span>

*   <span data-ttu-id="59995-398">String</span><span class="sxs-lookup"><span data-stu-id="59995-398">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="59995-399">要求</span><span class="sxs-lookup"><span data-stu-id="59995-399">Requirements</span></span>

|<span data-ttu-id="59995-400">要求</span><span class="sxs-lookup"><span data-stu-id="59995-400">Requirement</span></span>|<span data-ttu-id="59995-401">值</span><span class="sxs-lookup"><span data-stu-id="59995-401">Value</span></span>|
|---|---|
|[<span data-ttu-id="59995-402">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="59995-402">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="59995-403">1.0</span><span class="sxs-lookup"><span data-stu-id="59995-403">1.0</span></span>|
|[<span data-ttu-id="59995-404">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="59995-404">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="59995-405">ReadItem</span><span class="sxs-lookup"><span data-stu-id="59995-405">ReadItem</span></span>|
|[<span data-ttu-id="59995-406">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="59995-406">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="59995-407">阅读</span><span class="sxs-lookup"><span data-stu-id="59995-407">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="59995-408">示例</span><span class="sxs-lookup"><span data-stu-id="59995-408">Example</span></span>

```javascript
var itemClass = Office.context.mailbox.item.itemClass;
console.log("Item class: " + itemClass);
```

#### <a name="nullable-itemid-string"></a><span data-ttu-id="59995-409">(nullable) itemId :String</span><span class="sxs-lookup"><span data-stu-id="59995-409">(nullable) itemId :String</span></span>

<span data-ttu-id="59995-p116">获取当前项目的 Exchange Web 服务项目标识符。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="59995-p116">Gets the Exchange Web Services item identifier for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="59995-412">`itemId` 属性返回的标识符与 Exchange Web 服务项目标识符相同。</span><span class="sxs-lookup"><span data-stu-id="59995-412">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="59995-413">`itemId` 属性与 Outlook 条目 ID 或 Outlook REST API 使用的 ID 不同。</span><span class="sxs-lookup"><span data-stu-id="59995-413">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="59995-414">使用此值进行 REST API 调用前，应使用 [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string) 对它进行转换。</span><span class="sxs-lookup"><span data-stu-id="59995-414">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="59995-415">有关详细信息，请参阅[从 Outlook 加载项使用 Outlook REST API](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id)。</span><span class="sxs-lookup"><span data-stu-id="59995-415">For more details, see [Use the Outlook REST APIs from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

<span data-ttu-id="59995-p118">`itemId` 属性在撰写模式下不可用。如果需要项目标识符，[`saveAsync`](#saveasyncoptions-callback) 方法可用于将项目保存到存储，这将在回调函数的 [`AsyncResult.value`](/javascript/api/office/office.asyncresult) 参数中返回项目标识符。</span><span class="sxs-lookup"><span data-stu-id="59995-p118">The `itemId` property is not available in compose mode. If an item identifier is required, the [`saveAsync`](#saveasyncoptions-callback) method can be used to save the item to the store, which will return the item identifier in the [`AsyncResult.value`](/javascript/api/office/office.asyncresult) parameter in the callback function.</span></span>

##### <a name="type"></a><span data-ttu-id="59995-418">Type</span><span class="sxs-lookup"><span data-stu-id="59995-418">Type</span></span>

*   <span data-ttu-id="59995-419">String</span><span class="sxs-lookup"><span data-stu-id="59995-419">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="59995-420">要求</span><span class="sxs-lookup"><span data-stu-id="59995-420">Requirements</span></span>

|<span data-ttu-id="59995-421">要求</span><span class="sxs-lookup"><span data-stu-id="59995-421">Requirement</span></span>|<span data-ttu-id="59995-422">值</span><span class="sxs-lookup"><span data-stu-id="59995-422">Value</span></span>|
|---|---|
|[<span data-ttu-id="59995-423">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="59995-423">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="59995-424">1.0</span><span class="sxs-lookup"><span data-stu-id="59995-424">1.0</span></span>|
|[<span data-ttu-id="59995-425">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="59995-425">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="59995-426">ReadItem</span><span class="sxs-lookup"><span data-stu-id="59995-426">ReadItem</span></span>|
|[<span data-ttu-id="59995-427">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="59995-427">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="59995-428">阅读</span><span class="sxs-lookup"><span data-stu-id="59995-428">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="59995-429">示例</span><span class="sxs-lookup"><span data-stu-id="59995-429">Example</span></span>

<span data-ttu-id="59995-p119">以下代码检查项目标识符是否存在。如果 `itemId` 属性返回 `null` 或 `undefined`，则将项目保存到存储，并从异步结果中获取项目标识符。</span><span class="sxs-lookup"><span data-stu-id="59995-p119">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

```javascript
var itemId = Office.context.mailbox.item.itemId;
if (itemId === null || itemId == undefined) {
  Office.context.mailbox.item.saveAsync(function(result) {
    itemId = result.value;
  });
}
```

####  <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlook17officemailboxenumsitemtype"></a><span data-ttu-id="59995-432">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_7/office.mailboxenums.itemtype)</span><span class="sxs-lookup"><span data-stu-id="59995-432">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_7/office.mailboxenums.itemtype)</span></span>

<span data-ttu-id="59995-433">获取实例表示的项的类型。</span><span class="sxs-lookup"><span data-stu-id="59995-433">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="59995-434">`itemType` 属性返回其中一个 `ItemType` 枚举值，指示 `item` 对象实例是邮件还是约会。</span><span class="sxs-lookup"><span data-stu-id="59995-434">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="59995-435">类型</span><span class="sxs-lookup"><span data-stu-id="59995-435">Type</span></span>

*   [<span data-ttu-id="59995-436">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="59995-436">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook_1_7/office.mailboxenums.itemtype)

##### <a name="requirements"></a><span data-ttu-id="59995-437">Requirements</span><span class="sxs-lookup"><span data-stu-id="59995-437">Requirements</span></span>

|<span data-ttu-id="59995-438">要求</span><span class="sxs-lookup"><span data-stu-id="59995-438">Requirement</span></span>|<span data-ttu-id="59995-439">值</span><span class="sxs-lookup"><span data-stu-id="59995-439">Value</span></span>|
|---|---|
|[<span data-ttu-id="59995-440">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="59995-440">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="59995-441">1.0</span><span class="sxs-lookup"><span data-stu-id="59995-441">1.0</span></span>|
|[<span data-ttu-id="59995-442">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="59995-442">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="59995-443">ReadItem</span><span class="sxs-lookup"><span data-stu-id="59995-443">ReadItem</span></span>|
|[<span data-ttu-id="59995-444">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="59995-444">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="59995-445">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="59995-445">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="59995-446">示例</span><span class="sxs-lookup"><span data-stu-id="59995-446">Example</span></span>

```javascript
if (Office.context.mailbox.item.itemType === Office.MailboxEnums.ItemType.Message) {
  // Do something.
} else {
  // Do something else.
}
```

####  <a name="location-stringlocationjavascriptapioutlook17officelocation"></a><span data-ttu-id="59995-447">location :String|[Location](/javascript/api/outlook_1_7/office.location)</span><span class="sxs-lookup"><span data-stu-id="59995-447">location :String|[Location](/javascript/api/outlook_1_7/office.location)</span></span>

<span data-ttu-id="59995-448">获取或设置约会的位置。</span><span class="sxs-lookup"><span data-stu-id="59995-448">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="59995-449">阅读模式</span><span class="sxs-lookup"><span data-stu-id="59995-449">Read mode</span></span>

<span data-ttu-id="59995-450">`location` 属性返回一个包含约会位置的字符串。</span><span class="sxs-lookup"><span data-stu-id="59995-450">The `location` property returns a string that contains the location of the appointment.</span></span>

```javascript
var location = Office.context.mailbox.item.location;
console.log("location: " + location);
```

##### <a name="compose-mode"></a><span data-ttu-id="59995-451">撰写模式</span><span class="sxs-lookup"><span data-stu-id="59995-451">Compose mode</span></span>

<span data-ttu-id="59995-452">`location` 属性返回一个 `Location` 对象，该对象提供用于获取和设置约会位置的方法。</span><span class="sxs-lookup"><span data-stu-id="59995-452">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

```javascript
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="59995-453">类型</span><span class="sxs-lookup"><span data-stu-id="59995-453">Type</span></span>

*   <span data-ttu-id="59995-454">String | [Location](/javascript/api/outlook_1_7/office.location)</span><span class="sxs-lookup"><span data-stu-id="59995-454">String | [Location](/javascript/api/outlook_1_7/office.location)</span></span>

##### <a name="requirements"></a><span data-ttu-id="59995-455">Requirements</span><span class="sxs-lookup"><span data-stu-id="59995-455">Requirements</span></span>

|<span data-ttu-id="59995-456">要求</span><span class="sxs-lookup"><span data-stu-id="59995-456">Requirement</span></span>|<span data-ttu-id="59995-457">值</span><span class="sxs-lookup"><span data-stu-id="59995-457">Value</span></span>|
|---|---|
|[<span data-ttu-id="59995-458">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="59995-458">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="59995-459">1.0</span><span class="sxs-lookup"><span data-stu-id="59995-459">1.0</span></span>|
|[<span data-ttu-id="59995-460">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="59995-460">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="59995-461">ReadItem</span><span class="sxs-lookup"><span data-stu-id="59995-461">ReadItem</span></span>|
|[<span data-ttu-id="59995-462">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="59995-462">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="59995-463">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="59995-463">Compose or Read</span></span>|

#### <a name="normalizedsubject-string"></a><span data-ttu-id="59995-464">normalizedSubject :String</span><span class="sxs-lookup"><span data-stu-id="59995-464">normalizedSubject :String</span></span>

<span data-ttu-id="59995-p120">获取删除了所有前缀（包括 `RE:` 和 `FWD:`）的项目主题。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="59995-p120">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="59995-p121">normalizedSubject 属性获取包含由电子邮件程序添加的任何标准前缀（如 `RE:` 和 `FW:`）的项目主题。若要获取包含完整前缀的项目主题，请使用 [`subject`](#subject-stringsubjectjavascriptapioutlook17officesubject) 属性。</span><span class="sxs-lookup"><span data-stu-id="59995-p121">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubjectjavascriptapioutlook17officesubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="59995-469">Type</span><span class="sxs-lookup"><span data-stu-id="59995-469">Type</span></span>

*   <span data-ttu-id="59995-470">String</span><span class="sxs-lookup"><span data-stu-id="59995-470">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="59995-471">要求</span><span class="sxs-lookup"><span data-stu-id="59995-471">Requirements</span></span>

|<span data-ttu-id="59995-472">要求</span><span class="sxs-lookup"><span data-stu-id="59995-472">Requirement</span></span>|<span data-ttu-id="59995-473">值</span><span class="sxs-lookup"><span data-stu-id="59995-473">Value</span></span>|
|---|---|
|[<span data-ttu-id="59995-474">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="59995-474">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="59995-475">1.0</span><span class="sxs-lookup"><span data-stu-id="59995-475">1.0</span></span>|
|[<span data-ttu-id="59995-476">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="59995-476">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="59995-477">ReadItem</span><span class="sxs-lookup"><span data-stu-id="59995-477">ReadItem</span></span>|
|[<span data-ttu-id="59995-478">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="59995-478">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="59995-479">阅读</span><span class="sxs-lookup"><span data-stu-id="59995-479">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="59995-480">示例</span><span class="sxs-lookup"><span data-stu-id="59995-480">Example</span></span>

```javascript
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
console.log("Normalized subject: " + normalizedSubject);
```

####  <a name="notificationmessages-notificationmessagesjavascriptapioutlook17officenotificationmessages"></a><span data-ttu-id="59995-481">notificationMessages :[NotificationMessages](/javascript/api/outlook_1_7/office.notificationmessages)</span><span class="sxs-lookup"><span data-stu-id="59995-481">notificationMessages :[NotificationMessages](/javascript/api/outlook_1_7/office.notificationmessages)</span></span>

<span data-ttu-id="59995-482">获取项目的通知邮件。</span><span class="sxs-lookup"><span data-stu-id="59995-482">Gets the notification messages for an item.</span></span>

##### <a name="type"></a><span data-ttu-id="59995-483">类型</span><span class="sxs-lookup"><span data-stu-id="59995-483">Type</span></span>

*   [<span data-ttu-id="59995-484">NotificationMessages</span><span class="sxs-lookup"><span data-stu-id="59995-484">NotificationMessages</span></span>](/javascript/api/outlook_1_7/office.notificationmessages)

##### <a name="requirements"></a><span data-ttu-id="59995-485">要求</span><span class="sxs-lookup"><span data-stu-id="59995-485">Requirements</span></span>

|<span data-ttu-id="59995-486">要求</span><span class="sxs-lookup"><span data-stu-id="59995-486">Requirement</span></span>|<span data-ttu-id="59995-487">值</span><span class="sxs-lookup"><span data-stu-id="59995-487">Value</span></span>|
|---|---|
|[<span data-ttu-id="59995-488">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="59995-488">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="59995-489">1.3</span><span class="sxs-lookup"><span data-stu-id="59995-489">1.3</span></span>|
|[<span data-ttu-id="59995-490">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="59995-490">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="59995-491">ReadItem</span><span class="sxs-lookup"><span data-stu-id="59995-491">ReadItem</span></span>|
|[<span data-ttu-id="59995-492">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="59995-492">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="59995-493">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="59995-493">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="59995-494">示例</span><span class="sxs-lookup"><span data-stu-id="59995-494">Example</span></span>

```javascript
// Get all notifications.
Office.context.mailbox.item.notificationMessages.getAllAsync(
  function (asyncResult) {
    console.log(JSON.stringify(asyncResult));
  }
);
```

####  <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsrecipientsjavascriptapioutlook17officerecipients"></a><span data-ttu-id="59995-495">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_7/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="59995-495">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_7/office.recipients)</span></span>

<span data-ttu-id="59995-496">提供对事件的可选与会者的访问权限。</span><span class="sxs-lookup"><span data-stu-id="59995-496">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="59995-497">对象的类型和访问级别取决于当前项目的模式。</span><span class="sxs-lookup"><span data-stu-id="59995-497">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="59995-498">阅读模式</span><span class="sxs-lookup"><span data-stu-id="59995-498">Read mode</span></span>

<span data-ttu-id="59995-499">`optionalAttendees` 属性返回一个数组，其中包含每个可选与会者的 `EmailAddressDetails` 对象。</span><span class="sxs-lookup"><span data-stu-id="59995-499">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span>

```javascript
var optionalAttendees = Office.context.mailbox.item.optionalAttendees;
console.log("Optional attendees: " + JSON.stringify(optionalAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="59995-500">撰写模式</span><span class="sxs-lookup"><span data-stu-id="59995-500">Compose mode</span></span>

<span data-ttu-id="59995-501">`optionalAttendees` 属性返回一个 `Recipients` 对象，该对象提供用于获取或更新可选与会者的方法。</span><span class="sxs-lookup"><span data-stu-id="59995-501">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span>

```javascript
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="59995-502">类型</span><span class="sxs-lookup"><span data-stu-id="59995-502">Type</span></span>

*   <span data-ttu-id="59995-503">Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_7/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="59995-503">Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_7/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="59995-504">Requirements</span><span class="sxs-lookup"><span data-stu-id="59995-504">Requirements</span></span>

|<span data-ttu-id="59995-505">要求</span><span class="sxs-lookup"><span data-stu-id="59995-505">Requirement</span></span>|<span data-ttu-id="59995-506">值</span><span class="sxs-lookup"><span data-stu-id="59995-506">Value</span></span>|
|---|---|
|[<span data-ttu-id="59995-507">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="59995-507">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="59995-508">1.0</span><span class="sxs-lookup"><span data-stu-id="59995-508">1.0</span></span>|
|[<span data-ttu-id="59995-509">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="59995-509">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="59995-510">ReadItem</span><span class="sxs-lookup"><span data-stu-id="59995-510">ReadItem</span></span>|
|[<span data-ttu-id="59995-511">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="59995-511">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="59995-512">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="59995-512">Compose or Read</span></span>|

#### <a name="organizer-emailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsorganizerjavascriptapioutlook17officeorganizer"></a><span data-ttu-id="59995-513">organizer :[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)|[Organizer](/javascript/api/outlook_1_7/office.organizer)</span><span class="sxs-lookup"><span data-stu-id="59995-513">organizer :[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)|[Organizer](/javascript/api/outlook_1_7/office.organizer)</span></span>

<span data-ttu-id="59995-514">获取指定会议的组织者的电子邮件地址。</span><span class="sxs-lookup"><span data-stu-id="59995-514">Gets the email address of the organizer for a specified meeting.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="59995-515">阅读模式</span><span class="sxs-lookup"><span data-stu-id="59995-515">Read mode</span></span>

<span data-ttu-id="59995-516">`organizer` 属性返回 [EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails) 对象，它表示会议组织者。</span><span class="sxs-lookup"><span data-stu-id="59995-516">The `organizer` property returns an [EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails) object that represents the meeting organizer.</span></span>

```javascript
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
console.log("Organizer: " + organizerName + " (" + organizerAddress + ")");
```

##### <a name="compose-mode"></a><span data-ttu-id="59995-517">撰写模式</span><span class="sxs-lookup"><span data-stu-id="59995-517">Compose mode</span></span>

<span data-ttu-id="59995-518">`organizer` 属性返回 [Organizer](/javascript/api/outlook_1_7/office.organizer) 对象，该对象提供获取组织者值的方法。</span><span class="sxs-lookup"><span data-stu-id="59995-518">The `organizer` property returns an [Organizer](/javascript/api/outlook_1_7/office.organizer) object that provides a method to get the organizer value.</span></span>

```javascript
Office.context.mailbox.item.organizer.getAsync(
  function(asyncResult) {
    console.log(JSON.stringify(asyncResult));
  }
);
```

##### <a name="type"></a><span data-ttu-id="59995-519">类型</span><span class="sxs-lookup"><span data-stu-id="59995-519">Type</span></span>

*   <span data-ttu-id="59995-520">[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails) | [Organizer](/javascript/api/outlook_1_7/office.organizer)</span><span class="sxs-lookup"><span data-stu-id="59995-520">[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails) | [Organizer](/javascript/api/outlook_1_7/office.organizer)</span></span>

##### <a name="requirements"></a><span data-ttu-id="59995-521">要求</span><span class="sxs-lookup"><span data-stu-id="59995-521">Requirements</span></span>

|<span data-ttu-id="59995-522">要求</span><span class="sxs-lookup"><span data-stu-id="59995-522">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="59995-523">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="59995-523">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="59995-524">1.0</span><span class="sxs-lookup"><span data-stu-id="59995-524">1.0</span></span>|<span data-ttu-id="59995-525">1.7</span><span class="sxs-lookup"><span data-stu-id="59995-525">1.7</span></span>|
|[<span data-ttu-id="59995-526">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="59995-526">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="59995-527">ReadItem</span><span class="sxs-lookup"><span data-stu-id="59995-527">ReadItem</span></span>|<span data-ttu-id="59995-528">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="59995-528">ReadWriteItem</span></span>|
|[<span data-ttu-id="59995-529">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="59995-529">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="59995-530">阅读</span><span class="sxs-lookup"><span data-stu-id="59995-530">Read</span></span>|<span data-ttu-id="59995-531">撰写</span><span class="sxs-lookup"><span data-stu-id="59995-531">Compose</span></span>|

#### <a name="nullable-recurrence-recurrencejavascriptapioutlook17officerecurrence"></a><span data-ttu-id="59995-532">(nullable) recurrence :[Recurrence](/javascript/api/outlook_1_7/office.recurrence)</span><span class="sxs-lookup"><span data-stu-id="59995-532">(nullable) recurrence :[Recurrence](/javascript/api/outlook_1_7/office.recurrence)</span></span>

<span data-ttu-id="59995-533">获取或设置约会的定期模式。</span><span class="sxs-lookup"><span data-stu-id="59995-533">Gets or sets the recurrence pattern of an appointment.</span></span> <span data-ttu-id="59995-534">获取或设置会议请求的定期模式。</span><span class="sxs-lookup"><span data-stu-id="59995-534">Gets the recurrence pattern of a meeting request.</span></span> <span data-ttu-id="59995-535">阅读撰写约会项目的模式。</span><span class="sxs-lookup"><span data-stu-id="59995-535">Read and compose modes for appointment items.</span></span> <span data-ttu-id="59995-536">会议请求项目的阅读模式。</span><span class="sxs-lookup"><span data-stu-id="59995-536">Read mode for meeting request items.</span></span>

<span data-ttu-id="59995-537">如果项目是一个系列或系列中的一个实例，则 `recurrence` 属性将返回定期约会的 [recurrence](/javascript/api/outlook_1_7/office.recurrence) 对象或会议请求。</span><span class="sxs-lookup"><span data-stu-id="59995-537">The `recurrence` property returns a [recurrence](/javascript/api/outlook_1_7/office.recurrence) object for recurring appointments or meetings requests if an item is a series or an instance in a series.</span></span> <span data-ttu-id="59995-538">针对单个约会和单个约会的会议请求返回 `null`。</span><span class="sxs-lookup"><span data-stu-id="59995-538">`null` is returned for single appointments and meeting requests of single appointments.</span></span> <span data-ttu-id="59995-539">针对非会议请求的邮件返回 `undefined`。</span><span class="sxs-lookup"><span data-stu-id="59995-539">`undefined` is returned for messages that are not meeting requests.</span></span>

> <span data-ttu-id="59995-540">注意：会议请求的 `itemClass` 值为 IPM.Schedule.Meeting.Request。</span><span class="sxs-lookup"><span data-stu-id="59995-540">Note: Meeting requests have an `itemClass` value of IPM.Schedule.Meeting.Request.</span></span>

> <span data-ttu-id="59995-541">注意：如果 recurrence 对象为 `null`，则这表示对象是单个约会或单个约会的会议请求，而不是系列的一部分。</span><span class="sxs-lookup"><span data-stu-id="59995-541">Note: If the recurrence object is `null`, this indicates that the object is a single appointment or a meeting request of a single appointment and NOT a part of a series.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="59995-542">阅读模式</span><span class="sxs-lookup"><span data-stu-id="59995-542">Read mode</span></span>

<span data-ttu-id="59995-543">该`recurrence`属性返回一个代表约会定期的[定期](/javascript/api/outlook_1_7/office.recurrence)对象。</span><span class="sxs-lookup"><span data-stu-id="59995-543">The `recurrence` property returns a [Recurrence](/javascript/api/outlook_1_7/office.recurrence) object that represents the appointment recurrence.</span></span> <span data-ttu-id="59995-544">此功能适用于约会和会议请求。</span><span class="sxs-lookup"><span data-stu-id="59995-544">This is available for appointments and meeting requests.</span></span>

```javascript
var recurrence = Office.context.mailbox.item.recurrence;
console.log("Recurrence: " + JSON.stringify(recurrence));
```

##### <a name="compose-mode"></a><span data-ttu-id="59995-545">撰写模式</span><span class="sxs-lookup"><span data-stu-id="59995-545">Compose mode</span></span>

<span data-ttu-id="59995-546">该`recurrence`属性返回一个[定期](/javascript/api/outlook_1_7/office.recurrence)对象, 该对象提供用于管理约会周期的方法。</span><span class="sxs-lookup"><span data-stu-id="59995-546">The `recurrence` property returns a [Recurrence](/javascript/api/outlook_1_7/office.recurrence) object that provides methods to manage the appointment recurrence.</span></span> <span data-ttu-id="59995-547">这可用于约会。</span><span class="sxs-lookup"><span data-stu-id="59995-547">This is available for appointments.</span></span>

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

##### <a name="type"></a><span data-ttu-id="59995-548">类型</span><span class="sxs-lookup"><span data-stu-id="59995-548">Type</span></span>

* [<span data-ttu-id="59995-549">Recurrence</span><span class="sxs-lookup"><span data-stu-id="59995-549">Recurrence</span></span>](/javascript/api/outlook_1_7/office.recurrence)

|<span data-ttu-id="59995-550">要求</span><span class="sxs-lookup"><span data-stu-id="59995-550">Requirement</span></span>|<span data-ttu-id="59995-551">值</span><span class="sxs-lookup"><span data-stu-id="59995-551">Value</span></span>|
|---|---|
|[<span data-ttu-id="59995-552">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="59995-552">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="59995-553">1.7</span><span class="sxs-lookup"><span data-stu-id="59995-553">1.7</span></span>|
|[<span data-ttu-id="59995-554">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="59995-554">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="59995-555">ReadItem</span><span class="sxs-lookup"><span data-stu-id="59995-555">ReadItem</span></span>|
|[<span data-ttu-id="59995-556">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="59995-556">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="59995-557">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="59995-557">Compose or Read</span></span>|

####  <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsrecipientsjavascriptapioutlook17officerecipients"></a><span data-ttu-id="59995-558">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_7/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="59995-558">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_7/office.recipients)</span></span>

<span data-ttu-id="59995-559">提供对事件的必需与会者的访问权限。</span><span class="sxs-lookup"><span data-stu-id="59995-559">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="59995-560">对象的类型和访问级别取决于当前项目的模式。</span><span class="sxs-lookup"><span data-stu-id="59995-560">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="59995-561">读取模式</span><span class="sxs-lookup"><span data-stu-id="59995-561">Read mode</span></span>

<span data-ttu-id="59995-562">`requiredAttendees` 属性返回一个数组，其中包含每个必需与会者的 `EmailAddressDetails` 对象。</span><span class="sxs-lookup"><span data-stu-id="59995-562">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span>

```javascript
var requiredAttendees = Office.context.mailbox.item.requiredAttendees;
console.log("Required attendees: " + JSON.stringify(requiredAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="59995-563">撰写模式</span><span class="sxs-lookup"><span data-stu-id="59995-563">Compose mode</span></span>

<span data-ttu-id="59995-564">`requiredAttendees` 属性返回一个 `Recipients` 对象，该对象提供用于获取或更新必需与会者的方法。</span><span class="sxs-lookup"><span data-stu-id="59995-564">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span>

```javascript
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
  console.log(JSON.stringify(arrayOfRequiredAttendeesRecipients));
}
```

##### <a name="type"></a><span data-ttu-id="59995-565">类型</span><span class="sxs-lookup"><span data-stu-id="59995-565">Type</span></span>

*   <span data-ttu-id="59995-566">Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_7/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="59995-566">Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_7/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="59995-567">Requirements</span><span class="sxs-lookup"><span data-stu-id="59995-567">Requirements</span></span>

|<span data-ttu-id="59995-568">要求</span><span class="sxs-lookup"><span data-stu-id="59995-568">Requirement</span></span>|<span data-ttu-id="59995-569">值</span><span class="sxs-lookup"><span data-stu-id="59995-569">Value</span></span>|
|---|---|
|[<span data-ttu-id="59995-570">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="59995-570">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="59995-571">1.0</span><span class="sxs-lookup"><span data-stu-id="59995-571">1.0</span></span>|
|[<span data-ttu-id="59995-572">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="59995-572">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="59995-573">ReadItem</span><span class="sxs-lookup"><span data-stu-id="59995-573">ReadItem</span></span>|
|[<span data-ttu-id="59995-574">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="59995-574">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="59995-575">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="59995-575">Compose or Read</span></span>|

#### <a name="sender-emailaddressdetailsjavascriptapioutlook17officeemailaddressdetails"></a><span data-ttu-id="59995-576">sender :[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="59995-576">sender :[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)</span></span>

<span data-ttu-id="59995-p128">获取电子邮件发件人的电子邮件地址。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="59995-p128">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="59995-p129">[`from`](#from-emailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsfromjavascriptapioutlook17officefrom) 和 `sender` 属性表示同一个人，邮件由代理人发送的除外。在此情况下，`from` 属性表示代理程序，而 sender 属性表示代理人。</span><span class="sxs-lookup"><span data-stu-id="59995-p129">The [`from`](#from-emailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsfromjavascriptapioutlook17officefrom) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="59995-581">`sender` 属性中 `EmailAddressDetails` 对象的 `recipientType` 属性为 `undefined`。</span><span class="sxs-lookup"><span data-stu-id="59995-581">The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="59995-582">类型</span><span class="sxs-lookup"><span data-stu-id="59995-582">Type</span></span>

*   [<span data-ttu-id="59995-583">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="59995-583">EmailAddressDetails</span></span>](/javascript/api/outlook_1_7/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="59995-584">Requirements</span><span class="sxs-lookup"><span data-stu-id="59995-584">Requirements</span></span>

|<span data-ttu-id="59995-585">要求</span><span class="sxs-lookup"><span data-stu-id="59995-585">Requirement</span></span>|<span data-ttu-id="59995-586">值</span><span class="sxs-lookup"><span data-stu-id="59995-586">Value</span></span>|
|---|---|
|[<span data-ttu-id="59995-587">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="59995-587">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="59995-588">1.0</span><span class="sxs-lookup"><span data-stu-id="59995-588">1.0</span></span>|
|[<span data-ttu-id="59995-589">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="59995-589">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="59995-590">ReadItem</span><span class="sxs-lookup"><span data-stu-id="59995-590">ReadItem</span></span>|
|[<span data-ttu-id="59995-591">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="59995-591">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="59995-592">阅读</span><span class="sxs-lookup"><span data-stu-id="59995-592">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="59995-593">示例</span><span class="sxs-lookup"><span data-stu-id="59995-593">Example</span></span>

```javascript
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
console.log("Sender: " + senderName + " (" + senderAddress + ")");
```

#### <a name="nullable-seriesid-string"></a><span data-ttu-id="59995-594">(nullable) seriesId :String</span><span class="sxs-lookup"><span data-stu-id="59995-594">(nullable) seriesId :String</span></span>

<span data-ttu-id="59995-595">获取实例所属的系列的 ID。</span><span class="sxs-lookup"><span data-stu-id="59995-595">Gets the id of the series that an instance belongs to.</span></span>

<span data-ttu-id="59995-596">在 OWA 和 Outlook 中，`seriesId` 返回此项目所属的父（系列）项目的 Exchange Web 服务 (EWS) ID。</span><span class="sxs-lookup"><span data-stu-id="59995-596">In OWA and Outlook, the `seriesId` returns the Exchange Web Services (EWS) ID of the parent (series) item that this item belongs to.</span></span> <span data-ttu-id="59995-597">但是，在 iOS 和 Android 中，`seriesId` 返回父项目的其余部分 ID。</span><span class="sxs-lookup"><span data-stu-id="59995-597">However, in iOS and Android, the `seriesId` returns the REST ID of the parent item.</span></span>

> [!NOTE]
> <span data-ttu-id="59995-598">`seriesId` 属性返回的标识符与 Exchange Web 服务项目标识符相同。</span><span class="sxs-lookup"><span data-stu-id="59995-598">The identifier returned by the `seriesId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="59995-599">`seriesId` 属性与 Outlook REST API 使用的 Outlook ID 不同。</span><span class="sxs-lookup"><span data-stu-id="59995-599">The `seriesId` property is not identical to the Outlook IDs used by the Outlook REST API.</span></span> <span data-ttu-id="59995-600">使用此值进行 REST API 调用前，应使用 [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string) 对它进行转换。</span><span class="sxs-lookup"><span data-stu-id="59995-600">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="59995-601">有关详细信息，请参阅[从 Outlook 加载项使用 Outlook REST API](https://docs.microsoft.com/outlook/add-ins/use-rest-api)。</span><span class="sxs-lookup"><span data-stu-id="59995-601">For more details, see [Use the Outlook REST APIs from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/use-rest-api).</span></span>

<span data-ttu-id="59995-602">`seriesId` 属性对于没有父项目（如单个约会、系列项目或会议请求）的项目返回 `null`，对于非会议请求的任何其他项目，返回 `undefined`。</span><span class="sxs-lookup"><span data-stu-id="59995-602">The `seriesId` property returns `null` for items that do not have parent items such as single appointments, series items, or meeting requests and returns `undefined` for any other items that are not meeting requests.</span></span>

##### <a name="type"></a><span data-ttu-id="59995-603">Type</span><span class="sxs-lookup"><span data-stu-id="59995-603">Type</span></span>

* <span data-ttu-id="59995-604">String</span><span class="sxs-lookup"><span data-stu-id="59995-604">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="59995-605">要求</span><span class="sxs-lookup"><span data-stu-id="59995-605">Requirements</span></span>

|<span data-ttu-id="59995-606">要求</span><span class="sxs-lookup"><span data-stu-id="59995-606">Requirement</span></span>|<span data-ttu-id="59995-607">值</span><span class="sxs-lookup"><span data-stu-id="59995-607">Value</span></span>|
|---|---|
|[<span data-ttu-id="59995-608">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="59995-608">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="59995-609">1.7</span><span class="sxs-lookup"><span data-stu-id="59995-609">1.7</span></span>|
|[<span data-ttu-id="59995-610">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="59995-610">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="59995-611">ReadItem</span><span class="sxs-lookup"><span data-stu-id="59995-611">ReadItem</span></span>|
|[<span data-ttu-id="59995-612">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="59995-612">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="59995-613">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="59995-613">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="59995-614">示例</span><span class="sxs-lookup"><span data-stu-id="59995-614">Example</span></span>

```javascript
var seriesId = Office.context.mailbox.item.seriesId;

// The seriesId property returns null for items that do
// not have parent items (such as single appointments,
// series items, or meeting requests) and returns
// undefined for messages that are not meeting requests.
var isSeriesInstance = (seriesId != null);
console.log("SeriesId is " + seriesId + " and isSeriesInstance is " + isSeriesInstance);
```

####  <a name="start-datetimejavascriptapioutlook17officetime"></a><span data-ttu-id="59995-615">start :Date|[Time](/javascript/api/outlook_1_7/office.time)</span><span class="sxs-lookup"><span data-stu-id="59995-615">start :Date|[Time](/javascript/api/outlook_1_7/office.time)</span></span>

<span data-ttu-id="59995-616">获取或设置约会开始的日期和时间。</span><span class="sxs-lookup"><span data-stu-id="59995-616">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="59995-p132">将 `start` 属性表示为协调世界时 (UTC) 的日期和时间值。可使用 [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook17officelocalclienttime) 方法将该值转换为客户端的本地日期和时间。</span><span class="sxs-lookup"><span data-stu-id="59995-p132">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook17officelocalclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="59995-619">阅读模式</span><span class="sxs-lookup"><span data-stu-id="59995-619">Read mode</span></span>

<span data-ttu-id="59995-620">`start` 属性返回 `Date` 对象。</span><span class="sxs-lookup"><span data-stu-id="59995-620">The `start` property returns a `Date` object.</span></span>

```javascript
var start = Office.context.mailbox.item.start;
console.log("Appointment start: " + JSON.stringify(start));
```

##### <a name="compose-mode"></a><span data-ttu-id="59995-621">撰写模式</span><span class="sxs-lookup"><span data-stu-id="59995-621">Compose mode</span></span>

<span data-ttu-id="59995-622">`start` 属性返回 `Time` 对象。</span><span class="sxs-lookup"><span data-stu-id="59995-622">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="59995-623">使用 [`Time.setAsync`](/javascript/api/outlook_1_7/office.time#setasync-datetime--options--callback-) 方法设置开始时间时，应使用 [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) 方法将客户端的本地时间转换为服务器的 UTC。</span><span class="sxs-lookup"><span data-stu-id="59995-623">When you use the [`Time.setAsync`](/javascript/api/outlook_1_7/office.time#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="59995-624">以下示例通过使用 `Time` 对象的 [`setAsync`](/javascript/api/outlook_1_7/office.time#setasync-datetime--options--callback-) 方法，设置撰写模式下约会的开始时间。</span><span class="sxs-lookup"><span data-stu-id="59995-624">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_7/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="59995-625">类型</span><span class="sxs-lookup"><span data-stu-id="59995-625">Type</span></span>

*   <span data-ttu-id="59995-626">Date | [Time](/javascript/api/outlook_1_7/office.time)</span><span class="sxs-lookup"><span data-stu-id="59995-626">Date | [Time](/javascript/api/outlook_1_7/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="59995-627">Requirements</span><span class="sxs-lookup"><span data-stu-id="59995-627">Requirements</span></span>

|<span data-ttu-id="59995-628">要求</span><span class="sxs-lookup"><span data-stu-id="59995-628">Requirement</span></span>|<span data-ttu-id="59995-629">值</span><span class="sxs-lookup"><span data-stu-id="59995-629">Value</span></span>|
|---|---|
|[<span data-ttu-id="59995-630">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="59995-630">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="59995-631">1.0</span><span class="sxs-lookup"><span data-stu-id="59995-631">1.0</span></span>|
|[<span data-ttu-id="59995-632">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="59995-632">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="59995-633">ReadItem</span><span class="sxs-lookup"><span data-stu-id="59995-633">ReadItem</span></span>|
|[<span data-ttu-id="59995-634">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="59995-634">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="59995-635">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="59995-635">Compose or Read</span></span>|

####  <a name="subject-stringsubjectjavascriptapioutlook17officesubject"></a><span data-ttu-id="59995-636">subject :String|[Subject](/javascript/api/outlook_1_7/office.subject)</span><span class="sxs-lookup"><span data-stu-id="59995-636">subject :String|[Subject](/javascript/api/outlook_1_7/office.subject)</span></span>

<span data-ttu-id="59995-637">获取或设置显示在项目的主题字段中的说明。</span><span class="sxs-lookup"><span data-stu-id="59995-637">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="59995-638">`subject` 属性获取或设置由电子邮件服务器发送项目时的整个主题。</span><span class="sxs-lookup"><span data-stu-id="59995-638">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="59995-639">阅读模式</span><span class="sxs-lookup"><span data-stu-id="59995-639">Read mode</span></span>

<span data-ttu-id="59995-p133">`subject` 属性返回一个字符串。使用 [`normalizedSubject`](#normalizedsubject-string) 属性获取不带任何前导前缀（如 `RE:` 和 `FW:`）的主题。</span><span class="sxs-lookup"><span data-stu-id="59995-p133">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

<span data-ttu-id="59995-642">以下 JavaScript 代码示例显示了如何访问 Outlook 中当前项目的 `subject` 属性。</span><span class="sxs-lookup"><span data-stu-id="59995-642">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

```javascript
var subject = Office.context.mailbox.item.subject;
console.log(subject);
```

##### <a name="compose-mode"></a><span data-ttu-id="59995-643">撰写模式</span><span class="sxs-lookup"><span data-stu-id="59995-643">Compose mode</span></span>

<span data-ttu-id="59995-644">`subject` 属性返回一个 `Subject` 对象，该对象提供用于获取和设置主题的方法。</span><span class="sxs-lookup"><span data-stu-id="59995-644">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```javascript
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
  console.log(subject);
}
```

##### <a name="type"></a><span data-ttu-id="59995-645">类型</span><span class="sxs-lookup"><span data-stu-id="59995-645">Type</span></span>

*   <span data-ttu-id="59995-646">String | [Subject](/javascript/api/outlook_1_7/office.subject)</span><span class="sxs-lookup"><span data-stu-id="59995-646">String | [Subject](/javascript/api/outlook_1_7/office.subject)</span></span>

##### <a name="requirements"></a><span data-ttu-id="59995-647">要求</span><span class="sxs-lookup"><span data-stu-id="59995-647">Requirements</span></span>

|<span data-ttu-id="59995-648">要求</span><span class="sxs-lookup"><span data-stu-id="59995-648">Requirement</span></span>|<span data-ttu-id="59995-649">值</span><span class="sxs-lookup"><span data-stu-id="59995-649">Value</span></span>|
|---|---|
|[<span data-ttu-id="59995-650">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="59995-650">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="59995-651">1.0</span><span class="sxs-lookup"><span data-stu-id="59995-651">1.0</span></span>|
|[<span data-ttu-id="59995-652">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="59995-652">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="59995-653">ReadItem</span><span class="sxs-lookup"><span data-stu-id="59995-653">ReadItem</span></span>|
|[<span data-ttu-id="59995-654">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="59995-654">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="59995-655">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="59995-655">Compose or Read</span></span>|

####  <a name="to-arrayemailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsrecipientsjavascriptapioutlook17officerecipients"></a><span data-ttu-id="59995-656">to :Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_7/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="59995-656">to :Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_7/office.recipients)</span></span>

<span data-ttu-id="59995-657">提供对邮件的“**收件人**”行上的收件人的访问权限。</span><span class="sxs-lookup"><span data-stu-id="59995-657">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="59995-658">对象的类型和访问级别取决于当前项目的模式。</span><span class="sxs-lookup"><span data-stu-id="59995-658">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="59995-659">阅读模式</span><span class="sxs-lookup"><span data-stu-id="59995-659">Read mode</span></span>

<span data-ttu-id="59995-p135">`to` 属性返回包含邮件的**收件人**行上所列的每个收件人的 `EmailAddressDetails` 对象的数组。集合上限为 100 个成员。</span><span class="sxs-lookup"><span data-stu-id="59995-p135">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.</span></span>

```javascript
console.log(JSON.stringify(Office.context.mailbox.item.to));
```

##### <a name="compose-mode"></a><span data-ttu-id="59995-662">撰写模式</span><span class="sxs-lookup"><span data-stu-id="59995-662">Compose mode</span></span>

<span data-ttu-id="59995-663">`to` 属性返回一个 `Recipients` 对象，该对象提供用于获取或更新邮件的“**收件人**”行上收件人的方法。</span><span class="sxs-lookup"><span data-stu-id="59995-663">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span>

```javascript
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="59995-664">类型</span><span class="sxs-lookup"><span data-stu-id="59995-664">Type</span></span>

*   <span data-ttu-id="59995-665">Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_7/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="59995-665">Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_7/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="59995-666">Requirements</span><span class="sxs-lookup"><span data-stu-id="59995-666">Requirements</span></span>

|<span data-ttu-id="59995-667">要求</span><span class="sxs-lookup"><span data-stu-id="59995-667">Requirement</span></span>|<span data-ttu-id="59995-668">值</span><span class="sxs-lookup"><span data-stu-id="59995-668">Value</span></span>|
|---|---|
|[<span data-ttu-id="59995-669">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="59995-669">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="59995-670">1.0</span><span class="sxs-lookup"><span data-stu-id="59995-670">1.0</span></span>|
|[<span data-ttu-id="59995-671">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="59995-671">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="59995-672">ReadItem</span><span class="sxs-lookup"><span data-stu-id="59995-672">ReadItem</span></span>|
|[<span data-ttu-id="59995-673">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="59995-673">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="59995-674">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="59995-674">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="59995-675">方法</span><span class="sxs-lookup"><span data-stu-id="59995-675">Methods</span></span>

####  <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="59995-676">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="59995-676">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="59995-677">将文件作为附件添加到邮件或约会。</span><span class="sxs-lookup"><span data-stu-id="59995-677">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="59995-678">`addFileAttachmentAsync` 方法在指定的 URI 上载文件并将其附加到撰写窗体中的项目。</span><span class="sxs-lookup"><span data-stu-id="59995-678">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="59995-679">你随后可以将该标识符与 [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) 方法一同使用，以删除同一个会话中的附件。</span><span class="sxs-lookup"><span data-stu-id="59995-679">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="59995-680">Parameters</span><span class="sxs-lookup"><span data-stu-id="59995-680">Parameters</span></span>
|<span data-ttu-id="59995-681">名称</span><span class="sxs-lookup"><span data-stu-id="59995-681">Name</span></span>|<span data-ttu-id="59995-682">类型</span><span class="sxs-lookup"><span data-stu-id="59995-682">Type</span></span>|<span data-ttu-id="59995-683">属性</span><span class="sxs-lookup"><span data-stu-id="59995-683">Attributes</span></span>|<span data-ttu-id="59995-684">说明</span><span class="sxs-lookup"><span data-stu-id="59995-684">Description</span></span>|
|---|---|---|---|
|`uri`|<span data-ttu-id="59995-685">String</span><span class="sxs-lookup"><span data-stu-id="59995-685">String</span></span>||<span data-ttu-id="59995-p136">提供附加到邮件或约会的文件的位置的 URI。最大长度为 2048 个字符。</span><span class="sxs-lookup"><span data-stu-id="59995-p136">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`|<span data-ttu-id="59995-688">字符串</span><span class="sxs-lookup"><span data-stu-id="59995-688">String</span></span>||<span data-ttu-id="59995-p137">在附件上载过程中显示的附件名称。最大长度为 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="59995-p137">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="59995-691">Object</span><span class="sxs-lookup"><span data-stu-id="59995-691">Object</span></span>|<span data-ttu-id="59995-692">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="59995-692">&lt;optional&gt;</span></span>|<span data-ttu-id="59995-693">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="59995-693">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="59995-694">对象</span><span class="sxs-lookup"><span data-stu-id="59995-694">Object</span></span>|<span data-ttu-id="59995-695">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="59995-695">&lt;optional&gt;</span></span>|<span data-ttu-id="59995-696">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="59995-696">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.isInline`|<span data-ttu-id="59995-697">布尔值</span><span class="sxs-lookup"><span data-stu-id="59995-697">Boolean</span></span>|<span data-ttu-id="59995-698">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="59995-698">&lt;optional&gt;</span></span>|<span data-ttu-id="59995-699">如果为 `true`，则表示附件将在邮件正文中内联显示，并且不应显示在附件列表中。</span><span class="sxs-lookup"><span data-stu-id="59995-699">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`callback`|<span data-ttu-id="59995-700">函数</span><span class="sxs-lookup"><span data-stu-id="59995-700">function</span></span>|<span data-ttu-id="59995-701">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="59995-701">&lt;optional&gt;</span></span>|<span data-ttu-id="59995-702">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="59995-702">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="59995-703">如果成功，附件标识符将在 `asyncResult.value` 属性中提供。</span><span class="sxs-lookup"><span data-stu-id="59995-703">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="59995-704">如果上传附件失败，`asyncResult` 对象将包含一个提供错误说明的 `Error` 对象。</span><span class="sxs-lookup"><span data-stu-id="59995-704">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="59995-705">错误</span><span class="sxs-lookup"><span data-stu-id="59995-705">Errors</span></span>

|<span data-ttu-id="59995-706">错误代码</span><span class="sxs-lookup"><span data-stu-id="59995-706">Error code</span></span>|<span data-ttu-id="59995-707">说明</span><span class="sxs-lookup"><span data-stu-id="59995-707">Description</span></span>|
|------------|-------------|
|`AttachmentSizeExceeded`|<span data-ttu-id="59995-708">附件大小超过了允许的大小。</span><span class="sxs-lookup"><span data-stu-id="59995-708">The attachment is larger than allowed.</span></span>|
|`FileTypeNotSupported`|<span data-ttu-id="59995-709">该附件的扩展名不是允许的扩展名。</span><span class="sxs-lookup"><span data-stu-id="59995-709">The attachment has an extension that is not allowed.</span></span>|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="59995-710">邮件或约会具有的附件过多。</span><span class="sxs-lookup"><span data-stu-id="59995-710">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="59995-711">要求</span><span class="sxs-lookup"><span data-stu-id="59995-711">Requirements</span></span>

|<span data-ttu-id="59995-712">要求</span><span class="sxs-lookup"><span data-stu-id="59995-712">Requirement</span></span>|<span data-ttu-id="59995-713">值</span><span class="sxs-lookup"><span data-stu-id="59995-713">Value</span></span>|
|---|---|
|[<span data-ttu-id="59995-714">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="59995-714">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="59995-715">1.1</span><span class="sxs-lookup"><span data-stu-id="59995-715">1.1</span></span>|
|[<span data-ttu-id="59995-716">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="59995-716">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="59995-717">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="59995-717">ReadWriteItem</span></span>|
|[<span data-ttu-id="59995-718">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="59995-718">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="59995-719">撰写</span><span class="sxs-lookup"><span data-stu-id="59995-719">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="59995-720">示例</span><span class="sxs-lookup"><span data-stu-id="59995-720">Examples</span></span>

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

<span data-ttu-id="59995-721">以下示例将图像文件添加为内联附件，并在邮件正文中引用该附件。</span><span class="sxs-lookup"><span data-stu-id="59995-721">The following example adds an image file as an inline attachment and references the attachment in the message body.</span></span>

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

####  <a name="addhandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="59995-722">addHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="59995-722">addHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="59995-723">添加支持事件的事件处理程序。</span><span class="sxs-lookup"><span data-stu-id="59995-723">Adds an event handler for a supported event.</span></span>

<span data-ttu-id="59995-724">当前，支持的事件类型是 `Office.EventType.AppointmentTimeChanged`、`Office.EventType.RecipientsChanged` 和 `Office.EventType.RecurrenceChanged`</span><span class="sxs-lookup"><span data-stu-id="59995-724">Currently the supported event types are `Office.EventType.AppointmentTimeChanged`, `Office.EventType.RecipientsChanged`, and `Office.EventType.RecurrenceChanged`</span></span>

##### <a name="parameters"></a><span data-ttu-id="59995-725">Parameters</span><span class="sxs-lookup"><span data-stu-id="59995-725">Parameters</span></span>

| <span data-ttu-id="59995-726">名称</span><span class="sxs-lookup"><span data-stu-id="59995-726">Name</span></span> | <span data-ttu-id="59995-727">类型</span><span class="sxs-lookup"><span data-stu-id="59995-727">Type</span></span> | <span data-ttu-id="59995-728">属性</span><span class="sxs-lookup"><span data-stu-id="59995-728">Attributes</span></span> | <span data-ttu-id="59995-729">说明</span><span class="sxs-lookup"><span data-stu-id="59995-729">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="59995-730">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="59995-730">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="59995-731">应调用处理程序的事件。</span><span class="sxs-lookup"><span data-stu-id="59995-731">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="59995-732">函数</span><span class="sxs-lookup"><span data-stu-id="59995-732">Function</span></span> || <span data-ttu-id="59995-p138">用于处理事件的函数。此函数必须接受一个参数，即对象文本。参数上的 `type` 属性将匹配传递给 `addHandlerAsync` 的 `eventType` 参数。</span><span class="sxs-lookup"><span data-stu-id="59995-p138">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `addHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="59995-736">Object</span><span class="sxs-lookup"><span data-stu-id="59995-736">Object</span></span> | <span data-ttu-id="59995-737">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="59995-737">&lt;optional&gt;</span></span> | <span data-ttu-id="59995-738">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="59995-738">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="59995-739">对象</span><span class="sxs-lookup"><span data-stu-id="59995-739">Object</span></span> | <span data-ttu-id="59995-740">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="59995-740">&lt;optional&gt;</span></span> | <span data-ttu-id="59995-741">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="59995-741">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="59995-742">函数</span><span class="sxs-lookup"><span data-stu-id="59995-742">function</span></span>| <span data-ttu-id="59995-743">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="59995-743">&lt;optional&gt;</span></span>|<span data-ttu-id="59995-744">方法完成后，使用单个参数 `callback`（一个 [`asyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `AsyncResult` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="59995-744">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="59995-745">Requirements</span><span class="sxs-lookup"><span data-stu-id="59995-745">Requirements</span></span>

|<span data-ttu-id="59995-746">要求</span><span class="sxs-lookup"><span data-stu-id="59995-746">Requirement</span></span>| <span data-ttu-id="59995-747">值</span><span class="sxs-lookup"><span data-stu-id="59995-747">Value</span></span>|
|---|---|
|[<span data-ttu-id="59995-748">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="59995-748">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="59995-749">1.7</span><span class="sxs-lookup"><span data-stu-id="59995-749">1.7</span></span> |
|[<span data-ttu-id="59995-750">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="59995-750">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="59995-751">ReadItem</span><span class="sxs-lookup"><span data-stu-id="59995-751">ReadItem</span></span> |
|[<span data-ttu-id="59995-752">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="59995-752">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="59995-753">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="59995-753">Compose or Read</span></span> |

##### <a name="example"></a><span data-ttu-id="59995-754">示例</span><span class="sxs-lookup"><span data-stu-id="59995-754">Example</span></span>

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

####  <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="59995-755">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="59995-755">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="59995-756">将 Exchange 项目（如邮件）作为附件添加到邮件或约会。</span><span class="sxs-lookup"><span data-stu-id="59995-756">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="59995-p139">`addItemAttachmentAsync` 方法将包含指定 Exchange 标识符的项目附加到撰写窗体中的项目。如果指定一个回调方法，此方法使用 `asyncResult` 参数调用，该参数包含一个附件标识符或代码，指示附加项目过程中出现的任何错误。可以使用 `options` 参数将状态信息传递给回调方法（如果需要）。</span><span class="sxs-lookup"><span data-stu-id="59995-p139">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="59995-760">随后可以将该标识符与 [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) 方法一同使用，以删除同一个会话中的附件。</span><span class="sxs-lookup"><span data-stu-id="59995-760">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="59995-761">如果 Office 加载项在 Outlook Web App 中运行，则 `addItemAttachmentAsync` 方法可以将项目附加到项目（正在编辑的项目除外）中；然而，不支持也不建议这样做。</span><span class="sxs-lookup"><span data-stu-id="59995-761">If your Office Add-in is running in Outlook Web App, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="59995-762">Parameters</span><span class="sxs-lookup"><span data-stu-id="59995-762">Parameters</span></span>

|<span data-ttu-id="59995-763">名称</span><span class="sxs-lookup"><span data-stu-id="59995-763">Name</span></span>|<span data-ttu-id="59995-764">类型</span><span class="sxs-lookup"><span data-stu-id="59995-764">Type</span></span>|<span data-ttu-id="59995-765">属性</span><span class="sxs-lookup"><span data-stu-id="59995-765">Attributes</span></span>|<span data-ttu-id="59995-766">说明</span><span class="sxs-lookup"><span data-stu-id="59995-766">Description</span></span>|
|---|---|---|---|
|`itemId`|<span data-ttu-id="59995-767">String</span><span class="sxs-lookup"><span data-stu-id="59995-767">String</span></span>||<span data-ttu-id="59995-p140">要附加的项目的 Exchange 标识符。最大长度为 100 个字符。</span><span class="sxs-lookup"><span data-stu-id="59995-p140">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`|<span data-ttu-id="59995-770">字符串</span><span class="sxs-lookup"><span data-stu-id="59995-770">String</span></span>||<span data-ttu-id="59995-771">要附加的项目的主题。</span><span class="sxs-lookup"><span data-stu-id="59995-771">The subject of the item to be attached.</span></span> <span data-ttu-id="59995-772">最大长度为 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="59995-772">The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="59995-773">对象</span><span class="sxs-lookup"><span data-stu-id="59995-773">Object</span></span>|<span data-ttu-id="59995-774">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="59995-774">&lt;optional&gt;</span></span>|<span data-ttu-id="59995-775">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="59995-775">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="59995-776">对象</span><span class="sxs-lookup"><span data-stu-id="59995-776">Object</span></span>|<span data-ttu-id="59995-777">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="59995-777">&lt;optional&gt;</span></span>|<span data-ttu-id="59995-778">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="59995-778">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="59995-779">函数</span><span class="sxs-lookup"><span data-stu-id="59995-779">function</span></span>|<span data-ttu-id="59995-780">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="59995-780">&lt;optional&gt;</span></span>|<span data-ttu-id="59995-781">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="59995-781">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="59995-782">如果成功，附件标识符将在 `asyncResult.value` 属性中提供。</span><span class="sxs-lookup"><span data-stu-id="59995-782">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="59995-783">如果添加附件失败，`asyncResult` 对象将包含一个提供错误说明的 `Error` 对象。</span><span class="sxs-lookup"><span data-stu-id="59995-783">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="59995-784">错误</span><span class="sxs-lookup"><span data-stu-id="59995-784">Errors</span></span>

|<span data-ttu-id="59995-785">错误代码</span><span class="sxs-lookup"><span data-stu-id="59995-785">Error code</span></span>|<span data-ttu-id="59995-786">说明</span><span class="sxs-lookup"><span data-stu-id="59995-786">Description</span></span>|
|------------|-------------|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="59995-787">邮件或约会具有的附件过多。</span><span class="sxs-lookup"><span data-stu-id="59995-787">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="59995-788">Requirements</span><span class="sxs-lookup"><span data-stu-id="59995-788">Requirements</span></span>

|<span data-ttu-id="59995-789">要求</span><span class="sxs-lookup"><span data-stu-id="59995-789">Requirement</span></span>|<span data-ttu-id="59995-790">值</span><span class="sxs-lookup"><span data-stu-id="59995-790">Value</span></span>|
|---|---|
|[<span data-ttu-id="59995-791">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="59995-791">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="59995-792">1.1</span><span class="sxs-lookup"><span data-stu-id="59995-792">1.1</span></span>|
|[<span data-ttu-id="59995-793">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="59995-793">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="59995-794">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="59995-794">ReadWriteItem</span></span>|
|[<span data-ttu-id="59995-795">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="59995-795">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="59995-796">撰写</span><span class="sxs-lookup"><span data-stu-id="59995-796">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="59995-797">示例</span><span class="sxs-lookup"><span data-stu-id="59995-797">Example</span></span>

<span data-ttu-id="59995-798">以下示例将现有的 Outlook 项目添加为名为 `My Attachment` 的附件。</span><span class="sxs-lookup"><span data-stu-id="59995-798">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

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

####  <a name="close"></a><span data-ttu-id="59995-799">close()</span><span class="sxs-lookup"><span data-stu-id="59995-799">close()</span></span>

<span data-ttu-id="59995-800">关闭当前正在撰写的项目。</span><span class="sxs-lookup"><span data-stu-id="59995-800">Closes the current item that is being composed.</span></span>

<span data-ttu-id="59995-p142">
            \`close\` 方法的行为取决于要撰写的项目的当前状态。如果项目具有未保存的更改，客户端将提示用户保存、放弃或取消关闭操作。</span><span class="sxs-lookup"><span data-stu-id="59995-p142">The behavior of the `close` method depends on the current state of the item being composed. If the item has unsaved changes, the client prompts the user to save, discard, or cancel the close action.</span></span>

> [!NOTE]
> <span data-ttu-id="59995-803">在 Outlook 网页版中，如果该项目是约会并且之前已使用 `saveAsync` 保存，则即使自上次保存项目后未发生任何更改，也会提示用户保存、放弃或取消。</span><span class="sxs-lookup"><span data-stu-id="59995-803">In Outlook on the web, if the item is an appointment and it has previously been saved using `saveAsync`, the user is prompted to save, discard, or cancel even if no changes have occurred since the item was last saved.</span></span>

<span data-ttu-id="59995-804">在 Outlook 桌面客户端中，如果邮件是内联答复，`close` 方法不起作用。</span><span class="sxs-lookup"><span data-stu-id="59995-804">In the Outlook desktop client, if the message is an inline reply, the `close` method has no effect.</span></span>

##### <a name="requirements"></a><span data-ttu-id="59995-805">要求</span><span class="sxs-lookup"><span data-stu-id="59995-805">Requirements</span></span>

|<span data-ttu-id="59995-806">要求</span><span class="sxs-lookup"><span data-stu-id="59995-806">Requirement</span></span>|<span data-ttu-id="59995-807">值</span><span class="sxs-lookup"><span data-stu-id="59995-807">Value</span></span>|
|---|---|
|[<span data-ttu-id="59995-808">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="59995-808">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="59995-809">1.3</span><span class="sxs-lookup"><span data-stu-id="59995-809">1.3</span></span>|
|[<span data-ttu-id="59995-810">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="59995-810">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="59995-811">受限</span><span class="sxs-lookup"><span data-stu-id="59995-811">Restricted</span></span>|
|[<span data-ttu-id="59995-812">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="59995-812">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="59995-813">撰写</span><span class="sxs-lookup"><span data-stu-id="59995-813">Compose</span></span>|

#### <a name="displayreplyallformformdata-callback"></a><span data-ttu-id="59995-814">displayReplyAllForm (formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="59995-814">displayReplyAllForm(formData, [callback])</span></span>

<span data-ttu-id="59995-815">显示答复窗体，其中包括所选邮件的发件人和所有收件人或所选约会的组织者和所有与会者。</span><span class="sxs-lookup"><span data-stu-id="59995-815">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="59995-816">在 Outlook for iOS 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="59995-816">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="59995-817">在 Outlook Web App 中，答复窗体显示为包含 3 列视图的弹出式窗体以及包含 2 列或 1 列视图的弹出式窗体。</span><span class="sxs-lookup"><span data-stu-id="59995-817">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="59995-818">如果任意字符串参数超出其限制，`displayReplyAllForm` 将引发异常。</span><span class="sxs-lookup"><span data-stu-id="59995-818">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="59995-p143">当在 `formData.attachments` 参数中指定附件时，Outlook 和 Outlook Web App 尝试下载所有附件并将其附加到答复窗体。如果无法添加任何附件，则在窗体 UI 中显示错误。如果这不可能，则不引发错误消息。</span><span class="sxs-lookup"><span data-stu-id="59995-p143">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="59995-822">Parameters</span><span class="sxs-lookup"><span data-stu-id="59995-822">Parameters</span></span>

|<span data-ttu-id="59995-823">名称</span><span class="sxs-lookup"><span data-stu-id="59995-823">Name</span></span>|<span data-ttu-id="59995-824">类型</span><span class="sxs-lookup"><span data-stu-id="59995-824">Type</span></span>|<span data-ttu-id="59995-825">属性</span><span class="sxs-lookup"><span data-stu-id="59995-825">Attributes</span></span>|<span data-ttu-id="59995-826">说明</span><span class="sxs-lookup"><span data-stu-id="59995-826">Description</span></span>|
|---|---|---|---|
|`formData`|<span data-ttu-id="59995-827">字符串 &#124; 对象</span><span class="sxs-lookup"><span data-stu-id="59995-827">String &#124; Object</span></span>||<span data-ttu-id="59995-p144">一个包含文本和 HTML 且表示答复窗体的正文的字符串。字符串限制为 32 KB。</span><span class="sxs-lookup"><span data-stu-id="59995-p144">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="59995-830">**或**</span><span class="sxs-lookup"><span data-stu-id="59995-830">**OR**</span></span><br/><span data-ttu-id="59995-p145">包含正文或附件数据和回调函数的对象。对象定义如下。</span><span class="sxs-lookup"><span data-stu-id="59995-p145">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span>|
|`formData.htmlBody`|<span data-ttu-id="59995-833">String</span><span class="sxs-lookup"><span data-stu-id="59995-833">String</span></span>|<span data-ttu-id="59995-834">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="59995-834">&lt;optional&gt;</span></span>|<span data-ttu-id="59995-p146">一个包含文本和 HTML 且表示答复窗体的正文的字符串。字符串限制为 32 KB。</span><span class="sxs-lookup"><span data-stu-id="59995-p146">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
|`formData.attachments`|<span data-ttu-id="59995-837">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="59995-837">Array.&lt;Object&gt;</span></span>|<span data-ttu-id="59995-838">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="59995-838">&lt;optional&gt;</span></span>|<span data-ttu-id="59995-839">JSON 对象（是文件或项目附件）的数组。</span><span class="sxs-lookup"><span data-stu-id="59995-839">An array of JSON objects that are either file or item attachments.</span></span>|
|`formData.attachments.type`|<span data-ttu-id="59995-840">String</span><span class="sxs-lookup"><span data-stu-id="59995-840">String</span></span>||<span data-ttu-id="59995-p147">指示附件的类型。必须是文件附件的 `file` 或项目附件的 `item`。</span><span class="sxs-lookup"><span data-stu-id="59995-p147">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span>|
|`formData.attachments.name`|<span data-ttu-id="59995-843">字符串</span><span class="sxs-lookup"><span data-stu-id="59995-843">String</span></span>||<span data-ttu-id="59995-844">一个包含附件的名称的字符串，最多包含 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="59995-844">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
|`formData.attachments.url`|<span data-ttu-id="59995-845">String</span><span class="sxs-lookup"><span data-stu-id="59995-845">String</span></span>||<span data-ttu-id="59995-p148">仅在将 `type` 设置为 `file` 时使用。文件的位置的 URI。</span><span class="sxs-lookup"><span data-stu-id="59995-p148">Only used if `type` is set to `file`. The URI of the location for the file.</span></span>|
|`formData.attachments.isInline`|<span data-ttu-id="59995-848">Boolean</span><span class="sxs-lookup"><span data-stu-id="59995-848">Boolean</span></span>||<span data-ttu-id="59995-p149">仅在将 `type` 设置为 `file` 时使用。如果为 `true`，则表示附件将在邮件正文中内联显示，并且不应显示在附件列表中。</span><span class="sxs-lookup"><span data-stu-id="59995-p149">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`formData.attachments.itemId`|<span data-ttu-id="59995-851">String</span><span class="sxs-lookup"><span data-stu-id="59995-851">String</span></span>||<span data-ttu-id="59995-p150">仅在将 `type` 设置为 `item` 时使用。附件的 EWS 项目 ID。字符串最长为 100 个字符。</span><span class="sxs-lookup"><span data-stu-id="59995-p150">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span>|
|`callback`|<span data-ttu-id="59995-855">函数</span><span class="sxs-lookup"><span data-stu-id="59995-855">function</span></span>|<span data-ttu-id="59995-856">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="59995-856">&lt;optional&gt;</span></span>|<span data-ttu-id="59995-857">方法完成后，使用单个参数 `asyncResult`（一个 [AsyncResult](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="59995-857">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="59995-858">Requirements</span><span class="sxs-lookup"><span data-stu-id="59995-858">Requirements</span></span>

|<span data-ttu-id="59995-859">要求</span><span class="sxs-lookup"><span data-stu-id="59995-859">Requirement</span></span>|<span data-ttu-id="59995-860">值</span><span class="sxs-lookup"><span data-stu-id="59995-860">Value</span></span>|
|---|---|
|[<span data-ttu-id="59995-861">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="59995-861">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="59995-862">1.0</span><span class="sxs-lookup"><span data-stu-id="59995-862">1.0</span></span>|
|[<span data-ttu-id="59995-863">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="59995-863">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="59995-864">ReadItem</span><span class="sxs-lookup"><span data-stu-id="59995-864">ReadItem</span></span>|
|[<span data-ttu-id="59995-865">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="59995-865">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="59995-866">阅读</span><span class="sxs-lookup"><span data-stu-id="59995-866">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="59995-867">示例</span><span class="sxs-lookup"><span data-stu-id="59995-867">Examples</span></span>

<span data-ttu-id="59995-868">以下代码将一个字符串传递到 `displayReplyAllForm` 函数。</span><span class="sxs-lookup"><span data-stu-id="59995-868">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="59995-869">使用空白正文答复。</span><span class="sxs-lookup"><span data-stu-id="59995-869">Reply with an empty body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="59995-870">仅使用正文答复。</span><span class="sxs-lookup"><span data-stu-id="59995-870">Reply with just a body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="59995-871">使用正文和文件附件答复。</span><span class="sxs-lookup"><span data-stu-id="59995-871">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="59995-872">使用正文和项目附件答复。</span><span class="sxs-lookup"><span data-stu-id="59995-872">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="59995-873">使用正文、文件附件、项目附件和回调答复。</span><span class="sxs-lookup"><span data-stu-id="59995-873">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="displayreplyformformdata-callback"></a><span data-ttu-id="59995-874">displayReplyForm (formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="59995-874">displayReplyForm(formData, [callback])</span></span>

<span data-ttu-id="59995-875">显示答复窗体，其中仅包括所选邮件的发件人或所选约会的组织者。</span><span class="sxs-lookup"><span data-stu-id="59995-875">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="59995-876">在 Outlook for iOS 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="59995-876">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="59995-877">在 Outlook Web App 中，答复窗体显示为包含 3 列视图的弹出式窗体以及包含 2 列或 1 列视图的弹出式窗体。</span><span class="sxs-lookup"><span data-stu-id="59995-877">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="59995-878">如果任意字符串参数超出其限制，`displayReplyForm` 将引发异常。</span><span class="sxs-lookup"><span data-stu-id="59995-878">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="59995-p151">当在 `formData.attachments` 参数中指定附件时，Outlook 和 Outlook Web App 尝试下载所有附件并将其附加到答复窗体。如果无法添加任何附件，则在窗体 UI 中显示错误。如果这不可能，则不引发错误消息。</span><span class="sxs-lookup"><span data-stu-id="59995-p151">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="59995-882">Parameters</span><span class="sxs-lookup"><span data-stu-id="59995-882">Parameters</span></span>

|<span data-ttu-id="59995-883">名称</span><span class="sxs-lookup"><span data-stu-id="59995-883">Name</span></span>|<span data-ttu-id="59995-884">类型</span><span class="sxs-lookup"><span data-stu-id="59995-884">Type</span></span>|<span data-ttu-id="59995-885">属性</span><span class="sxs-lookup"><span data-stu-id="59995-885">Attributes</span></span>|<span data-ttu-id="59995-886">说明</span><span class="sxs-lookup"><span data-stu-id="59995-886">Description</span></span>|
|---|---|---|---|
|`formData`|<span data-ttu-id="59995-887">字符串 &#124; 对象</span><span class="sxs-lookup"><span data-stu-id="59995-887">String &#124; Object</span></span>||<span data-ttu-id="59995-p152">一个包含文本和 HTML 且表示答复窗体的正文的字符串。字符串限制为 32 KB。</span><span class="sxs-lookup"><span data-stu-id="59995-p152">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="59995-890">**或**</span><span class="sxs-lookup"><span data-stu-id="59995-890">**OR**</span></span><br/><span data-ttu-id="59995-p153">包含正文或附件数据和回调函数的对象。对象定义如下。</span><span class="sxs-lookup"><span data-stu-id="59995-p153">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span>|
|`formData.htmlBody`|<span data-ttu-id="59995-893">String</span><span class="sxs-lookup"><span data-stu-id="59995-893">String</span></span>|<span data-ttu-id="59995-894">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="59995-894">&lt;optional&gt;</span></span>|<span data-ttu-id="59995-p154">一个包含文本和 HTML 且表示答复窗体的正文的字符串。字符串限制为 32 KB。</span><span class="sxs-lookup"><span data-stu-id="59995-p154">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
|`formData.attachments`|<span data-ttu-id="59995-897">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="59995-897">Array.&lt;Object&gt;</span></span>|<span data-ttu-id="59995-898">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="59995-898">&lt;optional&gt;</span></span>|<span data-ttu-id="59995-899">JSON 对象（是文件或项目附件）的数组。</span><span class="sxs-lookup"><span data-stu-id="59995-899">An array of JSON objects that are either file or item attachments.</span></span>|
|`formData.attachments.type`|<span data-ttu-id="59995-900">String</span><span class="sxs-lookup"><span data-stu-id="59995-900">String</span></span>||<span data-ttu-id="59995-p155">指示附件的类型。必须是文件附件的 `file` 或项目附件的 `item`。</span><span class="sxs-lookup"><span data-stu-id="59995-p155">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span>|
|`formData.attachments.name`|<span data-ttu-id="59995-903">字符串</span><span class="sxs-lookup"><span data-stu-id="59995-903">String</span></span>||<span data-ttu-id="59995-904">一个包含附件的名称的字符串，最多包含 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="59995-904">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
|`formData.attachments.url`|<span data-ttu-id="59995-905">String</span><span class="sxs-lookup"><span data-stu-id="59995-905">String</span></span>||<span data-ttu-id="59995-p156">仅在将 `type` 设置为 `file` 时使用。文件的位置的 URI。</span><span class="sxs-lookup"><span data-stu-id="59995-p156">Only used if `type` is set to `file`. The URI of the location for the file.</span></span>|
|`formData.attachments.isInline`|<span data-ttu-id="59995-908">Boolean</span><span class="sxs-lookup"><span data-stu-id="59995-908">Boolean</span></span>||<span data-ttu-id="59995-p157">仅在将 `type` 设置为 `file` 时使用。如果为 `true`，则表示附件将在邮件正文中内联显示，并且不应显示在附件列表中。</span><span class="sxs-lookup"><span data-stu-id="59995-p157">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`formData.attachments.itemId`|<span data-ttu-id="59995-911">String</span><span class="sxs-lookup"><span data-stu-id="59995-911">String</span></span>||<span data-ttu-id="59995-p158">仅在将 `type` 设置为 `item` 时使用。附件的 EWS 项目 ID。字符串最长为 100 个字符。</span><span class="sxs-lookup"><span data-stu-id="59995-p158">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span>|
|`callback`|<span data-ttu-id="59995-915">函数</span><span class="sxs-lookup"><span data-stu-id="59995-915">function</span></span>|<span data-ttu-id="59995-916">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="59995-916">&lt;optional&gt;</span></span>|<span data-ttu-id="59995-917">方法完成后，使用单个参数 `asyncResult`（一个 [AsyncResult](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="59995-917">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="59995-918">Requirements</span><span class="sxs-lookup"><span data-stu-id="59995-918">Requirements</span></span>

|<span data-ttu-id="59995-919">要求</span><span class="sxs-lookup"><span data-stu-id="59995-919">Requirement</span></span>|<span data-ttu-id="59995-920">值</span><span class="sxs-lookup"><span data-stu-id="59995-920">Value</span></span>|
|---|---|
|[<span data-ttu-id="59995-921">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="59995-921">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="59995-922">1.0</span><span class="sxs-lookup"><span data-stu-id="59995-922">1.0</span></span>|
|[<span data-ttu-id="59995-923">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="59995-923">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="59995-924">ReadItem</span><span class="sxs-lookup"><span data-stu-id="59995-924">ReadItem</span></span>|
|[<span data-ttu-id="59995-925">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="59995-925">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="59995-926">阅读</span><span class="sxs-lookup"><span data-stu-id="59995-926">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="59995-927">示例</span><span class="sxs-lookup"><span data-stu-id="59995-927">Examples</span></span>

<span data-ttu-id="59995-928">以下代码将一个字符串传递到 `displayReplyForm` 函数。</span><span class="sxs-lookup"><span data-stu-id="59995-928">The following code passes a string to the `displayReplyForm` function.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="59995-929">使用空白正文答复。</span><span class="sxs-lookup"><span data-stu-id="59995-929">Reply with an empty body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="59995-930">仅使用正文答复。</span><span class="sxs-lookup"><span data-stu-id="59995-930">Reply with just a body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="59995-931">使用正文和文件附件答复。</span><span class="sxs-lookup"><span data-stu-id="59995-931">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="59995-932">使用正文和项目附件答复。</span><span class="sxs-lookup"><span data-stu-id="59995-932">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="59995-933">使用正文、文件附件、项目附件和回调答复。</span><span class="sxs-lookup"><span data-stu-id="59995-933">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="getentities--entitiesjavascriptapioutlook17officeentities"></a><span data-ttu-id="59995-934">getEntities() → {[Entities](/javascript/api/outlook_1_7/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="59995-934">getEntities() → {[Entities](/javascript/api/outlook_1_7/office.entities)}</span></span>

<span data-ttu-id="59995-935">获取在所选项目的正文中找到的实体。</span><span class="sxs-lookup"><span data-stu-id="59995-935">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="59995-936">在 Outlook for iOS 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="59995-936">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="59995-937">要求</span><span class="sxs-lookup"><span data-stu-id="59995-937">Requirements</span></span>

|<span data-ttu-id="59995-938">要求</span><span class="sxs-lookup"><span data-stu-id="59995-938">Requirement</span></span>|<span data-ttu-id="59995-939">值</span><span class="sxs-lookup"><span data-stu-id="59995-939">Value</span></span>|
|---|---|
|[<span data-ttu-id="59995-940">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="59995-940">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="59995-941">1.0</span><span class="sxs-lookup"><span data-stu-id="59995-941">1.0</span></span>|
|[<span data-ttu-id="59995-942">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="59995-942">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="59995-943">ReadItem</span><span class="sxs-lookup"><span data-stu-id="59995-943">ReadItem</span></span>|
|[<span data-ttu-id="59995-944">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="59995-944">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="59995-945">阅读</span><span class="sxs-lookup"><span data-stu-id="59995-945">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="59995-946">返回：</span><span class="sxs-lookup"><span data-stu-id="59995-946">Returns:</span></span>

<span data-ttu-id="59995-947">类型：[Entities](/javascript/api/outlook_1_7/office.entities)</span><span class="sxs-lookup"><span data-stu-id="59995-947">Type: [Entities](/javascript/api/outlook_1_7/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="59995-948">示例</span><span class="sxs-lookup"><span data-stu-id="59995-948">Example</span></span>

<span data-ttu-id="59995-949">以下示例访问当前项目的正文中的联系人实体。</span><span class="sxs-lookup"><span data-stu-id="59995-949">The following example accesses the contacts entities in the current item's body.</span></span>

```javascript
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlook17officecontactmeetingsuggestionjavascriptapioutlook17officemeetingsuggestionphonenumberjavascriptapioutlook17officephonenumbertasksuggestionjavascriptapioutlook17officetasksuggestion"></a><span data-ttu-id="59995-950">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_7/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_7/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_7/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_7/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="59995-950">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_7/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_7/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_7/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_7/office.tasksuggestion))>}</span></span>

<span data-ttu-id="59995-951">获取所选项目的正文中找到的指定实体类型的所有实体的数组。</span><span class="sxs-lookup"><span data-stu-id="59995-951">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="59995-952">在 Outlook for iOS 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="59995-952">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="59995-953">Parameters</span><span class="sxs-lookup"><span data-stu-id="59995-953">Parameters</span></span>

|<span data-ttu-id="59995-954">名称</span><span class="sxs-lookup"><span data-stu-id="59995-954">Name</span></span>|<span data-ttu-id="59995-955">说明</span><span class="sxs-lookup"><span data-stu-id="59995-955">Type</span></span>|<span data-ttu-id="59995-956">说明</span><span class="sxs-lookup"><span data-stu-id="59995-956">Description</span></span>|
|---|---|---|
|`entityType`|[<span data-ttu-id="59995-957">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="59995-957">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook_1_7/office.mailboxenums.entitytype)|<span data-ttu-id="59995-958">EntityType 枚举值之一。</span><span class="sxs-lookup"><span data-stu-id="59995-958">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="59995-959">Requirements</span><span class="sxs-lookup"><span data-stu-id="59995-959">Requirements</span></span>

|<span data-ttu-id="59995-960">要求</span><span class="sxs-lookup"><span data-stu-id="59995-960">Requirement</span></span>|<span data-ttu-id="59995-961">值</span><span class="sxs-lookup"><span data-stu-id="59995-961">Value</span></span>|
|---|---|
|[<span data-ttu-id="59995-962">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="59995-962">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="59995-963">1.0</span><span class="sxs-lookup"><span data-stu-id="59995-963">1.0</span></span>|
|[<span data-ttu-id="59995-964">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="59995-964">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="59995-965">受限</span><span class="sxs-lookup"><span data-stu-id="59995-965">Restricted</span></span>|
|[<span data-ttu-id="59995-966">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="59995-966">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="59995-967">阅读</span><span class="sxs-lookup"><span data-stu-id="59995-967">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="59995-968">返回：</span><span class="sxs-lookup"><span data-stu-id="59995-968">Returns:</span></span>

<span data-ttu-id="59995-969">如果在 `entityType` 中传递的值不是 `EntityType` 枚举的有效成员，该方法返回 null。</span><span class="sxs-lookup"><span data-stu-id="59995-969">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="59995-970">如果指定类型的任何实体都不存在于该项目的正文中，该方法将返回空数组。</span><span class="sxs-lookup"><span data-stu-id="59995-970">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="59995-971">否则，返回的数组中对象的类型取决于 `entityType` 参数中请求实体的类型。</span><span class="sxs-lookup"><span data-stu-id="59995-971">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="59995-972">当使用此方法的最低权限级别**受限**时，某些实体类型需要 **ReadItem** 才能进行访问，如下表中所指定。</span><span class="sxs-lookup"><span data-stu-id="59995-972">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

|<span data-ttu-id="59995-973">`entityType` 的值</span><span class="sxs-lookup"><span data-stu-id="59995-973">Value of `entityType`</span></span>|<span data-ttu-id="59995-974">返回的数组中对象的类型</span><span class="sxs-lookup"><span data-stu-id="59995-974">Type of objects in returned array</span></span>|<span data-ttu-id="59995-975">所需权限级别</span><span class="sxs-lookup"><span data-stu-id="59995-975">Required Permission Level</span></span>|
|---|---|---|
|`Address`|<span data-ttu-id="59995-976">字符串</span><span class="sxs-lookup"><span data-stu-id="59995-976">String</span></span>|<span data-ttu-id="59995-977">**受限**</span><span class="sxs-lookup"><span data-stu-id="59995-977">**Restricted**</span></span>|
|`Contact`|<span data-ttu-id="59995-978">Contact</span><span class="sxs-lookup"><span data-stu-id="59995-978">Contact</span></span>|<span data-ttu-id="59995-979">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="59995-979">**ReadItem**</span></span>|
|`EmailAddress`|<span data-ttu-id="59995-980">String</span><span class="sxs-lookup"><span data-stu-id="59995-980">String</span></span>|<span data-ttu-id="59995-981">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="59995-981">**ReadItem**</span></span>|
|`MeetingSuggestion`|<span data-ttu-id="59995-982">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="59995-982">MeetingSuggestion</span></span>|<span data-ttu-id="59995-983">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="59995-983">**ReadItem**</span></span>|
|`PhoneNumber`|<span data-ttu-id="59995-984">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="59995-984">PhoneNumber</span></span>|<span data-ttu-id="59995-985">**受限**</span><span class="sxs-lookup"><span data-stu-id="59995-985">**Restricted**</span></span>|
|`TaskSuggestion`|<span data-ttu-id="59995-986">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="59995-986">TaskSuggestion</span></span>|<span data-ttu-id="59995-987">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="59995-987">**ReadItem**</span></span>|
|`URL`|<span data-ttu-id="59995-988">String</span><span class="sxs-lookup"><span data-stu-id="59995-988">String</span></span>|<span data-ttu-id="59995-989">**受限**</span><span class="sxs-lookup"><span data-stu-id="59995-989">**Restricted**</span></span>|

<span data-ttu-id="59995-990">类型：Array.<(String|[Contact](/javascript/api/outlook_1_7/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_7/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_7/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_7/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="59995-990">Type: Array.<(String|[Contact](/javascript/api/outlook_1_7/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_7/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_7/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_7/office.tasksuggestion))></span></span>

##### <a name="example"></a><span data-ttu-id="59995-991">示例</span><span class="sxs-lookup"><span data-stu-id="59995-991">Example</span></span>

<span data-ttu-id="59995-992">以下示例显示了如何访问表示当前项目的正文中的邮政地址的字符串数组。</span><span class="sxs-lookup"><span data-stu-id="59995-992">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlook17officecontactmeetingsuggestionjavascriptapioutlook17officemeetingsuggestionphonenumberjavascriptapioutlook17officephonenumbertasksuggestionjavascriptapioutlook17officetasksuggestion"></a><span data-ttu-id="59995-993">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_7/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_7/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_7/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_7/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="59995-993">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_7/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_7/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_7/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_7/office.tasksuggestion))>}</span></span>

<span data-ttu-id="59995-994">返回传递清单 XML 文件中定义的命名筛选器的所选项目中的已知实体。</span><span class="sxs-lookup"><span data-stu-id="59995-994">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="59995-995">在 Outlook for iOS 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="59995-995">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="59995-996">`getFilteredEntitiesByName` 方法返回匹配在具有指定 `FilterName` 元素值的清单 XML 文件中的 [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) 规则元素中定义的正则表达式的实体。</span><span class="sxs-lookup"><span data-stu-id="59995-996">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="59995-997">Parameters</span><span class="sxs-lookup"><span data-stu-id="59995-997">Parameters</span></span>

|<span data-ttu-id="59995-998">名称</span><span class="sxs-lookup"><span data-stu-id="59995-998">Name</span></span>|<span data-ttu-id="59995-999">说明</span><span class="sxs-lookup"><span data-stu-id="59995-999">Type</span></span>|<span data-ttu-id="59995-1000">说明</span><span class="sxs-lookup"><span data-stu-id="59995-1000">Description</span></span>|
|---|---|---|
|`name`|<span data-ttu-id="59995-1001">字符串</span><span class="sxs-lookup"><span data-stu-id="59995-1001">String</span></span>|<span data-ttu-id="59995-1002">定义筛选器匹配的 `ItemHasKnownEntity` 规则元素的名称。</span><span class="sxs-lookup"><span data-stu-id="59995-1002">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="59995-1003">要求</span><span class="sxs-lookup"><span data-stu-id="59995-1003">Requirements</span></span>

|<span data-ttu-id="59995-1004">要求</span><span class="sxs-lookup"><span data-stu-id="59995-1004">Requirement</span></span>|<span data-ttu-id="59995-1005">值</span><span class="sxs-lookup"><span data-stu-id="59995-1005">Value</span></span>|
|---|---|
|[<span data-ttu-id="59995-1006">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="59995-1006">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="59995-1007">1.0</span><span class="sxs-lookup"><span data-stu-id="59995-1007">1.0</span></span>|
|[<span data-ttu-id="59995-1008">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="59995-1008">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="59995-1009">ReadItem</span><span class="sxs-lookup"><span data-stu-id="59995-1009">ReadItem</span></span>|
|[<span data-ttu-id="59995-1010">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="59995-1010">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="59995-1011">阅读</span><span class="sxs-lookup"><span data-stu-id="59995-1011">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="59995-1012">返回：</span><span class="sxs-lookup"><span data-stu-id="59995-1012">Returns:</span></span>

<span data-ttu-id="59995-p160">如果具有匹配 `name` 参数的 `FilterName` 元素值的清单中没有任何 `ItemHasKnownEntity` 元素，则该方法返回 `null`。如果 `name` 参数匹配清单中的 `ItemHasKnownEntity` 元素，但在匹配的当前项目中没有实体，则该方法返回一个空数组。</span><span class="sxs-lookup"><span data-stu-id="59995-p160">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="59995-1015">类型：Array.<(String|[Contact](/javascript/api/outlook_1_7/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_7/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_7/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_7/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="59995-1015">Type: Array.<(String|[Contact](/javascript/api/outlook_1_7/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_7/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_7/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_7/office.tasksuggestion))></span></span>

#### <a name="getregexmatches--object"></a><span data-ttu-id="59995-1016">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="59995-1016">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="59995-1017">返回所选项目中匹配在清单 XML 文件中定义的正则表达式的字符串值。</span><span class="sxs-lookup"><span data-stu-id="59995-1017">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="59995-1018">在 Outlook for iOS 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="59995-1018">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="59995-p161">`getRegExMatches` 方法返回匹配在清单 XML 文件中的每个 `ItemHasRegularExpressionMatch` 或 `ItemHasKnownEntity` 规则元素中定义的正则表达式的字符串。对于 `ItemHasRegularExpressionMatch` 规则，匹配字符串必须发生在该规则指定的项目的属性中。`PropertyName` 简单类型定义支持的属性。</span><span class="sxs-lookup"><span data-stu-id="59995-p161">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="59995-1022">例如，考虑一个外接程序清单具有以下 `Rule` 元素：</span><span class="sxs-lookup"><span data-stu-id="59995-1022">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="59995-1023">从 `getRegExMatches` 返回的对象应有两个属性：`fruits` 和 `veggies`。</span><span class="sxs-lookup"><span data-stu-id="59995-1023">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="59995-p162">如果在项目的正文属性上指定 `ItemHasRegularExpressionMatch` 规则，则正则表达式应进一步筛选正文，不应尝试返回该项目的整个正文。使用正则表达式（如 `.*`）获取项目的整个正文并不总是返回预期的结果。而是使用 [`Body.getAsync`](/javascript/api/outlook_1_7/office.body#getasync-coerciontype--options--callback-) 方法检索整个正文。</span><span class="sxs-lookup"><span data-stu-id="59995-p162">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook_1_7/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="59995-1027">Requirements</span><span class="sxs-lookup"><span data-stu-id="59995-1027">Requirements</span></span>

|<span data-ttu-id="59995-1028">要求</span><span class="sxs-lookup"><span data-stu-id="59995-1028">Requirement</span></span>|<span data-ttu-id="59995-1029">值</span><span class="sxs-lookup"><span data-stu-id="59995-1029">Value</span></span>|
|---|---|
|[<span data-ttu-id="59995-1030">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="59995-1030">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="59995-1031">1.0</span><span class="sxs-lookup"><span data-stu-id="59995-1031">1.0</span></span>|
|[<span data-ttu-id="59995-1032">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="59995-1032">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="59995-1033">ReadItem</span><span class="sxs-lookup"><span data-stu-id="59995-1033">ReadItem</span></span>|
|[<span data-ttu-id="59995-1034">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="59995-1034">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="59995-1035">阅读</span><span class="sxs-lookup"><span data-stu-id="59995-1035">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="59995-1036">返回：</span><span class="sxs-lookup"><span data-stu-id="59995-1036">Returns:</span></span>

<span data-ttu-id="59995-p163">一个包含匹配在清单 XML 文件中定义的正则表达式的字符串数组的对象。每个数组的名称等于匹配 `ItemHasRegularExpressionMatch` 规则的 `RegExName` 属性或匹配 `ItemHasKnownEntity` 规则的 `FilterName` 属性的相应值。</span><span class="sxs-lookup"><span data-stu-id="59995-p163">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<dl class="param-type"><span data-ttu-id="59995-1039">

<dt>
类型</dt>


</span><span class="sxs-lookup"><span data-stu-id="59995-1039">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="59995-1040">对象</span><span class="sxs-lookup"><span data-stu-id="59995-1040">Object</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="59995-1041">示例</span><span class="sxs-lookup"><span data-stu-id="59995-1041">Example</span></span>

<span data-ttu-id="59995-1042">以下示例显示了如何访问正则表达式规则元素 `fruits` 和 `veggies` 的匹配项的数组，这些元素在清单中指定。</span><span class="sxs-lookup"><span data-stu-id="59995-1042">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```javascript
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veggies = allMatches.veggies;
```

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="59995-1043">getRegExMatchesByName(name) → (nullable) {Array.<String>}</span><span class="sxs-lookup"><span data-stu-id="59995-1043">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="59995-1044">返回匹配在清单 XML 文件中定义的命名正则表达式的所选项目中的字符串值。</span><span class="sxs-lookup"><span data-stu-id="59995-1044">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="59995-1045">在 Outlook for iOS 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="59995-1045">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="59995-1046">`getRegExMatchesByName` 方法返回匹配在具有指定 `RegExName` 元素值的清单 XML 文件中的 `ItemHasRegularExpressionMatch` 规则元素中定义的正则表达式的字符串。</span><span class="sxs-lookup"><span data-stu-id="59995-1046">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="59995-p164">如果在项目的正文属性上指定 `ItemHasRegularExpressionMatch` 规则，则正则表达式应进一步筛选正文，不应尝试返回该项目的整个正文。使用正则表达式（如 `.*`）获取项目的整个正文并不总是返回预期的结果。</span><span class="sxs-lookup"><span data-stu-id="59995-p164">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="59995-1049">Parameters</span><span class="sxs-lookup"><span data-stu-id="59995-1049">Parameters</span></span>

|<span data-ttu-id="59995-1050">名称</span><span class="sxs-lookup"><span data-stu-id="59995-1050">Name</span></span>|<span data-ttu-id="59995-1051">说明</span><span class="sxs-lookup"><span data-stu-id="59995-1051">Type</span></span>|<span data-ttu-id="59995-1052">说明</span><span class="sxs-lookup"><span data-stu-id="59995-1052">Description</span></span>|
|---|---|---|
|`name`|<span data-ttu-id="59995-1053">字符串</span><span class="sxs-lookup"><span data-stu-id="59995-1053">String</span></span>|<span data-ttu-id="59995-1054">定义筛选器匹配的 `ItemHasRegularExpressionMatch` 规则元素的名称。</span><span class="sxs-lookup"><span data-stu-id="59995-1054">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="59995-1055">要求</span><span class="sxs-lookup"><span data-stu-id="59995-1055">Requirements</span></span>

|<span data-ttu-id="59995-1056">要求</span><span class="sxs-lookup"><span data-stu-id="59995-1056">Requirement</span></span>|<span data-ttu-id="59995-1057">值</span><span class="sxs-lookup"><span data-stu-id="59995-1057">Value</span></span>|
|---|---|
|[<span data-ttu-id="59995-1058">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="59995-1058">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="59995-1059">1.0</span><span class="sxs-lookup"><span data-stu-id="59995-1059">1.0</span></span>|
|[<span data-ttu-id="59995-1060">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="59995-1060">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="59995-1061">ReadItem</span><span class="sxs-lookup"><span data-stu-id="59995-1061">ReadItem</span></span>|
|[<span data-ttu-id="59995-1062">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="59995-1062">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="59995-1063">阅读</span><span class="sxs-lookup"><span data-stu-id="59995-1063">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="59995-1064">返回：</span><span class="sxs-lookup"><span data-stu-id="59995-1064">Returns:</span></span>

<span data-ttu-id="59995-1065">一个包含匹配在清单 XML 文件中定义的正则表达式的字符串的数组。</span><span class="sxs-lookup"><span data-stu-id="59995-1065">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<dl class="param-type"><span data-ttu-id="59995-1066">

<dt>类型</dt>

</span><span class="sxs-lookup"><span data-stu-id="59995-1066">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="59995-1067">Array.< String ></span><span class="sxs-lookup"><span data-stu-id="59995-1067">Array.< String ></span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="59995-1068">示例</span><span class="sxs-lookup"><span data-stu-id="59995-1068">Example</span></span>

```javascript
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

####  <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="59995-1069">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="59995-1069">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="59995-1070">以异步方式返回邮件的主题或正文中选定的数据。</span><span class="sxs-lookup"><span data-stu-id="59995-1070">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="59995-p165">如果没有选定内容，但光标位于正文或主题中，此方法将会为所选数据返回 null。如果选定的是字段，而不是正文或主题，则此方法返回 `InvalidSelection` 错误。</span><span class="sxs-lookup"><span data-stu-id="59995-p165">If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="59995-1073">Parameters</span><span class="sxs-lookup"><span data-stu-id="59995-1073">Parameters</span></span>

|<span data-ttu-id="59995-1074">名称</span><span class="sxs-lookup"><span data-stu-id="59995-1074">Name</span></span>|<span data-ttu-id="59995-1075">类型</span><span class="sxs-lookup"><span data-stu-id="59995-1075">Type</span></span>|<span data-ttu-id="59995-1076">属性</span><span class="sxs-lookup"><span data-stu-id="59995-1076">Attributes</span></span>|<span data-ttu-id="59995-1077">说明</span><span class="sxs-lookup"><span data-stu-id="59995-1077">Description</span></span>|
|---|---|---|---|
|`coercionType`|[<span data-ttu-id="59995-1078">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="59995-1078">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="59995-p166">请求数据的格式。如果为文本，则此方法返回纯文本作为字符串，删除任何显示的 HTML 标记。如果为 HTML，则此方法返回所选文本，不论是纯文本还是 HTML。</span><span class="sxs-lookup"><span data-stu-id="59995-p166">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`|<span data-ttu-id="59995-1082">对象</span><span class="sxs-lookup"><span data-stu-id="59995-1082">Object</span></span>|<span data-ttu-id="59995-1083">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="59995-1083">&lt;optional&gt;</span></span>|<span data-ttu-id="59995-1084">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="59995-1084">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="59995-1085">对象</span><span class="sxs-lookup"><span data-stu-id="59995-1085">Object</span></span>|<span data-ttu-id="59995-1086">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="59995-1086">&lt;optional&gt;</span></span>|<span data-ttu-id="59995-1087">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="59995-1087">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="59995-1088">函数</span><span class="sxs-lookup"><span data-stu-id="59995-1088">function</span></span>||<span data-ttu-id="59995-1089">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="59995-1089">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="59995-1090">若要从回调方法访问所选数据，请调用 `asyncResult.value.data`。</span><span class="sxs-lookup"><span data-stu-id="59995-1090">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="59995-1091">若要访问选定内容的源属性，请调用 `asyncResult.value.sourceProperty`，这将为 `body` 或 `subject`。</span><span class="sxs-lookup"><span data-stu-id="59995-1091">To access the source property that the selection comes from, call `asyncResult.value.sourceProperty`, which will be either `body` or `subject`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="59995-1092">要求</span><span class="sxs-lookup"><span data-stu-id="59995-1092">Requirements</span></span>

|<span data-ttu-id="59995-1093">要求</span><span class="sxs-lookup"><span data-stu-id="59995-1093">Requirement</span></span>|<span data-ttu-id="59995-1094">值</span><span class="sxs-lookup"><span data-stu-id="59995-1094">Value</span></span>|
|---|---|
|[<span data-ttu-id="59995-1095">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="59995-1095">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="59995-1096">1.2</span><span class="sxs-lookup"><span data-stu-id="59995-1096">1.2</span></span>|
|[<span data-ttu-id="59995-1097">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="59995-1097">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="59995-1098">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="59995-1098">ReadWriteItem</span></span>|
|[<span data-ttu-id="59995-1099">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="59995-1099">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="59995-1100">撰写</span><span class="sxs-lookup"><span data-stu-id="59995-1100">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="59995-1101">返回：</span><span class="sxs-lookup"><span data-stu-id="59995-1101">Returns:</span></span>

<span data-ttu-id="59995-1102">作为字符串的所选数据的格式由 `coercionType` 确定。</span><span class="sxs-lookup"><span data-stu-id="59995-1102">The selected data as a string with format determined by `coercionType`.</span></span>

<dl class="param-type"><span data-ttu-id="59995-1103">

<dt>
类型</dt>


</span><span class="sxs-lookup"><span data-stu-id="59995-1103">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="59995-1104">String</span><span class="sxs-lookup"><span data-stu-id="59995-1104">String</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="59995-1105">示例</span><span class="sxs-lookup"><span data-stu-id="59995-1105">Example</span></span>

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

#### <a name="getselectedentities--entitiesjavascriptapioutlook17officeentities"></a><span data-ttu-id="59995-1106">getSelectedEntities() → {[Entities](/javascript/api/outlook_1_7/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="59995-1106">getSelectedEntities() → {[Entities](/javascript/api/outlook_1_7/office.entities)}</span></span>

<span data-ttu-id="59995-p168">获取在用户已选择的突出显示匹配项中找到的实体。突出显示匹配项适用于[上下文加载项](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins)。</span><span class="sxs-lookup"><span data-stu-id="59995-p168">Gets the entities found in a highlighted match a user has selected. Highlighted matches apply to [contextual add-ins](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="59995-1109">在 Outlook for iOS 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="59995-1109">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="59995-1110">Requirements</span><span class="sxs-lookup"><span data-stu-id="59995-1110">Requirements</span></span>

|<span data-ttu-id="59995-1111">要求</span><span class="sxs-lookup"><span data-stu-id="59995-1111">Requirement</span></span>|<span data-ttu-id="59995-1112">值</span><span class="sxs-lookup"><span data-stu-id="59995-1112">Value</span></span>|
|---|---|
|[<span data-ttu-id="59995-1113">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="59995-1113">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="59995-1114">1.6</span><span class="sxs-lookup"><span data-stu-id="59995-1114">1.6</span></span>|
|[<span data-ttu-id="59995-1115">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="59995-1115">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="59995-1116">ReadItem</span><span class="sxs-lookup"><span data-stu-id="59995-1116">ReadItem</span></span>|
|[<span data-ttu-id="59995-1117">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="59995-1117">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="59995-1118">阅读</span><span class="sxs-lookup"><span data-stu-id="59995-1118">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="59995-1119">返回：</span><span class="sxs-lookup"><span data-stu-id="59995-1119">Returns:</span></span>

<span data-ttu-id="59995-1120">类型：[Entities](/javascript/api/outlook_1_7/office.entities)</span><span class="sxs-lookup"><span data-stu-id="59995-1120">Type: [Entities](/javascript/api/outlook_1_7/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="59995-1121">示例</span><span class="sxs-lookup"><span data-stu-id="59995-1121">Example</span></span>

<span data-ttu-id="59995-1122">以下示例访问用户选择的突出显示匹配项中的地址实体。</span><span class="sxs-lookup"><span data-stu-id="59995-1122">The following example accesses the addresses entities in the highlighted match selected by the user.</span></span>

```javascript
var contacts = Office.context.mailbox.item.getSelectedEntities().addresses;
```

#### <a name="getselectedregexmatches--object"></a><span data-ttu-id="59995-1123">getSelectedRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="59995-1123">getSelectedRegExMatches() → {Object}</span></span>

<span data-ttu-id="59995-p169">返回突出显示匹配项中匹配在清单 XML 文件中定义的正则表达式的字符串值。突出显示匹配项适用于[上下文外接程序](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins)。</span><span class="sxs-lookup"><span data-stu-id="59995-p169">Returns string values in a highlighted match that match the regular expressions defined in the manifest XML file. Highlighted matches apply to [contextual add-ins](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="59995-1126">在 Outlook for iOS 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="59995-1126">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="59995-p170">`getSelectedRegExMatches` 方法返回匹配在清单 XML 文件中的每个 `ItemHasRegularExpressionMatch` 或 `ItemHasKnownEntity` 规则元素中定义的正则表达式的字符串。对于 `ItemHasRegularExpressionMatch` 规则，匹配字符串必须发生在该规则指定的项目的属性中。`PropertyName` 简单类型定义支持的属性。</span><span class="sxs-lookup"><span data-stu-id="59995-p170">The `getSelectedRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="59995-1130">例如，考虑一个外接程序清单具有以下 `Rule` 元素：</span><span class="sxs-lookup"><span data-stu-id="59995-1130">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="59995-1131">从 `getRegExMatches` 返回的对象应有两个属性：`fruits` 和 `veggies`。</span><span class="sxs-lookup"><span data-stu-id="59995-1131">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="59995-p171">如果在项目的正文属性上指定 `ItemHasRegularExpressionMatch` 规则，则正则表达式应进一步筛选正文，不应尝试返回该项目的整个正文。使用正则表达式（如 `.*`）获取项目的整个正文并不总是返回预期的结果。而是使用 [`Body.getAsync`](/javascript/api/outlook_1_7/office.body#getasync-coerciontype--options--callback-) 方法检索整个正文。</span><span class="sxs-lookup"><span data-stu-id="59995-p171">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook_1_7/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="59995-1135">Requirements</span><span class="sxs-lookup"><span data-stu-id="59995-1135">Requirements</span></span>

|<span data-ttu-id="59995-1136">要求</span><span class="sxs-lookup"><span data-stu-id="59995-1136">Requirement</span></span>|<span data-ttu-id="59995-1137">值</span><span class="sxs-lookup"><span data-stu-id="59995-1137">Value</span></span>|
|---|---|
|[<span data-ttu-id="59995-1138">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="59995-1138">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="59995-1139">1.6</span><span class="sxs-lookup"><span data-stu-id="59995-1139">1.6</span></span>|
|[<span data-ttu-id="59995-1140">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="59995-1140">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="59995-1141">ReadItem</span><span class="sxs-lookup"><span data-stu-id="59995-1141">ReadItem</span></span>|
|[<span data-ttu-id="59995-1142">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="59995-1142">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="59995-1143">阅读</span><span class="sxs-lookup"><span data-stu-id="59995-1143">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="59995-1144">返回：</span><span class="sxs-lookup"><span data-stu-id="59995-1144">Returns:</span></span>

<span data-ttu-id="59995-p172">一个包含匹配在清单 XML 文件中定义的正则表达式的字符串数组的对象。每个数组的名称等于匹配 `ItemHasRegularExpressionMatch` 规则的 `RegExName` 属性或匹配 `ItemHasKnownEntity` 规则的 `FilterName` 属性的相应值。</span><span class="sxs-lookup"><span data-stu-id="59995-p172">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

##### <a name="example"></a><span data-ttu-id="59995-1147">示例</span><span class="sxs-lookup"><span data-stu-id="59995-1147">Example</span></span>

<span data-ttu-id="59995-1148">以下示例显示了如何访问正则表达式规则元素 `fruits` 和 `veggies` 的匹配项的数组，这些元素在清单中指定。</span><span class="sxs-lookup"><span data-stu-id="59995-1148">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```javascript
var selectedMatches = Office.context.mailbox.item.getSelectedRegExMatches();
var fruits = selectedMatches.fruits;
var veggies = selectedMatches.veggies;
```

####  <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="59995-1149">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="59995-1149">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="59995-1150">异步加载所选项目上此外接程序的自定义属性。</span><span class="sxs-lookup"><span data-stu-id="59995-1150">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="59995-p173">自定义属性基于每个应用、每个项目存储为键/值对。此方法在回调中返回 `CustomProperties` 对象，该回调提供访问特定于当前项目和当前外接程序的自定义属性的方法。自定义属性未在项目上加密，因此这不应用作安全存储。</span><span class="sxs-lookup"><span data-stu-id="59995-p173">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="59995-1154">Parameters</span><span class="sxs-lookup"><span data-stu-id="59995-1154">Parameters</span></span>

|<span data-ttu-id="59995-1155">名称</span><span class="sxs-lookup"><span data-stu-id="59995-1155">Name</span></span>|<span data-ttu-id="59995-1156">类型</span><span class="sxs-lookup"><span data-stu-id="59995-1156">Type</span></span>|<span data-ttu-id="59995-1157">属性</span><span class="sxs-lookup"><span data-stu-id="59995-1157">Attributes</span></span>|<span data-ttu-id="59995-1158">说明</span><span class="sxs-lookup"><span data-stu-id="59995-1158">Description</span></span>|
|---|---|---|---|
|`callback`|<span data-ttu-id="59995-1159">函数</span><span class="sxs-lookup"><span data-stu-id="59995-1159">function</span></span>||<span data-ttu-id="59995-1160">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="59995-1160">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="59995-1161">自定义属性作为 `asyncResult.value` 属性中的 [`CustomProperties`](/javascript/api/outlook_1_7/office.customproperties) 对象提供。</span><span class="sxs-lookup"><span data-stu-id="59995-1161">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook_1_7/office.customproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="59995-1162">此对象可用于获取、设置以及从项目中删除自定义属性，并将自定义属性集的更改重新保存到服务器。</span><span class="sxs-lookup"><span data-stu-id="59995-1162">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`|<span data-ttu-id="59995-1163">对象</span><span class="sxs-lookup"><span data-stu-id="59995-1163">Object</span></span>|<span data-ttu-id="59995-1164">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="59995-1164">&lt;optional&gt;</span></span>|<span data-ttu-id="59995-1165">开发人员可以提供他们想要在回调函数中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="59995-1165">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="59995-1166">此对象可以通过回调函数中的 `asyncResult.asyncContext` 属性进行访问。</span><span class="sxs-lookup"><span data-stu-id="59995-1166">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="59995-1167">要求</span><span class="sxs-lookup"><span data-stu-id="59995-1167">Requirements</span></span>

|<span data-ttu-id="59995-1168">要求</span><span class="sxs-lookup"><span data-stu-id="59995-1168">Requirement</span></span>|<span data-ttu-id="59995-1169">值</span><span class="sxs-lookup"><span data-stu-id="59995-1169">Value</span></span>|
|---|---|
|[<span data-ttu-id="59995-1170">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="59995-1170">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="59995-1171">1.0</span><span class="sxs-lookup"><span data-stu-id="59995-1171">1.0</span></span>|
|[<span data-ttu-id="59995-1172">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="59995-1172">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="59995-1173">ReadItem</span><span class="sxs-lookup"><span data-stu-id="59995-1173">ReadItem</span></span>|
|[<span data-ttu-id="59995-1174">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="59995-1174">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="59995-1175">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="59995-1175">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="59995-1176">示例</span><span class="sxs-lookup"><span data-stu-id="59995-1176">Example</span></span>

<span data-ttu-id="59995-p176">以下代码示例显示了如何使用 `loadCustomPropertiesAsync` 方法异步加载特定于当前项目的自定义属性。该示例还显示了如何使用 `CustomProperties.saveAsync` 方法将这些属性重新保存到服务器。加载自定义属性后，该代码示例将使用 `CustomProperties.get` 方法读取自定义属性 `myProp`，使用 `CustomProperties.set` 方法写入自定义属性 `otherProp`，最后调用 `saveAsync` 方法保存这些自定义属性。</span><span class="sxs-lookup"><span data-stu-id="59995-p176">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

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

####  <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="59995-1180">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="59995-1180">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="59995-1181">将附件从邮件或约会中删除。</span><span class="sxs-lookup"><span data-stu-id="59995-1181">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="59995-p177">`removeAttachmentAsync` 方法删除项目中带指定标识符的附件。最佳做法是，仅当同一个邮件应用程序在同一会话中添加了一个附件时，你才应使用该附件标识符来删除该附件。在 Outlook Web App 和适用于设备的 OWA 中，附件标识符只在同一个会话中才有效。当用户关闭应用程序，或者如果用户开始在内嵌窗体中撰写，并在随后弹出的内嵌窗体中继续在单独的窗口撰写时，会话即结束。</span><span class="sxs-lookup"><span data-stu-id="59995-p177">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item. As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session. In Outlook Web App and OWA for Devices, the attachment identifier is valid only within the same session. A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="59995-1186">Parameters</span><span class="sxs-lookup"><span data-stu-id="59995-1186">Parameters</span></span>

|<span data-ttu-id="59995-1187">名称</span><span class="sxs-lookup"><span data-stu-id="59995-1187">Name</span></span>|<span data-ttu-id="59995-1188">类型</span><span class="sxs-lookup"><span data-stu-id="59995-1188">Type</span></span>|<span data-ttu-id="59995-1189">属性</span><span class="sxs-lookup"><span data-stu-id="59995-1189">Attributes</span></span>|<span data-ttu-id="59995-1190">说明</span><span class="sxs-lookup"><span data-stu-id="59995-1190">Description</span></span>|
|---|---|---|---|
|`attachmentId`|<span data-ttu-id="59995-1191">字符串</span><span class="sxs-lookup"><span data-stu-id="59995-1191">String</span></span>||<span data-ttu-id="59995-1192">要删除的附件的标识符。</span><span class="sxs-lookup"><span data-stu-id="59995-1192">The identifier of the attachment to remove.</span></span>|
|`options`|<span data-ttu-id="59995-1193">对象</span><span class="sxs-lookup"><span data-stu-id="59995-1193">Object</span></span>|<span data-ttu-id="59995-1194">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="59995-1194">&lt;optional&gt;</span></span>|<span data-ttu-id="59995-1195">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="59995-1195">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="59995-1196">对象</span><span class="sxs-lookup"><span data-stu-id="59995-1196">Object</span></span>|<span data-ttu-id="59995-1197">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="59995-1197">&lt;optional&gt;</span></span>|<span data-ttu-id="59995-1198">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="59995-1198">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="59995-1199">函数</span><span class="sxs-lookup"><span data-stu-id="59995-1199">function</span></span>|<span data-ttu-id="59995-1200">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="59995-1200">&lt;optional&gt;</span></span>|<span data-ttu-id="59995-1201">方法完成后，使用单个参数 `callback`（一个 [`asyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `AsyncResult` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="59995-1201">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="59995-1202">如果删除附件失败，`asyncResult.error` 属性将包含一个说明失败原因的错误代码。</span><span class="sxs-lookup"><span data-stu-id="59995-1202">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="59995-1203">错误</span><span class="sxs-lookup"><span data-stu-id="59995-1203">Errors</span></span>

|<span data-ttu-id="59995-1204">错误代码</span><span class="sxs-lookup"><span data-stu-id="59995-1204">Error code</span></span>|<span data-ttu-id="59995-1205">说明</span><span class="sxs-lookup"><span data-stu-id="59995-1205">Description</span></span>|
|------------|-------------|
|`InvalidAttachmentId`|<span data-ttu-id="59995-1206">附件标识符不存在。</span><span class="sxs-lookup"><span data-stu-id="59995-1206">The attachment identifier does not exist.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="59995-1207">Requirements</span><span class="sxs-lookup"><span data-stu-id="59995-1207">Requirements</span></span>

|<span data-ttu-id="59995-1208">要求</span><span class="sxs-lookup"><span data-stu-id="59995-1208">Requirement</span></span>|<span data-ttu-id="59995-1209">值</span><span class="sxs-lookup"><span data-stu-id="59995-1209">Value</span></span>|
|---|---|
|[<span data-ttu-id="59995-1210">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="59995-1210">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="59995-1211">1.1</span><span class="sxs-lookup"><span data-stu-id="59995-1211">1.1</span></span>|
|[<span data-ttu-id="59995-1212">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="59995-1212">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="59995-1213">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="59995-1213">ReadWriteItem</span></span>|
|[<span data-ttu-id="59995-1214">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="59995-1214">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="59995-1215">撰写</span><span class="sxs-lookup"><span data-stu-id="59995-1215">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="59995-1216">示例</span><span class="sxs-lookup"><span data-stu-id="59995-1216">Example</span></span>

<span data-ttu-id="59995-1217">以下代码删除包含标识符 '0' 的附件。</span><span class="sxs-lookup"><span data-stu-id="59995-1217">The following code removes an attachment with an identifier of '0'.</span></span>

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

####  <a name="removehandlerasynceventtype-options-callback"></a><span data-ttu-id="59995-1218">removeHandlerAsync(eventType, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="59995-1218">removeHandlerAsync(eventType, [options], [callback])</span></span>

<span data-ttu-id="59995-1219">删除受支持事件类型的事件处理程序。</span><span class="sxs-lookup"><span data-stu-id="59995-1219">Removes the event handlers for a supported event type.</span></span>

<span data-ttu-id="59995-1220">当前，支持的事件类型是 `Office.EventType.AppointmentTimeChanged`、`Office.EventType.RecipientsChanged` 和 `Office.EventType.RecurrenceChanged`</span><span class="sxs-lookup"><span data-stu-id="59995-1220">Currently the supported event types are `Office.EventType.AppointmentTimeChanged`, `Office.EventType.RecipientsChanged`, and `Office.EventType.RecurrenceChanged`</span></span>

##### <a name="parameters"></a><span data-ttu-id="59995-1221">Parameters</span><span class="sxs-lookup"><span data-stu-id="59995-1221">Parameters</span></span>

| <span data-ttu-id="59995-1222">名称</span><span class="sxs-lookup"><span data-stu-id="59995-1222">Name</span></span> | <span data-ttu-id="59995-1223">类型</span><span class="sxs-lookup"><span data-stu-id="59995-1223">Type</span></span> | <span data-ttu-id="59995-1224">属性</span><span class="sxs-lookup"><span data-stu-id="59995-1224">Attributes</span></span> | <span data-ttu-id="59995-1225">说明</span><span class="sxs-lookup"><span data-stu-id="59995-1225">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="59995-1226">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="59995-1226">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="59995-1227">应调用处理程序的事件。</span><span class="sxs-lookup"><span data-stu-id="59995-1227">The event that should invoke the handler.</span></span> |
| `options` | <span data-ttu-id="59995-1228">对象</span><span class="sxs-lookup"><span data-stu-id="59995-1228">Object</span></span> | <span data-ttu-id="59995-1229">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="59995-1229">&lt;optional&gt;</span></span> | <span data-ttu-id="59995-1230">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="59995-1230">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="59995-1231">对象</span><span class="sxs-lookup"><span data-stu-id="59995-1231">Object</span></span> | <span data-ttu-id="59995-1232">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="59995-1232">&lt;optional&gt;</span></span> | <span data-ttu-id="59995-1233">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="59995-1233">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="59995-1234">函数</span><span class="sxs-lookup"><span data-stu-id="59995-1234">function</span></span>| <span data-ttu-id="59995-1235">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="59995-1235">&lt;optional&gt;</span></span>|<span data-ttu-id="59995-1236">方法完成后，使用单个参数 `callback`（一个 [`asyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `AsyncResult` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="59995-1236">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="59995-1237">Requirements</span><span class="sxs-lookup"><span data-stu-id="59995-1237">Requirements</span></span>

|<span data-ttu-id="59995-1238">要求</span><span class="sxs-lookup"><span data-stu-id="59995-1238">Requirement</span></span>| <span data-ttu-id="59995-1239">值</span><span class="sxs-lookup"><span data-stu-id="59995-1239">Value</span></span>|
|---|---|
|[<span data-ttu-id="59995-1240">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="59995-1240">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="59995-1241">1.7</span><span class="sxs-lookup"><span data-stu-id="59995-1241">1.7</span></span> |
|[<span data-ttu-id="59995-1242">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="59995-1242">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="59995-1243">ReadItem</span><span class="sxs-lookup"><span data-stu-id="59995-1243">ReadItem</span></span> |
|[<span data-ttu-id="59995-1244">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="59995-1244">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="59995-1245">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="59995-1245">Compose or Read</span></span> |

##### <a name="example"></a><span data-ttu-id="59995-1246">示例</span><span class="sxs-lookup"><span data-stu-id="59995-1246">Example</span></span>

```javascript
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

####  <a name="saveasyncoptions-callback"></a><span data-ttu-id="59995-1247">saveAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="59995-1247">saveAsync([options], callback)</span></span>

<span data-ttu-id="59995-1248">异步保存项目。</span><span class="sxs-lookup"><span data-stu-id="59995-1248">Asynchronously saves an item.</span></span>

<span data-ttu-id="59995-p178">调用时，此方法将当前邮件保存为草稿，并通过回调方法返回项目 ID。在 Outlook Web App 或 Outlook 联机模式下，该项目被保存到服务器中。在 Outlook 缓存模式下，该项目被保存到本地缓存中。</span><span class="sxs-lookup"><span data-stu-id="59995-p178">When invoked, this method saves the current message as a draft and returns the item id via the callback method. In Outlook Web App or Outlook in online mode, the item is saved to the server. In Outlook in cached mode, the item is saved to the local cache.</span></span>

> [!NOTE]
> <span data-ttu-id="59995-1252">如果加载项在撰写模式下对某个项目调用 `saveAsync` 来获得 `itemId`，以便与 EWS 或 REST API 一同使用，请注意，当 Outlook 处于高速缓存模式时，可能需要一段时间项目才能真正同步到服务器。</span><span class="sxs-lookup"><span data-stu-id="59995-1252">If your add-in calls `saveAsync` on an item in compose mode in order to get an `itemId` to use with EWS or the REST API, be aware that when Outlook is in cached mode, it may take some time before the item is actually synced to the server.</span></span> <span data-ttu-id="59995-1253">在项目同步前，使用 `itemId` 将返回一个错误。</span><span class="sxs-lookup"><span data-stu-id="59995-1253">Until the item is synced, using the `itemId` will return an error.</span></span>

<span data-ttu-id="59995-p180">由于约会没有草稿状态，如果以撰写模式在约会中调用 `saveAsync`，则该项将被保存为用户日历中的正常约会。对于之前未保存过的新约会，则不会发送邀请。保存现有约会将向添加或删除的与会者发送更新。</span><span class="sxs-lookup"><span data-stu-id="59995-p180">Since appointments have no draft state, if `saveAsync` is called on an appointment in compose mode, the item will be saved as a normal appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent. Saving an existing appointment will send an update to added or removed attendees.</span></span>

> [!NOTE]
> <span data-ttu-id="59995-1257">以下客户端在撰写模式下对约会上的 `saveAsync` 具有不同的行为：</span><span class="sxs-lookup"><span data-stu-id="59995-1257">The following clients have different behavior for `saveAsync` on appointments in compose mode:</span></span>
>
> - <span data-ttu-id="59995-1258">Mac Outlook 不支持在撰写模式下对会议执行 `saveAsync` 操作。</span><span class="sxs-lookup"><span data-stu-id="59995-1258">Mac Outlook does not support `saveAsync` on a meeting in compose mode.</span></span> <span data-ttu-id="59995-1259">对 Mac Outlook 中的会议调用 `saveAsync` 将会返回一个错误。</span><span class="sxs-lookup"><span data-stu-id="59995-1259">Calling `saveAsync` on a meeting in Mac Outlook will return an error.</span></span>
> - <span data-ttu-id="59995-1260">在撰写模式下的约会上调用 `saveAsync` 时，Outlook 网页版始终发送邀请或更新。</span><span class="sxs-lookup"><span data-stu-id="59995-1260">Outlook on the web always sends an invitation or update when `saveAsync` is called on an appointment in compose mode.</span></span>

##### <a name="parameters"></a><span data-ttu-id="59995-1261">Parameters</span><span class="sxs-lookup"><span data-stu-id="59995-1261">Parameters</span></span>

|<span data-ttu-id="59995-1262">名称</span><span class="sxs-lookup"><span data-stu-id="59995-1262">Name</span></span>|<span data-ttu-id="59995-1263">类型</span><span class="sxs-lookup"><span data-stu-id="59995-1263">Type</span></span>|<span data-ttu-id="59995-1264">属性</span><span class="sxs-lookup"><span data-stu-id="59995-1264">Attributes</span></span>|<span data-ttu-id="59995-1265">说明</span><span class="sxs-lookup"><span data-stu-id="59995-1265">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="59995-1266">对象</span><span class="sxs-lookup"><span data-stu-id="59995-1266">Object</span></span>|<span data-ttu-id="59995-1267">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="59995-1267">&lt;optional&gt;</span></span>|<span data-ttu-id="59995-1268">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="59995-1268">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="59995-1269">对象</span><span class="sxs-lookup"><span data-stu-id="59995-1269">Object</span></span>|<span data-ttu-id="59995-1270">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="59995-1270">&lt;optional&gt;</span></span>|<span data-ttu-id="59995-1271">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="59995-1271">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="59995-1272">函数</span><span class="sxs-lookup"><span data-stu-id="59995-1272">function</span></span>||<span data-ttu-id="59995-1273">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="59995-1273">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="59995-1274">如果成功，该项目标识符将在 `asyncResult.value` 属性中提供。</span><span class="sxs-lookup"><span data-stu-id="59995-1274">On success, the item identifier is provided in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="59995-1275">要求</span><span class="sxs-lookup"><span data-stu-id="59995-1275">Requirements</span></span>

|<span data-ttu-id="59995-1276">要求</span><span class="sxs-lookup"><span data-stu-id="59995-1276">Requirement</span></span>|<span data-ttu-id="59995-1277">值</span><span class="sxs-lookup"><span data-stu-id="59995-1277">Value</span></span>|
|---|---|
|[<span data-ttu-id="59995-1278">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="59995-1278">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="59995-1279">1.3</span><span class="sxs-lookup"><span data-stu-id="59995-1279">1.3</span></span>|
|[<span data-ttu-id="59995-1280">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="59995-1280">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="59995-1281">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="59995-1281">ReadWriteItem</span></span>|
|[<span data-ttu-id="59995-1282">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="59995-1282">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="59995-1283">撰写</span><span class="sxs-lookup"><span data-stu-id="59995-1283">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="59995-1284">示例</span><span class="sxs-lookup"><span data-stu-id="59995-1284">Examples</span></span>

```javascript
Office.context.mailbox.item.saveAsync(
  function callback(result) {
    // Process the result.
  });
```

<span data-ttu-id="59995-p182">下面是传递给回调函数的 `result` 参数的示例。`value` 属性包含的项目的项目 ID。</span><span class="sxs-lookup"><span data-stu-id="59995-p182">The following is an example of the `result` parameter passed to the callback function. The `value` property contains the item ID of the item.</span></span>

```json
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

####  <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="59995-1287">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="59995-1287">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="59995-1288">以异步方式将数据插入到邮件的正文或主题中。</span><span class="sxs-lookup"><span data-stu-id="59995-1288">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="59995-p183">`setSelectedDataAsync` 方法将指定字符串插入到项目主题或正文的光标位置，或者，如果在编辑器中已选择文本，则该方法将替换选择的文本。如果光标不在正文或主题字段中，则返回错误。插入之后，光标会位于插入内容的末尾。</span><span class="sxs-lookup"><span data-stu-id="59995-p183">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="59995-1292">Parameters</span><span class="sxs-lookup"><span data-stu-id="59995-1292">Parameters</span></span>

|<span data-ttu-id="59995-1293">名称</span><span class="sxs-lookup"><span data-stu-id="59995-1293">Name</span></span>|<span data-ttu-id="59995-1294">类型</span><span class="sxs-lookup"><span data-stu-id="59995-1294">Type</span></span>|<span data-ttu-id="59995-1295">属性</span><span class="sxs-lookup"><span data-stu-id="59995-1295">Attributes</span></span>|<span data-ttu-id="59995-1296">说明</span><span class="sxs-lookup"><span data-stu-id="59995-1296">Description</span></span>|
|---|---|---|---|
|`data`|<span data-ttu-id="59995-1297">字符串</span><span class="sxs-lookup"><span data-stu-id="59995-1297">String</span></span>||<span data-ttu-id="59995-p184">要插入的数据。数据不得超过 1,000,000 个字符。如果传入的数据超过 1,000,000 个字符，则会引发 `ArgumentOutOfRange` 异常。</span><span class="sxs-lookup"><span data-stu-id="59995-p184">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`|<span data-ttu-id="59995-1301">Object</span><span class="sxs-lookup"><span data-stu-id="59995-1301">Object</span></span>|<span data-ttu-id="59995-1302">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="59995-1302">&lt;optional&gt;</span></span>|<span data-ttu-id="59995-1303">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="59995-1303">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="59995-1304">对象</span><span class="sxs-lookup"><span data-stu-id="59995-1304">Object</span></span>|<span data-ttu-id="59995-1305">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="59995-1305">&lt;optional&gt;</span></span>|<span data-ttu-id="59995-1306">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="59995-1306">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`|[<span data-ttu-id="59995-1307">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="59995-1307">Office.CoercionType</span></span>](office.md#coerciontype-string)|<span data-ttu-id="59995-1308">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="59995-1308">&lt;optional&gt;</span></span>|<span data-ttu-id="59995-p185">如果为 `text`，则在 Outlook Web App 和 Outlook 中应用当前样式。如果该字段是 HTML 编辑器，则仅插入文本数据，即使数据为 HTML。</span><span class="sxs-lookup"><span data-stu-id="59995-p185">If `text`, the current style is applied in Outlook Web App and Outlook. If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="59995-p186">如果 `html` 和该字段支持 HTML（主题不支持），则在 Outlook Web App 中应用当前样式，而在 Outlook 中应用默认样式。如果该字段是文本字段，则返回 `InvalidDataFormat` 错误。</span><span class="sxs-lookup"><span data-stu-id="59995-p186">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook Web App and the default style is applied in Outlook. If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="59995-1313">如果未设置 `coercionType`，则结果取决于该字段：如果该字段是 HTML，则使用 HTML；如果该字段是文本，则使用纯文本。</span><span class="sxs-lookup"><span data-stu-id="59995-1313">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`|<span data-ttu-id="59995-1314">function</span><span class="sxs-lookup"><span data-stu-id="59995-1314">function</span></span>||<span data-ttu-id="59995-1315">方法完成后，使用单个参数 `callback`（一个 [`asyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `AsyncResult` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="59995-1315">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="59995-1316">Requirements</span><span class="sxs-lookup"><span data-stu-id="59995-1316">Requirements</span></span>

|<span data-ttu-id="59995-1317">要求</span><span class="sxs-lookup"><span data-stu-id="59995-1317">Requirement</span></span>|<span data-ttu-id="59995-1318">值</span><span class="sxs-lookup"><span data-stu-id="59995-1318">Value</span></span>|
|---|---|
|[<span data-ttu-id="59995-1319">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="59995-1319">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="59995-1320">1.2</span><span class="sxs-lookup"><span data-stu-id="59995-1320">1.2</span></span>|
|[<span data-ttu-id="59995-1321">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="59995-1321">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="59995-1322">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="59995-1322">ReadWriteItem</span></span>|
|[<span data-ttu-id="59995-1323">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="59995-1323">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="59995-1324">撰写</span><span class="sxs-lookup"><span data-stu-id="59995-1324">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="59995-1325">示例</span><span class="sxs-lookup"><span data-stu-id="59995-1325">Example</span></span>

```javascript
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```
