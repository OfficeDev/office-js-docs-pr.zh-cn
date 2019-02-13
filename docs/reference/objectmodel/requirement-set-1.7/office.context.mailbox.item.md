---
title: Office.context.mailbox.item-要求设置 1.7
description: ''
ms.date: 01/30/2019
localization_priority: Normal
ms.openlocfilehash: e4bfbd9629913f775edff66f4592c220c4e5d580
ms.sourcegitcommit: a59f4e322238efa187f388a75b7709462c71e668
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 02/13/2019
ms.locfileid: "29982053"
---
# <a name="item"></a><span data-ttu-id="c115e-102">item</span><span class="sxs-lookup"><span data-stu-id="c115e-102">item</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmditem"></a><span data-ttu-id="c115e-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span><span class="sxs-lookup"><span data-stu-id="c115e-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span></span>

<span data-ttu-id="c115e-p101">`item` 命名空间用于访问当前选定的邮件、会议请求或约会。可以通过使用 [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlook17officemailboxenumsitemtype) 属性确定 `item` 的类型。</span><span class="sxs-lookup"><span data-stu-id="c115e-p101">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlook17officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="c115e-106">Requirements</span><span class="sxs-lookup"><span data-stu-id="c115e-106">Requirements</span></span>

|<span data-ttu-id="c115e-107">要求</span><span class="sxs-lookup"><span data-stu-id="c115e-107">Requirement</span></span>|<span data-ttu-id="c115e-108">值</span><span class="sxs-lookup"><span data-stu-id="c115e-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="c115e-109">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="c115e-109">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c115e-110">1.0</span><span class="sxs-lookup"><span data-stu-id="c115e-110">1.0</span></span>|
|[<span data-ttu-id="c115e-111">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="c115e-111">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c115e-112">受限</span><span class="sxs-lookup"><span data-stu-id="c115e-112">Restricted</span></span>|
|[<span data-ttu-id="c115e-113">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c115e-113">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="c115e-114">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="c115e-114">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="c115e-115">成员和方法</span><span class="sxs-lookup"><span data-stu-id="c115e-115">Members and methods</span></span>

| <span data-ttu-id="c115e-116">成员</span><span class="sxs-lookup"><span data-stu-id="c115e-116">Member</span></span> | <span data-ttu-id="c115e-117">类型</span><span class="sxs-lookup"><span data-stu-id="c115e-117">Type</span></span> |
|--------|------|
| [<span data-ttu-id="c115e-118">attachments</span><span class="sxs-lookup"><span data-stu-id="c115e-118">attachments</span></span>](#attachments-arrayattachmentdetailsjavascriptapioutlook17officeattachmentdetails) | <span data-ttu-id="c115e-119">成员</span><span class="sxs-lookup"><span data-stu-id="c115e-119">Member</span></span> |
| [<span data-ttu-id="c115e-120">bcc</span><span class="sxs-lookup"><span data-stu-id="c115e-120">bcc</span></span>](#bcc-recipientsjavascriptapioutlook17officerecipients) | <span data-ttu-id="c115e-121">成员</span><span class="sxs-lookup"><span data-stu-id="c115e-121">Member</span></span> |
| [<span data-ttu-id="c115e-122">body</span><span class="sxs-lookup"><span data-stu-id="c115e-122">body</span></span>](#body-bodyjavascriptapioutlook17officebody) | <span data-ttu-id="c115e-123">成员</span><span class="sxs-lookup"><span data-stu-id="c115e-123">Member</span></span> |
| [<span data-ttu-id="c115e-124">cc</span><span class="sxs-lookup"><span data-stu-id="c115e-124">cc</span></span>](#cc-arrayemailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsrecipientsjavascriptapioutlook17officerecipients) | <span data-ttu-id="c115e-125">成员</span><span class="sxs-lookup"><span data-stu-id="c115e-125">Member</span></span> |
| [<span data-ttu-id="c115e-126">conversationId</span><span class="sxs-lookup"><span data-stu-id="c115e-126">conversationId</span></span>](#nullable-conversationid-string) | <span data-ttu-id="c115e-127">成员</span><span class="sxs-lookup"><span data-stu-id="c115e-127">Member</span></span> |
| [<span data-ttu-id="c115e-128">dateTimeCreated</span><span class="sxs-lookup"><span data-stu-id="c115e-128">dateTimeCreated</span></span>](#datetimecreated-date) | <span data-ttu-id="c115e-129">成员</span><span class="sxs-lookup"><span data-stu-id="c115e-129">Member</span></span> |
| [<span data-ttu-id="c115e-130">dateTimeModified</span><span class="sxs-lookup"><span data-stu-id="c115e-130">dateTimeModified</span></span>](#datetimemodified-date) | <span data-ttu-id="c115e-131">成员</span><span class="sxs-lookup"><span data-stu-id="c115e-131">Member</span></span> |
| [<span data-ttu-id="c115e-132">end</span><span class="sxs-lookup"><span data-stu-id="c115e-132">end</span></span>](#end-datetimejavascriptapioutlook17officetime) | <span data-ttu-id="c115e-133">成员</span><span class="sxs-lookup"><span data-stu-id="c115e-133">Member</span></span> |
| [<span data-ttu-id="c115e-134">from</span><span class="sxs-lookup"><span data-stu-id="c115e-134">from</span></span>](#from-emailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsfromjavascriptapioutlook17officefrom) | <span data-ttu-id="c115e-135">成员</span><span class="sxs-lookup"><span data-stu-id="c115e-135">Member</span></span> |
| [<span data-ttu-id="c115e-136">internetMessageId</span><span class="sxs-lookup"><span data-stu-id="c115e-136">internetMessageId</span></span>](#internetmessageid-string) | <span data-ttu-id="c115e-137">成员</span><span class="sxs-lookup"><span data-stu-id="c115e-137">Member</span></span> |
| [<span data-ttu-id="c115e-138">itemClass</span><span class="sxs-lookup"><span data-stu-id="c115e-138">itemClass</span></span>](#itemclass-string) | <span data-ttu-id="c115e-139">成员</span><span class="sxs-lookup"><span data-stu-id="c115e-139">Member</span></span> |
| [<span data-ttu-id="c115e-140">itemId</span><span class="sxs-lookup"><span data-stu-id="c115e-140">itemId</span></span>](#nullable-itemid-string) | <span data-ttu-id="c115e-141">成员</span><span class="sxs-lookup"><span data-stu-id="c115e-141">Member</span></span> |
| [<span data-ttu-id="c115e-142">itemType</span><span class="sxs-lookup"><span data-stu-id="c115e-142">itemType</span></span>](#itemtype-officemailboxenumsitemtypejavascriptapioutlook17officemailboxenumsitemtype) | <span data-ttu-id="c115e-143">成员</span><span class="sxs-lookup"><span data-stu-id="c115e-143">Member</span></span> |
| [<span data-ttu-id="c115e-144">location</span><span class="sxs-lookup"><span data-stu-id="c115e-144">location</span></span>](#location-stringlocationjavascriptapioutlook17officelocation) | <span data-ttu-id="c115e-145">成员</span><span class="sxs-lookup"><span data-stu-id="c115e-145">Member</span></span> |
| [<span data-ttu-id="c115e-146">normalizedSubject</span><span class="sxs-lookup"><span data-stu-id="c115e-146">normalizedSubject</span></span>](#normalizedsubject-string) | <span data-ttu-id="c115e-147">成员</span><span class="sxs-lookup"><span data-stu-id="c115e-147">Member</span></span> |
| [<span data-ttu-id="c115e-148">notificationMessages</span><span class="sxs-lookup"><span data-stu-id="c115e-148">notificationMessages</span></span>](#notificationmessages-notificationmessagesjavascriptapioutlook17officenotificationmessages) | <span data-ttu-id="c115e-149">成员</span><span class="sxs-lookup"><span data-stu-id="c115e-149">Member</span></span> |
| [<span data-ttu-id="c115e-150">optionalAttendees</span><span class="sxs-lookup"><span data-stu-id="c115e-150">optionalAttendees</span></span>](#optionalattendees-arrayemailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsrecipientsjavascriptapioutlook17officerecipients) | <span data-ttu-id="c115e-151">成员</span><span class="sxs-lookup"><span data-stu-id="c115e-151">Member</span></span> |
| [<span data-ttu-id="c115e-152">organizer</span><span class="sxs-lookup"><span data-stu-id="c115e-152">organizer</span></span>](#organizer-emailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsorganizerjavascriptapioutlook17officeorganizer) | <span data-ttu-id="c115e-153">成员</span><span class="sxs-lookup"><span data-stu-id="c115e-153">Member</span></span> |
| [<span data-ttu-id="c115e-154">recurrence</span><span class="sxs-lookup"><span data-stu-id="c115e-154">recurrence</span></span>](#nullable-recurrence-recurrencejavascriptapioutlook17officerecurrence) | <span data-ttu-id="c115e-155">Member</span><span class="sxs-lookup"><span data-stu-id="c115e-155">Member</span></span> |
| [<span data-ttu-id="c115e-156">requiredAttendees</span><span class="sxs-lookup"><span data-stu-id="c115e-156">requiredAttendees</span></span>](#requiredattendees-arrayemailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsrecipientsjavascriptapioutlook17officerecipients) | <span data-ttu-id="c115e-157">成员</span><span class="sxs-lookup"><span data-stu-id="c115e-157">Member</span></span> |
| [<span data-ttu-id="c115e-158">sender</span><span class="sxs-lookup"><span data-stu-id="c115e-158">sender</span></span>](#sender-emailaddressdetailsjavascriptapioutlook17officeemailaddressdetails) | <span data-ttu-id="c115e-159">成员</span><span class="sxs-lookup"><span data-stu-id="c115e-159">Member</span></span> |
| [<span data-ttu-id="c115e-160">seriesId</span><span class="sxs-lookup"><span data-stu-id="c115e-160">seriesId</span></span>](#nullable-seriesid-string) | <span data-ttu-id="c115e-161">成员</span><span class="sxs-lookup"><span data-stu-id="c115e-161">Member</span></span> |
| [<span data-ttu-id="c115e-162">start</span><span class="sxs-lookup"><span data-stu-id="c115e-162">start</span></span>](#start-datetimejavascriptapioutlook17officetime) | <span data-ttu-id="c115e-163">成员</span><span class="sxs-lookup"><span data-stu-id="c115e-163">Member</span></span> |
| [<span data-ttu-id="c115e-164">subject</span><span class="sxs-lookup"><span data-stu-id="c115e-164">subject</span></span>](#subject-stringsubjectjavascriptapioutlook17officesubject) | <span data-ttu-id="c115e-165">Member</span><span class="sxs-lookup"><span data-stu-id="c115e-165">Member</span></span> |
| [<span data-ttu-id="c115e-166">to</span><span class="sxs-lookup"><span data-stu-id="c115e-166">to</span></span>](#to-arrayemailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsrecipientsjavascriptapioutlook17officerecipients) | <span data-ttu-id="c115e-167">成员</span><span class="sxs-lookup"><span data-stu-id="c115e-167">Member</span></span> |
| [<span data-ttu-id="c115e-168">addFileAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="c115e-168">addFileAttachmentAsync</span></span>](#addfileattachmentasyncuri-attachmentname-options-callback) | <span data-ttu-id="c115e-169">方法</span><span class="sxs-lookup"><span data-stu-id="c115e-169">Method</span></span> |
| [<span data-ttu-id="c115e-170">addHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="c115e-170">addHandlerAsync</span></span>](#addhandlerasynceventtype-handler-options-callback) | <span data-ttu-id="c115e-171">方法</span><span class="sxs-lookup"><span data-stu-id="c115e-171">Method</span></span> |
| [<span data-ttu-id="c115e-172">addItemAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="c115e-172">addItemAttachmentAsync</span></span>](#additemattachmentasyncitemid-attachmentname-options-callback) | <span data-ttu-id="c115e-173">方法</span><span class="sxs-lookup"><span data-stu-id="c115e-173">Method</span></span> |
| [<span data-ttu-id="c115e-174">close</span><span class="sxs-lookup"><span data-stu-id="c115e-174">close</span></span>](#close) | <span data-ttu-id="c115e-175">方法</span><span class="sxs-lookup"><span data-stu-id="c115e-175">Method</span></span> |
| [<span data-ttu-id="c115e-176">displayReplyAllForm</span><span class="sxs-lookup"><span data-stu-id="c115e-176">displayReplyAllForm</span></span>](#displayreplyallformformdata) | <span data-ttu-id="c115e-177">方法</span><span class="sxs-lookup"><span data-stu-id="c115e-177">Method</span></span> |
| [<span data-ttu-id="c115e-178">displayReplyForm</span><span class="sxs-lookup"><span data-stu-id="c115e-178">displayReplyForm</span></span>](#displayreplyformformdata) | <span data-ttu-id="c115e-179">方法</span><span class="sxs-lookup"><span data-stu-id="c115e-179">Method</span></span> |
| [<span data-ttu-id="c115e-180">getEntities</span><span class="sxs-lookup"><span data-stu-id="c115e-180">getEntities</span></span>](#getentities--entitiesjavascriptapioutlook17officeentities) | <span data-ttu-id="c115e-181">方法</span><span class="sxs-lookup"><span data-stu-id="c115e-181">Method</span></span> |
| [<span data-ttu-id="c115e-182">getEntitiesByType</span><span class="sxs-lookup"><span data-stu-id="c115e-182">getEntitiesByType</span></span>](#getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlook17officecontactmeetingsuggestionjavascriptapioutlook17officemeetingsuggestionphonenumberjavascriptapioutlook17officephonenumbertasksuggestionjavascriptapioutlook17officetasksuggestion) | <span data-ttu-id="c115e-183">方法</span><span class="sxs-lookup"><span data-stu-id="c115e-183">Method</span></span> |
| [<span data-ttu-id="c115e-184">getFilteredEntitiesByName</span><span class="sxs-lookup"><span data-stu-id="c115e-184">getFilteredEntitiesByName</span></span>](#getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlook17officecontactmeetingsuggestionjavascriptapioutlook17officemeetingsuggestionphonenumberjavascriptapioutlook17officephonenumbertasksuggestionjavascriptapioutlook17officetasksuggestion) | <span data-ttu-id="c115e-185">方法</span><span class="sxs-lookup"><span data-stu-id="c115e-185">Method</span></span> |
| [<span data-ttu-id="c115e-186">getRegExMatches</span><span class="sxs-lookup"><span data-stu-id="c115e-186">getRegExMatches</span></span>](#getregexmatches--object) | <span data-ttu-id="c115e-187">方法</span><span class="sxs-lookup"><span data-stu-id="c115e-187">Method</span></span> |
| [<span data-ttu-id="c115e-188">getRegExMatchesByName</span><span class="sxs-lookup"><span data-stu-id="c115e-188">getRegExMatchesByName</span></span>](#getregexmatchesbynamename--nullable-array-string-) | <span data-ttu-id="c115e-189">方法</span><span class="sxs-lookup"><span data-stu-id="c115e-189">Method</span></span> |
| [<span data-ttu-id="c115e-190">getSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="c115e-190">getSelectedDataAsync</span></span>](#getselecteddataasynccoerciontype-options-callback--string) | <span data-ttu-id="c115e-191">方法</span><span class="sxs-lookup"><span data-stu-id="c115e-191">Method</span></span> |
| [<span data-ttu-id="c115e-192">getSelectedEntities</span><span class="sxs-lookup"><span data-stu-id="c115e-192">getSelectedEntities</span></span>](#getselectedentities--entitiesjavascriptapioutlook17officeentities) | <span data-ttu-id="c115e-193">方法</span><span class="sxs-lookup"><span data-stu-id="c115e-193">Method</span></span> |
| [<span data-ttu-id="c115e-194">getSelectedRegExMatches</span><span class="sxs-lookup"><span data-stu-id="c115e-194">getSelectedRegExMatches</span></span>](#getselectedregexmatches--object) | <span data-ttu-id="c115e-195">方法</span><span class="sxs-lookup"><span data-stu-id="c115e-195">Method</span></span> |
| [<span data-ttu-id="c115e-196">loadCustomPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="c115e-196">loadCustomPropertiesAsync</span></span>](#loadcustompropertiesasynccallback-usercontext) | <span data-ttu-id="c115e-197">方法</span><span class="sxs-lookup"><span data-stu-id="c115e-197">Method</span></span> |
| [<span data-ttu-id="c115e-198">removeAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="c115e-198">removeAttachmentAsync</span></span>](#removeattachmentasyncattachmentid-options-callback) | <span data-ttu-id="c115e-199">方法</span><span class="sxs-lookup"><span data-stu-id="c115e-199">Method</span></span> |
| [<span data-ttu-id="c115e-200">removeHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="c115e-200">removeHandlerAsync</span></span>](#removehandlerasynceventtype-options-callback) | <span data-ttu-id="c115e-201">方法</span><span class="sxs-lookup"><span data-stu-id="c115e-201">Method</span></span> |
| [<span data-ttu-id="c115e-202">saveAsync</span><span class="sxs-lookup"><span data-stu-id="c115e-202">saveAsync</span></span>](#saveasyncoptions-callback) | <span data-ttu-id="c115e-203">方法</span><span class="sxs-lookup"><span data-stu-id="c115e-203">Method</span></span> |
| [<span data-ttu-id="c115e-204">setSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="c115e-204">setSelectedDataAsync</span></span>](#setselecteddataasyncdata-options-callback) | <span data-ttu-id="c115e-205">方法</span><span class="sxs-lookup"><span data-stu-id="c115e-205">Method</span></span> |

### <a name="example"></a><span data-ttu-id="c115e-206">示例</span><span class="sxs-lookup"><span data-stu-id="c115e-206">Example</span></span>

<span data-ttu-id="c115e-207">以下 JavaScript 代码示例显示了如何访问 Outlook 中当前项目的 `subject` 属性。</span><span class="sxs-lookup"><span data-stu-id="c115e-207">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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
}
```

### <a name="members"></a><span data-ttu-id="c115e-208">成员</span><span class="sxs-lookup"><span data-stu-id="c115e-208">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlook17officeattachmentdetails"></a><span data-ttu-id="c115e-209">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_7/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="c115e-209">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_7/office.attachmentdetails)></span></span>

<span data-ttu-id="c115e-p102">获取项目的附件数组。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="c115e-p102">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="c115e-212">由于潜在的安全问题，某些类型的文件会受到 Outlook 阻止，并且不会返回。</span><span class="sxs-lookup"><span data-stu-id="c115e-212">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="c115e-213">如需了解更多信息，请参阅 [Outlook 中阻止的附件](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519)。</span><span class="sxs-lookup"><span data-stu-id="c115e-213">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="c115e-214">类型：</span><span class="sxs-lookup"><span data-stu-id="c115e-214">Type:</span></span>

*   <span data-ttu-id="c115e-215">Array.<[AttachmentDetails](/javascript/api/outlook_1_7/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="c115e-215">Array.<[AttachmentDetails](/javascript/api/outlook_1_7/office.attachmentdetails)></span></span>

##### <a name="requirements"></a><span data-ttu-id="c115e-216">Requirements</span><span class="sxs-lookup"><span data-stu-id="c115e-216">Requirements</span></span>

|<span data-ttu-id="c115e-217">要求</span><span class="sxs-lookup"><span data-stu-id="c115e-217">Requirement</span></span>|<span data-ttu-id="c115e-218">值</span><span class="sxs-lookup"><span data-stu-id="c115e-218">Value</span></span>|
|---|---|
|[<span data-ttu-id="c115e-219">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="c115e-219">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c115e-220">1.0</span><span class="sxs-lookup"><span data-stu-id="c115e-220">1.0</span></span>|
|[<span data-ttu-id="c115e-221">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="c115e-221">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c115e-222">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c115e-222">ReadItem</span></span>|
|[<span data-ttu-id="c115e-223">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c115e-223">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="c115e-224">阅读</span><span class="sxs-lookup"><span data-stu-id="c115e-224">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c115e-225">示例</span><span class="sxs-lookup"><span data-stu-id="c115e-225">Example</span></span>

<span data-ttu-id="c115e-226">以下代码使用当前项目上所有附件的详细信息构成 HTML 字符串。</span><span class="sxs-lookup"><span data-stu-id="c115e-226">The following code builds an HTML string with details of all attachments on the current item.</span></span>

```js
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

####  <a name="bcc-recipientsjavascriptapioutlook17officerecipients"></a><span data-ttu-id="c115e-227">bcc :[Recipients](/javascript/api/outlook_1_7/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="c115e-227">bcc :[Recipients](/javascript/api/outlook_1_7/office.recipients)</span></span>

<span data-ttu-id="c115e-228">获取对象，该对象提供用于获取或更新邮件的密件抄送 (Bcc) 行上的收件人的方法。</span><span class="sxs-lookup"><span data-stu-id="c115e-228">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="c115e-229">仅限撰写模式。</span><span class="sxs-lookup"><span data-stu-id="c115e-229">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="c115e-230">类型:</span><span class="sxs-lookup"><span data-stu-id="c115e-230">Type:</span></span>

*   [<span data-ttu-id="c115e-231">收件人</span><span class="sxs-lookup"><span data-stu-id="c115e-231">Recipients</span></span>](/javascript/api/outlook_1_7/office.recipients)

##### <a name="requirements"></a><span data-ttu-id="c115e-232">Requirements</span><span class="sxs-lookup"><span data-stu-id="c115e-232">Requirements</span></span>

|<span data-ttu-id="c115e-233">要求</span><span class="sxs-lookup"><span data-stu-id="c115e-233">Requirement</span></span>|<span data-ttu-id="c115e-234">值</span><span class="sxs-lookup"><span data-stu-id="c115e-234">Value</span></span>|
|---|---|
|[<span data-ttu-id="c115e-235">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="c115e-235">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c115e-236">1.1</span><span class="sxs-lookup"><span data-stu-id="c115e-236">1.1</span></span>|
|[<span data-ttu-id="c115e-237">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="c115e-237">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c115e-238">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c115e-238">ReadItem</span></span>|
|[<span data-ttu-id="c115e-239">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c115e-239">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="c115e-240">撰写</span><span class="sxs-lookup"><span data-stu-id="c115e-240">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="c115e-241">示例</span><span class="sxs-lookup"><span data-stu-id="c115e-241">Example</span></span>

```js
Office.context.mailbox.item.bcc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.bcc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.bcc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfBccRecipients = asyncResult.value;
}
```

####  <a name="body-bodyjavascriptapioutlook17officebody"></a><span data-ttu-id="c115e-242">body :[Body](/javascript/api/outlook_1_7/office.body)</span><span class="sxs-lookup"><span data-stu-id="c115e-242">body :[Body](/javascript/api/outlook_1_7/office.body)</span></span>

<span data-ttu-id="c115e-243">获取一个提供用于处理项目正文的方法的对象。</span><span class="sxs-lookup"><span data-stu-id="c115e-243">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="c115e-244">类型：</span><span class="sxs-lookup"><span data-stu-id="c115e-244">Type:</span></span>

*   [<span data-ttu-id="c115e-245">Body</span><span class="sxs-lookup"><span data-stu-id="c115e-245">Body</span></span>](/javascript/api/outlook_1_7/office.body)

##### <a name="requirements"></a><span data-ttu-id="c115e-246">要求</span><span class="sxs-lookup"><span data-stu-id="c115e-246">Requirements</span></span>

|<span data-ttu-id="c115e-247">要求</span><span class="sxs-lookup"><span data-stu-id="c115e-247">Requirement</span></span>|<span data-ttu-id="c115e-248">值</span><span class="sxs-lookup"><span data-stu-id="c115e-248">Value</span></span>|
|---|---|
|[<span data-ttu-id="c115e-249">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="c115e-249">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c115e-250">1.1</span><span class="sxs-lookup"><span data-stu-id="c115e-250">1.1</span></span>|
|[<span data-ttu-id="c115e-251">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="c115e-251">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c115e-252">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c115e-252">ReadItem</span></span>|
|[<span data-ttu-id="c115e-253">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c115e-253">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="c115e-254">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="c115e-254">Compose or read</span></span>|

####  <a name="cc-arrayemailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsrecipientsjavascriptapioutlook17officerecipients"></a><span data-ttu-id="c115e-255">cc :Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_7/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="c115e-255">cc :Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_7/office.recipients)</span></span>

<span data-ttu-id="c115e-256">提供对邮件的抄送 (Cc) 收件人的访问权限。</span><span class="sxs-lookup"><span data-stu-id="c115e-256">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="c115e-257">对象的类型和访问级别取决于当前项目的模式。</span><span class="sxs-lookup"><span data-stu-id="c115e-257">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="c115e-258">阅读模式</span><span class="sxs-lookup"><span data-stu-id="c115e-258">Read mode</span></span>

<span data-ttu-id="c115e-p106">`cc` 属性返回包含邮件的**抄送**行上所列的每个收件人的 `EmailAddressDetails` 对象的数组。集合上限为 100 个成员。</span><span class="sxs-lookup"><span data-stu-id="c115e-p106">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="c115e-261">撰写模式</span><span class="sxs-lookup"><span data-stu-id="c115e-261">Compose mode</span></span>

<span data-ttu-id="c115e-262">`cc` 属性返回一个 `Recipients` 对象，该对象提供用于获取或更新邮件的“**抄送**”行上收件人的方法。</span><span class="sxs-lookup"><span data-stu-id="c115e-262">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="c115e-263">类型：</span><span class="sxs-lookup"><span data-stu-id="c115e-263">Type:</span></span>

*   <span data-ttu-id="c115e-264">Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_7/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="c115e-264">Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_7/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="c115e-265">要求</span><span class="sxs-lookup"><span data-stu-id="c115e-265">Requirements</span></span>

|<span data-ttu-id="c115e-266">要求</span><span class="sxs-lookup"><span data-stu-id="c115e-266">Requirement</span></span>|<span data-ttu-id="c115e-267">值</span><span class="sxs-lookup"><span data-stu-id="c115e-267">Value</span></span>|
|---|---|
|[<span data-ttu-id="c115e-268">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="c115e-268">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c115e-269">1.0</span><span class="sxs-lookup"><span data-stu-id="c115e-269">1.0</span></span>|
|[<span data-ttu-id="c115e-270">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="c115e-270">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c115e-271">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c115e-271">ReadItem</span></span>|
|[<span data-ttu-id="c115e-272">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c115e-272">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="c115e-273">Compose 或 Read</span><span class="sxs-lookup"><span data-stu-id="c115e-273">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="c115e-274">示例</span><span class="sxs-lookup"><span data-stu-id="c115e-274">Example</span></span>

```js
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

####  <a name="nullable-conversationid-string"></a><span data-ttu-id="c115e-275">(nullable) conversationId :String</span><span class="sxs-lookup"><span data-stu-id="c115e-275">(nullable) conversationId :String</span></span>

<span data-ttu-id="c115e-276">获取包含特定消息的电子邮件会话的标识符。</span><span class="sxs-lookup"><span data-stu-id="c115e-276">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="c115e-p107">如果在阅读窗体或撰写窗体的回复中激活邮件应用程序，则此属性可以获得一个整数值。如果用户随后更改了回复邮件的主题（若发送回复），则该邮件的对话 ID 将改变且之前获取的值将不适用。</span><span class="sxs-lookup"><span data-stu-id="c115e-p107">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="c115e-p108">对于撰写窗体的新项目，此属性获得一个 null 值。如果用户设置一个主题并保存该项目，`conversationId` 属性将返回一个值。</span><span class="sxs-lookup"><span data-stu-id="c115e-p108">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="c115e-281">类型：</span><span class="sxs-lookup"><span data-stu-id="c115e-281">Type:</span></span>

*   <span data-ttu-id="c115e-282">String</span><span class="sxs-lookup"><span data-stu-id="c115e-282">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="c115e-283">要求</span><span class="sxs-lookup"><span data-stu-id="c115e-283">Requirements</span></span>

|<span data-ttu-id="c115e-284">要求</span><span class="sxs-lookup"><span data-stu-id="c115e-284">Requirement</span></span>|<span data-ttu-id="c115e-285">值</span><span class="sxs-lookup"><span data-stu-id="c115e-285">Value</span></span>|
|---|---|
|[<span data-ttu-id="c115e-286">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="c115e-286">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c115e-287">1.0</span><span class="sxs-lookup"><span data-stu-id="c115e-287">1.0</span></span>|
|[<span data-ttu-id="c115e-288">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="c115e-288">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c115e-289">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c115e-289">ReadItem</span></span>|
|[<span data-ttu-id="c115e-290">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c115e-290">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="c115e-291">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="c115e-291">Compose or read</span></span>|

#### <a name="datetimecreated-date"></a><span data-ttu-id="c115e-292">dateTimeCreated :Date</span><span class="sxs-lookup"><span data-stu-id="c115e-292">dateTimeCreated :Date</span></span>

<span data-ttu-id="c115e-p109">获取项目创建的日期和时间。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="c115e-p109">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="c115e-295">类型:</span><span class="sxs-lookup"><span data-stu-id="c115e-295">Type:</span></span>

*   <span data-ttu-id="c115e-296">日期</span><span class="sxs-lookup"><span data-stu-id="c115e-296">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="c115e-297">Requirements</span><span class="sxs-lookup"><span data-stu-id="c115e-297">Requirements</span></span>

|<span data-ttu-id="c115e-298">要求</span><span class="sxs-lookup"><span data-stu-id="c115e-298">Requirement</span></span>|<span data-ttu-id="c115e-299">值</span><span class="sxs-lookup"><span data-stu-id="c115e-299">Value</span></span>|
|---|---|
|[<span data-ttu-id="c115e-300">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="c115e-300">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c115e-301">1.0</span><span class="sxs-lookup"><span data-stu-id="c115e-301">1.0</span></span>|
|[<span data-ttu-id="c115e-302">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="c115e-302">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c115e-303">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c115e-303">ReadItem</span></span>|
|[<span data-ttu-id="c115e-304">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c115e-304">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="c115e-305">阅读</span><span class="sxs-lookup"><span data-stu-id="c115e-305">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c115e-306">示例</span><span class="sxs-lookup"><span data-stu-id="c115e-306">Example</span></span>

```js
var created = Office.context.mailbox.item.dateTimeCreated;
```

#### <a name="datetimemodified-date"></a><span data-ttu-id="c115e-307">dateTimeModified :Date</span><span class="sxs-lookup"><span data-stu-id="c115e-307">dateTimeModified :Date</span></span>

<span data-ttu-id="c115e-p110">获取项目最近一次修改的日期和时间。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="c115e-p110">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="c115e-310">在 Outlook for iOS 或 Outlook for Android 中不支持此成员。</span><span class="sxs-lookup"><span data-stu-id="c115e-310">This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="type"></a><span data-ttu-id="c115e-311">类型:</span><span class="sxs-lookup"><span data-stu-id="c115e-311">Type:</span></span>

*   <span data-ttu-id="c115e-312">日期</span><span class="sxs-lookup"><span data-stu-id="c115e-312">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="c115e-313">Requirements</span><span class="sxs-lookup"><span data-stu-id="c115e-313">Requirements</span></span>

|<span data-ttu-id="c115e-314">要求</span><span class="sxs-lookup"><span data-stu-id="c115e-314">Requirement</span></span>|<span data-ttu-id="c115e-315">值</span><span class="sxs-lookup"><span data-stu-id="c115e-315">Value</span></span>|
|---|---|
|[<span data-ttu-id="c115e-316">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="c115e-316">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c115e-317">1.0</span><span class="sxs-lookup"><span data-stu-id="c115e-317">1.0</span></span>|
|[<span data-ttu-id="c115e-318">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="c115e-318">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c115e-319">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c115e-319">ReadItem</span></span>|
|[<span data-ttu-id="c115e-320">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c115e-320">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="c115e-321">阅读</span><span class="sxs-lookup"><span data-stu-id="c115e-321">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c115e-322">示例</span><span class="sxs-lookup"><span data-stu-id="c115e-322">Example</span></span>

```js
var modified = Office.context.mailbox.item.dateTimeModified;
```

####  <a name="end-datetimejavascriptapioutlook17officetime"></a><span data-ttu-id="c115e-323">end :Date|[Time](/javascript/api/outlook_1_7/office.time)</span><span class="sxs-lookup"><span data-stu-id="c115e-323">end :Date|[Time](/javascript/api/outlook_1_7/office.time)</span></span>

<span data-ttu-id="c115e-324">获取或设置约会结束的日期和时间。</span><span class="sxs-lookup"><span data-stu-id="c115e-324">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="c115e-p111">将 `end` 属性表示为协调世界时 (UTC) 的日期和时间值。可使用 [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook17officelocalclienttime) 方法将 end 属性值转换为客户端的本地日期和时间。</span><span class="sxs-lookup"><span data-stu-id="c115e-p111">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook17officelocalclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="c115e-327">阅读模式</span><span class="sxs-lookup"><span data-stu-id="c115e-327">Read mode</span></span>

<span data-ttu-id="c115e-328">`end` 属性返回 `Date` 对象。</span><span class="sxs-lookup"><span data-stu-id="c115e-328">The `end` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="c115e-329">撰写模式</span><span class="sxs-lookup"><span data-stu-id="c115e-329">Compose mode</span></span>

<span data-ttu-id="c115e-330">`end` 属性返回 `Time` 对象。</span><span class="sxs-lookup"><span data-stu-id="c115e-330">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="c115e-331">使用 [`Time.setAsync`](/javascript/api/outlook_1_7/office.time#setasync-datetime--options--callback-) 方法设置结束时间时，应使用 [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) 方法将客户端的本地时间转换为服务器的 UTC。</span><span class="sxs-lookup"><span data-stu-id="c115e-331">When you use the [`Time.setAsync`](/javascript/api/outlook_1_7/office.time#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="c115e-332">类型：</span><span class="sxs-lookup"><span data-stu-id="c115e-332">Type:</span></span>

*   <span data-ttu-id="c115e-333">Date | [Time](/javascript/api/outlook_1_7/office.time)</span><span class="sxs-lookup"><span data-stu-id="c115e-333">Date | [Time](/javascript/api/outlook_1_7/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="c115e-334">Requirements</span><span class="sxs-lookup"><span data-stu-id="c115e-334">Requirements</span></span>

|<span data-ttu-id="c115e-335">要求</span><span class="sxs-lookup"><span data-stu-id="c115e-335">Requirement</span></span>|<span data-ttu-id="c115e-336">值</span><span class="sxs-lookup"><span data-stu-id="c115e-336">Value</span></span>|
|---|---|
|[<span data-ttu-id="c115e-337">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="c115e-337">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c115e-338">1.0</span><span class="sxs-lookup"><span data-stu-id="c115e-338">1.0</span></span>|
|[<span data-ttu-id="c115e-339">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="c115e-339">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c115e-340">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c115e-340">ReadItem</span></span>|
|[<span data-ttu-id="c115e-341">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c115e-341">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="c115e-342">Compose 或 Read</span><span class="sxs-lookup"><span data-stu-id="c115e-342">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="c115e-343">示例</span><span class="sxs-lookup"><span data-stu-id="c115e-343">Example</span></span>

<span data-ttu-id="c115e-344">以下示例通过使用 `Time` 对象的 [`setAsync`](/javascript/api/outlook_1_7/office.time#setasync-datetime--options--callback-) 方法，设置撰写模式下约会的结束时间。</span><span class="sxs-lookup"><span data-stu-id="c115e-344">The following example sets the end time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_7/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

```js
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

#### <a name="from-emailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsfromjavascriptapioutlook17officefrom"></a><span data-ttu-id="c115e-345">from :[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)|[From](/javascript/api/outlook_1_7/office.from)</span><span class="sxs-lookup"><span data-stu-id="c115e-345">from :[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)|[From](/javascript/api/outlook_1_7/office.from)</span></span>

<span data-ttu-id="c115e-346">获取邮件发件人的电子邮件地址。</span><span class="sxs-lookup"><span data-stu-id="c115e-346">Gets the email address of the sender of a message.</span></span>

<span data-ttu-id="c115e-p112">`from` 和 [`sender`](#sender-emailaddressdetailsjavascriptapioutlook17officeemailaddressdetails) 属性表示同一个人，邮件由代理人发送的除外。在此情况下，`from` 属性表示代理程序，而 sender 属性表示代理人。</span><span class="sxs-lookup"><span data-stu-id="c115e-p112">The `from` and [`sender`](#sender-emailaddressdetailsjavascriptapioutlook17officeemailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="c115e-349">`from` 属性中 `EmailAddressDetails` 对象的 `recipientType` 属性为 `undefined`。</span><span class="sxs-lookup"><span data-stu-id="c115e-349">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="c115e-350">阅读模式</span><span class="sxs-lookup"><span data-stu-id="c115e-350">Read mode</span></span>

<span data-ttu-id="c115e-351">`from` 属性返回一个 `EmailAddressDetails` 对象。</span><span class="sxs-lookup"><span data-stu-id="c115e-351">The `from` property returns an `EmailAddressDetails` object.</span></span>

```js
var from = Office.context.mailbox.item.from;
```

##### <a name="compose-mode"></a><span data-ttu-id="c115e-352">撰写模式</span><span class="sxs-lookup"><span data-stu-id="c115e-352">Compose mode</span></span>

<span data-ttu-id="c115e-353">`from` 属性返回一个 `From` 对象，该对象提供从值中进行获取的方法。</span><span class="sxs-lookup"><span data-stu-id="c115e-353">The `from` property returns a `From` object that provides a method to get the from value.</span></span>

```js
Office.context.mailbox.item.from.getAsync(callback);

function callback(asyncResult) {
  var from = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="c115e-354">类型：</span><span class="sxs-lookup"><span data-stu-id="c115e-354">Type:</span></span>

*   <span data-ttu-id="c115e-355">[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails) | [From](/javascript/api/outlook_1_7/office.from)</span><span class="sxs-lookup"><span data-stu-id="c115e-355">[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails) | [From](/javascript/api/outlook_1_7/office.from)</span></span>

##### <a name="requirements"></a><span data-ttu-id="c115e-356">要求</span><span class="sxs-lookup"><span data-stu-id="c115e-356">Requirements</span></span>

|<span data-ttu-id="c115e-357">要求</span><span class="sxs-lookup"><span data-stu-id="c115e-357">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="c115e-358">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="c115e-358">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c115e-359">1.0</span><span class="sxs-lookup"><span data-stu-id="c115e-359">1.0</span></span>|<span data-ttu-id="c115e-360">1.7</span><span class="sxs-lookup"><span data-stu-id="c115e-360">1.7</span></span>|
|[<span data-ttu-id="c115e-361">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="c115e-361">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c115e-362">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c115e-362">ReadItem</span></span>|<span data-ttu-id="c115e-363">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="c115e-363">ReadWriteItem</span></span>|
|[<span data-ttu-id="c115e-364">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c115e-364">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="c115e-365">阅读</span><span class="sxs-lookup"><span data-stu-id="c115e-365">Read</span></span>|<span data-ttu-id="c115e-366">撰写</span><span class="sxs-lookup"><span data-stu-id="c115e-366">Compose</span></span>|

#### <a name="internetmessageid-string"></a><span data-ttu-id="c115e-367">internetMessageId :String</span><span class="sxs-lookup"><span data-stu-id="c115e-367">internetMessageId :String</span></span>

<span data-ttu-id="c115e-p113">获取电子邮件的 Internet 消息标识符。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="c115e-p113">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="c115e-370">类型：</span><span class="sxs-lookup"><span data-stu-id="c115e-370">Type:</span></span>

*   <span data-ttu-id="c115e-371">String</span><span class="sxs-lookup"><span data-stu-id="c115e-371">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="c115e-372">要求</span><span class="sxs-lookup"><span data-stu-id="c115e-372">Requirements</span></span>

|<span data-ttu-id="c115e-373">要求</span><span class="sxs-lookup"><span data-stu-id="c115e-373">Requirement</span></span>|<span data-ttu-id="c115e-374">值</span><span class="sxs-lookup"><span data-stu-id="c115e-374">Value</span></span>|
|---|---|
|[<span data-ttu-id="c115e-375">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="c115e-375">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c115e-376">1.0</span><span class="sxs-lookup"><span data-stu-id="c115e-376">1.0</span></span>|
|[<span data-ttu-id="c115e-377">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="c115e-377">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c115e-378">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c115e-378">ReadItem</span></span>|
|[<span data-ttu-id="c115e-379">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c115e-379">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="c115e-380">阅读</span><span class="sxs-lookup"><span data-stu-id="c115e-380">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c115e-381">示例</span><span class="sxs-lookup"><span data-stu-id="c115e-381">Example</span></span>

```js
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

#### <a name="itemclass-string"></a><span data-ttu-id="c115e-382">itemClass :String</span><span class="sxs-lookup"><span data-stu-id="c115e-382">itemClass :String</span></span>

<span data-ttu-id="c115e-p114">获取选定项目的 Exchange Web 服务项目类。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="c115e-p114">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="c115e-p115">`itemClass` 属性指定所选项目的邮件类别。以下是邮件或约会项目的默认邮件类别。</span><span class="sxs-lookup"><span data-stu-id="c115e-p115">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

|<span data-ttu-id="c115e-387">类型</span><span class="sxs-lookup"><span data-stu-id="c115e-387">Type</span></span>|<span data-ttu-id="c115e-388">说明</span><span class="sxs-lookup"><span data-stu-id="c115e-388">Description</span></span>|<span data-ttu-id="c115e-389">项目类</span><span class="sxs-lookup"><span data-stu-id="c115e-389">item class</span></span>|
|---|---|---|
|<span data-ttu-id="c115e-390">约会项目</span><span class="sxs-lookup"><span data-stu-id="c115e-390">Appointment items</span></span>|<span data-ttu-id="c115e-391">这些是项目类 `IPM.Appointment` 或 `IPM.Appointment.Occurrence` 的日历项目。</span><span class="sxs-lookup"><span data-stu-id="c115e-391">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurrence`.</span></span>|`IPM.Appointment`<br />`IPM.Appointment.Occurrence`|
|<span data-ttu-id="c115e-392">邮件项目</span><span class="sxs-lookup"><span data-stu-id="c115e-392">Message items</span></span>|<span data-ttu-id="c115e-393">这些项目包括具有默认邮件类别 `IPM.Note` 的电子邮件，以及使用 `IPM.Schedule.Meeting` 作为基础邮件类别的会议请求、响应和取消。</span><span class="sxs-lookup"><span data-stu-id="c115e-393">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span>|`IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled`|

<span data-ttu-id="c115e-394">你可以创建用于扩展默认邮件类别的自定义邮件类别，例如，自定义约会邮件类别 `IPM.Appointment.Contoso`。</span><span class="sxs-lookup"><span data-stu-id="c115e-394">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="c115e-395">类型：</span><span class="sxs-lookup"><span data-stu-id="c115e-395">Type:</span></span>

*   <span data-ttu-id="c115e-396">String</span><span class="sxs-lookup"><span data-stu-id="c115e-396">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="c115e-397">要求</span><span class="sxs-lookup"><span data-stu-id="c115e-397">Requirements</span></span>

|<span data-ttu-id="c115e-398">要求</span><span class="sxs-lookup"><span data-stu-id="c115e-398">Requirement</span></span>|<span data-ttu-id="c115e-399">值</span><span class="sxs-lookup"><span data-stu-id="c115e-399">Value</span></span>|
|---|---|
|[<span data-ttu-id="c115e-400">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="c115e-400">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c115e-401">1.0</span><span class="sxs-lookup"><span data-stu-id="c115e-401">1.0</span></span>|
|[<span data-ttu-id="c115e-402">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="c115e-402">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c115e-403">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c115e-403">ReadItem</span></span>|
|[<span data-ttu-id="c115e-404">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c115e-404">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="c115e-405">阅读</span><span class="sxs-lookup"><span data-stu-id="c115e-405">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c115e-406">示例</span><span class="sxs-lookup"><span data-stu-id="c115e-406">Example</span></span>

```js
var itemClass = Office.context.mailbox.item.itemClass;
```

#### <a name="nullable-itemid-string"></a><span data-ttu-id="c115e-407">(nullable) itemId :String</span><span class="sxs-lookup"><span data-stu-id="c115e-407">(nullable) itemId :String</span></span>

<span data-ttu-id="c115e-p116">获取当前项目的 Exchange Web 服务项目标识符。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="c115e-p116">Gets the Exchange Web Services item identifier for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="c115e-410">`itemId` 属性返回的标识符与 Exchange Web 服务项目标识符相同。</span><span class="sxs-lookup"><span data-stu-id="c115e-410">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="c115e-411">`itemId` 属性与 Outlook 条目 ID 或 Outlook REST API 使用的 ID 不同。</span><span class="sxs-lookup"><span data-stu-id="c115e-411">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="c115e-412">使用此值进行 REST API 调用前，应使用 [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string) 对它进行转换。</span><span class="sxs-lookup"><span data-stu-id="c115e-412">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="c115e-413">有关详细信息，请参阅[从 Outlook 加载项使用 Outlook REST API](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id)。</span><span class="sxs-lookup"><span data-stu-id="c115e-413">For more details, see [Use the Outlook REST APIs from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

<span data-ttu-id="c115e-p118">`itemId` 属性在撰写模式下不可用。如果需要项目标识符，[`saveAsync`](#saveasyncoptions-callback) 方法可用于将项目保存到存储，这将在回调函数的 [`AsyncResult.value`](/javascript/api/office/office.asyncresult) 参数中返回项目标识符。</span><span class="sxs-lookup"><span data-stu-id="c115e-p118">The `itemId` property is not available in compose mode. If an item identifier is required, the [`saveAsync`](#saveasyncoptions-callback) method can be used to save the item to the store, which will return the item identifier in the [`AsyncResult.value`](/javascript/api/office/office.asyncresult) parameter in the callback function.</span></span>

##### <a name="type"></a><span data-ttu-id="c115e-416">类型：</span><span class="sxs-lookup"><span data-stu-id="c115e-416">Type:</span></span>

*   <span data-ttu-id="c115e-417">String</span><span class="sxs-lookup"><span data-stu-id="c115e-417">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="c115e-418">要求</span><span class="sxs-lookup"><span data-stu-id="c115e-418">Requirements</span></span>

|<span data-ttu-id="c115e-419">要求</span><span class="sxs-lookup"><span data-stu-id="c115e-419">Requirement</span></span>|<span data-ttu-id="c115e-420">值</span><span class="sxs-lookup"><span data-stu-id="c115e-420">Value</span></span>|
|---|---|
|[<span data-ttu-id="c115e-421">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="c115e-421">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c115e-422">1.0</span><span class="sxs-lookup"><span data-stu-id="c115e-422">1.0</span></span>|
|[<span data-ttu-id="c115e-423">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="c115e-423">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c115e-424">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c115e-424">ReadItem</span></span>|
|[<span data-ttu-id="c115e-425">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c115e-425">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="c115e-426">阅读</span><span class="sxs-lookup"><span data-stu-id="c115e-426">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c115e-427">示例</span><span class="sxs-lookup"><span data-stu-id="c115e-427">Example</span></span>

<span data-ttu-id="c115e-p119">以下代码检查项目标识符是否存在。如果 `itemId` 属性返回 `null` 或 `undefined`，则将项目保存到存储，并从异步结果中获取项目标识符。</span><span class="sxs-lookup"><span data-stu-id="c115e-p119">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

```js
var itemId = Office.context.mailbox.item.itemId;
if (itemId === null || itemId == undefined) {
  Office.context.mailbox.item.saveAsync(function(result){
    itemId = result.value;
  });
}
```

####  <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlook17officemailboxenumsitemtype"></a><span data-ttu-id="c115e-430">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_7/office.mailboxenums.itemtype)</span><span class="sxs-lookup"><span data-stu-id="c115e-430">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_7/office.mailboxenums.itemtype)</span></span>

<span data-ttu-id="c115e-431">获取实例表示的项的类型。</span><span class="sxs-lookup"><span data-stu-id="c115e-431">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="c115e-432">`itemType` 属性返回其中一个 `ItemType` 枚举值，指示 `item` 对象实例是邮件还是约会。</span><span class="sxs-lookup"><span data-stu-id="c115e-432">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="c115e-433">类型：</span><span class="sxs-lookup"><span data-stu-id="c115e-433">Type:</span></span>

*   [<span data-ttu-id="c115e-434">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="c115e-434">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook_1_7/office.mailboxenums.itemtype)

##### <a name="requirements"></a><span data-ttu-id="c115e-435">Requirements</span><span class="sxs-lookup"><span data-stu-id="c115e-435">Requirements</span></span>

|<span data-ttu-id="c115e-436">要求</span><span class="sxs-lookup"><span data-stu-id="c115e-436">Requirement</span></span>|<span data-ttu-id="c115e-437">值</span><span class="sxs-lookup"><span data-stu-id="c115e-437">Value</span></span>|
|---|---|
|[<span data-ttu-id="c115e-438">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="c115e-438">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c115e-439">1.0</span><span class="sxs-lookup"><span data-stu-id="c115e-439">1.0</span></span>|
|[<span data-ttu-id="c115e-440">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="c115e-440">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c115e-441">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c115e-441">ReadItem</span></span>|
|[<span data-ttu-id="c115e-442">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c115e-442">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="c115e-443">Compose 或 Read</span><span class="sxs-lookup"><span data-stu-id="c115e-443">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="c115e-444">示例</span><span class="sxs-lookup"><span data-stu-id="c115e-444">Example</span></span>

```js
if (Office.context.mailbox.item.itemType == Office.MailboxEnums.ItemType.Message)
  // do something
else
  // do something else
```

####  <a name="location-stringlocationjavascriptapioutlook17officelocation"></a><span data-ttu-id="c115e-445">location :String|[Location](/javascript/api/outlook_1_7/office.location)</span><span class="sxs-lookup"><span data-stu-id="c115e-445">location :String|[Location](/javascript/api/outlook_1_7/office.location)</span></span>

<span data-ttu-id="c115e-446">获取或设置约会的位置。</span><span class="sxs-lookup"><span data-stu-id="c115e-446">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="c115e-447">阅读模式</span><span class="sxs-lookup"><span data-stu-id="c115e-447">Read mode</span></span>

<span data-ttu-id="c115e-448">`location` 属性返回一个包含约会位置的字符串。</span><span class="sxs-lookup"><span data-stu-id="c115e-448">The `location` property returns a string that contains the location of the appointment.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="c115e-449">撰写模式</span><span class="sxs-lookup"><span data-stu-id="c115e-449">Compose mode</span></span>

<span data-ttu-id="c115e-450">`location` 属性返回一个 `Location` 对象，该对象提供用于获取和设置约会位置的方法。</span><span class="sxs-lookup"><span data-stu-id="c115e-450">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="c115e-451">类型：</span><span class="sxs-lookup"><span data-stu-id="c115e-451">Type:</span></span>

*   <span data-ttu-id="c115e-452">String | [Location](/javascript/api/outlook_1_7/office.location)</span><span class="sxs-lookup"><span data-stu-id="c115e-452">String | [Location](/javascript/api/outlook_1_7/office.location)</span></span>

##### <a name="requirements"></a><span data-ttu-id="c115e-453">Requirements</span><span class="sxs-lookup"><span data-stu-id="c115e-453">Requirements</span></span>

|<span data-ttu-id="c115e-454">要求</span><span class="sxs-lookup"><span data-stu-id="c115e-454">Requirement</span></span>|<span data-ttu-id="c115e-455">值</span><span class="sxs-lookup"><span data-stu-id="c115e-455">Value</span></span>|
|---|---|
|[<span data-ttu-id="c115e-456">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="c115e-456">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c115e-457">1.0</span><span class="sxs-lookup"><span data-stu-id="c115e-457">1.0</span></span>|
|[<span data-ttu-id="c115e-458">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="c115e-458">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c115e-459">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c115e-459">ReadItem</span></span>|
|[<span data-ttu-id="c115e-460">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c115e-460">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="c115e-461">Compose 或 Read</span><span class="sxs-lookup"><span data-stu-id="c115e-461">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="c115e-462">示例</span><span class="sxs-lookup"><span data-stu-id="c115e-462">Example</span></span>

```js
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

#### <a name="normalizedsubject-string"></a><span data-ttu-id="c115e-463">normalizedSubject :String</span><span class="sxs-lookup"><span data-stu-id="c115e-463">normalizedSubject :String</span></span>

<span data-ttu-id="c115e-p120">获取删除了所有前缀（包括 `RE:` 和 `FWD:`）的项目主题。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="c115e-p120">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="c115e-p121">normalizedSubject 属性获取包含由电子邮件程序添加的任何标准前缀（如 `RE:` 和 `FW:`）的项目主题。若要获取包含完整前缀的项目主题，请使用 [`subject`](#subject-stringsubjectjavascriptapioutlook17officesubject) 属性。</span><span class="sxs-lookup"><span data-stu-id="c115e-p121">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubjectjavascriptapioutlook17officesubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="c115e-468">类型:</span><span class="sxs-lookup"><span data-stu-id="c115e-468">Type:</span></span>

*   <span data-ttu-id="c115e-469">String</span><span class="sxs-lookup"><span data-stu-id="c115e-469">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="c115e-470">要求</span><span class="sxs-lookup"><span data-stu-id="c115e-470">Requirements</span></span>

|<span data-ttu-id="c115e-471">要求</span><span class="sxs-lookup"><span data-stu-id="c115e-471">Requirement</span></span>|<span data-ttu-id="c115e-472">值</span><span class="sxs-lookup"><span data-stu-id="c115e-472">Value</span></span>|
|---|---|
|[<span data-ttu-id="c115e-473">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="c115e-473">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c115e-474">1.0</span><span class="sxs-lookup"><span data-stu-id="c115e-474">1.0</span></span>|
|[<span data-ttu-id="c115e-475">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="c115e-475">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c115e-476">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c115e-476">ReadItem</span></span>|
|[<span data-ttu-id="c115e-477">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c115e-477">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="c115e-478">阅读</span><span class="sxs-lookup"><span data-stu-id="c115e-478">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c115e-479">示例</span><span class="sxs-lookup"><span data-stu-id="c115e-479">Example</span></span>

```js
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
```

####  <a name="notificationmessages-notificationmessagesjavascriptapioutlook17officenotificationmessages"></a><span data-ttu-id="c115e-480">notificationMessages :[NotificationMessages](/javascript/api/outlook_1_7/office.notificationmessages)</span><span class="sxs-lookup"><span data-stu-id="c115e-480">notificationMessages :[NotificationMessages](/javascript/api/outlook_1_7/office.notificationmessages)</span></span>

<span data-ttu-id="c115e-481">获取项目的通知邮件。</span><span class="sxs-lookup"><span data-stu-id="c115e-481">Gets the notification messages for an item.</span></span>

##### <a name="type"></a><span data-ttu-id="c115e-482">类型：</span><span class="sxs-lookup"><span data-stu-id="c115e-482">Type:</span></span>

*   [<span data-ttu-id="c115e-483">NotificationMessages</span><span class="sxs-lookup"><span data-stu-id="c115e-483">NotificationMessages</span></span>](/javascript/api/outlook_1_7/office.notificationmessages)

##### <a name="requirements"></a><span data-ttu-id="c115e-484">要求</span><span class="sxs-lookup"><span data-stu-id="c115e-484">Requirements</span></span>

|<span data-ttu-id="c115e-485">要求</span><span class="sxs-lookup"><span data-stu-id="c115e-485">Requirement</span></span>|<span data-ttu-id="c115e-486">值</span><span class="sxs-lookup"><span data-stu-id="c115e-486">Value</span></span>|
|---|---|
|[<span data-ttu-id="c115e-487">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="c115e-487">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c115e-488">1.3</span><span class="sxs-lookup"><span data-stu-id="c115e-488">1.3</span></span>|
|[<span data-ttu-id="c115e-489">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="c115e-489">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c115e-490">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c115e-490">ReadItem</span></span>|
|[<span data-ttu-id="c115e-491">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c115e-491">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="c115e-492">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="c115e-492">Compose or read</span></span>|

####  <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsrecipientsjavascriptapioutlook17officerecipients"></a><span data-ttu-id="c115e-493">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_7/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="c115e-493">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_7/office.recipients)</span></span>

<span data-ttu-id="c115e-494">提供对事件的可选与会者的访问权限。</span><span class="sxs-lookup"><span data-stu-id="c115e-494">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="c115e-495">对象的类型和访问级别取决于当前项目的模式。</span><span class="sxs-lookup"><span data-stu-id="c115e-495">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="c115e-496">阅读模式</span><span class="sxs-lookup"><span data-stu-id="c115e-496">Read mode</span></span>

<span data-ttu-id="c115e-497">`optionalAttendees` 属性返回一个数组，其中包含每个可选与会者的 `EmailAddressDetails` 对象。</span><span class="sxs-lookup"><span data-stu-id="c115e-497">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="c115e-498">撰写模式</span><span class="sxs-lookup"><span data-stu-id="c115e-498">Compose mode</span></span>

<span data-ttu-id="c115e-499">`optionalAttendees` 属性返回一个 `Recipients` 对象，该对象提供用于获取或更新可选与会者的方法。</span><span class="sxs-lookup"><span data-stu-id="c115e-499">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="c115e-500">类型：</span><span class="sxs-lookup"><span data-stu-id="c115e-500">Type:</span></span>

*   <span data-ttu-id="c115e-501">Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_7/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="c115e-501">Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_7/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="c115e-502">Requirements</span><span class="sxs-lookup"><span data-stu-id="c115e-502">Requirements</span></span>

|<span data-ttu-id="c115e-503">要求</span><span class="sxs-lookup"><span data-stu-id="c115e-503">Requirement</span></span>|<span data-ttu-id="c115e-504">值</span><span class="sxs-lookup"><span data-stu-id="c115e-504">Value</span></span>|
|---|---|
|[<span data-ttu-id="c115e-505">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="c115e-505">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c115e-506">1.0</span><span class="sxs-lookup"><span data-stu-id="c115e-506">1.0</span></span>|
|[<span data-ttu-id="c115e-507">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="c115e-507">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c115e-508">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c115e-508">ReadItem</span></span>|
|[<span data-ttu-id="c115e-509">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c115e-509">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="c115e-510">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="c115e-510">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="c115e-511">示例</span><span class="sxs-lookup"><span data-stu-id="c115e-511">Example</span></span>

```js
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

#### <a name="organizer-emailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsorganizerjavascriptapioutlook17officeorganizer"></a><span data-ttu-id="c115e-512">organizer :[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)|[Organizer](/javascript/api/outlook_1_7/office.organizer)</span><span class="sxs-lookup"><span data-stu-id="c115e-512">organizer :[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)|[Organizer](/javascript/api/outlook_1_7/office.organizer)</span></span>

<span data-ttu-id="c115e-513">获取指定会议的组织者的电子邮件地址。</span><span class="sxs-lookup"><span data-stu-id="c115e-513">Gets the email address of the organizer for a specified meeting.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="c115e-514">阅读模式</span><span class="sxs-lookup"><span data-stu-id="c115e-514">Read mode</span></span>

<span data-ttu-id="c115e-515">`organizer` 属性返回 [EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails) 对象，它表示会议组织者。</span><span class="sxs-lookup"><span data-stu-id="c115e-515">The `organizer` property returns an [EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails) object that represents the meeting organizer.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="c115e-516">撰写模式</span><span class="sxs-lookup"><span data-stu-id="c115e-516">Compose mode</span></span>

<span data-ttu-id="c115e-517">`organizer` 属性返回 [Organizer](/javascript/api/outlook_1_7/office.organizer) 对象，该对象提供获取组织者值的方法。</span><span class="sxs-lookup"><span data-stu-id="c115e-517">The `organizer` property returns an [Organizer](/javascript/api/outlook_1_7/office.organizer) object that provides a method to get the organizer value.</span></span>

##### <a name="type"></a><span data-ttu-id="c115e-518">类型：</span><span class="sxs-lookup"><span data-stu-id="c115e-518">Type:</span></span>

*   <span data-ttu-id="c115e-519">[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails) | [Organizer](/javascript/api/outlook_1_7/office.organizer)</span><span class="sxs-lookup"><span data-stu-id="c115e-519">[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails) | [Organizer](/javascript/api/outlook_1_7/office.organizer)</span></span>

##### <a name="requirements"></a><span data-ttu-id="c115e-520">要求</span><span class="sxs-lookup"><span data-stu-id="c115e-520">Requirements</span></span>

|<span data-ttu-id="c115e-521">要求</span><span class="sxs-lookup"><span data-stu-id="c115e-521">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="c115e-522">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="c115e-522">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c115e-523">1.0</span><span class="sxs-lookup"><span data-stu-id="c115e-523">1.0</span></span>|<span data-ttu-id="c115e-524">1.7</span><span class="sxs-lookup"><span data-stu-id="c115e-524">1.7</span></span>|
|[<span data-ttu-id="c115e-525">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="c115e-525">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c115e-526">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c115e-526">ReadItem</span></span>|<span data-ttu-id="c115e-527">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="c115e-527">ReadWriteItem</span></span>|
|[<span data-ttu-id="c115e-528">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c115e-528">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="c115e-529">阅读</span><span class="sxs-lookup"><span data-stu-id="c115e-529">Read</span></span>|<span data-ttu-id="c115e-530">撰写</span><span class="sxs-lookup"><span data-stu-id="c115e-530">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="c115e-531">示例</span><span class="sxs-lookup"><span data-stu-id="c115e-531">Example</span></span>

```js
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
```

#### <a name="nullable-recurrence-recurrencejavascriptapioutlook17officerecurrence"></a><span data-ttu-id="c115e-532">(nullable) recurrence :[Recurrence](/javascript/api/outlook_1_7/office.recurrence)</span><span class="sxs-lookup"><span data-stu-id="c115e-532">(nullable) recurrence :[Recurrence](/javascript/api/outlook_1_7/office.recurrence)</span></span>

<span data-ttu-id="c115e-533">获取或设置约会的定期模式。</span><span class="sxs-lookup"><span data-stu-id="c115e-533">Gets or sets the recurrence pattern of an appointment.</span></span> <span data-ttu-id="c115e-534">获取或设置会议请求的定期模式。</span><span class="sxs-lookup"><span data-stu-id="c115e-534">Gets the recurrence pattern of a meeting request.</span></span> <span data-ttu-id="c115e-535">阅读撰写约会项目的模式。</span><span class="sxs-lookup"><span data-stu-id="c115e-535">Read and compose modes for appointment items.</span></span> <span data-ttu-id="c115e-536">会议请求项目的阅读模式。</span><span class="sxs-lookup"><span data-stu-id="c115e-536">Read mode for meeting request items.</span></span>

<span data-ttu-id="c115e-537">如果项目是一个系列或系列中的一个实例，则 `recurrence` 属性将返回定期约会的 [recurrence](/javascript/api/outlook_1_7/office.recurrence) 对象或会议请求。</span><span class="sxs-lookup"><span data-stu-id="c115e-537">The `recurrence` property returns a [recurrence](/javascript/api/outlook_1_7/office.recurrence) object for recurring appointments or meetings requests if an item is a series or an instance in a series.</span></span> <span data-ttu-id="c115e-538">针对单个约会和单个约会的会议请求返回 `null`。</span><span class="sxs-lookup"><span data-stu-id="c115e-538">`null` is returned for single appointments and meeting requests of single appointments.</span></span> <span data-ttu-id="c115e-539">针对非会议请求的邮件返回 `undefined`。</span><span class="sxs-lookup"><span data-stu-id="c115e-539">`undefined` is returned for messages that are not meeting requests.</span></span>

> <span data-ttu-id="c115e-540">注意：会议请求的 `itemClass` 值为 IPM.Schedule.Meeting.Request。</span><span class="sxs-lookup"><span data-stu-id="c115e-540">Note: Meeting requests have an `itemClass` value of IPM.Schedule.Meeting.Request.</span></span>

> <span data-ttu-id="c115e-541">注意：如果 recurrence 对象为 `null`，则这表示对象是单个约会或单个约会的会议请求，而不是系列的一部分。</span><span class="sxs-lookup"><span data-stu-id="c115e-541">Note: If the recurrence object is `null`, this indicates that the object is a single appointment or a meeting request of a single appointment and NOT a part of a series.</span></span>

##### <a name="type"></a><span data-ttu-id="c115e-542">类型：</span><span class="sxs-lookup"><span data-stu-id="c115e-542">Type:</span></span>

* [<span data-ttu-id="c115e-543">Recurrence</span><span class="sxs-lookup"><span data-stu-id="c115e-543">Recurrence</span></span>](/javascript/api/outlook_1_7/office.recurrence)

|<span data-ttu-id="c115e-544">要求</span><span class="sxs-lookup"><span data-stu-id="c115e-544">Requirement</span></span>|<span data-ttu-id="c115e-545">值</span><span class="sxs-lookup"><span data-stu-id="c115e-545">Value</span></span>|
|---|---|
|[<span data-ttu-id="c115e-546">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="c115e-546">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c115e-547">1.7</span><span class="sxs-lookup"><span data-stu-id="c115e-547">1.7</span></span>|
|[<span data-ttu-id="c115e-548">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="c115e-548">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c115e-549">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c115e-549">ReadItem</span></span>|
|[<span data-ttu-id="c115e-550">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c115e-550">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="c115e-551">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="c115e-551">Compose or read</span></span>|

####  <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsrecipientsjavascriptapioutlook17officerecipients"></a><span data-ttu-id="c115e-552">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_7/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="c115e-552">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_7/office.recipients)</span></span>

<span data-ttu-id="c115e-553">提供对事件的必需与会者的访问权限。</span><span class="sxs-lookup"><span data-stu-id="c115e-553">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="c115e-554">对象的类型和访问级别取决于当前项目的模式。</span><span class="sxs-lookup"><span data-stu-id="c115e-554">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="c115e-555">读取模式</span><span class="sxs-lookup"><span data-stu-id="c115e-555">Read mode</span></span>

<span data-ttu-id="c115e-556">`requiredAttendees` 属性返回一个数组，其中包含每个必需与会者的 `EmailAddressDetails` 对象。</span><span class="sxs-lookup"><span data-stu-id="c115e-556">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="c115e-557">撰写模式</span><span class="sxs-lookup"><span data-stu-id="c115e-557">Compose mode</span></span>

<span data-ttu-id="c115e-558">`requiredAttendees` 属性返回一个 `Recipients` 对象，该对象提供用于获取或更新必需与会者的方法。</span><span class="sxs-lookup"><span data-stu-id="c115e-558">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="c115e-559">类型：</span><span class="sxs-lookup"><span data-stu-id="c115e-559">Type:</span></span>

*   <span data-ttu-id="c115e-560">Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_7/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="c115e-560">Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_7/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="c115e-561">Requirements</span><span class="sxs-lookup"><span data-stu-id="c115e-561">Requirements</span></span>

|<span data-ttu-id="c115e-562">要求</span><span class="sxs-lookup"><span data-stu-id="c115e-562">Requirement</span></span>|<span data-ttu-id="c115e-563">值</span><span class="sxs-lookup"><span data-stu-id="c115e-563">Value</span></span>|
|---|---|
|[<span data-ttu-id="c115e-564">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="c115e-564">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c115e-565">1.0</span><span class="sxs-lookup"><span data-stu-id="c115e-565">1.0</span></span>|
|[<span data-ttu-id="c115e-566">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="c115e-566">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c115e-567">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c115e-567">ReadItem</span></span>|
|[<span data-ttu-id="c115e-568">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c115e-568">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="c115e-569">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="c115e-569">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="c115e-570">示例</span><span class="sxs-lookup"><span data-stu-id="c115e-570">Example</span></span>

```js
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
}
```

#### <a name="sender-emailaddressdetailsjavascriptapioutlook17officeemailaddressdetails"></a><span data-ttu-id="c115e-571">sender :[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="c115e-571">sender :[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)</span></span>

<span data-ttu-id="c115e-p126">获取电子邮件发件人的电子邮件地址。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="c115e-p126">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="c115e-p127">[`from`](#from-emailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsfromjavascriptapioutlook17officefrom) 和 `sender` 属性表示同一个人，邮件由代理人发送的除外。在此情况下，`from` 属性表示代理程序，而 sender 属性表示代理人。</span><span class="sxs-lookup"><span data-stu-id="c115e-p127">The [`from`](#from-emailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsfromjavascriptapioutlook17officefrom) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="c115e-576">`sender` 属性中 `EmailAddressDetails` 对象的 `recipientType` 属性为 `undefined`。</span><span class="sxs-lookup"><span data-stu-id="c115e-576">The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="c115e-577">类型:</span><span class="sxs-lookup"><span data-stu-id="c115e-577">Type:</span></span>

*   [<span data-ttu-id="c115e-578">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="c115e-578">EmailAddressDetails</span></span>](/javascript/api/outlook_1_7/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="c115e-579">Requirements</span><span class="sxs-lookup"><span data-stu-id="c115e-579">Requirements</span></span>

|<span data-ttu-id="c115e-580">要求</span><span class="sxs-lookup"><span data-stu-id="c115e-580">Requirement</span></span>|<span data-ttu-id="c115e-581">值</span><span class="sxs-lookup"><span data-stu-id="c115e-581">Value</span></span>|
|---|---|
|[<span data-ttu-id="c115e-582">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="c115e-582">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c115e-583">1.0</span><span class="sxs-lookup"><span data-stu-id="c115e-583">1.0</span></span>|
|[<span data-ttu-id="c115e-584">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="c115e-584">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c115e-585">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c115e-585">ReadItem</span></span>|
|[<span data-ttu-id="c115e-586">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c115e-586">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="c115e-587">阅读</span><span class="sxs-lookup"><span data-stu-id="c115e-587">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c115e-588">示例</span><span class="sxs-lookup"><span data-stu-id="c115e-588">Example</span></span>

```js
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
```

#### <a name="nullable-seriesid-string"></a><span data-ttu-id="c115e-589">(nullable) seriesId :String</span><span class="sxs-lookup"><span data-stu-id="c115e-589">(nullable) seriesId :String</span></span>

<span data-ttu-id="c115e-590">获取实例所属的系列的 ID。</span><span class="sxs-lookup"><span data-stu-id="c115e-590">Gets the id of the series that an instance belongs to.</span></span>

<span data-ttu-id="c115e-591">在 OWA 和 Outlook 中，`seriesId` 返回此项目所属的父（系列）项目的 Exchange Web 服务 (EWS) ID。</span><span class="sxs-lookup"><span data-stu-id="c115e-591">In OWA and Outlook, the `seriesId` returns the Exchange Web Services (EWS) ID of the parent (series) item that this item belongs to.</span></span> <span data-ttu-id="c115e-592">但是，在 iOS 和 Android 中，`seriesId` 返回父项目的其余部分 ID。</span><span class="sxs-lookup"><span data-stu-id="c115e-592">However, in iOS and Android, the `seriesId` returns the REST ID of the parent item.</span></span>

> [!NOTE]
> <span data-ttu-id="c115e-593">`seriesId` 属性返回的标识符与 Exchange Web 服务项目标识符相同。</span><span class="sxs-lookup"><span data-stu-id="c115e-593">The identifier returned by the `seriesId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="c115e-594">`seriesId` 属性与 Outlook REST API 使用的 Outlook ID 不同。</span><span class="sxs-lookup"><span data-stu-id="c115e-594">The `seriesId` property is not identical to the Outlook IDs used by the Outlook REST API.</span></span> <span data-ttu-id="c115e-595">使用此值进行 REST API 调用前，应使用 [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string) 对它进行转换。</span><span class="sxs-lookup"><span data-stu-id="c115e-595">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="c115e-596">有关详细信息，请参阅[从 Outlook 加载项使用 Outlook REST API](https://docs.microsoft.com/outlook/add-ins/use-rest-api)。</span><span class="sxs-lookup"><span data-stu-id="c115e-596">For more details, see [Use the Outlook REST APIs from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/use-rest-api).</span></span>

<span data-ttu-id="c115e-597">`seriesId` 属性对于没有父项目（如单个约会、系列项目或会议请求）的项目返回 `null`，对于非会议请求的任何其他项目，返回 `undefined`。</span><span class="sxs-lookup"><span data-stu-id="c115e-597">The `seriesId` property returns `null` for items that do not have parent items such as single appointments, series items, or meeting requests and returns `undefined` for any other items that are not meeting requests.</span></span>

##### <a name="type"></a><span data-ttu-id="c115e-598">类型：</span><span class="sxs-lookup"><span data-stu-id="c115e-598">Type:</span></span>

* <span data-ttu-id="c115e-599">String</span><span class="sxs-lookup"><span data-stu-id="c115e-599">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="c115e-600">要求</span><span class="sxs-lookup"><span data-stu-id="c115e-600">Requirements</span></span>

|<span data-ttu-id="c115e-601">要求</span><span class="sxs-lookup"><span data-stu-id="c115e-601">Requirement</span></span>|<span data-ttu-id="c115e-602">值</span><span class="sxs-lookup"><span data-stu-id="c115e-602">Value</span></span>|
|---|---|
|[<span data-ttu-id="c115e-603">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="c115e-603">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c115e-604">1.7</span><span class="sxs-lookup"><span data-stu-id="c115e-604">1.7</span></span>|
|[<span data-ttu-id="c115e-605">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="c115e-605">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c115e-606">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c115e-606">ReadItem</span></span>|
|[<span data-ttu-id="c115e-607">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c115e-607">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="c115e-608">Compose 或 Read</span><span class="sxs-lookup"><span data-stu-id="c115e-608">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="c115e-609">示例</span><span class="sxs-lookup"><span data-stu-id="c115e-609">Example</span></span>

```js
var seriesId = Office.context.mailbox.item.seriesId;
var isSeries = (seriesId == null);
```

####  <a name="start-datetimejavascriptapioutlook17officetime"></a><span data-ttu-id="c115e-610">start :Date|[Time](/javascript/api/outlook_1_7/office.time)</span><span class="sxs-lookup"><span data-stu-id="c115e-610">start :Date|[Time](/javascript/api/outlook_1_7/office.time)</span></span>

<span data-ttu-id="c115e-611">获取或设置约会开始的日期和时间。</span><span class="sxs-lookup"><span data-stu-id="c115e-611">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="c115e-p130">将 `start` 属性表示为协调世界时 (UTC) 的日期和时间值。可使用 [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook17officelocalclienttime) 方法将该值转换为客户端的本地日期和时间。</span><span class="sxs-lookup"><span data-stu-id="c115e-p130">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook17officelocalclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="c115e-614">阅读模式</span><span class="sxs-lookup"><span data-stu-id="c115e-614">Read mode</span></span>

<span data-ttu-id="c115e-615">`start` 属性返回 `Date` 对象。</span><span class="sxs-lookup"><span data-stu-id="c115e-615">The `start` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="c115e-616">撰写模式</span><span class="sxs-lookup"><span data-stu-id="c115e-616">Compose mode</span></span>

<span data-ttu-id="c115e-617">`start` 属性返回 `Time` 对象。</span><span class="sxs-lookup"><span data-stu-id="c115e-617">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="c115e-618">使用 [`Time.setAsync`](/javascript/api/outlook_1_7/office.time#setasync-datetime--options--callback-) 方法设置开始时间时，应使用 [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) 方法将客户端的本地时间转换为服务器的 UTC。</span><span class="sxs-lookup"><span data-stu-id="c115e-618">When you use the [`Time.setAsync`](/javascript/api/outlook_1_7/office.time#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="c115e-619">类型：</span><span class="sxs-lookup"><span data-stu-id="c115e-619">Type:</span></span>

*   <span data-ttu-id="c115e-620">Date | [Time](/javascript/api/outlook_1_7/office.time)</span><span class="sxs-lookup"><span data-stu-id="c115e-620">Date | [Time](/javascript/api/outlook_1_7/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="c115e-621">Requirements</span><span class="sxs-lookup"><span data-stu-id="c115e-621">Requirements</span></span>

|<span data-ttu-id="c115e-622">要求</span><span class="sxs-lookup"><span data-stu-id="c115e-622">Requirement</span></span>|<span data-ttu-id="c115e-623">值</span><span class="sxs-lookup"><span data-stu-id="c115e-623">Value</span></span>|
|---|---|
|[<span data-ttu-id="c115e-624">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="c115e-624">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c115e-625">1.0</span><span class="sxs-lookup"><span data-stu-id="c115e-625">1.0</span></span>|
|[<span data-ttu-id="c115e-626">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="c115e-626">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c115e-627">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c115e-627">ReadItem</span></span>|
|[<span data-ttu-id="c115e-628">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c115e-628">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="c115e-629">Compose 或 Read</span><span class="sxs-lookup"><span data-stu-id="c115e-629">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="c115e-630">示例</span><span class="sxs-lookup"><span data-stu-id="c115e-630">Example</span></span>

<span data-ttu-id="c115e-631">以下示例通过使用 `Time` 对象的 [`setAsync`](/javascript/api/outlook_1_7/office.time#setasync-datetime--options--callback-) 方法，设置撰写模式下约会的开始时间。</span><span class="sxs-lookup"><span data-stu-id="c115e-631">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_7/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

```js
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

####  <a name="subject-stringsubjectjavascriptapioutlook17officesubject"></a><span data-ttu-id="c115e-632">subject :String|[Subject](/javascript/api/outlook_1_7/office.subject)</span><span class="sxs-lookup"><span data-stu-id="c115e-632">subject :String|[Subject](/javascript/api/outlook_1_7/office.subject)</span></span>

<span data-ttu-id="c115e-633">获取或设置显示在项目的主题字段中的说明。</span><span class="sxs-lookup"><span data-stu-id="c115e-633">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="c115e-634">`subject` 属性获取或设置由电子邮件服务器发送项目时的整个主题。</span><span class="sxs-lookup"><span data-stu-id="c115e-634">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="c115e-635">阅读模式</span><span class="sxs-lookup"><span data-stu-id="c115e-635">Read mode</span></span>

<span data-ttu-id="c115e-p131">`subject` 属性返回一个字符串。使用 [`normalizedSubject`](#normalizedsubject-string) 属性获取不带任何前导前缀（如 `RE:` 和 `FW:`）的主题。</span><span class="sxs-lookup"><span data-stu-id="c115e-p131">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

```js
var subject = Office.context.mailbox.item.subject;
```

##### <a name="compose-mode"></a><span data-ttu-id="c115e-638">撰写模式</span><span class="sxs-lookup"><span data-stu-id="c115e-638">Compose mode</span></span>

<span data-ttu-id="c115e-639">`subject` 属性返回一个 `Subject` 对象，该对象提供用于获取和设置主题的方法。</span><span class="sxs-lookup"><span data-stu-id="c115e-639">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```js
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="c115e-640">类型：</span><span class="sxs-lookup"><span data-stu-id="c115e-640">Type:</span></span>

*   <span data-ttu-id="c115e-641">String | [Subject](/javascript/api/outlook_1_7/office.subject)</span><span class="sxs-lookup"><span data-stu-id="c115e-641">String | [Subject](/javascript/api/outlook_1_7/office.subject)</span></span>

##### <a name="requirements"></a><span data-ttu-id="c115e-642">要求</span><span class="sxs-lookup"><span data-stu-id="c115e-642">Requirements</span></span>

|<span data-ttu-id="c115e-643">要求</span><span class="sxs-lookup"><span data-stu-id="c115e-643">Requirement</span></span>|<span data-ttu-id="c115e-644">值</span><span class="sxs-lookup"><span data-stu-id="c115e-644">Value</span></span>|
|---|---|
|[<span data-ttu-id="c115e-645">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="c115e-645">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c115e-646">1.0</span><span class="sxs-lookup"><span data-stu-id="c115e-646">1.0</span></span>|
|[<span data-ttu-id="c115e-647">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="c115e-647">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c115e-648">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c115e-648">ReadItem</span></span>|
|[<span data-ttu-id="c115e-649">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c115e-649">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="c115e-650">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="c115e-650">Compose or read</span></span>|

####  <a name="to-arrayemailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsrecipientsjavascriptapioutlook17officerecipients"></a><span data-ttu-id="c115e-651">to :Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_7/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="c115e-651">to :Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_7/office.recipients)</span></span>

<span data-ttu-id="c115e-652">提供对邮件的“**收件人**”行上的收件人的访问权限。</span><span class="sxs-lookup"><span data-stu-id="c115e-652">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="c115e-653">对象的类型和访问级别取决于当前项目的模式。</span><span class="sxs-lookup"><span data-stu-id="c115e-653">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="c115e-654">阅读模式</span><span class="sxs-lookup"><span data-stu-id="c115e-654">Read mode</span></span>

<span data-ttu-id="c115e-p133">`to` 属性返回包含邮件的**收件人**行上所列的每个收件人的 `EmailAddressDetails` 对象的数组。集合上限为 100 个成员。</span><span class="sxs-lookup"><span data-stu-id="c115e-p133">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="c115e-657">撰写模式</span><span class="sxs-lookup"><span data-stu-id="c115e-657">Compose mode</span></span>

<span data-ttu-id="c115e-658">`to` 属性返回一个 `Recipients` 对象，该对象提供用于获取或更新邮件的“**收件人**”行上收件人的方法。</span><span class="sxs-lookup"><span data-stu-id="c115e-658">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="c115e-659">类型：</span><span class="sxs-lookup"><span data-stu-id="c115e-659">Type:</span></span>

*   <span data-ttu-id="c115e-660">Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_7/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="c115e-660">Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_7/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="c115e-661">Requirements</span><span class="sxs-lookup"><span data-stu-id="c115e-661">Requirements</span></span>

|<span data-ttu-id="c115e-662">要求</span><span class="sxs-lookup"><span data-stu-id="c115e-662">Requirement</span></span>|<span data-ttu-id="c115e-663">值</span><span class="sxs-lookup"><span data-stu-id="c115e-663">Value</span></span>|
|---|---|
|[<span data-ttu-id="c115e-664">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="c115e-664">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c115e-665">1.0</span><span class="sxs-lookup"><span data-stu-id="c115e-665">1.0</span></span>|
|[<span data-ttu-id="c115e-666">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="c115e-666">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c115e-667">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c115e-667">ReadItem</span></span>|
|[<span data-ttu-id="c115e-668">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c115e-668">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="c115e-669">Compose 或 Read</span><span class="sxs-lookup"><span data-stu-id="c115e-669">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="c115e-670">示例</span><span class="sxs-lookup"><span data-stu-id="c115e-670">Example</span></span>

```js
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

### <a name="methods"></a><span data-ttu-id="c115e-671">方法</span><span class="sxs-lookup"><span data-stu-id="c115e-671">Methods</span></span>

####  <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="c115e-672">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="c115e-672">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="c115e-673">将文件作为附件添加到邮件或约会。</span><span class="sxs-lookup"><span data-stu-id="c115e-673">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="c115e-674">`addFileAttachmentAsync` 方法在指定的 URI 上载文件并将其附加到撰写窗体中的项目。</span><span class="sxs-lookup"><span data-stu-id="c115e-674">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="c115e-675">你随后可以将该标识符与 [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) 方法一同使用，以删除同一个会话中的附件。</span><span class="sxs-lookup"><span data-stu-id="c115e-675">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c115e-676">参数：</span><span class="sxs-lookup"><span data-stu-id="c115e-676">Parameters:</span></span>
|<span data-ttu-id="c115e-677">名称</span><span class="sxs-lookup"><span data-stu-id="c115e-677">Name</span></span>|<span data-ttu-id="c115e-678">类型</span><span class="sxs-lookup"><span data-stu-id="c115e-678">Type</span></span>|<span data-ttu-id="c115e-679">属性</span><span class="sxs-lookup"><span data-stu-id="c115e-679">Attributes</span></span>|<span data-ttu-id="c115e-680">说明</span><span class="sxs-lookup"><span data-stu-id="c115e-680">Description</span></span>|
|---|---|---|---|
|`uri`|<span data-ttu-id="c115e-681">String</span><span class="sxs-lookup"><span data-stu-id="c115e-681">String</span></span>||<span data-ttu-id="c115e-p134">提供附加到邮件或约会的文件的位置的 URI。最大长度为 2048 个字符。</span><span class="sxs-lookup"><span data-stu-id="c115e-p134">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`|<span data-ttu-id="c115e-684">字符串</span><span class="sxs-lookup"><span data-stu-id="c115e-684">String</span></span>||<span data-ttu-id="c115e-p135">在附件上载过程中显示的附件名称。最大长度为 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="c115e-p135">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="c115e-687">Object</span><span class="sxs-lookup"><span data-stu-id="c115e-687">Object</span></span>|<span data-ttu-id="c115e-688">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="c115e-688">&lt;optional&gt;</span></span>|<span data-ttu-id="c115e-689">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="c115e-689">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="c115e-690">对象</span><span class="sxs-lookup"><span data-stu-id="c115e-690">Object</span></span>|<span data-ttu-id="c115e-691">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="c115e-691">&lt;optional&gt;</span></span>|<span data-ttu-id="c115e-692">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="c115e-692">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.isInline`|<span data-ttu-id="c115e-693">布尔值</span><span class="sxs-lookup"><span data-stu-id="c115e-693">Boolean</span></span>|<span data-ttu-id="c115e-694">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="c115e-694">&lt;optional&gt;</span></span>|<span data-ttu-id="c115e-695">如果为 `true`，则表示附件将在邮件正文中内联显示，并且不应显示在附件列表中。</span><span class="sxs-lookup"><span data-stu-id="c115e-695">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`callback`|<span data-ttu-id="c115e-696">函数</span><span class="sxs-lookup"><span data-stu-id="c115e-696">function</span></span>|<span data-ttu-id="c115e-697">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="c115e-697">&lt;optional&gt;</span></span>|<span data-ttu-id="c115e-698">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="c115e-698">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="c115e-699">如果成功，附件标识符将在 `asyncResult.value` 属性中提供。</span><span class="sxs-lookup"><span data-stu-id="c115e-699">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="c115e-700">如果上传附件失败，`asyncResult` 对象将包含一个提供错误说明的 `Error` 对象。</span><span class="sxs-lookup"><span data-stu-id="c115e-700">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="c115e-701">错误</span><span class="sxs-lookup"><span data-stu-id="c115e-701">Errors</span></span>

|<span data-ttu-id="c115e-702">错误代码</span><span class="sxs-lookup"><span data-stu-id="c115e-702">Error code</span></span>|<span data-ttu-id="c115e-703">说明</span><span class="sxs-lookup"><span data-stu-id="c115e-703">Description</span></span>|
|------------|-------------|
|`AttachmentSizeExceeded`|<span data-ttu-id="c115e-704">附件大小超过了允许的大小。</span><span class="sxs-lookup"><span data-stu-id="c115e-704">The attachment is larger than allowed.</span></span>|
|`FileTypeNotSupported`|<span data-ttu-id="c115e-705">该附件的扩展名不是允许的扩展名。</span><span class="sxs-lookup"><span data-stu-id="c115e-705">The attachment has an extension that is not allowed.</span></span>|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="c115e-706">邮件或约会具有的附件过多。</span><span class="sxs-lookup"><span data-stu-id="c115e-706">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c115e-707">要求</span><span class="sxs-lookup"><span data-stu-id="c115e-707">Requirements</span></span>

|<span data-ttu-id="c115e-708">要求</span><span class="sxs-lookup"><span data-stu-id="c115e-708">Requirement</span></span>|<span data-ttu-id="c115e-709">值</span><span class="sxs-lookup"><span data-stu-id="c115e-709">Value</span></span>|
|---|---|
|[<span data-ttu-id="c115e-710">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="c115e-710">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c115e-711">1.1</span><span class="sxs-lookup"><span data-stu-id="c115e-711">1.1</span></span>|
|[<span data-ttu-id="c115e-712">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="c115e-712">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c115e-713">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="c115e-713">ReadWriteItem</span></span>|
|[<span data-ttu-id="c115e-714">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c115e-714">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="c115e-715">撰写</span><span class="sxs-lookup"><span data-stu-id="c115e-715">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="c115e-716">示例</span><span class="sxs-lookup"><span data-stu-id="c115e-716">Examples</span></span>

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

<span data-ttu-id="c115e-717">以下示例将图像文件添加为内联附件，并在邮件正文中引用该附件。</span><span class="sxs-lookup"><span data-stu-id="c115e-717">The following example adds an image file as an inline attachment and references the attachment in the message body.</span></span>

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

####  <a name="addhandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="c115e-718">addHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="c115e-718">addHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="c115e-719">添加支持事件的事件处理程序。</span><span class="sxs-lookup"><span data-stu-id="c115e-719">Adds an event handler for a supported event.</span></span>

<span data-ttu-id="c115e-720">当前，支持的事件类型是 `Office.EventType.AppointmentTimeChanged`、`Office.EventType.RecipientsChanged` 和 `Office.EventType.RecurrenceChanged`</span><span class="sxs-lookup"><span data-stu-id="c115e-720">Currently the supported event types are `Office.EventType.AppointmentTimeChanged`, `Office.EventType.RecipientsChanged`, and `Office.EventType.RecurrenceChanged`</span></span>

##### <a name="parameters"></a><span data-ttu-id="c115e-721">参数：</span><span class="sxs-lookup"><span data-stu-id="c115e-721">Parameters:</span></span>

| <span data-ttu-id="c115e-722">名称</span><span class="sxs-lookup"><span data-stu-id="c115e-722">Name</span></span> | <span data-ttu-id="c115e-723">类型</span><span class="sxs-lookup"><span data-stu-id="c115e-723">Type</span></span> | <span data-ttu-id="c115e-724">属性</span><span class="sxs-lookup"><span data-stu-id="c115e-724">Attributes</span></span> | <span data-ttu-id="c115e-725">说明</span><span class="sxs-lookup"><span data-stu-id="c115e-725">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="c115e-726">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="c115e-726">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="c115e-727">应调用处理程序的事件。</span><span class="sxs-lookup"><span data-stu-id="c115e-727">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="c115e-728">函数</span><span class="sxs-lookup"><span data-stu-id="c115e-728">Function</span></span> || <span data-ttu-id="c115e-p136">用于处理事件的函数。此函数必须接受一个参数，即对象文本。参数上的 `type` 属性将匹配传递给 `addHandlerAsync` 的 `eventType` 参数。</span><span class="sxs-lookup"><span data-stu-id="c115e-p136">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `addHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="c115e-732">Object</span><span class="sxs-lookup"><span data-stu-id="c115e-732">Object</span></span> | <span data-ttu-id="c115e-733">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="c115e-733">&lt;optional&gt;</span></span> | <span data-ttu-id="c115e-734">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="c115e-734">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="c115e-735">对象</span><span class="sxs-lookup"><span data-stu-id="c115e-735">Object</span></span> | <span data-ttu-id="c115e-736">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="c115e-736">&lt;optional&gt;</span></span> | <span data-ttu-id="c115e-737">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="c115e-737">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="c115e-738">函数</span><span class="sxs-lookup"><span data-stu-id="c115e-738">function</span></span>| <span data-ttu-id="c115e-739">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="c115e-739">&lt;optional&gt;</span></span>|<span data-ttu-id="c115e-740">方法完成后，使用单个参数 `callback`（一个 [`asyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `AsyncResult` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="c115e-740">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c115e-741">Requirements</span><span class="sxs-lookup"><span data-stu-id="c115e-741">Requirements</span></span>

|<span data-ttu-id="c115e-742">要求</span><span class="sxs-lookup"><span data-stu-id="c115e-742">Requirement</span></span>| <span data-ttu-id="c115e-743">值</span><span class="sxs-lookup"><span data-stu-id="c115e-743">Value</span></span>|
|---|---|
|[<span data-ttu-id="c115e-744">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="c115e-744">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c115e-745">1.7</span><span class="sxs-lookup"><span data-stu-id="c115e-745">1.7</span></span> |
|[<span data-ttu-id="c115e-746">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="c115e-746">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c115e-747">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c115e-747">ReadItem</span></span> |
|[<span data-ttu-id="c115e-748">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c115e-748">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c115e-749">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="c115e-749">Compose or read</span></span> |

##### <a name="example"></a><span data-ttu-id="c115e-750">示例</span><span class="sxs-lookup"><span data-stu-id="c115e-750">Example</span></span>

```js
Office.initialize = function (reason) {
  $(document).ready(function () {
    Office.context.mailbox.item.addHandlerAsync(Office.EventType.RecurrenceChanged, loadNewItem, function (result) {
      if (result.status === Office.AsyncResultStatus.Failed) {
        // Handle error
      }
    });
  });
};

function loadNewItem(eventArgs) {
  // Load the properties of the newly selected item
  loadProps(Office.context.mailbox.item);
};
```

####  <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="c115e-751">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="c115e-751">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="c115e-752">将 Exchange 项目（如邮件）作为附件添加到邮件或约会。</span><span class="sxs-lookup"><span data-stu-id="c115e-752">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="c115e-p137">`addItemAttachmentAsync` 方法将包含指定 Exchange 标识符的项目附加到撰写窗体中的项目。如果指定一个回调方法，此方法使用 `asyncResult` 参数调用，该参数包含一个附件标识符或代码，指示附加项目过程中出现的任何错误。可以使用 `options` 参数将状态信息传递给回调方法（如果需要）。</span><span class="sxs-lookup"><span data-stu-id="c115e-p137">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="c115e-756">随后可以将该标识符与 [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) 方法一同使用，以删除同一个会话中的附件。</span><span class="sxs-lookup"><span data-stu-id="c115e-756">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="c115e-757">如果 Office 加载项在 Outlook Web App 中运行，则 `addItemAttachmentAsync` 方法可以将项目附加到项目（正在编辑的项目除外）中；然而，不支持也不建议这样做。</span><span class="sxs-lookup"><span data-stu-id="c115e-757">If your Office Add-in is running in Outlook Web App, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c115e-758">参数：</span><span class="sxs-lookup"><span data-stu-id="c115e-758">Parameters:</span></span>

|<span data-ttu-id="c115e-759">名称</span><span class="sxs-lookup"><span data-stu-id="c115e-759">Name</span></span>|<span data-ttu-id="c115e-760">类型</span><span class="sxs-lookup"><span data-stu-id="c115e-760">Type</span></span>|<span data-ttu-id="c115e-761">属性</span><span class="sxs-lookup"><span data-stu-id="c115e-761">Attributes</span></span>|<span data-ttu-id="c115e-762">说明</span><span class="sxs-lookup"><span data-stu-id="c115e-762">Description</span></span>|
|---|---|---|---|
|`itemId`|<span data-ttu-id="c115e-763">String</span><span class="sxs-lookup"><span data-stu-id="c115e-763">String</span></span>||<span data-ttu-id="c115e-p138">要附加的项目的 Exchange 标识符。最大长度为 100 个字符。</span><span class="sxs-lookup"><span data-stu-id="c115e-p138">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`|<span data-ttu-id="c115e-766">String</span><span class="sxs-lookup"><span data-stu-id="c115e-766">String</span></span>||<span data-ttu-id="c115e-p139">要附加的项目的主题。最大长度为 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="c115e-p139">The sujbect of the item to be attached. The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="c115e-769">对象</span><span class="sxs-lookup"><span data-stu-id="c115e-769">Object</span></span>|<span data-ttu-id="c115e-770">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="c115e-770">&lt;optional&gt;</span></span>|<span data-ttu-id="c115e-771">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="c115e-771">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="c115e-772">对象</span><span class="sxs-lookup"><span data-stu-id="c115e-772">Object</span></span>|<span data-ttu-id="c115e-773">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="c115e-773">&lt;optional&gt;</span></span>|<span data-ttu-id="c115e-774">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="c115e-774">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="c115e-775">函数</span><span class="sxs-lookup"><span data-stu-id="c115e-775">function</span></span>|<span data-ttu-id="c115e-776">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="c115e-776">&lt;optional&gt;</span></span>|<span data-ttu-id="c115e-777">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="c115e-777">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="c115e-778">如果成功，附件标识符将在 `asyncResult.value` 属性中提供。</span><span class="sxs-lookup"><span data-stu-id="c115e-778">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="c115e-779">如果添加附件失败，`asyncResult` 对象将包含一个提供错误说明的 `Error` 对象。</span><span class="sxs-lookup"><span data-stu-id="c115e-779">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="c115e-780">错误</span><span class="sxs-lookup"><span data-stu-id="c115e-780">Errors</span></span>

|<span data-ttu-id="c115e-781">错误代码</span><span class="sxs-lookup"><span data-stu-id="c115e-781">Error code</span></span>|<span data-ttu-id="c115e-782">说明</span><span class="sxs-lookup"><span data-stu-id="c115e-782">Description</span></span>|
|------------|-------------|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="c115e-783">邮件或约会具有的附件过多。</span><span class="sxs-lookup"><span data-stu-id="c115e-783">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c115e-784">Requirements</span><span class="sxs-lookup"><span data-stu-id="c115e-784">Requirements</span></span>

|<span data-ttu-id="c115e-785">要求</span><span class="sxs-lookup"><span data-stu-id="c115e-785">Requirement</span></span>|<span data-ttu-id="c115e-786">值</span><span class="sxs-lookup"><span data-stu-id="c115e-786">Value</span></span>|
|---|---|
|[<span data-ttu-id="c115e-787">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="c115e-787">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c115e-788">1.1</span><span class="sxs-lookup"><span data-stu-id="c115e-788">1.1</span></span>|
|[<span data-ttu-id="c115e-789">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="c115e-789">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c115e-790">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="c115e-790">ReadWriteItem</span></span>|
|[<span data-ttu-id="c115e-791">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c115e-791">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="c115e-792">撰写</span><span class="sxs-lookup"><span data-stu-id="c115e-792">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="c115e-793">示例</span><span class="sxs-lookup"><span data-stu-id="c115e-793">Example</span></span>

<span data-ttu-id="c115e-794">以下示例将现有的 Outlook 项目添加为名为 `My Attachment` 的附件。</span><span class="sxs-lookup"><span data-stu-id="c115e-794">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

```js
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

####  <a name="close"></a><span data-ttu-id="c115e-795">close()</span><span class="sxs-lookup"><span data-stu-id="c115e-795">close()</span></span>

<span data-ttu-id="c115e-796">关闭当前正在撰写的项目。</span><span class="sxs-lookup"><span data-stu-id="c115e-796">Closes the current item that is being composed.</span></span>

<span data-ttu-id="c115e-p140">
            \`close\` 方法的行为取决于要撰写的项目的当前状态。如果项目具有未保存的更改，客户端将提示用户保存、放弃或取消关闭操作。</span><span class="sxs-lookup"><span data-stu-id="c115e-p140">The behavior of the `close` method depends on the current state of the item being composed. If the item has unsaved changes, the client prompts the user to save, discard, or cancel the close action.</span></span>

> [!NOTE]
> <span data-ttu-id="c115e-799">在 Outlook 网页版中，如果该项目是约会并且之前已使用 `saveAsync` 保存，则即使自上次保存项目后未发生任何更改，也会提示用户保存、放弃或取消。</span><span class="sxs-lookup"><span data-stu-id="c115e-799">In Outlook on the web, if the item is an appointment and it has previously been saved using `saveAsync`, the user is prompted to save, discard, or cancel even if no changes have occurred since the item was last saved.</span></span>

<span data-ttu-id="c115e-800">在 Outlook 桌面客户端中，如果邮件是内联答复，`close` 方法不起作用。</span><span class="sxs-lookup"><span data-stu-id="c115e-800">In the Outlook desktop client, if the message is an inline reply, the `close` method has no effect.</span></span>

##### <a name="requirements"></a><span data-ttu-id="c115e-801">要求</span><span class="sxs-lookup"><span data-stu-id="c115e-801">Requirements</span></span>

|<span data-ttu-id="c115e-802">要求</span><span class="sxs-lookup"><span data-stu-id="c115e-802">Requirement</span></span>|<span data-ttu-id="c115e-803">值</span><span class="sxs-lookup"><span data-stu-id="c115e-803">Value</span></span>|
|---|---|
|[<span data-ttu-id="c115e-804">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="c115e-804">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c115e-805">1.3</span><span class="sxs-lookup"><span data-stu-id="c115e-805">1.3</span></span>|
|[<span data-ttu-id="c115e-806">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="c115e-806">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c115e-807">受限</span><span class="sxs-lookup"><span data-stu-id="c115e-807">Restricted</span></span>|
|[<span data-ttu-id="c115e-808">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c115e-808">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="c115e-809">撰写</span><span class="sxs-lookup"><span data-stu-id="c115e-809">Compose</span></span>|

#### <a name="displayreplyallformformdata"></a><span data-ttu-id="c115e-810">displayReplyAllForm(formData)</span><span class="sxs-lookup"><span data-stu-id="c115e-810">displayReplyAllForm(formData)</span></span>

<span data-ttu-id="c115e-811">显示答复窗体，其中包括所选邮件的发件人和所有收件人或所选约会的组织者和所有与会者。</span><span class="sxs-lookup"><span data-stu-id="c115e-811">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="c115e-812">在 Outlook for iOS 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="c115e-812">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="c115e-813">在 Outlook Web App 中，答复窗体显示为包含 3 列视图的弹出式窗体以及包含 2 列或 1 列视图的弹出式窗体。</span><span class="sxs-lookup"><span data-stu-id="c115e-813">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="c115e-814">如果任意字符串参数超出其限制，`displayReplyAllForm` 将引发异常。</span><span class="sxs-lookup"><span data-stu-id="c115e-814">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="c115e-p141">当在 `formData.attachments` 参数中指定附件时，Outlook 和 Outlook Web App 尝试下载所有附件并将其附加到答复窗体。如果无法添加任何附件，则在窗体 UI 中显示错误。如果这不可能，则不引发错误消息。</span><span class="sxs-lookup"><span data-stu-id="c115e-p141">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c115e-818">参数：</span><span class="sxs-lookup"><span data-stu-id="c115e-818">Parameters:</span></span>

|<span data-ttu-id="c115e-819">名称</span><span class="sxs-lookup"><span data-stu-id="c115e-819">Name</span></span>|<span data-ttu-id="c115e-820">类型</span><span class="sxs-lookup"><span data-stu-id="c115e-820">Type</span></span>|<span data-ttu-id="c115e-821">属性</span><span class="sxs-lookup"><span data-stu-id="c115e-821">Attributes</span></span>|<span data-ttu-id="c115e-822">说明</span><span class="sxs-lookup"><span data-stu-id="c115e-822">Description</span></span>|
|---|---|---|---|
|`formData`|<span data-ttu-id="c115e-823">字符串 &#124; 对象</span><span class="sxs-lookup"><span data-stu-id="c115e-823">String &#124; Object</span></span>||<span data-ttu-id="c115e-p142">一个包含文本和 HTML 且表示答复窗体的正文的字符串。字符串限制为 32 KB。</span><span class="sxs-lookup"><span data-stu-id="c115e-p142">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="c115e-826">**或**</span><span class="sxs-lookup"><span data-stu-id="c115e-826">**OR**</span></span><br/><span data-ttu-id="c115e-p143">包含正文或附件数据和回调函数的对象。对象定义如下。</span><span class="sxs-lookup"><span data-stu-id="c115e-p143">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span>|
|`formData.htmlBody`|<span data-ttu-id="c115e-829">String</span><span class="sxs-lookup"><span data-stu-id="c115e-829">String</span></span>|<span data-ttu-id="c115e-830">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="c115e-830">&lt;optional&gt;</span></span>|<span data-ttu-id="c115e-p144">一个包含文本和 HTML 且表示答复窗体的正文的字符串。字符串限制为 32 KB。</span><span class="sxs-lookup"><span data-stu-id="c115e-p144">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
|`formData.attachments`|<span data-ttu-id="c115e-833">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="c115e-833">Array.&lt;Object&gt;</span></span>|<span data-ttu-id="c115e-834">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="c115e-834">&lt;optional&gt;</span></span>|<span data-ttu-id="c115e-835">JSON 对象（是文件或项目附件）的数组。</span><span class="sxs-lookup"><span data-stu-id="c115e-835">An array of JSON objects that are either file or item attachments.</span></span>|
|`formData.attachments.type`|<span data-ttu-id="c115e-836">字符串</span><span class="sxs-lookup"><span data-stu-id="c115e-836">String</span></span>||<span data-ttu-id="c115e-p145">指示附件的类型。必须是文件附件的 `file` 或项目附件的 `item`。</span><span class="sxs-lookup"><span data-stu-id="c115e-p145">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span>|
|`formData.attachments.name`|<span data-ttu-id="c115e-839">字符串</span><span class="sxs-lookup"><span data-stu-id="c115e-839">String</span></span>||<span data-ttu-id="c115e-840">一个包含附件的名称的字符串，最多包含 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="c115e-840">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
|`formData.attachments.url`|<span data-ttu-id="c115e-841">String</span><span class="sxs-lookup"><span data-stu-id="c115e-841">String</span></span>||<span data-ttu-id="c115e-p146">仅在将 `type` 设置为 `file` 时使用。文件的位置的 URI。</span><span class="sxs-lookup"><span data-stu-id="c115e-p146">Only used if `type` is set to `file`. The URI of the location for the file.</span></span>|
|`formData.attachments.isInline`|<span data-ttu-id="c115e-844">Boolean</span><span class="sxs-lookup"><span data-stu-id="c115e-844">Boolean</span></span>||<span data-ttu-id="c115e-p147">仅在将 `type` 设置为 `file` 时使用。如果为 `true`，则表示附件将在邮件正文中内联显示，并且不应显示在附件列表中。</span><span class="sxs-lookup"><span data-stu-id="c115e-p147">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`formData.attachments.itemId`|<span data-ttu-id="c115e-847">String</span><span class="sxs-lookup"><span data-stu-id="c115e-847">String</span></span>||<span data-ttu-id="c115e-p148">仅在将 `type` 设置为 `item` 时使用。附件的 EWS 项目 ID。字符串最长为 100 个字符。</span><span class="sxs-lookup"><span data-stu-id="c115e-p148">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span>|
|`callback`|<span data-ttu-id="c115e-851">函数</span><span class="sxs-lookup"><span data-stu-id="c115e-851">function</span></span>|<span data-ttu-id="c115e-852">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="c115e-852">&lt;optional&gt;</span></span>|<span data-ttu-id="c115e-853">方法完成后，使用单个参数 `asyncResult`（一个 [AsyncResult](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="c115e-853">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c115e-854">Requirements</span><span class="sxs-lookup"><span data-stu-id="c115e-854">Requirements</span></span>

|<span data-ttu-id="c115e-855">要求</span><span class="sxs-lookup"><span data-stu-id="c115e-855">Requirement</span></span>|<span data-ttu-id="c115e-856">值</span><span class="sxs-lookup"><span data-stu-id="c115e-856">Value</span></span>|
|---|---|
|[<span data-ttu-id="c115e-857">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="c115e-857">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c115e-858">1.0</span><span class="sxs-lookup"><span data-stu-id="c115e-858">1.0</span></span>|
|[<span data-ttu-id="c115e-859">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="c115e-859">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c115e-860">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c115e-860">ReadItem</span></span>|
|[<span data-ttu-id="c115e-861">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c115e-861">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="c115e-862">阅读</span><span class="sxs-lookup"><span data-stu-id="c115e-862">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="c115e-863">示例</span><span class="sxs-lookup"><span data-stu-id="c115e-863">Examples</span></span>

<span data-ttu-id="c115e-864">以下代码将一个字符串传递到 `displayReplyAllForm` 函数。</span><span class="sxs-lookup"><span data-stu-id="c115e-864">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="c115e-865">使用空白正文答复。</span><span class="sxs-lookup"><span data-stu-id="c115e-865">Reply with an empty body.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="c115e-866">仅使用正文答复。</span><span class="sxs-lookup"><span data-stu-id="c115e-866">Reply with just a body.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="c115e-867">使用正文和文件附件答复。</span><span class="sxs-lookup"><span data-stu-id="c115e-867">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="c115e-868">使用正文和项目附件答复。</span><span class="sxs-lookup"><span data-stu-id="c115e-868">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="c115e-869">使用正文、文件附件、项目附件和回调答复。</span><span class="sxs-lookup"><span data-stu-id="c115e-869">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="displayreplyformformdata"></a><span data-ttu-id="c115e-870">displayReplyForm(formData)</span><span class="sxs-lookup"><span data-stu-id="c115e-870">displayReplyForm(formData)</span></span>

<span data-ttu-id="c115e-871">显示答复窗体，其中仅包括所选邮件的发件人或所选约会的组织者。</span><span class="sxs-lookup"><span data-stu-id="c115e-871">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="c115e-872">在 Outlook for iOS 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="c115e-872">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="c115e-873">在 Outlook Web App 中，答复窗体显示为包含 3 列视图的弹出式窗体以及包含 2 列或 1 列视图的弹出式窗体。</span><span class="sxs-lookup"><span data-stu-id="c115e-873">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="c115e-874">如果任意字符串参数超出其限制，`displayReplyForm` 将引发异常。</span><span class="sxs-lookup"><span data-stu-id="c115e-874">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="c115e-p149">当在 `formData.attachments` 参数中指定附件时，Outlook 和 Outlook Web App 尝试下载所有附件并将其附加到答复窗体。如果无法添加任何附件，则在窗体 UI 中显示错误。如果这不可能，则不引发错误消息。</span><span class="sxs-lookup"><span data-stu-id="c115e-p149">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c115e-878">参数：</span><span class="sxs-lookup"><span data-stu-id="c115e-878">Parameters:</span></span>

|<span data-ttu-id="c115e-879">名称</span><span class="sxs-lookup"><span data-stu-id="c115e-879">Name</span></span>|<span data-ttu-id="c115e-880">类型</span><span class="sxs-lookup"><span data-stu-id="c115e-880">Type</span></span>|<span data-ttu-id="c115e-881">属性</span><span class="sxs-lookup"><span data-stu-id="c115e-881">Attributes</span></span>|<span data-ttu-id="c115e-882">说明</span><span class="sxs-lookup"><span data-stu-id="c115e-882">Description</span></span>|
|---|---|---|---|
|`formData`|<span data-ttu-id="c115e-883">字符串 &#124; 对象</span><span class="sxs-lookup"><span data-stu-id="c115e-883">String &#124; Object</span></span>||<span data-ttu-id="c115e-p150">一个包含文本和 HTML 且表示答复窗体的正文的字符串。字符串限制为 32 KB。</span><span class="sxs-lookup"><span data-stu-id="c115e-p150">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="c115e-886">**或**</span><span class="sxs-lookup"><span data-stu-id="c115e-886">**OR**</span></span><br/><span data-ttu-id="c115e-p151">包含正文或附件数据和回调函数的对象。对象定义如下。</span><span class="sxs-lookup"><span data-stu-id="c115e-p151">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span>|
|`formData.htmlBody`|<span data-ttu-id="c115e-889">String</span><span class="sxs-lookup"><span data-stu-id="c115e-889">String</span></span>|<span data-ttu-id="c115e-890">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="c115e-890">&lt;optional&gt;</span></span>|<span data-ttu-id="c115e-p152">一个包含文本和 HTML 且表示答复窗体的正文的字符串。字符串限制为 32 KB。</span><span class="sxs-lookup"><span data-stu-id="c115e-p152">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
|`formData.attachments`|<span data-ttu-id="c115e-893">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="c115e-893">Array.&lt;Object&gt;</span></span>|<span data-ttu-id="c115e-894">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="c115e-894">&lt;optional&gt;</span></span>|<span data-ttu-id="c115e-895">JSON 对象（是文件或项目附件）的数组。</span><span class="sxs-lookup"><span data-stu-id="c115e-895">An array of JSON objects that are either file or item attachments.</span></span>|
|`formData.attachments.type`|<span data-ttu-id="c115e-896">String</span><span class="sxs-lookup"><span data-stu-id="c115e-896">String</span></span>||<span data-ttu-id="c115e-p153">指示附件的类型。必须是文件附件的 `file` 或项目附件的 `item`。</span><span class="sxs-lookup"><span data-stu-id="c115e-p153">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span>|
|`formData.attachments.name`|<span data-ttu-id="c115e-899">字符串</span><span class="sxs-lookup"><span data-stu-id="c115e-899">String</span></span>||<span data-ttu-id="c115e-900">一个包含附件的名称的字符串，最多包含 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="c115e-900">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
|`formData.attachments.url`|<span data-ttu-id="c115e-901">String</span><span class="sxs-lookup"><span data-stu-id="c115e-901">String</span></span>||<span data-ttu-id="c115e-p154">仅在将 `type` 设置为 `file` 时使用。文件的位置的 URI。</span><span class="sxs-lookup"><span data-stu-id="c115e-p154">Only used if `type` is set to `file`. The URI of the location for the file.</span></span>|
|`formData.attachments.isInline`|<span data-ttu-id="c115e-904">Boolean</span><span class="sxs-lookup"><span data-stu-id="c115e-904">Boolean</span></span>||<span data-ttu-id="c115e-p155">仅在将 `type` 设置为 `file` 时使用。如果为 `true`，则表示附件将在邮件正文中内联显示，并且不应显示在附件列表中。</span><span class="sxs-lookup"><span data-stu-id="c115e-p155">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`formData.attachments.itemId`|<span data-ttu-id="c115e-907">String</span><span class="sxs-lookup"><span data-stu-id="c115e-907">String</span></span>||<span data-ttu-id="c115e-p156">仅在将 `type` 设置为 `item` 时使用。附件的 EWS 项目 ID。字符串最长为 100 个字符。</span><span class="sxs-lookup"><span data-stu-id="c115e-p156">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span>|
|`callback`|<span data-ttu-id="c115e-911">函数</span><span class="sxs-lookup"><span data-stu-id="c115e-911">function</span></span>|<span data-ttu-id="c115e-912">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="c115e-912">&lt;optional&gt;</span></span>|<span data-ttu-id="c115e-913">方法完成后，使用单个参数 `asyncResult`（一个 [AsyncResult](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="c115e-913">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c115e-914">Requirements</span><span class="sxs-lookup"><span data-stu-id="c115e-914">Requirements</span></span>

|<span data-ttu-id="c115e-915">要求</span><span class="sxs-lookup"><span data-stu-id="c115e-915">Requirement</span></span>|<span data-ttu-id="c115e-916">值</span><span class="sxs-lookup"><span data-stu-id="c115e-916">Value</span></span>|
|---|---|
|[<span data-ttu-id="c115e-917">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="c115e-917">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c115e-918">1.0</span><span class="sxs-lookup"><span data-stu-id="c115e-918">1.0</span></span>|
|[<span data-ttu-id="c115e-919">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="c115e-919">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c115e-920">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c115e-920">ReadItem</span></span>|
|[<span data-ttu-id="c115e-921">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c115e-921">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="c115e-922">阅读</span><span class="sxs-lookup"><span data-stu-id="c115e-922">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="c115e-923">示例</span><span class="sxs-lookup"><span data-stu-id="c115e-923">Examples</span></span>

<span data-ttu-id="c115e-924">以下代码将一个字符串传递到 `displayReplyForm` 函数。</span><span class="sxs-lookup"><span data-stu-id="c115e-924">The following code passes a string to the `displayReplyForm` function.</span></span>

```js
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="c115e-925">使用空白正文答复。</span><span class="sxs-lookup"><span data-stu-id="c115e-925">Reply with an empty body.</span></span>

```js
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="c115e-926">仅使用正文答复。</span><span class="sxs-lookup"><span data-stu-id="c115e-926">Reply with just a body.</span></span>

```js
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="c115e-927">使用正文和文件附件答复。</span><span class="sxs-lookup"><span data-stu-id="c115e-927">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="c115e-928">使用正文和项目附件答复。</span><span class="sxs-lookup"><span data-stu-id="c115e-928">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="c115e-929">使用正文、文件附件、项目附件和回调答复。</span><span class="sxs-lookup"><span data-stu-id="c115e-929">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="getentities--entitiesjavascriptapioutlook17officeentities"></a><span data-ttu-id="c115e-930">getEntities() → {[Entities](/javascript/api/outlook_1_7/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="c115e-930">getEntities() → {[Entities](/javascript/api/outlook_1_7/office.entities)}</span></span>

<span data-ttu-id="c115e-931">获取在所选项目的正文中找到的实体。</span><span class="sxs-lookup"><span data-stu-id="c115e-931">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="c115e-932">在 Outlook for iOS 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="c115e-932">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="c115e-933">要求</span><span class="sxs-lookup"><span data-stu-id="c115e-933">Requirements</span></span>

|<span data-ttu-id="c115e-934">要求</span><span class="sxs-lookup"><span data-stu-id="c115e-934">Requirement</span></span>|<span data-ttu-id="c115e-935">值</span><span class="sxs-lookup"><span data-stu-id="c115e-935">Value</span></span>|
|---|---|
|[<span data-ttu-id="c115e-936">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="c115e-936">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c115e-937">1.0</span><span class="sxs-lookup"><span data-stu-id="c115e-937">1.0</span></span>|
|[<span data-ttu-id="c115e-938">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="c115e-938">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c115e-939">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c115e-939">ReadItem</span></span>|
|[<span data-ttu-id="c115e-940">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c115e-940">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="c115e-941">阅读</span><span class="sxs-lookup"><span data-stu-id="c115e-941">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="c115e-942">返回：</span><span class="sxs-lookup"><span data-stu-id="c115e-942">Returns:</span></span>

<span data-ttu-id="c115e-943">类型：[Entities](/javascript/api/outlook_1_7/office.entities)</span><span class="sxs-lookup"><span data-stu-id="c115e-943">Type: [Entities](/javascript/api/outlook_1_7/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="c115e-944">示例</span><span class="sxs-lookup"><span data-stu-id="c115e-944">Example</span></span>

<span data-ttu-id="c115e-945">以下示例访问当前项目的正文中的联系人实体。</span><span class="sxs-lookup"><span data-stu-id="c115e-945">The following example accesses the contacts entities in the current item's body.</span></span>

```js
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlook17officecontactmeetingsuggestionjavascriptapioutlook17officemeetingsuggestionphonenumberjavascriptapioutlook17officephonenumbertasksuggestionjavascriptapioutlook17officetasksuggestion"></a><span data-ttu-id="c115e-946">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_7/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_7/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_7/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_7/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="c115e-946">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_7/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_7/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_7/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_7/office.tasksuggestion))>}</span></span>

<span data-ttu-id="c115e-947">获取所选项目的正文中找到的指定实体类型的所有实体的数组。</span><span class="sxs-lookup"><span data-stu-id="c115e-947">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="c115e-948">在 Outlook for iOS 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="c115e-948">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c115e-949">参数：</span><span class="sxs-lookup"><span data-stu-id="c115e-949">Parameters:</span></span>

|<span data-ttu-id="c115e-950">名称</span><span class="sxs-lookup"><span data-stu-id="c115e-950">Name</span></span>|<span data-ttu-id="c115e-951">说明</span><span class="sxs-lookup"><span data-stu-id="c115e-951">Type</span></span>|<span data-ttu-id="c115e-952">说明</span><span class="sxs-lookup"><span data-stu-id="c115e-952">Description</span></span>|
|---|---|---|
|`entityType`|[<span data-ttu-id="c115e-953">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="c115e-953">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook_1_7/office.mailboxenums.entitytype)|<span data-ttu-id="c115e-954">EntityType 枚举值之一。</span><span class="sxs-lookup"><span data-stu-id="c115e-954">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c115e-955">Requirements</span><span class="sxs-lookup"><span data-stu-id="c115e-955">Requirements</span></span>

|<span data-ttu-id="c115e-956">要求</span><span class="sxs-lookup"><span data-stu-id="c115e-956">Requirement</span></span>|<span data-ttu-id="c115e-957">值</span><span class="sxs-lookup"><span data-stu-id="c115e-957">Value</span></span>|
|---|---|
|[<span data-ttu-id="c115e-958">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="c115e-958">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c115e-959">1.0</span><span class="sxs-lookup"><span data-stu-id="c115e-959">1.0</span></span>|
|[<span data-ttu-id="c115e-960">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="c115e-960">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c115e-961">受限</span><span class="sxs-lookup"><span data-stu-id="c115e-961">Restricted</span></span>|
|[<span data-ttu-id="c115e-962">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c115e-962">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="c115e-963">阅读</span><span class="sxs-lookup"><span data-stu-id="c115e-963">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="c115e-964">返回：</span><span class="sxs-lookup"><span data-stu-id="c115e-964">Returns:</span></span>

<span data-ttu-id="c115e-965">如果在 `entityType` 中传递的值不是 `EntityType` 枚举的有效成员，该方法返回 null。</span><span class="sxs-lookup"><span data-stu-id="c115e-965">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="c115e-966">如果指定类型的任何实体都不存在于该项目的正文中，该方法将返回空数组。</span><span class="sxs-lookup"><span data-stu-id="c115e-966">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="c115e-967">否则，返回的数组中对象的类型取决于 `entityType` 参数中请求实体的类型。</span><span class="sxs-lookup"><span data-stu-id="c115e-967">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="c115e-968">当使用此方法的最低权限级别**受限**时，某些实体类型需要 **ReadItem** 才能进行访问，如下表中所指定。</span><span class="sxs-lookup"><span data-stu-id="c115e-968">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

|<span data-ttu-id="c115e-969">`entityType` 的值</span><span class="sxs-lookup"><span data-stu-id="c115e-969">Value of `entityType`</span></span>|<span data-ttu-id="c115e-970">返回的数组中对象的类型</span><span class="sxs-lookup"><span data-stu-id="c115e-970">Type of objects in returned array</span></span>|<span data-ttu-id="c115e-971">所需权限级别</span><span class="sxs-lookup"><span data-stu-id="c115e-971">Required Permission Level</span></span>|
|---|---|---|
|`Address`|<span data-ttu-id="c115e-972">字符串</span><span class="sxs-lookup"><span data-stu-id="c115e-972">String</span></span>|<span data-ttu-id="c115e-973">**受限**</span><span class="sxs-lookup"><span data-stu-id="c115e-973">**Restricted**</span></span>|
|`Contact`|<span data-ttu-id="c115e-974">Contact</span><span class="sxs-lookup"><span data-stu-id="c115e-974">Contact</span></span>|<span data-ttu-id="c115e-975">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="c115e-975">**ReadItem**</span></span>|
|`EmailAddress`|<span data-ttu-id="c115e-976">String</span><span class="sxs-lookup"><span data-stu-id="c115e-976">String</span></span>|<span data-ttu-id="c115e-977">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="c115e-977">**ReadItem**</span></span>|
|`MeetingSuggestion`|<span data-ttu-id="c115e-978">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="c115e-978">MeetingSuggestion</span></span>|<span data-ttu-id="c115e-979">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="c115e-979">**ReadItem**</span></span>|
|`PhoneNumber`|<span data-ttu-id="c115e-980">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="c115e-980">PhoneNumber</span></span>|<span data-ttu-id="c115e-981">**受限**</span><span class="sxs-lookup"><span data-stu-id="c115e-981">**Restricted**</span></span>|
|`TaskSuggestion`|<span data-ttu-id="c115e-982">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="c115e-982">TaskSuggestion</span></span>|<span data-ttu-id="c115e-983">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="c115e-983">**ReadItem**</span></span>|
|`URL`|<span data-ttu-id="c115e-984">String</span><span class="sxs-lookup"><span data-stu-id="c115e-984">String</span></span>|<span data-ttu-id="c115e-985">**受限**</span><span class="sxs-lookup"><span data-stu-id="c115e-985">**Restricted**</span></span>|

<span data-ttu-id="c115e-986">类型：Array.<(String|[Contact](/javascript/api/outlook_1_7/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_7/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_7/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_7/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="c115e-986">Type: Array.<(String|[Contact](/javascript/api/outlook_1_7/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_7/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_7/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_7/office.tasksuggestion))></span></span>

##### <a name="example"></a><span data-ttu-id="c115e-987">示例</span><span class="sxs-lookup"><span data-stu-id="c115e-987">Example</span></span>

<span data-ttu-id="c115e-988">以下示例显示了如何访问表示当前项目的正文中的邮政地址的字符串数组。</span><span class="sxs-lookup"><span data-stu-id="c115e-988">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlook17officecontactmeetingsuggestionjavascriptapioutlook17officemeetingsuggestionphonenumberjavascriptapioutlook17officephonenumbertasksuggestionjavascriptapioutlook17officetasksuggestion"></a><span data-ttu-id="c115e-989">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_7/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_7/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_7/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_7/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="c115e-989">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_7/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_7/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_7/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_7/office.tasksuggestion))>}</span></span>

<span data-ttu-id="c115e-990">返回传递清单 XML 文件中定义的命名筛选器的所选项目中的已知实体。</span><span class="sxs-lookup"><span data-stu-id="c115e-990">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="c115e-991">在 Outlook for iOS 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="c115e-991">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="c115e-992">`getFilteredEntitiesByName` 方法返回匹配在具有指定 `FilterName` 元素值的清单 XML 文件中的 [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) 规则元素中定义的正则表达式的实体。</span><span class="sxs-lookup"><span data-stu-id="c115e-992">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c115e-993">参数：</span><span class="sxs-lookup"><span data-stu-id="c115e-993">Parameters:</span></span>

|<span data-ttu-id="c115e-994">名称</span><span class="sxs-lookup"><span data-stu-id="c115e-994">Name</span></span>|<span data-ttu-id="c115e-995">说明</span><span class="sxs-lookup"><span data-stu-id="c115e-995">Type</span></span>|<span data-ttu-id="c115e-996">描述</span><span class="sxs-lookup"><span data-stu-id="c115e-996">Description</span></span>|
|---|---|---|
|`name`|<span data-ttu-id="c115e-997">字符串</span><span class="sxs-lookup"><span data-stu-id="c115e-997">String</span></span>|<span data-ttu-id="c115e-998">定义筛选器匹配的 `ItemHasKnownEntity` 规则元素的名称。</span><span class="sxs-lookup"><span data-stu-id="c115e-998">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c115e-999">要求</span><span class="sxs-lookup"><span data-stu-id="c115e-999">Requirements</span></span>

|<span data-ttu-id="c115e-1000">要求</span><span class="sxs-lookup"><span data-stu-id="c115e-1000">Requirement</span></span>|<span data-ttu-id="c115e-1001">值</span><span class="sxs-lookup"><span data-stu-id="c115e-1001">Value</span></span>|
|---|---|
|[<span data-ttu-id="c115e-1002">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="c115e-1002">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c115e-1003">1.0</span><span class="sxs-lookup"><span data-stu-id="c115e-1003">1.0</span></span>|
|[<span data-ttu-id="c115e-1004">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="c115e-1004">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c115e-1005">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c115e-1005">ReadItem</span></span>|
|[<span data-ttu-id="c115e-1006">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c115e-1006">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="c115e-1007">阅读</span><span class="sxs-lookup"><span data-stu-id="c115e-1007">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="c115e-1008">返回：</span><span class="sxs-lookup"><span data-stu-id="c115e-1008">Returns:</span></span>

<span data-ttu-id="c115e-p158">如果具有匹配 `name` 参数的 `FilterName` 元素值的清单中没有任何 `ItemHasKnownEntity` 元素，则该方法返回 `null`。如果 `name` 参数匹配清单中的 `ItemHasKnownEntity` 元素，但在匹配的当前项目中没有实体，则该方法返回一个空数组。</span><span class="sxs-lookup"><span data-stu-id="c115e-p158">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="c115e-1011">类型：Array.<(String|[Contact](/javascript/api/outlook_1_7/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_7/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_7/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_7/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="c115e-1011">Type: Array.<(String|[Contact](/javascript/api/outlook_1_7/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_7/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_7/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_7/office.tasksuggestion))></span></span>

#### <a name="getregexmatches--object"></a><span data-ttu-id="c115e-1012">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="c115e-1012">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="c115e-1013">返回所选项目中匹配在清单 XML 文件中定义的正则表达式的字符串值。</span><span class="sxs-lookup"><span data-stu-id="c115e-1013">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="c115e-1014">在 Outlook for iOS 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="c115e-1014">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="c115e-p159">`getRegExMatches` 方法返回匹配在清单 XML 文件中的每个 `ItemHasRegularExpressionMatch` 或 `ItemHasKnownEntity` 规则元素中定义的正则表达式的字符串。对于 `ItemHasRegularExpressionMatch` 规则，匹配字符串必须发生在该规则指定的项目的属性中。`PropertyName` 简单类型定义支持的属性。</span><span class="sxs-lookup"><span data-stu-id="c115e-p159">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="c115e-1018">例如，考虑一个外接程序清单具有以下 `Rule` 元素：</span><span class="sxs-lookup"><span data-stu-id="c115e-1018">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="c115e-1019">从 `getRegExMatches` 返回的对象应有两个属性：`fruits` 和 `veggies`。</span><span class="sxs-lookup"><span data-stu-id="c115e-1019">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="c115e-p160">如果在项目的正文属性上指定 `ItemHasRegularExpressionMatch` 规则，则正则表达式应进一步筛选正文，不应尝试返回该项目的整个正文。使用正则表达式（如 `.*`）获取项目的整个正文并不总是返回预期的结果。而是使用 [`Body.getAsync`](/javascript/api/outlook_1_7/office.body#getasync-coerciontype--options--callback-) 方法检索整个正文。</span><span class="sxs-lookup"><span data-stu-id="c115e-p160">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook_1_7/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="c115e-1023">Requirements</span><span class="sxs-lookup"><span data-stu-id="c115e-1023">Requirements</span></span>

|<span data-ttu-id="c115e-1024">要求</span><span class="sxs-lookup"><span data-stu-id="c115e-1024">Requirement</span></span>|<span data-ttu-id="c115e-1025">值</span><span class="sxs-lookup"><span data-stu-id="c115e-1025">Value</span></span>|
|---|---|
|[<span data-ttu-id="c115e-1026">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="c115e-1026">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c115e-1027">1.0</span><span class="sxs-lookup"><span data-stu-id="c115e-1027">1.0</span></span>|
|[<span data-ttu-id="c115e-1028">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="c115e-1028">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c115e-1029">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c115e-1029">ReadItem</span></span>|
|[<span data-ttu-id="c115e-1030">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c115e-1030">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="c115e-1031">阅读</span><span class="sxs-lookup"><span data-stu-id="c115e-1031">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="c115e-1032">返回：</span><span class="sxs-lookup"><span data-stu-id="c115e-1032">Returns:</span></span>

<span data-ttu-id="c115e-p161">一个包含匹配在清单 XML 文件中定义的正则表达式的字符串数组的对象。每个数组的名称等于匹配 `ItemHasRegularExpressionMatch` 规则的 `RegExName` 属性或匹配 `ItemHasKnownEntity` 规则的 `FilterName` 属性的相应值。</span><span class="sxs-lookup"><span data-stu-id="c115e-p161">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<dl class="param-type"><span data-ttu-id="c115e-1035">

<dt>
类型</dt>


</span><span class="sxs-lookup"><span data-stu-id="c115e-1035">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="c115e-1036">对象</span><span class="sxs-lookup"><span data-stu-id="c115e-1036">Object</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="c115e-1037">示例</span><span class="sxs-lookup"><span data-stu-id="c115e-1037">Example</span></span>

<span data-ttu-id="c115e-1038">以下示例显示了如何访问正则表达式规则元素 `fruits` 和 `veggies` 的匹配项的数组，这些元素在清单中指定。</span><span class="sxs-lookup"><span data-stu-id="c115e-1038">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```js
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veges = allMatches.veggies;
```

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="c115e-1039">getRegExMatchesByName(name) → (nullable) {Array.<String>}</span><span class="sxs-lookup"><span data-stu-id="c115e-1039">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="c115e-1040">返回匹配在清单 XML 文件中定义的命名正则表达式的所选项目中的字符串值。</span><span class="sxs-lookup"><span data-stu-id="c115e-1040">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="c115e-1041">在 Outlook for iOS 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="c115e-1041">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="c115e-1042">`getRegExMatchesByName` 方法返回匹配在具有指定 `RegExName` 元素值的清单 XML 文件中的 `ItemHasRegularExpressionMatch` 规则元素中定义的正则表达式的字符串。</span><span class="sxs-lookup"><span data-stu-id="c115e-1042">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="c115e-p162">如果在项目的正文属性上指定 `ItemHasRegularExpressionMatch` 规则，则正则表达式应进一步筛选正文，不应尝试返回该项目的整个正文。使用正则表达式（如 `.*`）获取项目的整个正文并不总是返回预期的结果。</span><span class="sxs-lookup"><span data-stu-id="c115e-p162">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c115e-1045">参数：</span><span class="sxs-lookup"><span data-stu-id="c115e-1045">Parameters:</span></span>

|<span data-ttu-id="c115e-1046">名称</span><span class="sxs-lookup"><span data-stu-id="c115e-1046">Name</span></span>|<span data-ttu-id="c115e-1047">说明</span><span class="sxs-lookup"><span data-stu-id="c115e-1047">Type</span></span>|<span data-ttu-id="c115e-1048">描述</span><span class="sxs-lookup"><span data-stu-id="c115e-1048">Description</span></span>|
|---|---|---|
|`name`|<span data-ttu-id="c115e-1049">字符串</span><span class="sxs-lookup"><span data-stu-id="c115e-1049">String</span></span>|<span data-ttu-id="c115e-1050">定义筛选器匹配的 `ItemHasRegularExpressionMatch` 规则元素的名称。</span><span class="sxs-lookup"><span data-stu-id="c115e-1050">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c115e-1051">要求</span><span class="sxs-lookup"><span data-stu-id="c115e-1051">Requirements</span></span>

|<span data-ttu-id="c115e-1052">要求</span><span class="sxs-lookup"><span data-stu-id="c115e-1052">Requirement</span></span>|<span data-ttu-id="c115e-1053">值</span><span class="sxs-lookup"><span data-stu-id="c115e-1053">Value</span></span>|
|---|---|
|[<span data-ttu-id="c115e-1054">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="c115e-1054">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c115e-1055">1.0</span><span class="sxs-lookup"><span data-stu-id="c115e-1055">1.0</span></span>|
|[<span data-ttu-id="c115e-1056">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="c115e-1056">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c115e-1057">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c115e-1057">ReadItem</span></span>|
|[<span data-ttu-id="c115e-1058">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c115e-1058">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="c115e-1059">阅读</span><span class="sxs-lookup"><span data-stu-id="c115e-1059">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="c115e-1060">返回：</span><span class="sxs-lookup"><span data-stu-id="c115e-1060">Returns:</span></span>

<span data-ttu-id="c115e-1061">一个包含匹配在清单 XML 文件中定义的正则表达式的字符串的数组。</span><span class="sxs-lookup"><span data-stu-id="c115e-1061">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<dl class="param-type"><span data-ttu-id="c115e-1062">

<dt>类型</dt>

</span><span class="sxs-lookup"><span data-stu-id="c115e-1062">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="c115e-1063">Array.< String ></span><span class="sxs-lookup"><span data-stu-id="c115e-1063">Array.< String ></span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="c115e-1064">示例</span><span class="sxs-lookup"><span data-stu-id="c115e-1064">Example</span></span>

```js
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

####  <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="c115e-1065">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="c115e-1065">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="c115e-1066">以异步方式返回邮件的主题或正文中选定的数据。</span><span class="sxs-lookup"><span data-stu-id="c115e-1066">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="c115e-p163">如果没有选定内容，但光标位于正文或主题中，此方法将会为所选数据返回 null。如果选定的是字段，而不是正文或主题，则此方法返回 `InvalidSelection` 错误。</span><span class="sxs-lookup"><span data-stu-id="c115e-p163">If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c115e-1069">参数：</span><span class="sxs-lookup"><span data-stu-id="c115e-1069">Parameters:</span></span>

|<span data-ttu-id="c115e-1070">名称</span><span class="sxs-lookup"><span data-stu-id="c115e-1070">Name</span></span>|<span data-ttu-id="c115e-1071">类型</span><span class="sxs-lookup"><span data-stu-id="c115e-1071">Type</span></span>|<span data-ttu-id="c115e-1072">属性</span><span class="sxs-lookup"><span data-stu-id="c115e-1072">Attributes</span></span>|<span data-ttu-id="c115e-1073">说明</span><span class="sxs-lookup"><span data-stu-id="c115e-1073">Description</span></span>|
|---|---|---|---|
|`coercionType`|[<span data-ttu-id="c115e-1074">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="c115e-1074">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="c115e-p164">请求数据的格式。如果为文本，则此方法返回纯文本作为字符串，删除任何显示的 HTML 标记。如果为 HTML，则此方法返回所选文本，不论是纯文本还是 HTML。</span><span class="sxs-lookup"><span data-stu-id="c115e-p164">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`|<span data-ttu-id="c115e-1078">对象</span><span class="sxs-lookup"><span data-stu-id="c115e-1078">Object</span></span>|<span data-ttu-id="c115e-1079">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="c115e-1079">&lt;optional&gt;</span></span>|<span data-ttu-id="c115e-1080">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="c115e-1080">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="c115e-1081">对象</span><span class="sxs-lookup"><span data-stu-id="c115e-1081">Object</span></span>|<span data-ttu-id="c115e-1082">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="c115e-1082">&lt;optional&gt;</span></span>|<span data-ttu-id="c115e-1083">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="c115e-1083">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="c115e-1084">函数</span><span class="sxs-lookup"><span data-stu-id="c115e-1084">function</span></span>||<span data-ttu-id="c115e-1085">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="c115e-1085">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="c115e-1086">若要从回调方法访问所选数据，请调用 `asyncResult.value.data`。</span><span class="sxs-lookup"><span data-stu-id="c115e-1086">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="c115e-1087">若要访问选定内容的源属性，请调用 `asyncResult.value.sourceProperty`，这将为 `body` 或 `subject`。</span><span class="sxs-lookup"><span data-stu-id="c115e-1087">To access the source property that the selection comes from, call `asyncResult.value.sourceProperty`, which will be either `body` or `subject`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c115e-1088">要求</span><span class="sxs-lookup"><span data-stu-id="c115e-1088">Requirements</span></span>

|<span data-ttu-id="c115e-1089">要求</span><span class="sxs-lookup"><span data-stu-id="c115e-1089">Requirement</span></span>|<span data-ttu-id="c115e-1090">值</span><span class="sxs-lookup"><span data-stu-id="c115e-1090">Value</span></span>|
|---|---|
|[<span data-ttu-id="c115e-1091">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="c115e-1091">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c115e-1092">1.2</span><span class="sxs-lookup"><span data-stu-id="c115e-1092">1.2</span></span>|
|[<span data-ttu-id="c115e-1093">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="c115e-1093">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c115e-1094">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="c115e-1094">ReadWriteItem</span></span>|
|[<span data-ttu-id="c115e-1095">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c115e-1095">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="c115e-1096">撰写</span><span class="sxs-lookup"><span data-stu-id="c115e-1096">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="c115e-1097">返回：</span><span class="sxs-lookup"><span data-stu-id="c115e-1097">Returns:</span></span>

<span data-ttu-id="c115e-1098">作为字符串的所选数据的格式由 `coercionType` 确定。</span><span class="sxs-lookup"><span data-stu-id="c115e-1098">The selected data as a string with format determined by `coercionType`.</span></span>

<dl class="param-type"><span data-ttu-id="c115e-1099">

<dt>
类型</dt>


</span><span class="sxs-lookup"><span data-stu-id="c115e-1099">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="c115e-1100">String</span><span class="sxs-lookup"><span data-stu-id="c115e-1100">String</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="c115e-1101">示例</span><span class="sxs-lookup"><span data-stu-id="c115e-1101">Example</span></span>

```js
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

#### <a name="getselectedentities--entitiesjavascriptapioutlook17officeentities"></a><span data-ttu-id="c115e-1102">getSelectedEntities() → {[Entities](/javascript/api/outlook_1_7/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="c115e-1102">getSelectedEntities() → {[Entities](/javascript/api/outlook_1_7/office.entities)}</span></span>

<span data-ttu-id="c115e-p166">获取在用户已选择的突出显示匹配项中找到的实体。突出显示匹配项适用于[上下文加载项](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins)。</span><span class="sxs-lookup"><span data-stu-id="c115e-p166">Gets the entities found in a highlighted match a user has selected. Highlighted matches apply to [contextual add-ins](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="c115e-1105">在 Outlook for iOS 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="c115e-1105">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="c115e-1106">Requirements</span><span class="sxs-lookup"><span data-stu-id="c115e-1106">Requirements</span></span>

|<span data-ttu-id="c115e-1107">要求</span><span class="sxs-lookup"><span data-stu-id="c115e-1107">Requirement</span></span>|<span data-ttu-id="c115e-1108">值</span><span class="sxs-lookup"><span data-stu-id="c115e-1108">Value</span></span>|
|---|---|
|[<span data-ttu-id="c115e-1109">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="c115e-1109">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c115e-1110">1.6</span><span class="sxs-lookup"><span data-stu-id="c115e-1110">1.6</span></span>|
|[<span data-ttu-id="c115e-1111">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="c115e-1111">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c115e-1112">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c115e-1112">ReadItem</span></span>|
|[<span data-ttu-id="c115e-1113">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c115e-1113">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="c115e-1114">阅读</span><span class="sxs-lookup"><span data-stu-id="c115e-1114">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="c115e-1115">返回：</span><span class="sxs-lookup"><span data-stu-id="c115e-1115">Returns:</span></span>

<span data-ttu-id="c115e-1116">类型：[Entities](/javascript/api/outlook_1_7/office.entities)</span><span class="sxs-lookup"><span data-stu-id="c115e-1116">Type: [Entities](/javascript/api/outlook_1_7/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="c115e-1117">示例</span><span class="sxs-lookup"><span data-stu-id="c115e-1117">Example</span></span>

<span data-ttu-id="c115e-1118">以下示例访问用户选择的突出显示匹配项中的地址实体。</span><span class="sxs-lookup"><span data-stu-id="c115e-1118">The following example accesses the addresses entities in the highlighted match selected by the user.</span></span>

```js
var contacts = Office.context.mailbox.item.getSelectedEntities().addresses;
```

#### <a name="getselectedregexmatches--object"></a><span data-ttu-id="c115e-1119">getSelectedRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="c115e-1119">getSelectedRegExMatches() → {Object}</span></span>

<span data-ttu-id="c115e-p167">返回突出显示匹配项中匹配在清单 XML 文件中定义的正则表达式的字符串值。突出显示匹配项适用于[上下文外接程序](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins)。</span><span class="sxs-lookup"><span data-stu-id="c115e-p167">Returns string values in a highlighted match that match the regular expressions defined in the manifest XML file. Highlighted matches apply to [contextual add-ins](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="c115e-1122">在 Outlook for iOS 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="c115e-1122">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="c115e-p168">`getSelectedRegExMatches` 方法返回匹配在清单 XML 文件中的每个 `ItemHasRegularExpressionMatch` 或 `ItemHasKnownEntity` 规则元素中定义的正则表达式的字符串。对于 `ItemHasRegularExpressionMatch` 规则，匹配字符串必须发生在该规则指定的项目的属性中。`PropertyName` 简单类型定义支持的属性。</span><span class="sxs-lookup"><span data-stu-id="c115e-p168">The `getSelectedRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="c115e-1126">例如，考虑一个外接程序清单具有以下 `Rule` 元素：</span><span class="sxs-lookup"><span data-stu-id="c115e-1126">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="c115e-1127">从 `getRegExMatches` 返回的对象应有两个属性：`fruits` 和 `veggies`。</span><span class="sxs-lookup"><span data-stu-id="c115e-1127">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="c115e-p169">如果在项目的正文属性上指定 `ItemHasRegularExpressionMatch` 规则，则正则表达式应进一步筛选正文，不应尝试返回该项目的整个正文。使用正则表达式（如 `.*`）获取项目的整个正文并不总是返回预期的结果。而是使用 [`Body.getAsync`](/javascript/api/outlook_1_7/office.body#getasync-coerciontype--options--callback-) 方法检索整个正文。</span><span class="sxs-lookup"><span data-stu-id="c115e-p169">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook_1_7/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="c115e-1131">Requirements</span><span class="sxs-lookup"><span data-stu-id="c115e-1131">Requirements</span></span>

|<span data-ttu-id="c115e-1132">要求</span><span class="sxs-lookup"><span data-stu-id="c115e-1132">Requirement</span></span>|<span data-ttu-id="c115e-1133">值</span><span class="sxs-lookup"><span data-stu-id="c115e-1133">Value</span></span>|
|---|---|
|[<span data-ttu-id="c115e-1134">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="c115e-1134">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c115e-1135">1.6</span><span class="sxs-lookup"><span data-stu-id="c115e-1135">1.6</span></span>|
|[<span data-ttu-id="c115e-1136">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="c115e-1136">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c115e-1137">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c115e-1137">ReadItem</span></span>|
|[<span data-ttu-id="c115e-1138">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c115e-1138">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="c115e-1139">阅读</span><span class="sxs-lookup"><span data-stu-id="c115e-1139">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="c115e-1140">返回：</span><span class="sxs-lookup"><span data-stu-id="c115e-1140">Returns:</span></span>

<span data-ttu-id="c115e-p170">一个包含匹配在清单 XML 文件中定义的正则表达式的字符串数组的对象。每个数组的名称等于匹配 `ItemHasRegularExpressionMatch` 规则的 `RegExName` 属性或匹配 `ItemHasKnownEntity` 规则的 `FilterName` 属性的相应值。</span><span class="sxs-lookup"><span data-stu-id="c115e-p170">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

##### <a name="example"></a><span data-ttu-id="c115e-1143">示例</span><span class="sxs-lookup"><span data-stu-id="c115e-1143">Example</span></span>

<span data-ttu-id="c115e-1144">以下示例显示了如何访问正则表达式规则元素 `fruits` 和 `veggies` 的匹配项的数组，这些元素在清单中指定。</span><span class="sxs-lookup"><span data-stu-id="c115e-1144">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```js
var selectedMatches = Office.context.mailbox.item.getSelectedRegExMatches();
var fruits = selectedMatches.fruits;
var veggies = selectedMatches.veggies;
```

####  <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="c115e-1145">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="c115e-1145">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="c115e-1146">异步加载所选项目上此外接程序的自定义属性。</span><span class="sxs-lookup"><span data-stu-id="c115e-1146">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="c115e-p171">自定义属性基于每个应用、每个项目存储为键/值对。此方法在回调中返回 `CustomProperties` 对象，该回调提供访问特定于当前项目和当前外接程序的自定义属性的方法。自定义属性未在项目上加密，因此这不应用作安全存储。</span><span class="sxs-lookup"><span data-stu-id="c115e-p171">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c115e-1150">参数：</span><span class="sxs-lookup"><span data-stu-id="c115e-1150">Parameters:</span></span>

|<span data-ttu-id="c115e-1151">名称</span><span class="sxs-lookup"><span data-stu-id="c115e-1151">Name</span></span>|<span data-ttu-id="c115e-1152">类型</span><span class="sxs-lookup"><span data-stu-id="c115e-1152">Type</span></span>|<span data-ttu-id="c115e-1153">属性</span><span class="sxs-lookup"><span data-stu-id="c115e-1153">Attributes</span></span>|<span data-ttu-id="c115e-1154">说明</span><span class="sxs-lookup"><span data-stu-id="c115e-1154">Description</span></span>|
|---|---|---|---|
|`callback`|<span data-ttu-id="c115e-1155">函数</span><span class="sxs-lookup"><span data-stu-id="c115e-1155">function</span></span>||<span data-ttu-id="c115e-1156">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="c115e-1156">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="c115e-1157">自定义属性作为 `asyncResult.value` 属性中的 [`CustomProperties`](/javascript/api/outlook_1_7/office.customproperties) 对象提供。</span><span class="sxs-lookup"><span data-stu-id="c115e-1157">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook_1_7/office.customproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="c115e-1158">此对象可用于获取、设置以及从项目中删除自定义属性，并将自定义属性集的更改重新保存到服务器。</span><span class="sxs-lookup"><span data-stu-id="c115e-1158">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`|<span data-ttu-id="c115e-1159">对象</span><span class="sxs-lookup"><span data-stu-id="c115e-1159">Object</span></span>|<span data-ttu-id="c115e-1160">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="c115e-1160">&lt;optional&gt;</span></span>|<span data-ttu-id="c115e-1161">开发人员可以提供他们想要在回调函数中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="c115e-1161">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="c115e-1162">此对象可以通过回调函数中的 `asyncResult.asyncContext` 属性进行访问。</span><span class="sxs-lookup"><span data-stu-id="c115e-1162">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c115e-1163">要求</span><span class="sxs-lookup"><span data-stu-id="c115e-1163">Requirements</span></span>

|<span data-ttu-id="c115e-1164">要求</span><span class="sxs-lookup"><span data-stu-id="c115e-1164">Requirement</span></span>|<span data-ttu-id="c115e-1165">值</span><span class="sxs-lookup"><span data-stu-id="c115e-1165">Value</span></span>|
|---|---|
|[<span data-ttu-id="c115e-1166">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="c115e-1166">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c115e-1167">1.0</span><span class="sxs-lookup"><span data-stu-id="c115e-1167">1.0</span></span>|
|[<span data-ttu-id="c115e-1168">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="c115e-1168">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c115e-1169">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c115e-1169">ReadItem</span></span>|
|[<span data-ttu-id="c115e-1170">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c115e-1170">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="c115e-1171">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="c115e-1171">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="c115e-1172">示例</span><span class="sxs-lookup"><span data-stu-id="c115e-1172">Example</span></span>

<span data-ttu-id="c115e-p174">以下代码示例显示了如何使用 `loadCustomPropertiesAsync` 方法异步加载特定于当前项目的自定义属性。该示例还显示了如何使用 `CustomProperties.saveAsync` 方法将这些属性重新保存到服务器。加载自定义属性后，该代码示例将使用 `CustomProperties.get` 方法读取自定义属性 `myProp`，使用 `CustomProperties.set` 方法写入自定义属性 `otherProp`，最后调用 `saveAsync` 方法保存这些自定义属性。</span><span class="sxs-lookup"><span data-stu-id="c115e-p174">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

```js
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

####  <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="c115e-1176">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="c115e-1176">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="c115e-1177">将附件从邮件或约会中删除。</span><span class="sxs-lookup"><span data-stu-id="c115e-1177">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="c115e-p175">`removeAttachmentAsync` 方法删除项目中带指定标识符的附件。最佳做法是，仅当同一个邮件应用程序在同一会话中添加了一个附件时，你才应使用该附件标识符来删除该附件。在 Outlook Web App 和适用于设备的 OWA 中，附件标识符只在同一个会话中才有效。当用户关闭应用程序，或者如果用户开始在内嵌窗体中撰写，并在随后弹出的内嵌窗体中继续在单独的窗口撰写时，会话即结束。</span><span class="sxs-lookup"><span data-stu-id="c115e-p175">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item. As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session. In Outlook Web App and OWA for Devices, the attachment identifier is valid only within the same session. A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c115e-1182">参数：</span><span class="sxs-lookup"><span data-stu-id="c115e-1182">Parameters:</span></span>

|<span data-ttu-id="c115e-1183">名称</span><span class="sxs-lookup"><span data-stu-id="c115e-1183">Name</span></span>|<span data-ttu-id="c115e-1184">类型</span><span class="sxs-lookup"><span data-stu-id="c115e-1184">Type</span></span>|<span data-ttu-id="c115e-1185">属性</span><span class="sxs-lookup"><span data-stu-id="c115e-1185">Attributes</span></span>|<span data-ttu-id="c115e-1186">说明</span><span class="sxs-lookup"><span data-stu-id="c115e-1186">Description</span></span>|
|---|---|---|---|
|`attachmentId`|<span data-ttu-id="c115e-1187">字符串</span><span class="sxs-lookup"><span data-stu-id="c115e-1187">String</span></span>||<span data-ttu-id="c115e-1188">要删除的附件的标识符。</span><span class="sxs-lookup"><span data-stu-id="c115e-1188">The identifier of the attachment to remove.</span></span>|
|`options`|<span data-ttu-id="c115e-1189">对象</span><span class="sxs-lookup"><span data-stu-id="c115e-1189">Object</span></span>|<span data-ttu-id="c115e-1190">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="c115e-1190">&lt;optional&gt;</span></span>|<span data-ttu-id="c115e-1191">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="c115e-1191">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="c115e-1192">对象</span><span class="sxs-lookup"><span data-stu-id="c115e-1192">Object</span></span>|<span data-ttu-id="c115e-1193">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="c115e-1193">&lt;optional&gt;</span></span>|<span data-ttu-id="c115e-1194">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="c115e-1194">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="c115e-1195">function</span><span class="sxs-lookup"><span data-stu-id="c115e-1195">function</span></span>|<span data-ttu-id="c115e-1196">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="c115e-1196">&lt;optional&gt;</span></span>|<span data-ttu-id="c115e-1197">方法完成后，使用单个参数 `callback`（一个 [`asyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `AsyncResult` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="c115e-1197">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="c115e-1198">如果删除附件失败，`asyncResult.error` 属性将包含一个说明失败原因的错误代码。</span><span class="sxs-lookup"><span data-stu-id="c115e-1198">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="c115e-1199">错误</span><span class="sxs-lookup"><span data-stu-id="c115e-1199">Errors</span></span>

|<span data-ttu-id="c115e-1200">错误代码</span><span class="sxs-lookup"><span data-stu-id="c115e-1200">Error code</span></span>|<span data-ttu-id="c115e-1201">说明</span><span class="sxs-lookup"><span data-stu-id="c115e-1201">Description</span></span>|
|------------|-------------|
|`InvalidAttachmentId`|<span data-ttu-id="c115e-1202">附件标识符不存在。</span><span class="sxs-lookup"><span data-stu-id="c115e-1202">The attachment identifier does not exist.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c115e-1203">Requirements</span><span class="sxs-lookup"><span data-stu-id="c115e-1203">Requirements</span></span>

|<span data-ttu-id="c115e-1204">要求</span><span class="sxs-lookup"><span data-stu-id="c115e-1204">Requirement</span></span>|<span data-ttu-id="c115e-1205">值</span><span class="sxs-lookup"><span data-stu-id="c115e-1205">Value</span></span>|
|---|---|
|[<span data-ttu-id="c115e-1206">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="c115e-1206">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c115e-1207">1.1</span><span class="sxs-lookup"><span data-stu-id="c115e-1207">1.1</span></span>|
|[<span data-ttu-id="c115e-1208">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="c115e-1208">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c115e-1209">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="c115e-1209">ReadWriteItem</span></span>|
|[<span data-ttu-id="c115e-1210">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c115e-1210">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="c115e-1211">撰写</span><span class="sxs-lookup"><span data-stu-id="c115e-1211">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="c115e-1212">示例</span><span class="sxs-lookup"><span data-stu-id="c115e-1212">Example</span></span>

<span data-ttu-id="c115e-1213">以下代码删除包含标识符 '0' 的附件。</span><span class="sxs-lookup"><span data-stu-id="c115e-1213">The following code removes an attachment with an identifier of '0'.</span></span>

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

####  <a name="removehandlerasynceventtype-options-callback"></a><span data-ttu-id="c115e-1214">removeHandlerAsync(eventType, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="c115e-1214">removeHandlerAsync(eventType, [options], [callback])</span></span>

<span data-ttu-id="c115e-1215">删除受支持事件类型的事件处理程序。</span><span class="sxs-lookup"><span data-stu-id="c115e-1215">Removes the event handlers for a supported event type.</span></span>

<span data-ttu-id="c115e-1216">当前，支持的事件类型是 `Office.EventType.AppointmentTimeChanged`、`Office.EventType.RecipientsChanged` 和 `Office.EventType.RecurrenceChanged`</span><span class="sxs-lookup"><span data-stu-id="c115e-1216">Currently the supported event types are `Office.EventType.AppointmentTimeChanged`, `Office.EventType.RecipientsChanged`, and `Office.EventType.RecurrenceChanged`</span></span>

##### <a name="parameters"></a><span data-ttu-id="c115e-1217">参数：</span><span class="sxs-lookup"><span data-stu-id="c115e-1217">Parameters:</span></span>

| <span data-ttu-id="c115e-1218">名称</span><span class="sxs-lookup"><span data-stu-id="c115e-1218">Name</span></span> | <span data-ttu-id="c115e-1219">类型</span><span class="sxs-lookup"><span data-stu-id="c115e-1219">Type</span></span> | <span data-ttu-id="c115e-1220">属性</span><span class="sxs-lookup"><span data-stu-id="c115e-1220">Attributes</span></span> | <span data-ttu-id="c115e-1221">说明</span><span class="sxs-lookup"><span data-stu-id="c115e-1221">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="c115e-1222">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="c115e-1222">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="c115e-1223">应调用处理程序的事件。</span><span class="sxs-lookup"><span data-stu-id="c115e-1223">The event that should invoke the handler.</span></span> |
| `options` | <span data-ttu-id="c115e-1224">对象</span><span class="sxs-lookup"><span data-stu-id="c115e-1224">Object</span></span> | <span data-ttu-id="c115e-1225">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="c115e-1225">&lt;optional&gt;</span></span> | <span data-ttu-id="c115e-1226">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="c115e-1226">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="c115e-1227">对象</span><span class="sxs-lookup"><span data-stu-id="c115e-1227">Object</span></span> | <span data-ttu-id="c115e-1228">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="c115e-1228">&lt;optional&gt;</span></span> | <span data-ttu-id="c115e-1229">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="c115e-1229">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="c115e-1230">函数</span><span class="sxs-lookup"><span data-stu-id="c115e-1230">function</span></span>| <span data-ttu-id="c115e-1231">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="c115e-1231">&lt;optional&gt;</span></span>|<span data-ttu-id="c115e-1232">方法完成后，使用单个参数 `callback`（一个 [`asyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `AsyncResult` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="c115e-1232">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c115e-1233">Requirements</span><span class="sxs-lookup"><span data-stu-id="c115e-1233">Requirements</span></span>

|<span data-ttu-id="c115e-1234">要求</span><span class="sxs-lookup"><span data-stu-id="c115e-1234">Requirement</span></span>| <span data-ttu-id="c115e-1235">值</span><span class="sxs-lookup"><span data-stu-id="c115e-1235">Value</span></span>|
|---|---|
|[<span data-ttu-id="c115e-1236">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="c115e-1236">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c115e-1237">1.7</span><span class="sxs-lookup"><span data-stu-id="c115e-1237">1.7</span></span> |
|[<span data-ttu-id="c115e-1238">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="c115e-1238">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c115e-1239">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c115e-1239">ReadItem</span></span> |
|[<span data-ttu-id="c115e-1240">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c115e-1240">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c115e-1241">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="c115e-1241">Compose or read</span></span> |

##### <a name="example"></a><span data-ttu-id="c115e-1242">示例</span><span class="sxs-lookup"><span data-stu-id="c115e-1242">Example</span></span>

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

####  <a name="saveasyncoptions-callback"></a><span data-ttu-id="c115e-1243">saveAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="c115e-1243">saveAsync([options], callback)</span></span>

<span data-ttu-id="c115e-1244">异步保存项目。</span><span class="sxs-lookup"><span data-stu-id="c115e-1244">Asynchronously saves an item.</span></span>

<span data-ttu-id="c115e-p176">调用时，此方法将当前邮件保存为草稿，并通过回调方法返回项目 ID。在 Outlook Web App 或 Outlook 联机模式下，该项目被保存到服务器中。在 Outlook 缓存模式下，该项目被保存到本地缓存中。</span><span class="sxs-lookup"><span data-stu-id="c115e-p176">When invoked, this method saves the current message as a draft and returns the item id via the callback method. In Outlook Web App or Outlook in online mode, the item is saved to the server. In Outlook in cached mode, the item is saved to the local cache.</span></span>

> [!NOTE]
> <span data-ttu-id="c115e-1248">如果加载项在撰写模式下对某个项目调用 `saveAsync` 来获得 `itemId`，以便与 EWS 或 REST API 一同使用，请注意，当 Outlook 处于高速缓存模式时，可能需要一段时间项目才能真正同步到服务器。</span><span class="sxs-lookup"><span data-stu-id="c115e-1248">If your add-in calls `saveAsync` on an item in compose mode in order to get an `itemId` to use with EWS or the REST API, be aware that when Outlook is in cached mode, it may take some time before the item is actually synced to the server.</span></span> <span data-ttu-id="c115e-1249">在项目同步前，使用 `itemId` 将返回一个错误。</span><span class="sxs-lookup"><span data-stu-id="c115e-1249">Until the item is synced, using the `itemId` will return an error.</span></span>

<span data-ttu-id="c115e-p178">由于约会没有草稿状态，如果以撰写模式在约会中调用 `saveAsync`，则该项将被保存为用户日历中的正常约会。对于之前未保存过的新约会，则不会发送邀请。保存现有约会将向添加或删除的与会者发送更新。</span><span class="sxs-lookup"><span data-stu-id="c115e-p178">Since appointments have no draft state, if `saveAsync` is called on an appointment in compose mode, the item will be saved as a normal appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent. Saving an existing appointment will send an update to added or removed attendees.</span></span>

> [!NOTE]
> <span data-ttu-id="c115e-1253">以下客户端在撰写模式下对约会上的 `saveAsync` 具有不同的行为：</span><span class="sxs-lookup"><span data-stu-id="c115e-1253">The following clients have different behavior for `saveAsync` on appointments in compose mode:</span></span>
>
> - <span data-ttu-id="c115e-1254">Mac Outlook 不支持在撰写模式下对会议执行 `saveAsync` 操作。</span><span class="sxs-lookup"><span data-stu-id="c115e-1254">Mac Outlook does not support `saveAsync` on a meeting in compose mode.</span></span> <span data-ttu-id="c115e-1255">对 Mac Outlook 中的会议调用 `saveAsync` 将会返回一个错误。</span><span class="sxs-lookup"><span data-stu-id="c115e-1255">Calling `saveAsync` on a meeting in Mac Outlook will return an error.</span></span>
> - <span data-ttu-id="c115e-1256">在撰写模式下的约会上调用 `saveAsync` 时，Outlook 网页版始终发送邀请或更新。</span><span class="sxs-lookup"><span data-stu-id="c115e-1256">Outlook on the web always sends an invitation or update when `saveAsync` is called on an appointment in compose mode.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c115e-1257">参数：</span><span class="sxs-lookup"><span data-stu-id="c115e-1257">Parameters:</span></span>

|<span data-ttu-id="c115e-1258">名称</span><span class="sxs-lookup"><span data-stu-id="c115e-1258">Name</span></span>|<span data-ttu-id="c115e-1259">类型</span><span class="sxs-lookup"><span data-stu-id="c115e-1259">Type</span></span>|<span data-ttu-id="c115e-1260">属性</span><span class="sxs-lookup"><span data-stu-id="c115e-1260">Attributes</span></span>|<span data-ttu-id="c115e-1261">说明</span><span class="sxs-lookup"><span data-stu-id="c115e-1261">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="c115e-1262">对象</span><span class="sxs-lookup"><span data-stu-id="c115e-1262">Object</span></span>|<span data-ttu-id="c115e-1263">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="c115e-1263">&lt;optional&gt;</span></span>|<span data-ttu-id="c115e-1264">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="c115e-1264">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="c115e-1265">对象</span><span class="sxs-lookup"><span data-stu-id="c115e-1265">Object</span></span>|<span data-ttu-id="c115e-1266">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="c115e-1266">&lt;optional&gt;</span></span>|<span data-ttu-id="c115e-1267">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="c115e-1267">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="c115e-1268">函数</span><span class="sxs-lookup"><span data-stu-id="c115e-1268">function</span></span>||<span data-ttu-id="c115e-1269">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="c115e-1269">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="c115e-1270">如果成功，该项目标识符将在 `asyncResult.value` 属性中提供。</span><span class="sxs-lookup"><span data-stu-id="c115e-1270">On success, the item identifier is provided in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c115e-1271">要求</span><span class="sxs-lookup"><span data-stu-id="c115e-1271">Requirements</span></span>

|<span data-ttu-id="c115e-1272">要求</span><span class="sxs-lookup"><span data-stu-id="c115e-1272">Requirement</span></span>|<span data-ttu-id="c115e-1273">值</span><span class="sxs-lookup"><span data-stu-id="c115e-1273">Value</span></span>|
|---|---|
|[<span data-ttu-id="c115e-1274">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="c115e-1274">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c115e-1275">1.3</span><span class="sxs-lookup"><span data-stu-id="c115e-1275">1.3</span></span>|
|[<span data-ttu-id="c115e-1276">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="c115e-1276">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c115e-1277">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="c115e-1277">ReadWriteItem</span></span>|
|[<span data-ttu-id="c115e-1278">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c115e-1278">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="c115e-1279">撰写</span><span class="sxs-lookup"><span data-stu-id="c115e-1279">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="c115e-1280">示例</span><span class="sxs-lookup"><span data-stu-id="c115e-1280">Examples</span></span>

```js
Office.context.mailbox.item.saveAsync(
  function callback(result) {
    // Process the result
  });
```

<span data-ttu-id="c115e-p180">下面是传递给回调函数的 `result` 参数的示例。`value` 属性包含的项目的项目 ID。</span><span class="sxs-lookup"><span data-stu-id="c115e-p180">The following is an example of the `result` parameter passed to the callback function. The `value` property contains the item ID of the item.</span></span>

```json
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

####  <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="c115e-1283">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="c115e-1283">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="c115e-1284">以异步方式将数据插入到邮件的正文或主题中。</span><span class="sxs-lookup"><span data-stu-id="c115e-1284">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="c115e-p181">`setSelectedDataAsync` 方法将指定字符串插入到项目主题或正文的光标位置，或者，如果在编辑器中已选择文本，则该方法将替换选择的文本。如果光标不在正文或主题字段中，则返回错误。插入之后，光标会位于插入内容的末尾。</span><span class="sxs-lookup"><span data-stu-id="c115e-p181">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c115e-1288">参数：</span><span class="sxs-lookup"><span data-stu-id="c115e-1288">Parameters:</span></span>

|<span data-ttu-id="c115e-1289">名称</span><span class="sxs-lookup"><span data-stu-id="c115e-1289">Name</span></span>|<span data-ttu-id="c115e-1290">类型</span><span class="sxs-lookup"><span data-stu-id="c115e-1290">Type</span></span>|<span data-ttu-id="c115e-1291">属性</span><span class="sxs-lookup"><span data-stu-id="c115e-1291">Attributes</span></span>|<span data-ttu-id="c115e-1292">说明</span><span class="sxs-lookup"><span data-stu-id="c115e-1292">Description</span></span>|
|---|---|---|---|
|`data`|<span data-ttu-id="c115e-1293">字符串</span><span class="sxs-lookup"><span data-stu-id="c115e-1293">String</span></span>||<span data-ttu-id="c115e-p182">要插入的数据。数据不得超过 1,000,000 个字符。如果传入的数据超过 1,000,000 个字符，则会引发 `ArgumentOutOfRange` 异常。</span><span class="sxs-lookup"><span data-stu-id="c115e-p182">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`|<span data-ttu-id="c115e-1297">Object</span><span class="sxs-lookup"><span data-stu-id="c115e-1297">Object</span></span>|<span data-ttu-id="c115e-1298">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="c115e-1298">&lt;optional&gt;</span></span>|<span data-ttu-id="c115e-1299">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="c115e-1299">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="c115e-1300">对象</span><span class="sxs-lookup"><span data-stu-id="c115e-1300">Object</span></span>|<span data-ttu-id="c115e-1301">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="c115e-1301">&lt;optional&gt;</span></span>|<span data-ttu-id="c115e-1302">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="c115e-1302">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`|[<span data-ttu-id="c115e-1303">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="c115e-1303">Office.CoercionType</span></span>](office.md#coerciontype-string)|<span data-ttu-id="c115e-1304">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="c115e-1304">&lt;optional&gt;</span></span>|<span data-ttu-id="c115e-p183">如果为 `text`，则在 Outlook Web App 和 Outlook 中应用当前样式。如果该字段是 HTML 编辑器，则仅插入文本数据，即使数据为 HTML。</span><span class="sxs-lookup"><span data-stu-id="c115e-p183">If `text`, the current style is applied in Outlook Web App and Outlook. If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="c115e-p184">如果 `html` 和该字段支持 HTML（主题不支持），则在 Outlook Web App 中应用当前样式，而在 Outlook 中应用默认样式。如果该字段是文本字段，则返回 `InvalidDataFormat` 错误。</span><span class="sxs-lookup"><span data-stu-id="c115e-p184">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook Web App and the default style is applied in Outlook. If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="c115e-1309">如果未设置 `coercionType`，则结果取决于该字段：如果该字段是 HTML，则使用 HTML；如果该字段是文本，则使用纯文本。</span><span class="sxs-lookup"><span data-stu-id="c115e-1309">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`|<span data-ttu-id="c115e-1310">function</span><span class="sxs-lookup"><span data-stu-id="c115e-1310">function</span></span>||<span data-ttu-id="c115e-1311">方法完成后，使用单个参数 `callback`（一个 [`asyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `AsyncResult` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="c115e-1311">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c115e-1312">Requirements</span><span class="sxs-lookup"><span data-stu-id="c115e-1312">Requirements</span></span>

|<span data-ttu-id="c115e-1313">要求</span><span class="sxs-lookup"><span data-stu-id="c115e-1313">Requirement</span></span>|<span data-ttu-id="c115e-1314">值</span><span class="sxs-lookup"><span data-stu-id="c115e-1314">Value</span></span>|
|---|---|
|[<span data-ttu-id="c115e-1315">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="c115e-1315">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c115e-1316">1.2</span><span class="sxs-lookup"><span data-stu-id="c115e-1316">1.2</span></span>|
|[<span data-ttu-id="c115e-1317">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="c115e-1317">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c115e-1318">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="c115e-1318">ReadWriteItem</span></span>|
|[<span data-ttu-id="c115e-1319">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c115e-1319">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="c115e-1320">撰写</span><span class="sxs-lookup"><span data-stu-id="c115e-1320">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="c115e-1321">示例</span><span class="sxs-lookup"><span data-stu-id="c115e-1321">Example</span></span>

```js
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```
