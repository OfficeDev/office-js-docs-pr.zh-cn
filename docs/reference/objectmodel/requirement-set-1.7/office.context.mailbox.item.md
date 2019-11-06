---
title: "\"Context\"-\"邮箱\"。项目-要求集1。7"
description: ''
ms.date: 11/05/2019
localization_priority: Normal
ms.openlocfilehash: 0cd498efb11f759dfb97d60565e2eb0bb95fd2f5
ms.sourcegitcommit: 21aa084875c9e07a300b3bbe8852b3e5dd163e1d
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 11/06/2019
ms.locfileid: "38001563"
---
# <a name="item"></a><span data-ttu-id="d8eac-102">item</span><span class="sxs-lookup"><span data-stu-id="d8eac-102">item</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmditem"></a><span data-ttu-id="d8eac-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span><span class="sxs-lookup"><span data-stu-id="d8eac-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span></span>

<span data-ttu-id="d8eac-p101">`item` 命名空间用于访问当前选定的邮件、会议请求或约会。可以通过使用 [itemType](#itemtype-officemailboxenumsitemtype) 属性确定 `item` 的类型。</span><span class="sxs-lookup"><span data-stu-id="d8eac-p101">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="d8eac-106">Requirements</span><span class="sxs-lookup"><span data-stu-id="d8eac-106">Requirements</span></span>

|<span data-ttu-id="d8eac-107">要求</span><span class="sxs-lookup"><span data-stu-id="d8eac-107">Requirement</span></span>|<span data-ttu-id="d8eac-108">值</span><span class="sxs-lookup"><span data-stu-id="d8eac-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="d8eac-109">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="d8eac-109">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d8eac-110">1.0</span><span class="sxs-lookup"><span data-stu-id="d8eac-110">1.0</span></span>|
|[<span data-ttu-id="d8eac-111">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="d8eac-111">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d8eac-112">受限</span><span class="sxs-lookup"><span data-stu-id="d8eac-112">Restricted</span></span>|
|[<span data-ttu-id="d8eac-113">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="d8eac-113">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="d8eac-114">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="d8eac-114">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="d8eac-115">成员和方法</span><span class="sxs-lookup"><span data-stu-id="d8eac-115">Members and methods</span></span>

| <span data-ttu-id="d8eac-116">成员</span><span class="sxs-lookup"><span data-stu-id="d8eac-116">Member</span></span> | <span data-ttu-id="d8eac-117">类型</span><span class="sxs-lookup"><span data-stu-id="d8eac-117">Type</span></span> |
|--------|------|
| [<span data-ttu-id="d8eac-118">attachments</span><span class="sxs-lookup"><span data-stu-id="d8eac-118">attachments</span></span>](#attachments-arrayattachmentdetails) | <span data-ttu-id="d8eac-119">成员</span><span class="sxs-lookup"><span data-stu-id="d8eac-119">Member</span></span> |
| [<span data-ttu-id="d8eac-120">bcc</span><span class="sxs-lookup"><span data-stu-id="d8eac-120">bcc</span></span>](#bcc-recipients) | <span data-ttu-id="d8eac-121">成员</span><span class="sxs-lookup"><span data-stu-id="d8eac-121">Member</span></span> |
| [<span data-ttu-id="d8eac-122">body</span><span class="sxs-lookup"><span data-stu-id="d8eac-122">body</span></span>](#body-body) | <span data-ttu-id="d8eac-123">成员</span><span class="sxs-lookup"><span data-stu-id="d8eac-123">Member</span></span> |
| [<span data-ttu-id="d8eac-124">cc</span><span class="sxs-lookup"><span data-stu-id="d8eac-124">cc</span></span>](#cc-arrayemailaddressdetailsrecipients) | <span data-ttu-id="d8eac-125">成员</span><span class="sxs-lookup"><span data-stu-id="d8eac-125">Member</span></span> |
| [<span data-ttu-id="d8eac-126">conversationId</span><span class="sxs-lookup"><span data-stu-id="d8eac-126">conversationId</span></span>](#nullable-conversationid-string) | <span data-ttu-id="d8eac-127">成员</span><span class="sxs-lookup"><span data-stu-id="d8eac-127">Member</span></span> |
| [<span data-ttu-id="d8eac-128">dateTimeCreated</span><span class="sxs-lookup"><span data-stu-id="d8eac-128">dateTimeCreated</span></span>](#datetimecreated-date) | <span data-ttu-id="d8eac-129">成员</span><span class="sxs-lookup"><span data-stu-id="d8eac-129">Member</span></span> |
| [<span data-ttu-id="d8eac-130">dateTimeModified</span><span class="sxs-lookup"><span data-stu-id="d8eac-130">dateTimeModified</span></span>](#datetimemodified-date) | <span data-ttu-id="d8eac-131">成员</span><span class="sxs-lookup"><span data-stu-id="d8eac-131">Member</span></span> |
| [<span data-ttu-id="d8eac-132">end</span><span class="sxs-lookup"><span data-stu-id="d8eac-132">end</span></span>](#end-datetime) | <span data-ttu-id="d8eac-133">成员</span><span class="sxs-lookup"><span data-stu-id="d8eac-133">Member</span></span> |
| [<span data-ttu-id="d8eac-134">from</span><span class="sxs-lookup"><span data-stu-id="d8eac-134">from</span></span>](#from-emailaddressdetailsfrom) | <span data-ttu-id="d8eac-135">成员</span><span class="sxs-lookup"><span data-stu-id="d8eac-135">Member</span></span> |
| [<span data-ttu-id="d8eac-136">internetMessageId</span><span class="sxs-lookup"><span data-stu-id="d8eac-136">internetMessageId</span></span>](#internetmessageid-string) | <span data-ttu-id="d8eac-137">成员</span><span class="sxs-lookup"><span data-stu-id="d8eac-137">Member</span></span> |
| [<span data-ttu-id="d8eac-138">itemClass</span><span class="sxs-lookup"><span data-stu-id="d8eac-138">itemClass</span></span>](#itemclass-string) | <span data-ttu-id="d8eac-139">成员</span><span class="sxs-lookup"><span data-stu-id="d8eac-139">Member</span></span> |
| [<span data-ttu-id="d8eac-140">itemId</span><span class="sxs-lookup"><span data-stu-id="d8eac-140">itemId</span></span>](#nullable-itemid-string) | <span data-ttu-id="d8eac-141">成员</span><span class="sxs-lookup"><span data-stu-id="d8eac-141">Member</span></span> |
| [<span data-ttu-id="d8eac-142">itemType</span><span class="sxs-lookup"><span data-stu-id="d8eac-142">itemType</span></span>](#itemtype-officemailboxenumsitemtype) | <span data-ttu-id="d8eac-143">成员</span><span class="sxs-lookup"><span data-stu-id="d8eac-143">Member</span></span> |
| [<span data-ttu-id="d8eac-144">location</span><span class="sxs-lookup"><span data-stu-id="d8eac-144">location</span></span>](#location-stringlocation) | <span data-ttu-id="d8eac-145">成员</span><span class="sxs-lookup"><span data-stu-id="d8eac-145">Member</span></span> |
| [<span data-ttu-id="d8eac-146">normalizedSubject</span><span class="sxs-lookup"><span data-stu-id="d8eac-146">normalizedSubject</span></span>](#normalizedsubject-string) | <span data-ttu-id="d8eac-147">成员</span><span class="sxs-lookup"><span data-stu-id="d8eac-147">Member</span></span> |
| [<span data-ttu-id="d8eac-148">notificationMessages</span><span class="sxs-lookup"><span data-stu-id="d8eac-148">notificationMessages</span></span>](#notificationmessages-notificationmessages) | <span data-ttu-id="d8eac-149">成员</span><span class="sxs-lookup"><span data-stu-id="d8eac-149">Member</span></span> |
| [<span data-ttu-id="d8eac-150">optionalAttendees</span><span class="sxs-lookup"><span data-stu-id="d8eac-150">optionalAttendees</span></span>](#optionalattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="d8eac-151">成员</span><span class="sxs-lookup"><span data-stu-id="d8eac-151">Member</span></span> |
| [<span data-ttu-id="d8eac-152">organizer</span><span class="sxs-lookup"><span data-stu-id="d8eac-152">organizer</span></span>](#organizer-emailaddressdetailsorganizer) | <span data-ttu-id="d8eac-153">成员</span><span class="sxs-lookup"><span data-stu-id="d8eac-153">Member</span></span> |
| [<span data-ttu-id="d8eac-154">recurrence</span><span class="sxs-lookup"><span data-stu-id="d8eac-154">recurrence</span></span>](#nullable-recurrence-recurrence) | <span data-ttu-id="d8eac-155">Member</span><span class="sxs-lookup"><span data-stu-id="d8eac-155">Member</span></span> |
| [<span data-ttu-id="d8eac-156">requiredAttendees</span><span class="sxs-lookup"><span data-stu-id="d8eac-156">requiredAttendees</span></span>](#requiredattendees-arrayemailaddressdetailsrecipients) | <span data-ttu-id="d8eac-157">成员</span><span class="sxs-lookup"><span data-stu-id="d8eac-157">Member</span></span> |
| [<span data-ttu-id="d8eac-158">sender</span><span class="sxs-lookup"><span data-stu-id="d8eac-158">sender</span></span>](#sender-emailaddressdetails) | <span data-ttu-id="d8eac-159">成员</span><span class="sxs-lookup"><span data-stu-id="d8eac-159">Member</span></span> |
| [<span data-ttu-id="d8eac-160">Webcasts&seriesid</span><span class="sxs-lookup"><span data-stu-id="d8eac-160">seriesId</span></span>](#nullable-seriesid-string) | <span data-ttu-id="d8eac-161">成员</span><span class="sxs-lookup"><span data-stu-id="d8eac-161">Member</span></span> |
| [<span data-ttu-id="d8eac-162">start</span><span class="sxs-lookup"><span data-stu-id="d8eac-162">start</span></span>](#start-datetime) | <span data-ttu-id="d8eac-163">成员</span><span class="sxs-lookup"><span data-stu-id="d8eac-163">Member</span></span> |
| [<span data-ttu-id="d8eac-164">subject</span><span class="sxs-lookup"><span data-stu-id="d8eac-164">subject</span></span>](#subject-stringsubject) | <span data-ttu-id="d8eac-165">成员</span><span class="sxs-lookup"><span data-stu-id="d8eac-165">Member</span></span> |
| [<span data-ttu-id="d8eac-166">to</span><span class="sxs-lookup"><span data-stu-id="d8eac-166">to</span></span>](#to-arrayemailaddressdetailsrecipients) | <span data-ttu-id="d8eac-167">成员</span><span class="sxs-lookup"><span data-stu-id="d8eac-167">Member</span></span> |
| [<span data-ttu-id="d8eac-168">addFileAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="d8eac-168">addFileAttachmentAsync</span></span>](#addfileattachmentasyncuri-attachmentname-options-callback) | <span data-ttu-id="d8eac-169">方法</span><span class="sxs-lookup"><span data-stu-id="d8eac-169">Method</span></span> |
| [<span data-ttu-id="d8eac-170">addHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="d8eac-170">addHandlerAsync</span></span>](#addhandlerasynceventtype-handler-options-callback) | <span data-ttu-id="d8eac-171">方法</span><span class="sxs-lookup"><span data-stu-id="d8eac-171">Method</span></span> |
| [<span data-ttu-id="d8eac-172">addItemAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="d8eac-172">addItemAttachmentAsync</span></span>](#additemattachmentasyncitemid-attachmentname-options-callback) | <span data-ttu-id="d8eac-173">方法</span><span class="sxs-lookup"><span data-stu-id="d8eac-173">Method</span></span> |
| [<span data-ttu-id="d8eac-174">close</span><span class="sxs-lookup"><span data-stu-id="d8eac-174">close</span></span>](#close) | <span data-ttu-id="d8eac-175">方法</span><span class="sxs-lookup"><span data-stu-id="d8eac-175">Method</span></span> |
| [<span data-ttu-id="d8eac-176">displayReplyAllForm</span><span class="sxs-lookup"><span data-stu-id="d8eac-176">displayReplyAllForm</span></span>](#displayreplyallformformdata-callback) | <span data-ttu-id="d8eac-177">方法</span><span class="sxs-lookup"><span data-stu-id="d8eac-177">Method</span></span> |
| [<span data-ttu-id="d8eac-178">displayReplyForm</span><span class="sxs-lookup"><span data-stu-id="d8eac-178">displayReplyForm</span></span>](#displayreplyformformdata-callback) | <span data-ttu-id="d8eac-179">方法</span><span class="sxs-lookup"><span data-stu-id="d8eac-179">Method</span></span> |
| [<span data-ttu-id="d8eac-180">getEntities</span><span class="sxs-lookup"><span data-stu-id="d8eac-180">getEntities</span></span>](#getentities--entities) | <span data-ttu-id="d8eac-181">方法</span><span class="sxs-lookup"><span data-stu-id="d8eac-181">Method</span></span> |
| [<span data-ttu-id="d8eac-182">getEntitiesByType</span><span class="sxs-lookup"><span data-stu-id="d8eac-182">getEntitiesByType</span></span>](#getentitiesbytypeentitytype--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="d8eac-183">方法</span><span class="sxs-lookup"><span data-stu-id="d8eac-183">Method</span></span> |
| [<span data-ttu-id="d8eac-184">getFilteredEntitiesByName</span><span class="sxs-lookup"><span data-stu-id="d8eac-184">getFilteredEntitiesByName</span></span>](#getfilteredentitiesbynamename--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | <span data-ttu-id="d8eac-185">方法</span><span class="sxs-lookup"><span data-stu-id="d8eac-185">Method</span></span> |
| [<span data-ttu-id="d8eac-186">getRegExMatches</span><span class="sxs-lookup"><span data-stu-id="d8eac-186">getRegExMatches</span></span>](#getregexmatches--object) | <span data-ttu-id="d8eac-187">方法</span><span class="sxs-lookup"><span data-stu-id="d8eac-187">Method</span></span> |
| [<span data-ttu-id="d8eac-188">getRegExMatchesByName</span><span class="sxs-lookup"><span data-stu-id="d8eac-188">getRegExMatchesByName</span></span>](#getregexmatchesbynamename--nullable-array-string-) | <span data-ttu-id="d8eac-189">方法</span><span class="sxs-lookup"><span data-stu-id="d8eac-189">Method</span></span> |
| [<span data-ttu-id="d8eac-190">getSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="d8eac-190">getSelectedDataAsync</span></span>](#getselecteddataasynccoerciontype-options-callback--string) | <span data-ttu-id="d8eac-191">方法</span><span class="sxs-lookup"><span data-stu-id="d8eac-191">Method</span></span> |
| [<span data-ttu-id="d8eac-192">Office.context.mailbox.item.getselectedentities</span><span class="sxs-lookup"><span data-stu-id="d8eac-192">getSelectedEntities</span></span>](#getselectedentities--entities) | <span data-ttu-id="d8eac-193">方法</span><span class="sxs-lookup"><span data-stu-id="d8eac-193">Method</span></span> |
| [<span data-ttu-id="d8eac-194">Office.context.mailbox.item.getselectedregexmatches</span><span class="sxs-lookup"><span data-stu-id="d8eac-194">getSelectedRegExMatches</span></span>](#getselectedregexmatches--object) | <span data-ttu-id="d8eac-195">Method</span><span class="sxs-lookup"><span data-stu-id="d8eac-195">Method</span></span> |
| [<span data-ttu-id="d8eac-196">loadCustomPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="d8eac-196">loadCustomPropertiesAsync</span></span>](#loadcustompropertiesasynccallback-usercontext) | <span data-ttu-id="d8eac-197">Method</span><span class="sxs-lookup"><span data-stu-id="d8eac-197">Method</span></span> |
| [<span data-ttu-id="d8eac-198">removeAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="d8eac-198">removeAttachmentAsync</span></span>](#removeattachmentasyncattachmentid-options-callback) | <span data-ttu-id="d8eac-199">方法</span><span class="sxs-lookup"><span data-stu-id="d8eac-199">Method</span></span> |
| [<span data-ttu-id="d8eac-200">removeHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="d8eac-200">removeHandlerAsync</span></span>](#removehandlerasynceventtype-options-callback) | <span data-ttu-id="d8eac-201">方法</span><span class="sxs-lookup"><span data-stu-id="d8eac-201">Method</span></span> |
| [<span data-ttu-id="d8eac-202">saveAsync</span><span class="sxs-lookup"><span data-stu-id="d8eac-202">saveAsync</span></span>](#saveasyncoptions-callback) | <span data-ttu-id="d8eac-203">Method</span><span class="sxs-lookup"><span data-stu-id="d8eac-203">Method</span></span> |
| [<span data-ttu-id="d8eac-204">setSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="d8eac-204">setSelectedDataAsync</span></span>](#setselecteddataasyncdata-options-callback) | <span data-ttu-id="d8eac-205">方法</span><span class="sxs-lookup"><span data-stu-id="d8eac-205">Method</span></span> |

### <a name="example"></a><span data-ttu-id="d8eac-206">示例</span><span class="sxs-lookup"><span data-stu-id="d8eac-206">Example</span></span>

<span data-ttu-id="d8eac-207">以下 JavaScript 代码示例显示了如何访问 Outlook 中当前项目的 `subject` 属性。</span><span class="sxs-lookup"><span data-stu-id="d8eac-207">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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

### <a name="members"></a><span data-ttu-id="d8eac-208">Members</span><span class="sxs-lookup"><span data-stu-id="d8eac-208">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetailsviewoutlook-js-17"></a><span data-ttu-id="d8eac-209">attachments: Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.7)></span><span class="sxs-lookup"><span data-stu-id="d8eac-209">attachments: Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.7)></span></span>

<span data-ttu-id="d8eac-p102">获取项目的附件数组。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="d8eac-p102">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="d8eac-212">由于潜在的安全问题，某些类型的文件会受到 Outlook 阻止，并且不会返回。</span><span class="sxs-lookup"><span data-stu-id="d8eac-212">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="d8eac-213">如需了解更多信息，请参阅 [Outlook 中阻止的附件](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519)。</span><span class="sxs-lookup"><span data-stu-id="d8eac-213">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="d8eac-214">类型</span><span class="sxs-lookup"><span data-stu-id="d8eac-214">Type</span></span>

*   <span data-ttu-id="d8eac-215">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.7)></span><span class="sxs-lookup"><span data-stu-id="d8eac-215">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.7)></span></span>

##### <a name="requirements"></a><span data-ttu-id="d8eac-216">Requirements</span><span class="sxs-lookup"><span data-stu-id="d8eac-216">Requirements</span></span>

|<span data-ttu-id="d8eac-217">要求</span><span class="sxs-lookup"><span data-stu-id="d8eac-217">Requirement</span></span>|<span data-ttu-id="d8eac-218">值</span><span class="sxs-lookup"><span data-stu-id="d8eac-218">Value</span></span>|
|---|---|
|[<span data-ttu-id="d8eac-219">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="d8eac-219">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d8eac-220">1.0</span><span class="sxs-lookup"><span data-stu-id="d8eac-220">1.0</span></span>|
|[<span data-ttu-id="d8eac-221">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="d8eac-221">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d8eac-222">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d8eac-222">ReadItem</span></span>|
|[<span data-ttu-id="d8eac-223">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="d8eac-223">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="d8eac-224">阅读</span><span class="sxs-lookup"><span data-stu-id="d8eac-224">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d8eac-225">示例</span><span class="sxs-lookup"><span data-stu-id="d8eac-225">Example</span></span>

<span data-ttu-id="d8eac-226">以下代码使用当前项目上所有附件的详细信息构成 HTML 字符串。</span><span class="sxs-lookup"><span data-stu-id="d8eac-226">The following code builds an HTML string with details of all attachments on the current item.</span></span>

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

#### <a name="bcc-recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-17"></a><span data-ttu-id="d8eac-227">bcc: [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="d8eac-227">bcc: [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.7)</span></span>

<span data-ttu-id="d8eac-228">获取对象，该对象提供用于获取或更新邮件的密件抄送 (Bcc) 行上的收件人的方法。</span><span class="sxs-lookup"><span data-stu-id="d8eac-228">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="d8eac-229">仅限撰写模式。</span><span class="sxs-lookup"><span data-stu-id="d8eac-229">Compose mode only.</span></span>

<span data-ttu-id="d8eac-230">默认情况下，集合上限为 100 个成员。</span><span class="sxs-lookup"><span data-stu-id="d8eac-230">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="d8eac-231">但是，在 Windows 和 Mac 上，以下限制适用。</span><span class="sxs-lookup"><span data-stu-id="d8eac-231">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="d8eac-232">最多包含 500 个成员。</span><span class="sxs-lookup"><span data-stu-id="d8eac-232">Get 500 members maximum.</span></span>
- <span data-ttu-id="d8eac-233">为每个呼叫最多设置 100 个成员，总共多达 500 个成员。</span><span class="sxs-lookup"><span data-stu-id="d8eac-233">Set a maximum of 100 members per call, up to 500 members total.</span></span>

##### <a name="type"></a><span data-ttu-id="d8eac-234">类型</span><span class="sxs-lookup"><span data-stu-id="d8eac-234">Type</span></span>

*   [<span data-ttu-id="d8eac-235">收件人</span><span class="sxs-lookup"><span data-stu-id="d8eac-235">Recipients</span></span>](/javascript/api/outlook/office.recipients?view=outlook-js-1.7)

##### <a name="requirements"></a><span data-ttu-id="d8eac-236">Requirements</span><span class="sxs-lookup"><span data-stu-id="d8eac-236">Requirements</span></span>

|<span data-ttu-id="d8eac-237">要求</span><span class="sxs-lookup"><span data-stu-id="d8eac-237">Requirement</span></span>|<span data-ttu-id="d8eac-238">值</span><span class="sxs-lookup"><span data-stu-id="d8eac-238">Value</span></span>|
|---|---|
|[<span data-ttu-id="d8eac-239">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="d8eac-239">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d8eac-240">1.1</span><span class="sxs-lookup"><span data-stu-id="d8eac-240">1.1</span></span>|
|[<span data-ttu-id="d8eac-241">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="d8eac-241">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d8eac-242">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d8eac-242">ReadItem</span></span>|
|[<span data-ttu-id="d8eac-243">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="d8eac-243">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="d8eac-244">撰写</span><span class="sxs-lookup"><span data-stu-id="d8eac-244">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="d8eac-245">示例</span><span class="sxs-lookup"><span data-stu-id="d8eac-245">Example</span></span>

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

#### <a name="body-bodyjavascriptapioutlookofficebodyviewoutlook-js-17"></a><span data-ttu-id="d8eac-246">body: [Body](/javascript/api/outlook/office.body?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="d8eac-246">body: [Body](/javascript/api/outlook/office.body?view=outlook-js-1.7)</span></span>

<span data-ttu-id="d8eac-247">获取一个提供用于处理项目正文的方法的对象。</span><span class="sxs-lookup"><span data-stu-id="d8eac-247">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="d8eac-248">类型</span><span class="sxs-lookup"><span data-stu-id="d8eac-248">Type</span></span>

*   [<span data-ttu-id="d8eac-249">Body</span><span class="sxs-lookup"><span data-stu-id="d8eac-249">Body</span></span>](/javascript/api/outlook/office.body?view=outlook-js-1.7)

##### <a name="requirements"></a><span data-ttu-id="d8eac-250">要求</span><span class="sxs-lookup"><span data-stu-id="d8eac-250">Requirements</span></span>

|<span data-ttu-id="d8eac-251">要求</span><span class="sxs-lookup"><span data-stu-id="d8eac-251">Requirement</span></span>|<span data-ttu-id="d8eac-252">值</span><span class="sxs-lookup"><span data-stu-id="d8eac-252">Value</span></span>|
|---|---|
|[<span data-ttu-id="d8eac-253">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="d8eac-253">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d8eac-254">1.1</span><span class="sxs-lookup"><span data-stu-id="d8eac-254">1.1</span></span>|
|[<span data-ttu-id="d8eac-255">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="d8eac-255">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d8eac-256">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d8eac-256">ReadItem</span></span>|
|[<span data-ttu-id="d8eac-257">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="d8eac-257">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="d8eac-258">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="d8eac-258">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d8eac-259">示例</span><span class="sxs-lookup"><span data-stu-id="d8eac-259">Example</span></span>

<span data-ttu-id="d8eac-260">本示例获取纯文本格式的邮件正文。</span><span class="sxs-lookup"><span data-stu-id="d8eac-260">This example gets the body of the message in plain text.</span></span>

```js
Office.context.mailbox.item.body.getAsync(
  "text",
  { asyncContext: "This is passed to the callback" },
  function callback(result) {
    // Do something with the result.
  });

```

<span data-ttu-id="d8eac-261">以下是传递到回调函数的结果参数的示例。</span><span class="sxs-lookup"><span data-stu-id="d8eac-261">The following is an example of the result parameter passed to the callback function.</span></span>

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

#### <a name="cc-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-17recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-17"></a><span data-ttu-id="d8eac-262">cc: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="d8eac-262">cc: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.7)</span></span>

<span data-ttu-id="d8eac-263">提供对邮件的抄送 (Cc) 收件人的访问权限。</span><span class="sxs-lookup"><span data-stu-id="d8eac-263">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="d8eac-264">对象的类型和访问级别取决于当前项目的模式。</span><span class="sxs-lookup"><span data-stu-id="d8eac-264">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="d8eac-265">阅读模式</span><span class="sxs-lookup"><span data-stu-id="d8eac-265">Read mode</span></span>

<span data-ttu-id="d8eac-266">`cc` 属性返回包含邮件的**抄送**行上所列的每个收件人的 `EmailAddressDetails` 对象的数组。</span><span class="sxs-lookup"><span data-stu-id="d8eac-266">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message.</span></span> <span data-ttu-id="d8eac-267">默认情况下，集合上限为 100 个成员。</span><span class="sxs-lookup"><span data-stu-id="d8eac-267">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="d8eac-268">但是，在 Windows 和 Mac 上，最多可包含 500 个成员。</span><span class="sxs-lookup"><span data-stu-id="d8eac-268">However, on Windows and Mac, you can get 500 members maximum.</span></span>

```js
console.log(JSON.stringify(Office.context.mailbox.item.cc));
```

##### <a name="compose-mode"></a><span data-ttu-id="d8eac-269">撰写模式</span><span class="sxs-lookup"><span data-stu-id="d8eac-269">Compose mode</span></span>

<span data-ttu-id="d8eac-270">`cc` 属性返回一个 `Recipients` 对象，该对象提供用于获取或更新邮件的“**抄送**”行上收件人的方法。</span><span class="sxs-lookup"><span data-stu-id="d8eac-270">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span> <span data-ttu-id="d8eac-271">默认情况下，集合上限为 100 个成员。</span><span class="sxs-lookup"><span data-stu-id="d8eac-271">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="d8eac-272">但是，在 Windows 和 Mac 上，以下限制适用。</span><span class="sxs-lookup"><span data-stu-id="d8eac-272">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="d8eac-273">最多包含 500 个成员。</span><span class="sxs-lookup"><span data-stu-id="d8eac-273">Get 500 members maximum.</span></span>
- <span data-ttu-id="d8eac-274">为每个呼叫最多设置 100 个成员，总共多达 500 个成员。</span><span class="sxs-lookup"><span data-stu-id="d8eac-274">Set a maximum of 100 members per call, up to 500 members total.</span></span>

```js
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="d8eac-275">类型</span><span class="sxs-lookup"><span data-stu-id="d8eac-275">Type</span></span>

*   <span data-ttu-id="d8eac-276">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="d8eac-276">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.7)</span></span>

##### <a name="requirements"></a><span data-ttu-id="d8eac-277">Requirements</span><span class="sxs-lookup"><span data-stu-id="d8eac-277">Requirements</span></span>

|<span data-ttu-id="d8eac-278">要求</span><span class="sxs-lookup"><span data-stu-id="d8eac-278">Requirement</span></span>|<span data-ttu-id="d8eac-279">值</span><span class="sxs-lookup"><span data-stu-id="d8eac-279">Value</span></span>|
|---|---|
|[<span data-ttu-id="d8eac-280">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="d8eac-280">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d8eac-281">1.0</span><span class="sxs-lookup"><span data-stu-id="d8eac-281">1.0</span></span>|
|[<span data-ttu-id="d8eac-282">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="d8eac-282">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d8eac-283">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d8eac-283">ReadItem</span></span>|
|[<span data-ttu-id="d8eac-284">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="d8eac-284">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="d8eac-285">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="d8eac-285">Compose or Read</span></span>|

<br>

---
---

#### <a name="nullable-conversationid-string"></a><span data-ttu-id="d8eac-286">(nullable) conversationId: String</span><span class="sxs-lookup"><span data-stu-id="d8eac-286">(nullable) conversationId: String</span></span>

<span data-ttu-id="d8eac-287">获取包含特定消息的电子邮件会话的标识符。</span><span class="sxs-lookup"><span data-stu-id="d8eac-287">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="d8eac-p109">如果在阅读窗体或撰写窗体的回复中激活邮件应用程序，则此属性可以获得一个整数值。如果用户随后更改了回复邮件的主题（若发送回复），则该邮件的对话 ID 将改变且之前获取的值将不适用。</span><span class="sxs-lookup"><span data-stu-id="d8eac-p109">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="d8eac-p110">对于撰写窗体的新项目，此属性获得一个 null 值。如果用户设置一个主题并保存该项目，`conversationId` 属性将返回一个值。</span><span class="sxs-lookup"><span data-stu-id="d8eac-p110">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="d8eac-292">Type</span><span class="sxs-lookup"><span data-stu-id="d8eac-292">Type</span></span>

*   <span data-ttu-id="d8eac-293">String</span><span class="sxs-lookup"><span data-stu-id="d8eac-293">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="d8eac-294">Requirements</span><span class="sxs-lookup"><span data-stu-id="d8eac-294">Requirements</span></span>

|<span data-ttu-id="d8eac-295">要求</span><span class="sxs-lookup"><span data-stu-id="d8eac-295">Requirement</span></span>|<span data-ttu-id="d8eac-296">值</span><span class="sxs-lookup"><span data-stu-id="d8eac-296">Value</span></span>|
|---|---|
|[<span data-ttu-id="d8eac-297">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="d8eac-297">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d8eac-298">1.0</span><span class="sxs-lookup"><span data-stu-id="d8eac-298">1.0</span></span>|
|[<span data-ttu-id="d8eac-299">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="d8eac-299">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d8eac-300">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d8eac-300">ReadItem</span></span>|
|[<span data-ttu-id="d8eac-301">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="d8eac-301">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="d8eac-302">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="d8eac-302">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d8eac-303">示例</span><span class="sxs-lookup"><span data-stu-id="d8eac-303">Example</span></span>

```js
var conversationId = Office.context.mailbox.item.conversationId;
console.log("conversationId: " + conversationId);
```

<br>

---
---

#### <a name="datetimecreated-date"></a><span data-ttu-id="d8eac-304">dateTimeCreated: Date</span><span class="sxs-lookup"><span data-stu-id="d8eac-304">dateTimeCreated: Date</span></span>

<span data-ttu-id="d8eac-p111">获取项目创建的日期和时间。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="d8eac-p111">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="d8eac-307">类型</span><span class="sxs-lookup"><span data-stu-id="d8eac-307">Type</span></span>

*   <span data-ttu-id="d8eac-308">日期</span><span class="sxs-lookup"><span data-stu-id="d8eac-308">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="d8eac-309">Requirements</span><span class="sxs-lookup"><span data-stu-id="d8eac-309">Requirements</span></span>

|<span data-ttu-id="d8eac-310">要求</span><span class="sxs-lookup"><span data-stu-id="d8eac-310">Requirement</span></span>|<span data-ttu-id="d8eac-311">值</span><span class="sxs-lookup"><span data-stu-id="d8eac-311">Value</span></span>|
|---|---|
|[<span data-ttu-id="d8eac-312">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="d8eac-312">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d8eac-313">1.0</span><span class="sxs-lookup"><span data-stu-id="d8eac-313">1.0</span></span>|
|[<span data-ttu-id="d8eac-314">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="d8eac-314">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d8eac-315">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d8eac-315">ReadItem</span></span>|
|[<span data-ttu-id="d8eac-316">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="d8eac-316">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="d8eac-317">阅读</span><span class="sxs-lookup"><span data-stu-id="d8eac-317">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d8eac-318">示例</span><span class="sxs-lookup"><span data-stu-id="d8eac-318">Example</span></span>

```js
var dateTimeCreated = Office.context.mailbox.item.dateTimeCreated;
console.log("Date and time created: " + dateTimeCreated);
```

<br>

---
---

#### <a name="datetimemodified-date"></a><span data-ttu-id="d8eac-319">dateTimeModified: Date</span><span class="sxs-lookup"><span data-stu-id="d8eac-319">dateTimeModified: Date</span></span>

<span data-ttu-id="d8eac-p112">获取项目最近一次修改的日期和时间。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="d8eac-p112">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="d8eac-322">iOS 版 Outlook 或 Android 版 Outlook 不支持此成员。</span><span class="sxs-lookup"><span data-stu-id="d8eac-322">This member is not supported in Outlook on iOS or Android.</span></span>

##### <a name="type"></a><span data-ttu-id="d8eac-323">类型</span><span class="sxs-lookup"><span data-stu-id="d8eac-323">Type</span></span>

*   <span data-ttu-id="d8eac-324">日期</span><span class="sxs-lookup"><span data-stu-id="d8eac-324">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="d8eac-325">Requirements</span><span class="sxs-lookup"><span data-stu-id="d8eac-325">Requirements</span></span>

|<span data-ttu-id="d8eac-326">要求</span><span class="sxs-lookup"><span data-stu-id="d8eac-326">Requirement</span></span>|<span data-ttu-id="d8eac-327">值</span><span class="sxs-lookup"><span data-stu-id="d8eac-327">Value</span></span>|
|---|---|
|[<span data-ttu-id="d8eac-328">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="d8eac-328">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d8eac-329">1.0</span><span class="sxs-lookup"><span data-stu-id="d8eac-329">1.0</span></span>|
|[<span data-ttu-id="d8eac-330">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="d8eac-330">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d8eac-331">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d8eac-331">ReadItem</span></span>|
|[<span data-ttu-id="d8eac-332">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="d8eac-332">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="d8eac-333">阅读</span><span class="sxs-lookup"><span data-stu-id="d8eac-333">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d8eac-334">示例</span><span class="sxs-lookup"><span data-stu-id="d8eac-334">Example</span></span>

```js
var dateTimeModified = Office.context.mailbox.item.dateTimeModified;
console.log("Date and time modified: " + dateTimeModified);
```

<br>

---
---

#### <a name="end-datetimejavascriptapioutlookofficetimeviewoutlook-js-17"></a><span data-ttu-id="d8eac-335">end: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="d8eac-335">end: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.7)</span></span>

<span data-ttu-id="d8eac-336">获取或设置约会结束的日期和时间。</span><span class="sxs-lookup"><span data-stu-id="d8eac-336">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="d8eac-p113">将 `end` 属性表示为协调世界时 (UTC) 的日期和时间值。可使用 [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) 方法将 end 属性值转换为客户端的本地日期和时间。</span><span class="sxs-lookup"><span data-stu-id="d8eac-p113">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="d8eac-339">阅读模式</span><span class="sxs-lookup"><span data-stu-id="d8eac-339">Read mode</span></span>

<span data-ttu-id="d8eac-340">`end` 属性返回 `Date` 对象。</span><span class="sxs-lookup"><span data-stu-id="d8eac-340">The `end` property returns a `Date` object.</span></span>

```js
var end = Office.context.mailbox.item.end;
console.log("Appointment end: " + end);
```

##### <a name="compose-mode"></a><span data-ttu-id="d8eac-341">撰写模式</span><span class="sxs-lookup"><span data-stu-id="d8eac-341">Compose mode</span></span>

<span data-ttu-id="d8eac-342">`end` 属性返回 `Time` 对象。</span><span class="sxs-lookup"><span data-stu-id="d8eac-342">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="d8eac-343">使用 [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.7#setasync-datetime--options--callback-) 方法设置结束时间时，应使用 [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) 方法将客户端的本地时间转换为服务器的 UTC。</span><span class="sxs-lookup"><span data-stu-id="d8eac-343">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.7#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="d8eac-344">以下示例使用 `Time` 对象的 [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.7#setasync-datetime--options--callback-) 方法设置约会的结束时间。</span><span class="sxs-lookup"><span data-stu-id="d8eac-344">The following example sets the end time of an appointment by using the [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.7#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="d8eac-345">类型</span><span class="sxs-lookup"><span data-stu-id="d8eac-345">Type</span></span>

*   <span data-ttu-id="d8eac-346">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="d8eac-346">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.7)</span></span>

##### <a name="requirements"></a><span data-ttu-id="d8eac-347">要求</span><span class="sxs-lookup"><span data-stu-id="d8eac-347">Requirements</span></span>

|<span data-ttu-id="d8eac-348">要求</span><span class="sxs-lookup"><span data-stu-id="d8eac-348">Requirement</span></span>|<span data-ttu-id="d8eac-349">值</span><span class="sxs-lookup"><span data-stu-id="d8eac-349">Value</span></span>|
|---|---|
|[<span data-ttu-id="d8eac-350">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="d8eac-350">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d8eac-351">1.0</span><span class="sxs-lookup"><span data-stu-id="d8eac-351">1.0</span></span>|
|[<span data-ttu-id="d8eac-352">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="d8eac-352">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d8eac-353">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d8eac-353">ReadItem</span></span>|
|[<span data-ttu-id="d8eac-354">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="d8eac-354">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="d8eac-355">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="d8eac-355">Compose or Read</span></span>|

<br>

---
---

#### <a name="from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-17fromjavascriptapioutlookofficefromviewoutlook-js-17"></a><span data-ttu-id="d8eac-356">发件人： [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)|[From](/javascript/api/outlook/office.from?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="d8eac-356">from: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)|[From](/javascript/api/outlook/office.from?view=outlook-js-1.7)</span></span>

<span data-ttu-id="d8eac-357">获取邮件发件人的电子邮件地址。</span><span class="sxs-lookup"><span data-stu-id="d8eac-357">Gets the email address of the sender of a message.</span></span>

<span data-ttu-id="d8eac-p114">`from` 和 [`sender`](#sender-emailaddressdetails) 属性表示同一个人，邮件由代理人发送的除外。在此情况下，`from` 属性表示代理程序，而 sender 属性表示代理人。</span><span class="sxs-lookup"><span data-stu-id="d8eac-p114">The `from` and [`sender`](#sender-emailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="d8eac-360">`from` 属性中 `EmailAddressDetails` 对象的 `recipientType` 属性为 `undefined`。</span><span class="sxs-lookup"><span data-stu-id="d8eac-360">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="d8eac-361">阅读模式</span><span class="sxs-lookup"><span data-stu-id="d8eac-361">Read mode</span></span>

<span data-ttu-id="d8eac-362">`from`属性返回一个`EmailAddressDetails`对象。</span><span class="sxs-lookup"><span data-stu-id="d8eac-362">The `from` property returns an `EmailAddressDetails` object.</span></span>

```js
var from = Office.context.mailbox.item.from;
console.log("From " + from);
```

##### <a name="compose-mode"></a><span data-ttu-id="d8eac-363">撰写模式</span><span class="sxs-lookup"><span data-stu-id="d8eac-363">Compose mode</span></span>

<span data-ttu-id="d8eac-364">`from`属性返回一个`From`对象，该对象提供用于获取 "起始" 值的方法。</span><span class="sxs-lookup"><span data-stu-id="d8eac-364">The `from` property returns a `From` object that provides a method to get the from value.</span></span>

```js
Office.context.mailbox.item.from.getAsync(callback);

function callback(asyncResult) {
  var from = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="d8eac-365">类型</span><span class="sxs-lookup"><span data-stu-id="d8eac-365">Type</span></span>

*   <span data-ttu-id="d8eac-366">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7) | [](/javascript/api/outlook/office.from?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="d8eac-366">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7) | [From](/javascript/api/outlook/office.from?view=outlook-js-1.7)</span></span>

##### <a name="requirements"></a><span data-ttu-id="d8eac-367">Requirements</span><span class="sxs-lookup"><span data-stu-id="d8eac-367">Requirements</span></span>

|<span data-ttu-id="d8eac-368">要求</span><span class="sxs-lookup"><span data-stu-id="d8eac-368">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="d8eac-369">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="d8eac-369">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d8eac-370">1.0</span><span class="sxs-lookup"><span data-stu-id="d8eac-370">1.0</span></span>|<span data-ttu-id="d8eac-371">1.7</span><span class="sxs-lookup"><span data-stu-id="d8eac-371">1.7</span></span>|
|[<span data-ttu-id="d8eac-372">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="d8eac-372">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d8eac-373">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d8eac-373">ReadItem</span></span>|<span data-ttu-id="d8eac-374">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="d8eac-374">ReadWriteItem</span></span>|
|[<span data-ttu-id="d8eac-375">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="d8eac-375">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="d8eac-376">阅读</span><span class="sxs-lookup"><span data-stu-id="d8eac-376">Read</span></span>|<span data-ttu-id="d8eac-377">撰写</span><span class="sxs-lookup"><span data-stu-id="d8eac-377">Compose</span></span>|

<br>

---
---

#### <a name="internetmessageid-string"></a><span data-ttu-id="d8eac-378">internetMessageId: String</span><span class="sxs-lookup"><span data-stu-id="d8eac-378">internetMessageId: String</span></span>

<span data-ttu-id="d8eac-p115">获取电子邮件的 Internet 消息标识符。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="d8eac-p115">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="d8eac-381">Type</span><span class="sxs-lookup"><span data-stu-id="d8eac-381">Type</span></span>

*   <span data-ttu-id="d8eac-382">String</span><span class="sxs-lookup"><span data-stu-id="d8eac-382">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="d8eac-383">Requirements</span><span class="sxs-lookup"><span data-stu-id="d8eac-383">Requirements</span></span>

|<span data-ttu-id="d8eac-384">要求</span><span class="sxs-lookup"><span data-stu-id="d8eac-384">Requirement</span></span>|<span data-ttu-id="d8eac-385">值</span><span class="sxs-lookup"><span data-stu-id="d8eac-385">Value</span></span>|
|---|---|
|[<span data-ttu-id="d8eac-386">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="d8eac-386">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d8eac-387">1.0</span><span class="sxs-lookup"><span data-stu-id="d8eac-387">1.0</span></span>|
|[<span data-ttu-id="d8eac-388">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="d8eac-388">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d8eac-389">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d8eac-389">ReadItem</span></span>|
|[<span data-ttu-id="d8eac-390">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="d8eac-390">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="d8eac-391">阅读</span><span class="sxs-lookup"><span data-stu-id="d8eac-391">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d8eac-392">示例</span><span class="sxs-lookup"><span data-stu-id="d8eac-392">Example</span></span>

```js
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

<br>

---
---

#### <a name="itemclass-string"></a><span data-ttu-id="d8eac-393">itemClass: String</span><span class="sxs-lookup"><span data-stu-id="d8eac-393">itemClass: String</span></span>

<span data-ttu-id="d8eac-p116">获取选定项目的 Exchange Web 服务项目类。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="d8eac-p116">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="d8eac-p117">`itemClass` 属性指定所选项目的邮件类别。以下是邮件或约会项目的默认邮件类别。</span><span class="sxs-lookup"><span data-stu-id="d8eac-p117">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

|<span data-ttu-id="d8eac-398">类型</span><span class="sxs-lookup"><span data-stu-id="d8eac-398">Type</span></span>|<span data-ttu-id="d8eac-399">说明</span><span class="sxs-lookup"><span data-stu-id="d8eac-399">Description</span></span>|<span data-ttu-id="d8eac-400">项目类</span><span class="sxs-lookup"><span data-stu-id="d8eac-400">item class</span></span>|
|---|---|---|
|<span data-ttu-id="d8eac-401">约会项目</span><span class="sxs-lookup"><span data-stu-id="d8eac-401">Appointment items</span></span>|<span data-ttu-id="d8eac-402">这些是项目类 `IPM.Appointment` 或 `IPM.Appointment.Occurrence` 的日历项目。</span><span class="sxs-lookup"><span data-stu-id="d8eac-402">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurrence`.</span></span>|`IPM.Appointment`<br />`IPM.Appointment.Occurrence`|
|<span data-ttu-id="d8eac-403">邮件项目</span><span class="sxs-lookup"><span data-stu-id="d8eac-403">Message items</span></span>|<span data-ttu-id="d8eac-404">这些项目包括具有默认邮件类别 `IPM.Note` 的电子邮件，以及使用 `IPM.Schedule.Meeting` 作为基础邮件类别的会议请求、响应和取消。</span><span class="sxs-lookup"><span data-stu-id="d8eac-404">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span>|`IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled`|

<span data-ttu-id="d8eac-405">你可以创建用于扩展默认邮件类别的自定义邮件类别，例如，自定义约会邮件类别 `IPM.Appointment.Contoso`。</span><span class="sxs-lookup"><span data-stu-id="d8eac-405">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="d8eac-406">Type</span><span class="sxs-lookup"><span data-stu-id="d8eac-406">Type</span></span>

*   <span data-ttu-id="d8eac-407">String</span><span class="sxs-lookup"><span data-stu-id="d8eac-407">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="d8eac-408">Requirements</span><span class="sxs-lookup"><span data-stu-id="d8eac-408">Requirements</span></span>

|<span data-ttu-id="d8eac-409">要求</span><span class="sxs-lookup"><span data-stu-id="d8eac-409">Requirement</span></span>|<span data-ttu-id="d8eac-410">值</span><span class="sxs-lookup"><span data-stu-id="d8eac-410">Value</span></span>|
|---|---|
|[<span data-ttu-id="d8eac-411">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="d8eac-411">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d8eac-412">1.0</span><span class="sxs-lookup"><span data-stu-id="d8eac-412">1.0</span></span>|
|[<span data-ttu-id="d8eac-413">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="d8eac-413">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d8eac-414">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d8eac-414">ReadItem</span></span>|
|[<span data-ttu-id="d8eac-415">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="d8eac-415">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="d8eac-416">阅读</span><span class="sxs-lookup"><span data-stu-id="d8eac-416">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d8eac-417">示例</span><span class="sxs-lookup"><span data-stu-id="d8eac-417">Example</span></span>

```js
var itemClass = Office.context.mailbox.item.itemClass;
console.log("Item class: " + itemClass);
```

<br>

---
---

#### <a name="nullable-itemid-string"></a><span data-ttu-id="d8eac-418">(nullable) itemId: String</span><span class="sxs-lookup"><span data-stu-id="d8eac-418">(nullable) itemId: String</span></span>

<span data-ttu-id="d8eac-419">获取当前项的[Exchange Web 服务项标识符](/exchange/client-developer/exchange-web-services/ews-identifiers-in-exchange)。</span><span class="sxs-lookup"><span data-stu-id="d8eac-419">Gets the [Exchange Web Services item identifier](/exchange/client-developer/exchange-web-services/ews-identifiers-in-exchange) for the current item.</span></span> <span data-ttu-id="d8eac-420">仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="d8eac-420">Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="d8eac-421">`itemId`属性返回的标识符与[Exchange Web 服务项目标识符](/exchange/client-developer/exchange-web-services/ews-identifiers-in-exchange)相同。</span><span class="sxs-lookup"><span data-stu-id="d8eac-421">The identifier returned by the `itemId` property is the same as the [Exchange Web Services item identifier](/exchange/client-developer/exchange-web-services/ews-identifiers-in-exchange).</span></span> <span data-ttu-id="d8eac-422">`itemId` 属性与 Outlook 条目 ID 或 Outlook REST API 使用的 ID 不同。</span><span class="sxs-lookup"><span data-stu-id="d8eac-422">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="d8eac-423">使用此值进行 REST API 调用前，应使用 [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string) 对它进行转换。</span><span class="sxs-lookup"><span data-stu-id="d8eac-423">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="d8eac-424">有关详细信息，请参阅[从 Outlook 加载项使用 Outlook REST API](/outlook/add-ins/use-rest-api#get-the-item-id)。</span><span class="sxs-lookup"><span data-stu-id="d8eac-424">For more details, see [Use the Outlook REST APIs from an Outlook add-in](/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

<span data-ttu-id="d8eac-p120">`itemId` 属性在撰写模式下不可用。如果需要项目标识符，[`saveAsync`](#saveasyncoptions-callback) 方法可用于将项目保存到存储，这将在回调函数的 [`AsyncResult.value`](/javascript/api/office/office.asyncresult) 参数中返回项目标识符。</span><span class="sxs-lookup"><span data-stu-id="d8eac-p120">The `itemId` property is not available in compose mode. If an item identifier is required, the [`saveAsync`](#saveasyncoptions-callback) method can be used to save the item to the store, which will return the item identifier in the [`AsyncResult.value`](/javascript/api/office/office.asyncresult) parameter in the callback function.</span></span>

##### <a name="type"></a><span data-ttu-id="d8eac-427">Type</span><span class="sxs-lookup"><span data-stu-id="d8eac-427">Type</span></span>

*   <span data-ttu-id="d8eac-428">String</span><span class="sxs-lookup"><span data-stu-id="d8eac-428">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="d8eac-429">Requirements</span><span class="sxs-lookup"><span data-stu-id="d8eac-429">Requirements</span></span>

|<span data-ttu-id="d8eac-430">要求</span><span class="sxs-lookup"><span data-stu-id="d8eac-430">Requirement</span></span>|<span data-ttu-id="d8eac-431">值</span><span class="sxs-lookup"><span data-stu-id="d8eac-431">Value</span></span>|
|---|---|
|[<span data-ttu-id="d8eac-432">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="d8eac-432">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d8eac-433">1.0</span><span class="sxs-lookup"><span data-stu-id="d8eac-433">1.0</span></span>|
|[<span data-ttu-id="d8eac-434">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="d8eac-434">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d8eac-435">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d8eac-435">ReadItem</span></span>|
|[<span data-ttu-id="d8eac-436">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="d8eac-436">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="d8eac-437">阅读</span><span class="sxs-lookup"><span data-stu-id="d8eac-437">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d8eac-438">示例</span><span class="sxs-lookup"><span data-stu-id="d8eac-438">Example</span></span>

<span data-ttu-id="d8eac-p121">以下代码检查项目标识符是否存在。如果 `itemId` 属性返回 `null` 或 `undefined`，则将项目保存到存储，并从异步结果中获取项目标识符。</span><span class="sxs-lookup"><span data-stu-id="d8eac-p121">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

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

#### <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtypeviewoutlook-js-17"></a><span data-ttu-id="d8eac-441">itemType: [Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="d8eac-441">itemType: [Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.7)</span></span>

<span data-ttu-id="d8eac-442">获取实例表示的项的类型。</span><span class="sxs-lookup"><span data-stu-id="d8eac-442">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="d8eac-443">`itemType` 属性返回其中一个 `ItemType` 枚举值，指示 `item` 对象实例是邮件还是约会。</span><span class="sxs-lookup"><span data-stu-id="d8eac-443">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="d8eac-444">类型</span><span class="sxs-lookup"><span data-stu-id="d8eac-444">Type</span></span>

*   [<span data-ttu-id="d8eac-445">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="d8eac-445">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.7)

##### <a name="requirements"></a><span data-ttu-id="d8eac-446">Requirements</span><span class="sxs-lookup"><span data-stu-id="d8eac-446">Requirements</span></span>

|<span data-ttu-id="d8eac-447">要求</span><span class="sxs-lookup"><span data-stu-id="d8eac-447">Requirement</span></span>|<span data-ttu-id="d8eac-448">值</span><span class="sxs-lookup"><span data-stu-id="d8eac-448">Value</span></span>|
|---|---|
|[<span data-ttu-id="d8eac-449">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="d8eac-449">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d8eac-450">1.0</span><span class="sxs-lookup"><span data-stu-id="d8eac-450">1.0</span></span>|
|[<span data-ttu-id="d8eac-451">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="d8eac-451">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d8eac-452">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d8eac-452">ReadItem</span></span>|
|[<span data-ttu-id="d8eac-453">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="d8eac-453">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="d8eac-454">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="d8eac-454">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d8eac-455">示例</span><span class="sxs-lookup"><span data-stu-id="d8eac-455">Example</span></span>

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

#### <a name="location-stringlocationjavascriptapioutlookofficelocationviewoutlook-js-17"></a><span data-ttu-id="d8eac-456">location: String|[Location](/javascript/api/outlook/office.location?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="d8eac-456">location: String|[Location](/javascript/api/outlook/office.location?view=outlook-js-1.7)</span></span>

<span data-ttu-id="d8eac-457">获取或设置约会的位置。</span><span class="sxs-lookup"><span data-stu-id="d8eac-457">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="d8eac-458">阅读模式</span><span class="sxs-lookup"><span data-stu-id="d8eac-458">Read mode</span></span>

<span data-ttu-id="d8eac-459">`location` 属性返回一个包含约会位置的字符串。</span><span class="sxs-lookup"><span data-stu-id="d8eac-459">The `location` property returns a string that contains the location of the appointment.</span></span>

```js
var location = Office.context.mailbox.item.location;
console.log("location: " + location);
```

##### <a name="compose-mode"></a><span data-ttu-id="d8eac-460">撰写模式</span><span class="sxs-lookup"><span data-stu-id="d8eac-460">Compose mode</span></span>

<span data-ttu-id="d8eac-461">`location` 属性返回一个 `Location` 对象，该对象提供用于获取和设置约会位置的方法。</span><span class="sxs-lookup"><span data-stu-id="d8eac-461">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

```js
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="d8eac-462">类型</span><span class="sxs-lookup"><span data-stu-id="d8eac-462">Type</span></span>

*   <span data-ttu-id="d8eac-463">String | [Location](/javascript/api/outlook/office.location?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="d8eac-463">String | [Location](/javascript/api/outlook/office.location?view=outlook-js-1.7)</span></span>

##### <a name="requirements"></a><span data-ttu-id="d8eac-464">Requirements</span><span class="sxs-lookup"><span data-stu-id="d8eac-464">Requirements</span></span>

|<span data-ttu-id="d8eac-465">要求</span><span class="sxs-lookup"><span data-stu-id="d8eac-465">Requirement</span></span>|<span data-ttu-id="d8eac-466">值</span><span class="sxs-lookup"><span data-stu-id="d8eac-466">Value</span></span>|
|---|---|
|[<span data-ttu-id="d8eac-467">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="d8eac-467">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d8eac-468">1.0</span><span class="sxs-lookup"><span data-stu-id="d8eac-468">1.0</span></span>|
|[<span data-ttu-id="d8eac-469">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="d8eac-469">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d8eac-470">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d8eac-470">ReadItem</span></span>|
|[<span data-ttu-id="d8eac-471">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="d8eac-471">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="d8eac-472">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="d8eac-472">Compose or Read</span></span>|

<br>

---
---

#### <a name="normalizedsubject-string"></a><span data-ttu-id="d8eac-473">normalizedSubject: String</span><span class="sxs-lookup"><span data-stu-id="d8eac-473">normalizedSubject: String</span></span>

<span data-ttu-id="d8eac-p122">获取删除了所有前缀（包括 `RE:` 和 `FWD:`）的项目主题。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="d8eac-p122">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="d8eac-p123">normalizedSubject 属性获取包含由电子邮件程序添加的任何标准前缀（如 `RE:` 和 `FW:`）的项目主题。若要获取包含完整前缀的项目主题，请使用 [`subject`](#subject-stringsubject) 属性。</span><span class="sxs-lookup"><span data-stu-id="d8eac-p123">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="d8eac-478">Type</span><span class="sxs-lookup"><span data-stu-id="d8eac-478">Type</span></span>

*   <span data-ttu-id="d8eac-479">String</span><span class="sxs-lookup"><span data-stu-id="d8eac-479">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="d8eac-480">Requirements</span><span class="sxs-lookup"><span data-stu-id="d8eac-480">Requirements</span></span>

|<span data-ttu-id="d8eac-481">要求</span><span class="sxs-lookup"><span data-stu-id="d8eac-481">Requirement</span></span>|<span data-ttu-id="d8eac-482">值</span><span class="sxs-lookup"><span data-stu-id="d8eac-482">Value</span></span>|
|---|---|
|[<span data-ttu-id="d8eac-483">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="d8eac-483">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d8eac-484">1.0</span><span class="sxs-lookup"><span data-stu-id="d8eac-484">1.0</span></span>|
|[<span data-ttu-id="d8eac-485">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="d8eac-485">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d8eac-486">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d8eac-486">ReadItem</span></span>|
|[<span data-ttu-id="d8eac-487">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="d8eac-487">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="d8eac-488">阅读</span><span class="sxs-lookup"><span data-stu-id="d8eac-488">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d8eac-489">示例</span><span class="sxs-lookup"><span data-stu-id="d8eac-489">Example</span></span>

```js
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
console.log("Normalized subject: " + normalizedSubject);
```

<br>

---
---

#### <a name="notificationmessages-notificationmessagesjavascriptapioutlookofficenotificationmessagesviewoutlook-js-17"></a><span data-ttu-id="d8eac-490">notificationMessages: [NotificationMessages](/javascript/api/outlook/office.notificationmessages?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="d8eac-490">notificationMessages: [NotificationMessages](/javascript/api/outlook/office.notificationmessages?view=outlook-js-1.7)</span></span>

<span data-ttu-id="d8eac-491">获取项目的通知邮件。</span><span class="sxs-lookup"><span data-stu-id="d8eac-491">Gets the notification messages for an item.</span></span>

##### <a name="type"></a><span data-ttu-id="d8eac-492">类型</span><span class="sxs-lookup"><span data-stu-id="d8eac-492">Type</span></span>

*   [<span data-ttu-id="d8eac-493">NotificationMessages</span><span class="sxs-lookup"><span data-stu-id="d8eac-493">NotificationMessages</span></span>](/javascript/api/outlook/office.notificationmessages?view=outlook-js-1.7)

##### <a name="requirements"></a><span data-ttu-id="d8eac-494">要求</span><span class="sxs-lookup"><span data-stu-id="d8eac-494">Requirements</span></span>

|<span data-ttu-id="d8eac-495">要求</span><span class="sxs-lookup"><span data-stu-id="d8eac-495">Requirement</span></span>|<span data-ttu-id="d8eac-496">值</span><span class="sxs-lookup"><span data-stu-id="d8eac-496">Value</span></span>|
|---|---|
|[<span data-ttu-id="d8eac-497">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="d8eac-497">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d8eac-498">1.3</span><span class="sxs-lookup"><span data-stu-id="d8eac-498">1.3</span></span>|
|[<span data-ttu-id="d8eac-499">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="d8eac-499">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d8eac-500">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d8eac-500">ReadItem</span></span>|
|[<span data-ttu-id="d8eac-501">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="d8eac-501">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="d8eac-502">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="d8eac-502">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d8eac-503">示例</span><span class="sxs-lookup"><span data-stu-id="d8eac-503">Example</span></span>

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

#### <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-17recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-17"></a><span data-ttu-id="d8eac-504">optionalAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="d8eac-504">optionalAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.7)</span></span>

<span data-ttu-id="d8eac-505">提供对事件的可选与会者的访问权限。</span><span class="sxs-lookup"><span data-stu-id="d8eac-505">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="d8eac-506">对象的类型和访问级别取决于当前项目的模式。</span><span class="sxs-lookup"><span data-stu-id="d8eac-506">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="d8eac-507">阅读模式</span><span class="sxs-lookup"><span data-stu-id="d8eac-507">Read mode</span></span>

<span data-ttu-id="d8eac-508">`optionalAttendees` 属性返回一个数组，其中包含每个可选与会者的 `EmailAddressDetails` 对象。</span><span class="sxs-lookup"><span data-stu-id="d8eac-508">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span> <span data-ttu-id="d8eac-509">默认情况下，集合上限为 100 个成员。</span><span class="sxs-lookup"><span data-stu-id="d8eac-509">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="d8eac-510">但是，在 Windows 和 Mac 上，最多可包含 500 个成员。</span><span class="sxs-lookup"><span data-stu-id="d8eac-510">However, on Windows and Mac, you can get 500 members maximum.</span></span>

```js
var optionalAttendees = Office.context.mailbox.item.optionalAttendees;
console.log("Optional attendees: " + JSON.stringify(optionalAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="d8eac-511">撰写模式</span><span class="sxs-lookup"><span data-stu-id="d8eac-511">Compose mode</span></span>

<span data-ttu-id="d8eac-512">`optionalAttendees` 属性返回一个 `Recipients` 对象，该对象提供用于获取或更新可选与会者的方法。</span><span class="sxs-lookup"><span data-stu-id="d8eac-512">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span> <span data-ttu-id="d8eac-513">默认情况下，集合上限为 100 个成员。</span><span class="sxs-lookup"><span data-stu-id="d8eac-513">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="d8eac-514">但是，在 Windows 和 Mac 上，以下限制适用。</span><span class="sxs-lookup"><span data-stu-id="d8eac-514">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="d8eac-515">最多包含 500 个成员。</span><span class="sxs-lookup"><span data-stu-id="d8eac-515">Get 500 members maximum.</span></span>
- <span data-ttu-id="d8eac-516">为每个呼叫最多设置 100 个成员，总共多达 500 个成员。</span><span class="sxs-lookup"><span data-stu-id="d8eac-516">Set a maximum of 100 members per call, up to 500 members total.</span></span>

```js
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="d8eac-517">类型</span><span class="sxs-lookup"><span data-stu-id="d8eac-517">Type</span></span>

*   <span data-ttu-id="d8eac-518">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="d8eac-518">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.7)</span></span>

##### <a name="requirements"></a><span data-ttu-id="d8eac-519">要求</span><span class="sxs-lookup"><span data-stu-id="d8eac-519">Requirements</span></span>

|<span data-ttu-id="d8eac-520">要求</span><span class="sxs-lookup"><span data-stu-id="d8eac-520">Requirement</span></span>|<span data-ttu-id="d8eac-521">值</span><span class="sxs-lookup"><span data-stu-id="d8eac-521">Value</span></span>|
|---|---|
|[<span data-ttu-id="d8eac-522">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="d8eac-522">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d8eac-523">1.0</span><span class="sxs-lookup"><span data-stu-id="d8eac-523">1.0</span></span>|
|[<span data-ttu-id="d8eac-524">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="d8eac-524">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d8eac-525">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d8eac-525">ReadItem</span></span>|
|[<span data-ttu-id="d8eac-526">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="d8eac-526">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="d8eac-527">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="d8eac-527">Compose or Read</span></span>|

<br>

---
---

#### <a name="organizer-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-17organizerjavascriptapioutlookofficeorganizerviewoutlook-js-17"></a><span data-ttu-id="d8eac-528">组织者： [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)|[组织者](/javascript/api/outlook/office.organizer?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="d8eac-528">organizer: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)|[Organizer](/javascript/api/outlook/office.organizer?view=outlook-js-1.7)</span></span>

<span data-ttu-id="d8eac-529">获取指定会议的组织者的电子邮件地址。</span><span class="sxs-lookup"><span data-stu-id="d8eac-529">Gets the email address of the organizer for a specified meeting.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="d8eac-530">阅读模式</span><span class="sxs-lookup"><span data-stu-id="d8eac-530">Read mode</span></span>

<span data-ttu-id="d8eac-531">该`organizer`属性返回一个[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)对象，该对象代表会议组织者。</span><span class="sxs-lookup"><span data-stu-id="d8eac-531">The `organizer` property returns an [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7) object that represents the meeting organizer.</span></span>

```js
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
console.log("Organizer: " + organizerName + " (" + organizerAddress + ")");
```

##### <a name="compose-mode"></a><span data-ttu-id="d8eac-532">撰写模式</span><span class="sxs-lookup"><span data-stu-id="d8eac-532">Compose mode</span></span>

<span data-ttu-id="d8eac-533">该`organizer`属性返回一个[管理](/javascript/api/outlook/office.organizer?view=outlook-js-1.7)器对象，该对象提供获取组织者值的方法。</span><span class="sxs-lookup"><span data-stu-id="d8eac-533">The `organizer` property returns an [Organizer](/javascript/api/outlook/office.organizer?view=outlook-js-1.7) object that provides a method to get the organizer value.</span></span>

```js
Office.context.mailbox.item.organizer.getAsync(
  function(asyncResult) {
    console.log(JSON.stringify(asyncResult));
  }
);
```

##### <a name="type"></a><span data-ttu-id="d8eac-534">类型</span><span class="sxs-lookup"><span data-stu-id="d8eac-534">Type</span></span>

*   <span data-ttu-id="d8eac-535">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7) | [组织者](/javascript/api/outlook/office.organizer?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="d8eac-535">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7) | [Organizer](/javascript/api/outlook/office.organizer?view=outlook-js-1.7)</span></span>

##### <a name="requirements"></a><span data-ttu-id="d8eac-536">Requirements</span><span class="sxs-lookup"><span data-stu-id="d8eac-536">Requirements</span></span>

|<span data-ttu-id="d8eac-537">要求</span><span class="sxs-lookup"><span data-stu-id="d8eac-537">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="d8eac-538">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="d8eac-538">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d8eac-539">1.0</span><span class="sxs-lookup"><span data-stu-id="d8eac-539">1.0</span></span>|<span data-ttu-id="d8eac-540">1.7</span><span class="sxs-lookup"><span data-stu-id="d8eac-540">1.7</span></span>|
|[<span data-ttu-id="d8eac-541">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="d8eac-541">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d8eac-542">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d8eac-542">ReadItem</span></span>|<span data-ttu-id="d8eac-543">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="d8eac-543">ReadWriteItem</span></span>|
|[<span data-ttu-id="d8eac-544">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="d8eac-544">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="d8eac-545">阅读</span><span class="sxs-lookup"><span data-stu-id="d8eac-545">Read</span></span>|<span data-ttu-id="d8eac-546">撰写</span><span class="sxs-lookup"><span data-stu-id="d8eac-546">Compose</span></span>|

<br>

---
---

#### <a name="nullable-recurrence-recurrencejavascriptapioutlookofficerecurrenceviewoutlook-js-17"></a><span data-ttu-id="d8eac-547">（可以为 null）定期：[定期](/javascript/api/outlook/office.recurrence?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="d8eac-547">(nullable) recurrence: [Recurrence](/javascript/api/outlook/office.recurrence?view=outlook-js-1.7)</span></span>

<span data-ttu-id="d8eac-548">获取或设置约会的定期模式。</span><span class="sxs-lookup"><span data-stu-id="d8eac-548">Gets or sets the recurrence pattern of an appointment.</span></span> <span data-ttu-id="d8eac-549">获取会议请求的定期模式。</span><span class="sxs-lookup"><span data-stu-id="d8eac-549">Gets the recurrence pattern of a meeting request.</span></span> <span data-ttu-id="d8eac-550">约会项目的阅读和撰写模式。</span><span class="sxs-lookup"><span data-stu-id="d8eac-550">Read and compose modes for appointment items.</span></span> <span data-ttu-id="d8eac-551">会议请求项目的阅读模式。</span><span class="sxs-lookup"><span data-stu-id="d8eac-551">Read mode for meeting request items.</span></span>

<span data-ttu-id="d8eac-552">如果`recurrence`项目是系列中的一个系列或一个实例，则该属性返回定期约会或会议请求的[定期](/javascript/api/outlook/office.recurrence?view=outlook-js-1.7)对象。</span><span class="sxs-lookup"><span data-stu-id="d8eac-552">The `recurrence` property returns a [recurrence](/javascript/api/outlook/office.recurrence?view=outlook-js-1.7) object for recurring appointments or meetings requests if an item is a series or an instance in a series.</span></span> <span data-ttu-id="d8eac-553">`null`返回单个约会的单个约会和会议请求。</span><span class="sxs-lookup"><span data-stu-id="d8eac-553">`null` is returned for single appointments and meeting requests of single appointments.</span></span> <span data-ttu-id="d8eac-554">`undefined`对于不是会议请求的邮件，将返回。</span><span class="sxs-lookup"><span data-stu-id="d8eac-554">`undefined` is returned for messages that are not meeting requests.</span></span>

> <span data-ttu-id="d8eac-555">注意：会议请求的`itemClass`值为 IPM。Schedule. 会议请求。</span><span class="sxs-lookup"><span data-stu-id="d8eac-555">Note: Meeting requests have an `itemClass` value of IPM.Schedule.Meeting.Request.</span></span>

> <span data-ttu-id="d8eac-556">注意：如果定期对象为`null`，则表示该对象是单个约会的单个约会或会议请求，而不是某个系列的一部分。</span><span class="sxs-lookup"><span data-stu-id="d8eac-556">Note: If the recurrence object is `null`, this indicates that the object is a single appointment or a meeting request of a single appointment and NOT a part of a series.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="d8eac-557">阅读模式</span><span class="sxs-lookup"><span data-stu-id="d8eac-557">Read mode</span></span>

<span data-ttu-id="d8eac-558">该`recurrence`属性返回一个代表约会定期的[定期](/javascript/api/outlook/office.recurrence?view=outlook-js-1.7)对象。</span><span class="sxs-lookup"><span data-stu-id="d8eac-558">The `recurrence` property returns a [Recurrence](/javascript/api/outlook/office.recurrence?view=outlook-js-1.7) object that represents the appointment recurrence.</span></span> <span data-ttu-id="d8eac-559">此功能适用于约会和会议请求。</span><span class="sxs-lookup"><span data-stu-id="d8eac-559">This is available for appointments and meeting requests.</span></span>

```js
var recurrence = Office.context.mailbox.item.recurrence;
console.log("Recurrence: " + JSON.stringify(recurrence));
```

##### <a name="compose-mode"></a><span data-ttu-id="d8eac-560">撰写模式</span><span class="sxs-lookup"><span data-stu-id="d8eac-560">Compose mode</span></span>

<span data-ttu-id="d8eac-561">该`recurrence`属性返回一个[定期](/javascript/api/outlook/office.recurrence?view=outlook-js-1.7)对象，该对象提供用于管理约会周期的方法。</span><span class="sxs-lookup"><span data-stu-id="d8eac-561">The `recurrence` property returns a [Recurrence](/javascript/api/outlook/office.recurrence?view=outlook-js-1.7) object that provides methods to manage the appointment recurrence.</span></span> <span data-ttu-id="d8eac-562">这可用于约会。</span><span class="sxs-lookup"><span data-stu-id="d8eac-562">This is available for appointments.</span></span>

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

##### <a name="type"></a><span data-ttu-id="d8eac-563">类型</span><span class="sxs-lookup"><span data-stu-id="d8eac-563">Type</span></span>

* [<span data-ttu-id="d8eac-564">循环</span><span class="sxs-lookup"><span data-stu-id="d8eac-564">Recurrence</span></span>](/javascript/api/outlook/office.recurrence?view=outlook-js-1.7)

|<span data-ttu-id="d8eac-565">要求</span><span class="sxs-lookup"><span data-stu-id="d8eac-565">Requirement</span></span>|<span data-ttu-id="d8eac-566">值</span><span class="sxs-lookup"><span data-stu-id="d8eac-566">Value</span></span>|
|---|---|
|[<span data-ttu-id="d8eac-567">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="d8eac-567">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d8eac-568">1.7</span><span class="sxs-lookup"><span data-stu-id="d8eac-568">1.7</span></span>|
|[<span data-ttu-id="d8eac-569">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="d8eac-569">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d8eac-570">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d8eac-570">ReadItem</span></span>|
|[<span data-ttu-id="d8eac-571">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="d8eac-571">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="d8eac-572">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="d8eac-572">Compose or Read</span></span>|

<br>

---
---

#### <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-17recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-17"></a><span data-ttu-id="d8eac-573">requiredAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="d8eac-573">requiredAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.7)</span></span>

<span data-ttu-id="d8eac-574">提供对事件的必需与会者的访问权限。</span><span class="sxs-lookup"><span data-stu-id="d8eac-574">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="d8eac-575">对象的类型和访问级别取决于当前项目的模式。</span><span class="sxs-lookup"><span data-stu-id="d8eac-575">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="d8eac-576">阅读模式</span><span class="sxs-lookup"><span data-stu-id="d8eac-576">Read mode</span></span>

<span data-ttu-id="d8eac-577">`requiredAttendees` 属性返回一个数组，其中包含每个必需与会者的 `EmailAddressDetails` 对象。</span><span class="sxs-lookup"><span data-stu-id="d8eac-577">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span> <span data-ttu-id="d8eac-578">默认情况下，集合上限为 100 个成员。</span><span class="sxs-lookup"><span data-stu-id="d8eac-578">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="d8eac-579">但是，在 Windows 和 Mac 上，最多可包含 500 个成员。</span><span class="sxs-lookup"><span data-stu-id="d8eac-579">However, on Windows and Mac, you can get 500 members maximum.</span></span>

```js
var requiredAttendees = Office.context.mailbox.item.requiredAttendees;
console.log("Required attendees: " + JSON.stringify(requiredAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="d8eac-580">撰写模式</span><span class="sxs-lookup"><span data-stu-id="d8eac-580">Compose mode</span></span>

<span data-ttu-id="d8eac-581">`requiredAttendees` 属性返回一个 `Recipients` 对象，该对象提供用于获取或更新必需与会者的方法。</span><span class="sxs-lookup"><span data-stu-id="d8eac-581">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span> <span data-ttu-id="d8eac-582">默认情况下，集合上限为 100 个成员。</span><span class="sxs-lookup"><span data-stu-id="d8eac-582">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="d8eac-583">但是，在 Windows 和 Mac 上，以下限制适用。</span><span class="sxs-lookup"><span data-stu-id="d8eac-583">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="d8eac-584">最多包含 500 个成员。</span><span class="sxs-lookup"><span data-stu-id="d8eac-584">Get 500 members maximum.</span></span>
- <span data-ttu-id="d8eac-585">为每个呼叫最多设置 100 个成员，总共多达 500 个成员。</span><span class="sxs-lookup"><span data-stu-id="d8eac-585">Set a maximum of 100 members per call, up to 500 members total.</span></span>

```js
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
  console.log(JSON.stringify(arrayOfRequiredAttendeesRecipients));
}
```

##### <a name="type"></a><span data-ttu-id="d8eac-586">类型</span><span class="sxs-lookup"><span data-stu-id="d8eac-586">Type</span></span>

*   <span data-ttu-id="d8eac-587">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="d8eac-587">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.7)</span></span>

##### <a name="requirements"></a><span data-ttu-id="d8eac-588">要求</span><span class="sxs-lookup"><span data-stu-id="d8eac-588">Requirements</span></span>

|<span data-ttu-id="d8eac-589">要求</span><span class="sxs-lookup"><span data-stu-id="d8eac-589">Requirement</span></span>|<span data-ttu-id="d8eac-590">值</span><span class="sxs-lookup"><span data-stu-id="d8eac-590">Value</span></span>|
|---|---|
|[<span data-ttu-id="d8eac-591">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="d8eac-591">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d8eac-592">1.0</span><span class="sxs-lookup"><span data-stu-id="d8eac-592">1.0</span></span>|
|[<span data-ttu-id="d8eac-593">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="d8eac-593">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d8eac-594">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d8eac-594">ReadItem</span></span>|
|[<span data-ttu-id="d8eac-595">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="d8eac-595">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="d8eac-596">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="d8eac-596">Compose or Read</span></span>|

<br>

---
---

#### <a name="sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-17"></a><span data-ttu-id="d8eac-597">sender: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="d8eac-597">sender: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)</span></span>

<span data-ttu-id="d8eac-p134">获取电子邮件发件人的电子邮件地址。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="d8eac-p134">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="d8eac-p135">[`from`](#from-emailaddressdetailsfrom) 和 `sender` 属性表示同一个人，邮件由代理人发送的除外。在此情况下，`from` 属性表示代理程序，而 sender 属性表示代理人。</span><span class="sxs-lookup"><span data-stu-id="d8eac-p135">The [`from`](#from-emailaddressdetailsfrom) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="d8eac-602">`sender` 属性中 `EmailAddressDetails` 对象的 `recipientType` 属性为 `undefined`。</span><span class="sxs-lookup"><span data-stu-id="d8eac-602">The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="d8eac-603">类型</span><span class="sxs-lookup"><span data-stu-id="d8eac-603">Type</span></span>

*   [<span data-ttu-id="d8eac-604">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="d8eac-604">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)

##### <a name="requirements"></a><span data-ttu-id="d8eac-605">Requirements</span><span class="sxs-lookup"><span data-stu-id="d8eac-605">Requirements</span></span>

|<span data-ttu-id="d8eac-606">要求</span><span class="sxs-lookup"><span data-stu-id="d8eac-606">Requirement</span></span>|<span data-ttu-id="d8eac-607">值</span><span class="sxs-lookup"><span data-stu-id="d8eac-607">Value</span></span>|
|---|---|
|[<span data-ttu-id="d8eac-608">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="d8eac-608">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d8eac-609">1.0</span><span class="sxs-lookup"><span data-stu-id="d8eac-609">1.0</span></span>|
|[<span data-ttu-id="d8eac-610">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="d8eac-610">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d8eac-611">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d8eac-611">ReadItem</span></span>|
|[<span data-ttu-id="d8eac-612">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="d8eac-612">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="d8eac-613">阅读</span><span class="sxs-lookup"><span data-stu-id="d8eac-613">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d8eac-614">示例</span><span class="sxs-lookup"><span data-stu-id="d8eac-614">Example</span></span>

```js
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
console.log("Sender: " + senderName + " (" + senderAddress + ")");
```

<br>

---
---

#### <a name="nullable-seriesid-string"></a><span data-ttu-id="d8eac-615">（可以为 null） Webcasts&seriesid： String</span><span class="sxs-lookup"><span data-stu-id="d8eac-615">(nullable) seriesId: String</span></span>

<span data-ttu-id="d8eac-616">获取实例所属的系列的 id。</span><span class="sxs-lookup"><span data-stu-id="d8eac-616">Gets the id of the series that an instance belongs to.</span></span>

<span data-ttu-id="d8eac-617">在 web 上的 Outlook 和桌面客户端中`seriesId` ，返回此项所属的父（系列）项的 Exchange web 服务（EWS） ID。</span><span class="sxs-lookup"><span data-stu-id="d8eac-617">In Outlook on the web and desktop clients, the `seriesId` returns the Exchange Web Services (EWS) ID of the parent (series) item that this item belongs to.</span></span> <span data-ttu-id="d8eac-618">但是，在 iOS 和 Android 中， `seriesId`将返回父项的 REST ID。</span><span class="sxs-lookup"><span data-stu-id="d8eac-618">However, in iOS and Android, the `seriesId` returns the REST ID of the parent item.</span></span>

> [!NOTE]
> <span data-ttu-id="d8eac-619">`seriesId` 属性返回的标识符与 Exchange Web 服务项目标识符相同。</span><span class="sxs-lookup"><span data-stu-id="d8eac-619">The identifier returned by the `seriesId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="d8eac-620">`seriesId`属性与 OUTLOOK REST API 使用的 outlook id 不相同。</span><span class="sxs-lookup"><span data-stu-id="d8eac-620">The `seriesId` property is not identical to the Outlook IDs used by the Outlook REST API.</span></span> <span data-ttu-id="d8eac-621">使用此值进行 REST API 调用前，应使用 [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string) 对它进行转换。</span><span class="sxs-lookup"><span data-stu-id="d8eac-621">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="d8eac-622">有关详细信息，请参阅[从 Outlook 加载项使用 Outlook REST API](/outlook/add-ins/use-rest-api)。</span><span class="sxs-lookup"><span data-stu-id="d8eac-622">For more details, see [Use the Outlook REST APIs from an Outlook add-in](/outlook/add-ins/use-rest-api).</span></span>

<span data-ttu-id="d8eac-623">对于`seriesId`不包含`null`父项（如单个约会、系列项或会议请求）的项，该属性将返回， `undefined`对于不是会议请求的任何其他项，该属性返回。</span><span class="sxs-lookup"><span data-stu-id="d8eac-623">The `seriesId` property returns `null` for items that do not have parent items such as single appointments, series items, or meeting requests and returns `undefined` for any other items that are not meeting requests.</span></span>

##### <a name="type"></a><span data-ttu-id="d8eac-624">Type</span><span class="sxs-lookup"><span data-stu-id="d8eac-624">Type</span></span>

* <span data-ttu-id="d8eac-625">String</span><span class="sxs-lookup"><span data-stu-id="d8eac-625">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="d8eac-626">Requirements</span><span class="sxs-lookup"><span data-stu-id="d8eac-626">Requirements</span></span>

|<span data-ttu-id="d8eac-627">要求</span><span class="sxs-lookup"><span data-stu-id="d8eac-627">Requirement</span></span>|<span data-ttu-id="d8eac-628">值</span><span class="sxs-lookup"><span data-stu-id="d8eac-628">Value</span></span>|
|---|---|
|[<span data-ttu-id="d8eac-629">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="d8eac-629">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d8eac-630">1.7</span><span class="sxs-lookup"><span data-stu-id="d8eac-630">1.7</span></span>|
|[<span data-ttu-id="d8eac-631">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="d8eac-631">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d8eac-632">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d8eac-632">ReadItem</span></span>|
|[<span data-ttu-id="d8eac-633">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="d8eac-633">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="d8eac-634">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="d8eac-634">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d8eac-635">示例</span><span class="sxs-lookup"><span data-stu-id="d8eac-635">Example</span></span>

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

#### <a name="start-datetimejavascriptapioutlookofficetimeviewoutlook-js-17"></a><span data-ttu-id="d8eac-636">start: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="d8eac-636">start: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.7)</span></span>

<span data-ttu-id="d8eac-637">获取或设置约会开始的日期和时间。</span><span class="sxs-lookup"><span data-stu-id="d8eac-637">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="d8eac-p138">将 `start` 属性表示为协调世界时 (UTC) 的日期和时间值。可使用 [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) 方法将该值转换为客户端的本地日期和时间。</span><span class="sxs-lookup"><span data-stu-id="d8eac-p138">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="d8eac-640">阅读模式</span><span class="sxs-lookup"><span data-stu-id="d8eac-640">Read mode</span></span>

<span data-ttu-id="d8eac-641">`start` 属性返回 `Date` 对象。</span><span class="sxs-lookup"><span data-stu-id="d8eac-641">The `start` property returns a `Date` object.</span></span>

```js
var start = Office.context.mailbox.item.start;
console.log("Appointment start: " + JSON.stringify(start));
```

##### <a name="compose-mode"></a><span data-ttu-id="d8eac-642">撰写模式</span><span class="sxs-lookup"><span data-stu-id="d8eac-642">Compose mode</span></span>

<span data-ttu-id="d8eac-643">`start` 属性返回 `Time` 对象。</span><span class="sxs-lookup"><span data-stu-id="d8eac-643">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="d8eac-644">使用 [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.7#setasync-datetime--options--callback-) 方法设置开始时间时，应使用 [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) 方法将客户端的本地时间转换为服务器的 UTC。</span><span class="sxs-lookup"><span data-stu-id="d8eac-644">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.7#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="d8eac-645">以下示例通过使用 `Time` 对象的 [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.7#setasync-datetime--options--callback-) 方法，设置撰写模式下约会的开始时间。</span><span class="sxs-lookup"><span data-stu-id="d8eac-645">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.7#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="d8eac-646">类型</span><span class="sxs-lookup"><span data-stu-id="d8eac-646">Type</span></span>

*   <span data-ttu-id="d8eac-647">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="d8eac-647">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.7)</span></span>

##### <a name="requirements"></a><span data-ttu-id="d8eac-648">Requirements</span><span class="sxs-lookup"><span data-stu-id="d8eac-648">Requirements</span></span>

|<span data-ttu-id="d8eac-649">要求</span><span class="sxs-lookup"><span data-stu-id="d8eac-649">Requirement</span></span>|<span data-ttu-id="d8eac-650">值</span><span class="sxs-lookup"><span data-stu-id="d8eac-650">Value</span></span>|
|---|---|
|[<span data-ttu-id="d8eac-651">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="d8eac-651">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d8eac-652">1.0</span><span class="sxs-lookup"><span data-stu-id="d8eac-652">1.0</span></span>|
|[<span data-ttu-id="d8eac-653">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="d8eac-653">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d8eac-654">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d8eac-654">ReadItem</span></span>|
|[<span data-ttu-id="d8eac-655">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="d8eac-655">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="d8eac-656">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="d8eac-656">Compose or Read</span></span>|

<br>

---
---

#### <a name="subject-stringsubjectjavascriptapioutlookofficesubjectviewoutlook-js-17"></a><span data-ttu-id="d8eac-657">subject: String|[Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="d8eac-657">subject: String|[Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.7)</span></span>

<span data-ttu-id="d8eac-658">获取或设置显示在项目的主题字段中的说明。</span><span class="sxs-lookup"><span data-stu-id="d8eac-658">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="d8eac-659">`subject` 属性获取或设置由电子邮件服务器发送项目时的整个主题。</span><span class="sxs-lookup"><span data-stu-id="d8eac-659">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="d8eac-660">阅读模式</span><span class="sxs-lookup"><span data-stu-id="d8eac-660">Read mode</span></span>

<span data-ttu-id="d8eac-p139">`subject` 属性返回一个字符串。使用 [`normalizedSubject`](#normalizedsubject-string) 属性获取不带任何前导前缀（如 `RE:` 和 `FW:`）的主题。</span><span class="sxs-lookup"><span data-stu-id="d8eac-p139">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

<span data-ttu-id="d8eac-663">以下 JavaScript 代码示例显示了如何访问 Outlook 中当前项目的 `subject` 属性。</span><span class="sxs-lookup"><span data-stu-id="d8eac-663">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

```js
var subject = Office.context.mailbox.item.subject;
console.log(subject);
```

##### <a name="compose-mode"></a><span data-ttu-id="d8eac-664">撰写模式</span><span class="sxs-lookup"><span data-stu-id="d8eac-664">Compose mode</span></span>

<span data-ttu-id="d8eac-665">`subject` 属性返回一个 `Subject` 对象，该对象提供用于获取和设置主题的方法。</span><span class="sxs-lookup"><span data-stu-id="d8eac-665">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```js
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
  console.log(subject);
}
```

##### <a name="type"></a><span data-ttu-id="d8eac-666">类型</span><span class="sxs-lookup"><span data-stu-id="d8eac-666">Type</span></span>

*   <span data-ttu-id="d8eac-667">String | [Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="d8eac-667">String | [Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.7)</span></span>

##### <a name="requirements"></a><span data-ttu-id="d8eac-668">要求</span><span class="sxs-lookup"><span data-stu-id="d8eac-668">Requirements</span></span>

|<span data-ttu-id="d8eac-669">要求</span><span class="sxs-lookup"><span data-stu-id="d8eac-669">Requirement</span></span>|<span data-ttu-id="d8eac-670">值</span><span class="sxs-lookup"><span data-stu-id="d8eac-670">Value</span></span>|
|---|---|
|[<span data-ttu-id="d8eac-671">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="d8eac-671">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d8eac-672">1.0</span><span class="sxs-lookup"><span data-stu-id="d8eac-672">1.0</span></span>|
|[<span data-ttu-id="d8eac-673">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="d8eac-673">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d8eac-674">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d8eac-674">ReadItem</span></span>|
|[<span data-ttu-id="d8eac-675">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="d8eac-675">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="d8eac-676">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="d8eac-676">Compose or Read</span></span>|

<br>

---
---

#### <a name="to-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-17recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-17"></a><span data-ttu-id="d8eac-677">to: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="d8eac-677">to: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.7)</span></span>

<span data-ttu-id="d8eac-678">提供对邮件的“**收件人**”行上的收件人的访问权限。</span><span class="sxs-lookup"><span data-stu-id="d8eac-678">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="d8eac-679">对象的类型和访问级别取决于当前项目的模式。</span><span class="sxs-lookup"><span data-stu-id="d8eac-679">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="d8eac-680">阅读模式</span><span class="sxs-lookup"><span data-stu-id="d8eac-680">Read mode</span></span>

<span data-ttu-id="d8eac-681">`to` 属性返回包含邮件的**收件人**行上所列的每个收件人的 `EmailAddressDetails` 对象的数组。</span><span class="sxs-lookup"><span data-stu-id="d8eac-681">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message.</span></span> <span data-ttu-id="d8eac-682">默认情况下，集合上限为 100 个成员。</span><span class="sxs-lookup"><span data-stu-id="d8eac-682">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="d8eac-683">但是，在 Windows 和 Mac 上，最多可包含 500 个成员。</span><span class="sxs-lookup"><span data-stu-id="d8eac-683">However, on Windows and Mac, you can get 500 members maximum.</span></span>

```js
console.log(JSON.stringify(Office.context.mailbox.item.to));
```

##### <a name="compose-mode"></a><span data-ttu-id="d8eac-684">撰写模式</span><span class="sxs-lookup"><span data-stu-id="d8eac-684">Compose mode</span></span>

<span data-ttu-id="d8eac-685">`to` 属性返回一个 `Recipients` 对象，该对象提供用于获取或更新邮件的“**收件人**”行上收件人的方法。</span><span class="sxs-lookup"><span data-stu-id="d8eac-685">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span> <span data-ttu-id="d8eac-686">默认情况下，集合上限为 100 个成员。</span><span class="sxs-lookup"><span data-stu-id="d8eac-686">By default, the collection is limited to a maximum of 100 members.</span></span> <span data-ttu-id="d8eac-687">但是，在 Windows 和 Mac 上，以下限制适用。</span><span class="sxs-lookup"><span data-stu-id="d8eac-687">However, on Windows and Mac, the following limits apply.</span></span>

- <span data-ttu-id="d8eac-688">最多包含 500 个成员。</span><span class="sxs-lookup"><span data-stu-id="d8eac-688">Get 500 members maximum.</span></span>
- <span data-ttu-id="d8eac-689">为每个呼叫最多设置 100 个成员，总共多达 500 个成员。</span><span class="sxs-lookup"><span data-stu-id="d8eac-689">Set a maximum of 100 members per call, up to 500 members total.</span></span>

```js
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="d8eac-690">类型</span><span class="sxs-lookup"><span data-stu-id="d8eac-690">Type</span></span>

*   <span data-ttu-id="d8eac-691">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="d8eac-691">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.7)</span></span>

##### <a name="requirements"></a><span data-ttu-id="d8eac-692">Requirements</span><span class="sxs-lookup"><span data-stu-id="d8eac-692">Requirements</span></span>

|<span data-ttu-id="d8eac-693">要求</span><span class="sxs-lookup"><span data-stu-id="d8eac-693">Requirement</span></span>|<span data-ttu-id="d8eac-694">值</span><span class="sxs-lookup"><span data-stu-id="d8eac-694">Value</span></span>|
|---|---|
|[<span data-ttu-id="d8eac-695">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="d8eac-695">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d8eac-696">1.0</span><span class="sxs-lookup"><span data-stu-id="d8eac-696">1.0</span></span>|
|[<span data-ttu-id="d8eac-697">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="d8eac-697">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d8eac-698">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d8eac-698">ReadItem</span></span>|
|[<span data-ttu-id="d8eac-699">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="d8eac-699">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="d8eac-700">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="d8eac-700">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="d8eac-701">方法</span><span class="sxs-lookup"><span data-stu-id="d8eac-701">Methods</span></span>

#### <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="d8eac-702">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="d8eac-702">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="d8eac-703">将文件作为附件添加到邮件或约会。</span><span class="sxs-lookup"><span data-stu-id="d8eac-703">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="d8eac-704">`addFileAttachmentAsync` 方法在指定的 URI 上载文件并将其附加到撰写窗体中的项目。</span><span class="sxs-lookup"><span data-stu-id="d8eac-704">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="d8eac-705">随后可以将该标识符与 [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) 方法一同使用，以删除同一个会话中的附件。</span><span class="sxs-lookup"><span data-stu-id="d8eac-705">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d8eac-706">参数</span><span class="sxs-lookup"><span data-stu-id="d8eac-706">Parameters</span></span>
|<span data-ttu-id="d8eac-707">名称</span><span class="sxs-lookup"><span data-stu-id="d8eac-707">Name</span></span>|<span data-ttu-id="d8eac-708">类型</span><span class="sxs-lookup"><span data-stu-id="d8eac-708">Type</span></span>|<span data-ttu-id="d8eac-709">属性</span><span class="sxs-lookup"><span data-stu-id="d8eac-709">Attributes</span></span>|<span data-ttu-id="d8eac-710">说明</span><span class="sxs-lookup"><span data-stu-id="d8eac-710">Description</span></span>|
|---|---|---|---|
|`uri`|<span data-ttu-id="d8eac-711">String</span><span class="sxs-lookup"><span data-stu-id="d8eac-711">String</span></span>||<span data-ttu-id="d8eac-p143">提供附加到邮件或约会的文件的位置的 URI。最大长度为 2048 个字符。</span><span class="sxs-lookup"><span data-stu-id="d8eac-p143">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`|<span data-ttu-id="d8eac-714">字符串</span><span class="sxs-lookup"><span data-stu-id="d8eac-714">String</span></span>||<span data-ttu-id="d8eac-p144">在附件上载过程中显示的附件名称。最大长度为 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="d8eac-p144">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="d8eac-717">对象</span><span class="sxs-lookup"><span data-stu-id="d8eac-717">Object</span></span>|<span data-ttu-id="d8eac-718">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="d8eac-718">&lt;optional&gt;</span></span>|<span data-ttu-id="d8eac-719">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="d8eac-719">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="d8eac-720">对象</span><span class="sxs-lookup"><span data-stu-id="d8eac-720">Object</span></span>|<span data-ttu-id="d8eac-721">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="d8eac-721">&lt;optional&gt;</span></span>|<span data-ttu-id="d8eac-722">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="d8eac-722">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.isInline`|<span data-ttu-id="d8eac-723">布尔值</span><span class="sxs-lookup"><span data-stu-id="d8eac-723">Boolean</span></span>|<span data-ttu-id="d8eac-724">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="d8eac-724">&lt;optional&gt;</span></span>|<span data-ttu-id="d8eac-725">如果为 `true`，则表示附件将在邮件正文中内联显示，并且不应显示在附件列表中。</span><span class="sxs-lookup"><span data-stu-id="d8eac-725">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`callback`|<span data-ttu-id="d8eac-726">函数</span><span class="sxs-lookup"><span data-stu-id="d8eac-726">function</span></span>|<span data-ttu-id="d8eac-727">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="d8eac-727">&lt;optional&gt;</span></span>|<span data-ttu-id="d8eac-728">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="d8eac-728">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="d8eac-729">如果成功，附件标识符将在 `asyncResult.value` 属性中提供。</span><span class="sxs-lookup"><span data-stu-id="d8eac-729">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="d8eac-730">如果上传附件失败，`asyncResult` 对象将包含一个提供错误说明的 `Error` 对象。</span><span class="sxs-lookup"><span data-stu-id="d8eac-730">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="d8eac-731">错误</span><span class="sxs-lookup"><span data-stu-id="d8eac-731">Errors</span></span>

|<span data-ttu-id="d8eac-732">错误代码</span><span class="sxs-lookup"><span data-stu-id="d8eac-732">Error code</span></span>|<span data-ttu-id="d8eac-733">说明</span><span class="sxs-lookup"><span data-stu-id="d8eac-733">Description</span></span>|
|------------|-------------|
|`AttachmentSizeExceeded`|<span data-ttu-id="d8eac-734">附件大小超过了允许的大小。</span><span class="sxs-lookup"><span data-stu-id="d8eac-734">The attachment is larger than allowed.</span></span>|
|`FileTypeNotSupported`|<span data-ttu-id="d8eac-735">该附件的扩展名不是允许的扩展名。</span><span class="sxs-lookup"><span data-stu-id="d8eac-735">The attachment has an extension that is not allowed.</span></span>|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="d8eac-736">邮件或约会具有的附件过多。</span><span class="sxs-lookup"><span data-stu-id="d8eac-736">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d8eac-737">要求</span><span class="sxs-lookup"><span data-stu-id="d8eac-737">Requirements</span></span>

|<span data-ttu-id="d8eac-738">要求</span><span class="sxs-lookup"><span data-stu-id="d8eac-738">Requirement</span></span>|<span data-ttu-id="d8eac-739">值</span><span class="sxs-lookup"><span data-stu-id="d8eac-739">Value</span></span>|
|---|---|
|[<span data-ttu-id="d8eac-740">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="d8eac-740">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d8eac-741">1.1</span><span class="sxs-lookup"><span data-stu-id="d8eac-741">1.1</span></span>|
|[<span data-ttu-id="d8eac-742">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="d8eac-742">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d8eac-743">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="d8eac-743">ReadWriteItem</span></span>|
|[<span data-ttu-id="d8eac-744">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="d8eac-744">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="d8eac-745">撰写</span><span class="sxs-lookup"><span data-stu-id="d8eac-745">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="d8eac-746">示例</span><span class="sxs-lookup"><span data-stu-id="d8eac-746">Examples</span></span>

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

<span data-ttu-id="d8eac-747">以下示例将图像文件添加为内联附件，并在邮件正文中引用该附件。</span><span class="sxs-lookup"><span data-stu-id="d8eac-747">The following example adds an image file as an inline attachment and references the attachment in the message body.</span></span>

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

#### <a name="addhandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="d8eac-748">addHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="d8eac-748">addHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="d8eac-749">添加支持事件的事件处理程序。</span><span class="sxs-lookup"><span data-stu-id="d8eac-749">Adds an event handler for a supported event.</span></span>

<span data-ttu-id="d8eac-750">目前，受支持的事件`Office.EventType.AppointmentTimeChanged`类型`Office.EventType.RecipientsChanged`是、和`Office.EventType.RecurrenceChanged`</span><span class="sxs-lookup"><span data-stu-id="d8eac-750">Currently the supported event types are `Office.EventType.AppointmentTimeChanged`, `Office.EventType.RecipientsChanged`, and `Office.EventType.RecurrenceChanged`</span></span>

##### <a name="parameters"></a><span data-ttu-id="d8eac-751">Parameters</span><span class="sxs-lookup"><span data-stu-id="d8eac-751">Parameters</span></span>

| <span data-ttu-id="d8eac-752">名称</span><span class="sxs-lookup"><span data-stu-id="d8eac-752">Name</span></span> | <span data-ttu-id="d8eac-753">类型</span><span class="sxs-lookup"><span data-stu-id="d8eac-753">Type</span></span> | <span data-ttu-id="d8eac-754">属性</span><span class="sxs-lookup"><span data-stu-id="d8eac-754">Attributes</span></span> | <span data-ttu-id="d8eac-755">说明</span><span class="sxs-lookup"><span data-stu-id="d8eac-755">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="d8eac-756">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="d8eac-756">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="d8eac-757">应调用处理程序的事件。</span><span class="sxs-lookup"><span data-stu-id="d8eac-757">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="d8eac-758">函数</span><span class="sxs-lookup"><span data-stu-id="d8eac-758">Function</span></span> || <span data-ttu-id="d8eac-p145">用于处理事件的函数。此函数必须接受一个参数，即对象文本。参数上的 `type` 属性将匹配传递给 `addHandlerAsync` 的 `eventType` 参数。</span><span class="sxs-lookup"><span data-stu-id="d8eac-p145">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `addHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="d8eac-762">Object</span><span class="sxs-lookup"><span data-stu-id="d8eac-762">Object</span></span> | <span data-ttu-id="d8eac-763">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="d8eac-763">&lt;optional&gt;</span></span> | <span data-ttu-id="d8eac-764">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="d8eac-764">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="d8eac-765">对象</span><span class="sxs-lookup"><span data-stu-id="d8eac-765">Object</span></span> | <span data-ttu-id="d8eac-766">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="d8eac-766">&lt;optional&gt;</span></span> | <span data-ttu-id="d8eac-767">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="d8eac-767">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="d8eac-768">函数</span><span class="sxs-lookup"><span data-stu-id="d8eac-768">function</span></span>| <span data-ttu-id="d8eac-769">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="d8eac-769">&lt;optional&gt;</span></span>|<span data-ttu-id="d8eac-770">方法完成后，使用单个参数 `callback`（一个 [`asyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `AsyncResult` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="d8eac-770">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d8eac-771">Requirements</span><span class="sxs-lookup"><span data-stu-id="d8eac-771">Requirements</span></span>

|<span data-ttu-id="d8eac-772">要求</span><span class="sxs-lookup"><span data-stu-id="d8eac-772">Requirement</span></span>| <span data-ttu-id="d8eac-773">值</span><span class="sxs-lookup"><span data-stu-id="d8eac-773">Value</span></span>|
|---|---|
|[<span data-ttu-id="d8eac-774">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="d8eac-774">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d8eac-775">1.7</span><span class="sxs-lookup"><span data-stu-id="d8eac-775">1.7</span></span> |
|[<span data-ttu-id="d8eac-776">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="d8eac-776">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d8eac-777">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d8eac-777">ReadItem</span></span> |
|[<span data-ttu-id="d8eac-778">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="d8eac-778">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d8eac-779">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="d8eac-779">Compose or Read</span></span> |

##### <a name="example"></a><span data-ttu-id="d8eac-780">示例</span><span class="sxs-lookup"><span data-stu-id="d8eac-780">Example</span></span>

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

#### <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="d8eac-781">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="d8eac-781">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="d8eac-782">将 Exchange 项目（如邮件）作为附件添加到邮件或约会。</span><span class="sxs-lookup"><span data-stu-id="d8eac-782">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="d8eac-p146">`addItemAttachmentAsync` 方法将包含指定 Exchange 标识符的项目附加到撰写窗体中的项目。如果指定一个回调方法，此方法使用 `asyncResult` 参数调用，该参数包含一个附件标识符或代码，指示附加项目过程中出现的任何错误。可以使用 `options` 参数将状态信息传递给回调方法（如果需要）。</span><span class="sxs-lookup"><span data-stu-id="d8eac-p146">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="d8eac-786">随后可以将该标识符与 [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) 方法一同使用，以删除同一个会话中的附件。</span><span class="sxs-lookup"><span data-stu-id="d8eac-786">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="d8eac-787">如果 Office 加载项是在 Outlook 网页版中运行，`addItemAttachmentAsync` 方法可以将项附加到除正在编辑的项外的项；但既不支持也不建议这样做。</span><span class="sxs-lookup"><span data-stu-id="d8eac-787">If your Office Add-in is running in Outlook on the web, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d8eac-788">Parameters</span><span class="sxs-lookup"><span data-stu-id="d8eac-788">Parameters</span></span>

|<span data-ttu-id="d8eac-789">名称</span><span class="sxs-lookup"><span data-stu-id="d8eac-789">Name</span></span>|<span data-ttu-id="d8eac-790">类型</span><span class="sxs-lookup"><span data-stu-id="d8eac-790">Type</span></span>|<span data-ttu-id="d8eac-791">属性</span><span class="sxs-lookup"><span data-stu-id="d8eac-791">Attributes</span></span>|<span data-ttu-id="d8eac-792">说明</span><span class="sxs-lookup"><span data-stu-id="d8eac-792">Description</span></span>|
|---|---|---|---|
|`itemId`|<span data-ttu-id="d8eac-793">String</span><span class="sxs-lookup"><span data-stu-id="d8eac-793">String</span></span>||<span data-ttu-id="d8eac-p147">要附加的项目的 Exchange 标识符。最大长度为 100 个字符。</span><span class="sxs-lookup"><span data-stu-id="d8eac-p147">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`|<span data-ttu-id="d8eac-796">String</span><span class="sxs-lookup"><span data-stu-id="d8eac-796">String</span></span>||<span data-ttu-id="d8eac-797">要附加的项目的主题。</span><span class="sxs-lookup"><span data-stu-id="d8eac-797">The subject of the item to be attached.</span></span> <span data-ttu-id="d8eac-798">最大长度为 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="d8eac-798">The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="d8eac-799">对象</span><span class="sxs-lookup"><span data-stu-id="d8eac-799">Object</span></span>|<span data-ttu-id="d8eac-800">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="d8eac-800">&lt;optional&gt;</span></span>|<span data-ttu-id="d8eac-801">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="d8eac-801">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="d8eac-802">对象</span><span class="sxs-lookup"><span data-stu-id="d8eac-802">Object</span></span>|<span data-ttu-id="d8eac-803">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="d8eac-803">&lt;optional&gt;</span></span>|<span data-ttu-id="d8eac-804">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="d8eac-804">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="d8eac-805">函数</span><span class="sxs-lookup"><span data-stu-id="d8eac-805">function</span></span>|<span data-ttu-id="d8eac-806">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="d8eac-806">&lt;optional&gt;</span></span>|<span data-ttu-id="d8eac-807">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="d8eac-807">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="d8eac-808">如果成功，附件标识符将在 `asyncResult.value` 属性中提供。</span><span class="sxs-lookup"><span data-stu-id="d8eac-808">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="d8eac-809">如果添加附件失败，`asyncResult` 对象将包含一个提供错误说明的 `Error` 对象。</span><span class="sxs-lookup"><span data-stu-id="d8eac-809">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="d8eac-810">错误</span><span class="sxs-lookup"><span data-stu-id="d8eac-810">Errors</span></span>

|<span data-ttu-id="d8eac-811">错误代码</span><span class="sxs-lookup"><span data-stu-id="d8eac-811">Error code</span></span>|<span data-ttu-id="d8eac-812">说明</span><span class="sxs-lookup"><span data-stu-id="d8eac-812">Description</span></span>|
|------------|-------------|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="d8eac-813">邮件或约会具有的附件过多。</span><span class="sxs-lookup"><span data-stu-id="d8eac-813">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d8eac-814">要求</span><span class="sxs-lookup"><span data-stu-id="d8eac-814">Requirements</span></span>

|<span data-ttu-id="d8eac-815">要求</span><span class="sxs-lookup"><span data-stu-id="d8eac-815">Requirement</span></span>|<span data-ttu-id="d8eac-816">值</span><span class="sxs-lookup"><span data-stu-id="d8eac-816">Value</span></span>|
|---|---|
|[<span data-ttu-id="d8eac-817">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="d8eac-817">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d8eac-818">1.1</span><span class="sxs-lookup"><span data-stu-id="d8eac-818">1.1</span></span>|
|[<span data-ttu-id="d8eac-819">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="d8eac-819">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d8eac-820">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="d8eac-820">ReadWriteItem</span></span>|
|[<span data-ttu-id="d8eac-821">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="d8eac-821">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="d8eac-822">撰写</span><span class="sxs-lookup"><span data-stu-id="d8eac-822">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="d8eac-823">示例</span><span class="sxs-lookup"><span data-stu-id="d8eac-823">Example</span></span>

<span data-ttu-id="d8eac-824">以下示例将现有的 Outlook 项目添加为名为 `My Attachment` 的附件。</span><span class="sxs-lookup"><span data-stu-id="d8eac-824">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

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

#### <a name="close"></a><span data-ttu-id="d8eac-825">close()</span><span class="sxs-lookup"><span data-stu-id="d8eac-825">close()</span></span>

<span data-ttu-id="d8eac-826">关闭当前正在撰写的项目。</span><span class="sxs-lookup"><span data-stu-id="d8eac-826">Closes the current item that is being composed.</span></span>

<span data-ttu-id="d8eac-p149">
            \`close\` 方法的行为取决于要撰写的项目的当前状态。如果项目具有未保存的更改，客户端将提示用户保存、放弃或取消关闭操作。</span><span class="sxs-lookup"><span data-stu-id="d8eac-p149">The behavior of the `close` method depends on the current state of the item being composed. If the item has unsaved changes, the client prompts the user to save, discard, or cancel the close action.</span></span>

> [!NOTE]
> <span data-ttu-id="d8eac-829">在 Outlook 网页版中，如果该项目是约会并且之前已使用 `saveAsync` 保存，则即使自上次保存项目后未发生任何更改，也会提示用户保存、放弃或取消。</span><span class="sxs-lookup"><span data-stu-id="d8eac-829">In Outlook on the web, if the item is an appointment and it has previously been saved using `saveAsync`, the user is prompted to save, discard, or cancel even if no changes have occurred since the item was last saved.</span></span>

<span data-ttu-id="d8eac-830">在 Outlook 桌面客户端中，如果邮件是内联答复，`close` 方法不起作用。</span><span class="sxs-lookup"><span data-stu-id="d8eac-830">In the Outlook desktop client, if the message is an inline reply, the `close` method has no effect.</span></span>

##### <a name="requirements"></a><span data-ttu-id="d8eac-831">Requirements</span><span class="sxs-lookup"><span data-stu-id="d8eac-831">Requirements</span></span>

|<span data-ttu-id="d8eac-832">要求</span><span class="sxs-lookup"><span data-stu-id="d8eac-832">Requirement</span></span>|<span data-ttu-id="d8eac-833">值</span><span class="sxs-lookup"><span data-stu-id="d8eac-833">Value</span></span>|
|---|---|
|[<span data-ttu-id="d8eac-834">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="d8eac-834">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d8eac-835">1.3</span><span class="sxs-lookup"><span data-stu-id="d8eac-835">1.3</span></span>|
|[<span data-ttu-id="d8eac-836">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="d8eac-836">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d8eac-837">受限</span><span class="sxs-lookup"><span data-stu-id="d8eac-837">Restricted</span></span>|
|[<span data-ttu-id="d8eac-838">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="d8eac-838">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="d8eac-839">撰写</span><span class="sxs-lookup"><span data-stu-id="d8eac-839">Compose</span></span>|

<br>

---
---

#### <a name="displayreplyallformformdata-callback"></a><span data-ttu-id="d8eac-840">displayReplyAllForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="d8eac-840">displayReplyAllForm(formData, [callback])</span></span>

<span data-ttu-id="d8eac-841">显示答复窗体，其中包括所选邮件的发件人和所有收件人或所选约会的组织者和所有与会者。</span><span class="sxs-lookup"><span data-stu-id="d8eac-841">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="d8eac-842">iOS 版 Outlook 或 Android 版 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="d8eac-842">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="d8eac-843">在 Outlook 网页版中，答复窗体显示为包含 3 列视图的弹出式窗体，以及包含 2 列或 1 列视图的弹出式窗体。</span><span class="sxs-lookup"><span data-stu-id="d8eac-843">In Outlook on the web, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="d8eac-844">如果任意字符串参数超出其限制，`displayReplyAllForm` 将引发异常。</span><span class="sxs-lookup"><span data-stu-id="d8eac-844">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="d8eac-p150">如果附件已在 `formData.attachments` 参数中指定，Outlook 网页版和 Outlook 桌面版客户端会尝试下载所有附件，并将它们附加到答复窗体。如果无法添加任何附件，窗体 UI 中会显示错误。如果此操作是不可能完成的，系统不会抛出任何错误消息。</span><span class="sxs-lookup"><span data-stu-id="d8eac-p150">When attachments are specified in the `formData.attachments` parameter, Outlook on the web and desktop clients attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d8eac-848">Parameters</span><span class="sxs-lookup"><span data-stu-id="d8eac-848">Parameters</span></span>

|<span data-ttu-id="d8eac-849">名称</span><span class="sxs-lookup"><span data-stu-id="d8eac-849">Name</span></span>|<span data-ttu-id="d8eac-850">类型</span><span class="sxs-lookup"><span data-stu-id="d8eac-850">Type</span></span>|<span data-ttu-id="d8eac-851">属性</span><span class="sxs-lookup"><span data-stu-id="d8eac-851">Attributes</span></span>|<span data-ttu-id="d8eac-852">说明</span><span class="sxs-lookup"><span data-stu-id="d8eac-852">Description</span></span>|
|---|---|---|---|
|`formData`|<span data-ttu-id="d8eac-853">字符串 &#124; 对象</span><span class="sxs-lookup"><span data-stu-id="d8eac-853">String &#124; Object</span></span>||<span data-ttu-id="d8eac-p151">一个包含文本和 HTML 且表示答复窗体的正文的字符串。字符串限制为 32 KB。</span><span class="sxs-lookup"><span data-stu-id="d8eac-p151">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="d8eac-856">**或**</span><span class="sxs-lookup"><span data-stu-id="d8eac-856">**OR**</span></span><br/><span data-ttu-id="d8eac-p152">包含正文或附件数据和回调函数的对象。对象定义如下。</span><span class="sxs-lookup"><span data-stu-id="d8eac-p152">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span>|
|`formData.htmlBody`|<span data-ttu-id="d8eac-859">String</span><span class="sxs-lookup"><span data-stu-id="d8eac-859">String</span></span>|<span data-ttu-id="d8eac-860">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="d8eac-860">&lt;optional&gt;</span></span>|<span data-ttu-id="d8eac-p153">一个包含文本和 HTML 且表示答复窗体的正文的字符串。字符串限制为 32 KB。</span><span class="sxs-lookup"><span data-stu-id="d8eac-p153">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
|`formData.attachments`|<span data-ttu-id="d8eac-863">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="d8eac-863">Array.&lt;Object&gt;</span></span>|<span data-ttu-id="d8eac-864">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="d8eac-864">&lt;optional&gt;</span></span>|<span data-ttu-id="d8eac-865">JSON 对象（是文件或项目附件）的数组。</span><span class="sxs-lookup"><span data-stu-id="d8eac-865">An array of JSON objects that are either file or item attachments.</span></span>|
|`formData.attachments.type`|<span data-ttu-id="d8eac-866">String</span><span class="sxs-lookup"><span data-stu-id="d8eac-866">String</span></span>||<span data-ttu-id="d8eac-p154">指示附件的类型。必须是文件附件的 `file` 或项目附件的 `item`。</span><span class="sxs-lookup"><span data-stu-id="d8eac-p154">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span>|
|`formData.attachments.name`|<span data-ttu-id="d8eac-869">字符串</span><span class="sxs-lookup"><span data-stu-id="d8eac-869">String</span></span>||<span data-ttu-id="d8eac-870">一个包含附件的名称的字符串，最多包含 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="d8eac-870">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
|`formData.attachments.url`|<span data-ttu-id="d8eac-871">String</span><span class="sxs-lookup"><span data-stu-id="d8eac-871">String</span></span>||<span data-ttu-id="d8eac-p155">仅在将 `type` 设置为 `file` 时使用。文件的位置的 URI。</span><span class="sxs-lookup"><span data-stu-id="d8eac-p155">Only used if `type` is set to `file`. The URI of the location for the file.</span></span>|
|`formData.attachments.isInline`|<span data-ttu-id="d8eac-874">布尔</span><span class="sxs-lookup"><span data-stu-id="d8eac-874">Boolean</span></span>||<span data-ttu-id="d8eac-p156">仅在将 `type` 设置为 `file` 时使用。如果为 `true`，则表示附件将在邮件正文中内联显示，并且不应显示在附件列表中。</span><span class="sxs-lookup"><span data-stu-id="d8eac-p156">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`formData.attachments.itemId`|<span data-ttu-id="d8eac-877">String</span><span class="sxs-lookup"><span data-stu-id="d8eac-877">String</span></span>||<span data-ttu-id="d8eac-p157">仅在将 `type` 设置为 `item` 时使用。附件的 EWS 项目 ID。字符串最长为 100 个字符。</span><span class="sxs-lookup"><span data-stu-id="d8eac-p157">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span>|
|`callback`|<span data-ttu-id="d8eac-881">函数</span><span class="sxs-lookup"><span data-stu-id="d8eac-881">function</span></span>|<span data-ttu-id="d8eac-882">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="d8eac-882">&lt;optional&gt;</span></span>|<span data-ttu-id="d8eac-883">方法完成后，使用单个参数 `asyncResult`（一个 [AsyncResult](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="d8eac-883">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d8eac-884">要求</span><span class="sxs-lookup"><span data-stu-id="d8eac-884">Requirements</span></span>

|<span data-ttu-id="d8eac-885">要求</span><span class="sxs-lookup"><span data-stu-id="d8eac-885">Requirement</span></span>|<span data-ttu-id="d8eac-886">值</span><span class="sxs-lookup"><span data-stu-id="d8eac-886">Value</span></span>|
|---|---|
|[<span data-ttu-id="d8eac-887">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="d8eac-887">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d8eac-888">1.0</span><span class="sxs-lookup"><span data-stu-id="d8eac-888">1.0</span></span>|
|[<span data-ttu-id="d8eac-889">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="d8eac-889">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d8eac-890">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d8eac-890">ReadItem</span></span>|
|[<span data-ttu-id="d8eac-891">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="d8eac-891">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="d8eac-892">阅读</span><span class="sxs-lookup"><span data-stu-id="d8eac-892">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="d8eac-893">示例</span><span class="sxs-lookup"><span data-stu-id="d8eac-893">Examples</span></span>

<span data-ttu-id="d8eac-894">以下代码将一个字符串传递到 `displayReplyAllForm` 函数。</span><span class="sxs-lookup"><span data-stu-id="d8eac-894">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="d8eac-895">使用空白正文答复。</span><span class="sxs-lookup"><span data-stu-id="d8eac-895">Reply with an empty body.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="d8eac-896">仅使用正文答复。</span><span class="sxs-lookup"><span data-stu-id="d8eac-896">Reply with just a body.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="d8eac-897">使用正文和文件附件答复。</span><span class="sxs-lookup"><span data-stu-id="d8eac-897">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="d8eac-898">使用正文和项目附件答复。</span><span class="sxs-lookup"><span data-stu-id="d8eac-898">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="d8eac-899">使用正文、文件附件、项目附件和回调答复。</span><span class="sxs-lookup"><span data-stu-id="d8eac-899">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="displayreplyformformdata-callback"></a><span data-ttu-id="d8eac-900">displayReplyForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="d8eac-900">displayReplyForm(formData, [callback])</span></span>

<span data-ttu-id="d8eac-901">显示答复窗体，其中仅包括所选邮件的发件人或所选约会的组织者。</span><span class="sxs-lookup"><span data-stu-id="d8eac-901">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="d8eac-902">iOS 版 Outlook 或 Android 版 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="d8eac-902">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="d8eac-903">在 Outlook 网页版中，答复窗体显示为包含 3 列视图的弹出式窗体，以及包含 2 列或 1 列视图的弹出式窗体。</span><span class="sxs-lookup"><span data-stu-id="d8eac-903">In Outlook on the web, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="d8eac-904">如果任意字符串参数超出其限制，`displayReplyForm` 将引发异常。</span><span class="sxs-lookup"><span data-stu-id="d8eac-904">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="d8eac-p158">如果附件已在 `formData.attachments` 参数中指定，Outlook 网页版和 Outlook 桌面版客户端会尝试下载所有附件，并将它们附加到答复窗体。如果无法添加任何附件，窗体 UI 中会显示错误。如果此操作是不可能完成的，系统不会抛出任何错误消息。</span><span class="sxs-lookup"><span data-stu-id="d8eac-p158">When attachments are specified in the `formData.attachments` parameter, Outlook on the web and desktop clients attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d8eac-908">Parameters</span><span class="sxs-lookup"><span data-stu-id="d8eac-908">Parameters</span></span>

|<span data-ttu-id="d8eac-909">名称</span><span class="sxs-lookup"><span data-stu-id="d8eac-909">Name</span></span>|<span data-ttu-id="d8eac-910">类型</span><span class="sxs-lookup"><span data-stu-id="d8eac-910">Type</span></span>|<span data-ttu-id="d8eac-911">属性</span><span class="sxs-lookup"><span data-stu-id="d8eac-911">Attributes</span></span>|<span data-ttu-id="d8eac-912">说明</span><span class="sxs-lookup"><span data-stu-id="d8eac-912">Description</span></span>|
|---|---|---|---|
|`formData`|<span data-ttu-id="d8eac-913">字符串 &#124; 对象</span><span class="sxs-lookup"><span data-stu-id="d8eac-913">String &#124; Object</span></span>||<span data-ttu-id="d8eac-p159">一个包含文本和 HTML 且表示答复窗体的正文的字符串。字符串限制为 32 KB。</span><span class="sxs-lookup"><span data-stu-id="d8eac-p159">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="d8eac-916">**或**</span><span class="sxs-lookup"><span data-stu-id="d8eac-916">**OR**</span></span><br/><span data-ttu-id="d8eac-p160">包含正文或附件数据和回调函数的对象。对象定义如下。</span><span class="sxs-lookup"><span data-stu-id="d8eac-p160">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span>|
|`formData.htmlBody`|<span data-ttu-id="d8eac-919">String</span><span class="sxs-lookup"><span data-stu-id="d8eac-919">String</span></span>|<span data-ttu-id="d8eac-920">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="d8eac-920">&lt;optional&gt;</span></span>|<span data-ttu-id="d8eac-p161">一个包含文本和 HTML 且表示答复窗体的正文的字符串。字符串限制为 32 KB。</span><span class="sxs-lookup"><span data-stu-id="d8eac-p161">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
|`formData.attachments`|<span data-ttu-id="d8eac-923">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="d8eac-923">Array.&lt;Object&gt;</span></span>|<span data-ttu-id="d8eac-924">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="d8eac-924">&lt;optional&gt;</span></span>|<span data-ttu-id="d8eac-925">JSON 对象（是文件或项目附件）的数组。</span><span class="sxs-lookup"><span data-stu-id="d8eac-925">An array of JSON objects that are either file or item attachments.</span></span>|
|`formData.attachments.type`|<span data-ttu-id="d8eac-926">字符串</span><span class="sxs-lookup"><span data-stu-id="d8eac-926">String</span></span>||<span data-ttu-id="d8eac-p162">指示附件的类型。必须是文件附件的 `file` 或项目附件的 `item`。</span><span class="sxs-lookup"><span data-stu-id="d8eac-p162">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span>|
|`formData.attachments.name`|<span data-ttu-id="d8eac-929">字符串</span><span class="sxs-lookup"><span data-stu-id="d8eac-929">String</span></span>||<span data-ttu-id="d8eac-930">一个包含附件的名称的字符串，最多包含 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="d8eac-930">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
|`formData.attachments.url`|<span data-ttu-id="d8eac-931">String</span><span class="sxs-lookup"><span data-stu-id="d8eac-931">String</span></span>||<span data-ttu-id="d8eac-p163">仅在将 `type` 设置为 `file` 时使用。文件的位置的 URI。</span><span class="sxs-lookup"><span data-stu-id="d8eac-p163">Only used if `type` is set to `file`. The URI of the location for the file.</span></span>|
|`formData.attachments.isInline`|<span data-ttu-id="d8eac-934">布尔</span><span class="sxs-lookup"><span data-stu-id="d8eac-934">Boolean</span></span>||<span data-ttu-id="d8eac-p164">仅在将 `type` 设置为 `file` 时使用。如果为 `true`，则表示附件将在邮件正文中内联显示，并且不应显示在附件列表中。</span><span class="sxs-lookup"><span data-stu-id="d8eac-p164">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`formData.attachments.itemId`|<span data-ttu-id="d8eac-937">String</span><span class="sxs-lookup"><span data-stu-id="d8eac-937">String</span></span>||<span data-ttu-id="d8eac-p165">仅在将 `type` 设置为 `item` 时使用。附件的 EWS 项目 ID。字符串最长为 100 个字符。</span><span class="sxs-lookup"><span data-stu-id="d8eac-p165">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span>|
|`callback`|<span data-ttu-id="d8eac-941">函数</span><span class="sxs-lookup"><span data-stu-id="d8eac-941">function</span></span>|<span data-ttu-id="d8eac-942">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="d8eac-942">&lt;optional&gt;</span></span>|<span data-ttu-id="d8eac-943">方法完成后，使用单个参数 `asyncResult`（一个 [AsyncResult](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="d8eac-943">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d8eac-944">要求</span><span class="sxs-lookup"><span data-stu-id="d8eac-944">Requirements</span></span>

|<span data-ttu-id="d8eac-945">要求</span><span class="sxs-lookup"><span data-stu-id="d8eac-945">Requirement</span></span>|<span data-ttu-id="d8eac-946">值</span><span class="sxs-lookup"><span data-stu-id="d8eac-946">Value</span></span>|
|---|---|
|[<span data-ttu-id="d8eac-947">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="d8eac-947">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d8eac-948">1.0</span><span class="sxs-lookup"><span data-stu-id="d8eac-948">1.0</span></span>|
|[<span data-ttu-id="d8eac-949">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="d8eac-949">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d8eac-950">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d8eac-950">ReadItem</span></span>|
|[<span data-ttu-id="d8eac-951">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="d8eac-951">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="d8eac-952">阅读</span><span class="sxs-lookup"><span data-stu-id="d8eac-952">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="d8eac-953">示例</span><span class="sxs-lookup"><span data-stu-id="d8eac-953">Examples</span></span>

<span data-ttu-id="d8eac-954">以下代码将一个字符串传递到 `displayReplyForm` 函数。</span><span class="sxs-lookup"><span data-stu-id="d8eac-954">The following code passes a string to the `displayReplyForm` function.</span></span>

```js
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="d8eac-955">使用空白正文答复。</span><span class="sxs-lookup"><span data-stu-id="d8eac-955">Reply with an empty body.</span></span>

```js
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="d8eac-956">仅使用正文答复。</span><span class="sxs-lookup"><span data-stu-id="d8eac-956">Reply with just a body.</span></span>

```js
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="d8eac-957">使用正文和文件附件答复。</span><span class="sxs-lookup"><span data-stu-id="d8eac-957">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="d8eac-958">使用正文和项目附件答复。</span><span class="sxs-lookup"><span data-stu-id="d8eac-958">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="d8eac-959">使用正文、文件附件、项目附件和回调答复。</span><span class="sxs-lookup"><span data-stu-id="d8eac-959">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="getentities--entitiesjavascriptapioutlookofficeentitiesviewoutlook-js-17"></a><span data-ttu-id="d8eac-960">getEntities() → {[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.7)}</span><span class="sxs-lookup"><span data-stu-id="d8eac-960">getEntities() → {[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.7)}</span></span>

<span data-ttu-id="d8eac-961">获取在所选项目的正文中找到的实体。</span><span class="sxs-lookup"><span data-stu-id="d8eac-961">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="d8eac-962">iOS 版 Outlook 或 Android 版 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="d8eac-962">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="d8eac-963">要求</span><span class="sxs-lookup"><span data-stu-id="d8eac-963">Requirements</span></span>

|<span data-ttu-id="d8eac-964">要求</span><span class="sxs-lookup"><span data-stu-id="d8eac-964">Requirement</span></span>|<span data-ttu-id="d8eac-965">值</span><span class="sxs-lookup"><span data-stu-id="d8eac-965">Value</span></span>|
|---|---|
|[<span data-ttu-id="d8eac-966">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="d8eac-966">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d8eac-967">1.0</span><span class="sxs-lookup"><span data-stu-id="d8eac-967">1.0</span></span>|
|[<span data-ttu-id="d8eac-968">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="d8eac-968">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d8eac-969">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d8eac-969">ReadItem</span></span>|
|[<span data-ttu-id="d8eac-970">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="d8eac-970">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="d8eac-971">阅读</span><span class="sxs-lookup"><span data-stu-id="d8eac-971">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="d8eac-972">返回：</span><span class="sxs-lookup"><span data-stu-id="d8eac-972">Returns:</span></span>

<span data-ttu-id="d8eac-973">类型：[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="d8eac-973">Type: [Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.7)</span></span>

##### <a name="example"></a><span data-ttu-id="d8eac-974">示例</span><span class="sxs-lookup"><span data-stu-id="d8eac-974">Example</span></span>

<span data-ttu-id="d8eac-975">以下示例访问当前项目的正文中的联系人实体。</span><span class="sxs-lookup"><span data-stu-id="d8eac-975">The following example accesses the contacts entities in the current item's body.</span></span>

```js
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

<br>

---
---

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlookofficecontactviewoutlook-js-17meetingsuggestionjavascriptapioutlookofficemeetingsuggestionviewoutlook-js-17phonenumberjavascriptapioutlookofficephonenumberviewoutlook-js-17tasksuggestionjavascriptapioutlookofficetasksuggestionviewoutlook-js-17"></a><span data-ttu-id="d8eac-976">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.7)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.7)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.7)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.7))>}</span><span class="sxs-lookup"><span data-stu-id="d8eac-976">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.7)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.7)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.7)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.7))>}</span></span>

<span data-ttu-id="d8eac-977">获取所选项目的正文中找到的指定实体类型的所有实体的数组。</span><span class="sxs-lookup"><span data-stu-id="d8eac-977">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="d8eac-978">iOS 版 Outlook 或 Android 版 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="d8eac-978">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d8eac-979">Parameters</span><span class="sxs-lookup"><span data-stu-id="d8eac-979">Parameters</span></span>

|<span data-ttu-id="d8eac-980">名称</span><span class="sxs-lookup"><span data-stu-id="d8eac-980">Name</span></span>|<span data-ttu-id="d8eac-981">类型</span><span class="sxs-lookup"><span data-stu-id="d8eac-981">Type</span></span>|<span data-ttu-id="d8eac-982">说明</span><span class="sxs-lookup"><span data-stu-id="d8eac-982">Description</span></span>|
|---|---|---|
|`entityType`|[<span data-ttu-id="d8eac-983">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="d8eac-983">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook/office.mailboxenums.entitytype?view=outlook-js-1.7)|<span data-ttu-id="d8eac-984">EntityType 枚举值之一。</span><span class="sxs-lookup"><span data-stu-id="d8eac-984">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d8eac-985">Requirements</span><span class="sxs-lookup"><span data-stu-id="d8eac-985">Requirements</span></span>

|<span data-ttu-id="d8eac-986">要求</span><span class="sxs-lookup"><span data-stu-id="d8eac-986">Requirement</span></span>|<span data-ttu-id="d8eac-987">值</span><span class="sxs-lookup"><span data-stu-id="d8eac-987">Value</span></span>|
|---|---|
|[<span data-ttu-id="d8eac-988">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="d8eac-988">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d8eac-989">1.0</span><span class="sxs-lookup"><span data-stu-id="d8eac-989">1.0</span></span>|
|[<span data-ttu-id="d8eac-990">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="d8eac-990">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d8eac-991">受限</span><span class="sxs-lookup"><span data-stu-id="d8eac-991">Restricted</span></span>|
|[<span data-ttu-id="d8eac-992">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="d8eac-992">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="d8eac-993">阅读</span><span class="sxs-lookup"><span data-stu-id="d8eac-993">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="d8eac-994">返回：</span><span class="sxs-lookup"><span data-stu-id="d8eac-994">Returns:</span></span>

<span data-ttu-id="d8eac-995">如果在 `entityType` 中传递的值不是 `EntityType` 枚举的有效成员，该方法返回 null。</span><span class="sxs-lookup"><span data-stu-id="d8eac-995">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="d8eac-996">如果指定类型的任何实体都不存在于该项目的正文中，该方法将返回空数组。</span><span class="sxs-lookup"><span data-stu-id="d8eac-996">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="d8eac-997">否则，返回的数组中对象的类型取决于 `entityType` 参数中请求实体的类型。</span><span class="sxs-lookup"><span data-stu-id="d8eac-997">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="d8eac-998">当使用此方法的最低权限级别**受限**时，某些实体类型需要 **ReadItem** 才能进行访问，如下表中所指定。</span><span class="sxs-lookup"><span data-stu-id="d8eac-998">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

|<span data-ttu-id="d8eac-999">`entityType` 的值</span><span class="sxs-lookup"><span data-stu-id="d8eac-999">Value of `entityType`</span></span>|<span data-ttu-id="d8eac-1000">返回的数组中对象的类型</span><span class="sxs-lookup"><span data-stu-id="d8eac-1000">Type of objects in returned array</span></span>|<span data-ttu-id="d8eac-1001">所需权限级别</span><span class="sxs-lookup"><span data-stu-id="d8eac-1001">Required Permission Level</span></span>|
|---|---|---|
|`Address`|<span data-ttu-id="d8eac-1002">字符串</span><span class="sxs-lookup"><span data-stu-id="d8eac-1002">String</span></span>|<span data-ttu-id="d8eac-1003">**受限**</span><span class="sxs-lookup"><span data-stu-id="d8eac-1003">**Restricted**</span></span>|
|`Contact`|<span data-ttu-id="d8eac-1004">Contact</span><span class="sxs-lookup"><span data-stu-id="d8eac-1004">Contact</span></span>|<span data-ttu-id="d8eac-1005">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="d8eac-1005">**ReadItem**</span></span>|
|`EmailAddress`|<span data-ttu-id="d8eac-1006">String</span><span class="sxs-lookup"><span data-stu-id="d8eac-1006">String</span></span>|<span data-ttu-id="d8eac-1007">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="d8eac-1007">**ReadItem**</span></span>|
|`MeetingSuggestion`|<span data-ttu-id="d8eac-1008">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="d8eac-1008">MeetingSuggestion</span></span>|<span data-ttu-id="d8eac-1009">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="d8eac-1009">**ReadItem**</span></span>|
|`PhoneNumber`|<span data-ttu-id="d8eac-1010">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="d8eac-1010">PhoneNumber</span></span>|<span data-ttu-id="d8eac-1011">**受限**</span><span class="sxs-lookup"><span data-stu-id="d8eac-1011">**Restricted**</span></span>|
|`TaskSuggestion`|<span data-ttu-id="d8eac-1012">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="d8eac-1012">TaskSuggestion</span></span>|<span data-ttu-id="d8eac-1013">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="d8eac-1013">**ReadItem**</span></span>|
|`URL`|<span data-ttu-id="d8eac-1014">字符串</span><span class="sxs-lookup"><span data-stu-id="d8eac-1014">String</span></span>|<span data-ttu-id="d8eac-1015">**受限**</span><span class="sxs-lookup"><span data-stu-id="d8eac-1015">**Restricted**</span></span>|

<span data-ttu-id="d8eac-1016">类型：Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.7)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.7)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.7)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.7))></span><span class="sxs-lookup"><span data-stu-id="d8eac-1016">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.7)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.7)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.7)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.7))></span></span>

##### <a name="example"></a><span data-ttu-id="d8eac-1017">示例</span><span class="sxs-lookup"><span data-stu-id="d8eac-1017">Example</span></span>

<span data-ttu-id="d8eac-1018">以下示例显示了如何访问表示当前项目的正文中的邮政地址的字符串数组。</span><span class="sxs-lookup"><span data-stu-id="d8eac-1018">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlookofficecontactviewoutlook-js-17meetingsuggestionjavascriptapioutlookofficemeetingsuggestionviewoutlook-js-17phonenumberjavascriptapioutlookofficephonenumberviewoutlook-js-17tasksuggestionjavascriptapioutlookofficetasksuggestionviewoutlook-js-17"></a><span data-ttu-id="d8eac-1019">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.7)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.7)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.7)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.7))>}</span><span class="sxs-lookup"><span data-stu-id="d8eac-1019">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.7)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.7)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.7)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.7))>}</span></span>

<span data-ttu-id="d8eac-1020">返回传递清单 XML 文件中定义的命名筛选器的所选项目中的已知实体。</span><span class="sxs-lookup"><span data-stu-id="d8eac-1020">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="d8eac-1021">iOS 版 Outlook 或 Android 版 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="d8eac-1021">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="d8eac-1022">`getFilteredEntitiesByName` 方法返回匹配在具有指定 `FilterName` 元素值的清单 XML 文件中的 [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) 规则元素中定义的正则表达式的实体。</span><span class="sxs-lookup"><span data-stu-id="d8eac-1022">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d8eac-1023">参数</span><span class="sxs-lookup"><span data-stu-id="d8eac-1023">Parameters</span></span>

|<span data-ttu-id="d8eac-1024">名称</span><span class="sxs-lookup"><span data-stu-id="d8eac-1024">Name</span></span>|<span data-ttu-id="d8eac-1025">类型</span><span class="sxs-lookup"><span data-stu-id="d8eac-1025">Type</span></span>|<span data-ttu-id="d8eac-1026">说明</span><span class="sxs-lookup"><span data-stu-id="d8eac-1026">Description</span></span>|
|---|---|---|
|`name`|<span data-ttu-id="d8eac-1027">字符串</span><span class="sxs-lookup"><span data-stu-id="d8eac-1027">String</span></span>|<span data-ttu-id="d8eac-1028">定义筛选器匹配的 `ItemHasKnownEntity` 规则元素的名称。</span><span class="sxs-lookup"><span data-stu-id="d8eac-1028">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d8eac-1029">Requirements</span><span class="sxs-lookup"><span data-stu-id="d8eac-1029">Requirements</span></span>

|<span data-ttu-id="d8eac-1030">要求</span><span class="sxs-lookup"><span data-stu-id="d8eac-1030">Requirement</span></span>|<span data-ttu-id="d8eac-1031">值</span><span class="sxs-lookup"><span data-stu-id="d8eac-1031">Value</span></span>|
|---|---|
|[<span data-ttu-id="d8eac-1032">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="d8eac-1032">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d8eac-1033">1.0</span><span class="sxs-lookup"><span data-stu-id="d8eac-1033">1.0</span></span>|
|[<span data-ttu-id="d8eac-1034">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="d8eac-1034">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d8eac-1035">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d8eac-1035">ReadItem</span></span>|
|[<span data-ttu-id="d8eac-1036">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="d8eac-1036">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="d8eac-1037">阅读</span><span class="sxs-lookup"><span data-stu-id="d8eac-1037">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="d8eac-1038">返回：</span><span class="sxs-lookup"><span data-stu-id="d8eac-1038">Returns:</span></span>

<span data-ttu-id="d8eac-p167">如果具有匹配 `name` 参数的 `FilterName` 元素值的清单中没有任何 `ItemHasKnownEntity` 元素，则该方法返回 `null`。如果 `name` 参数匹配清单中的 `ItemHasKnownEntity` 元素，但在匹配的当前项目中没有实体，则该方法返回一个空数组。</span><span class="sxs-lookup"><span data-stu-id="d8eac-p167">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="d8eac-1041">类型：Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.7)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.7)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.7)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.7))></span><span class="sxs-lookup"><span data-stu-id="d8eac-1041">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact?view=outlook-js-1.7)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion?view=outlook-js-1.7)|[PhoneNumber](/javascript/api/outlook/office.phonenumber?view=outlook-js-1.7)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion?view=outlook-js-1.7))></span></span>

<br>

---
---

#### <a name="getregexmatches--object"></a><span data-ttu-id="d8eac-1042">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="d8eac-1042">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="d8eac-1043">返回所选项目中匹配在清单 XML 文件中定义的正则表达式的字符串值。</span><span class="sxs-lookup"><span data-stu-id="d8eac-1043">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="d8eac-1044">iOS 版 Outlook 或 Android 版 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="d8eac-1044">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="d8eac-p168">`getRegExMatches` 方法返回匹配在清单 XML 文件中的每个 `ItemHasRegularExpressionMatch` 或 `ItemHasKnownEntity` 规则元素中定义的正则表达式的字符串。对于 `ItemHasRegularExpressionMatch` 规则，匹配字符串必须发生在该规则指定的项目的属性中。`PropertyName` 简单类型定义支持的属性。</span><span class="sxs-lookup"><span data-stu-id="d8eac-p168">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="d8eac-1048">例如，考虑一个外接程序清单具有以下 `Rule` 元素：</span><span class="sxs-lookup"><span data-stu-id="d8eac-1048">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="d8eac-1049">从 `getRegExMatches` 返回的对象应有两个属性：`fruits` 和 `veggies`。</span><span class="sxs-lookup"><span data-stu-id="d8eac-1049">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="d8eac-p169">如果在项目的正文属性上指定 `ItemHasRegularExpressionMatch` 规则，则正则表达式应进一步筛选正文，不应尝试返回该项目的整个正文。使用正则表达式（如 `.*`）获取项目的整个正文并不总是返回预期的结果。而是使用 [`Body.getAsync`](/javascript/api/outlook/office.body?view=outlook-js-1.7#getasync-coerciontype--options--callback-) 方法检索整个正文。</span><span class="sxs-lookup"><span data-stu-id="d8eac-p169">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook/office.body?view=outlook-js-1.7#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="d8eac-1053">Requirements</span><span class="sxs-lookup"><span data-stu-id="d8eac-1053">Requirements</span></span>

|<span data-ttu-id="d8eac-1054">要求</span><span class="sxs-lookup"><span data-stu-id="d8eac-1054">Requirement</span></span>|<span data-ttu-id="d8eac-1055">值</span><span class="sxs-lookup"><span data-stu-id="d8eac-1055">Value</span></span>|
|---|---|
|[<span data-ttu-id="d8eac-1056">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="d8eac-1056">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d8eac-1057">1.0</span><span class="sxs-lookup"><span data-stu-id="d8eac-1057">1.0</span></span>|
|[<span data-ttu-id="d8eac-1058">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="d8eac-1058">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d8eac-1059">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d8eac-1059">ReadItem</span></span>|
|[<span data-ttu-id="d8eac-1060">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="d8eac-1060">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="d8eac-1061">阅读</span><span class="sxs-lookup"><span data-stu-id="d8eac-1061">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="d8eac-1062">返回：</span><span class="sxs-lookup"><span data-stu-id="d8eac-1062">Returns:</span></span>

<span data-ttu-id="d8eac-p170">一个包含匹配在清单 XML 文件中定义的正则表达式的字符串数组的对象。每个数组的名称等于匹配 `ItemHasRegularExpressionMatch` 规则的 `RegExName` 属性或匹配 `ItemHasKnownEntity` 规则的 `FilterName` 属性的相应值。</span><span class="sxs-lookup"><span data-stu-id="d8eac-p170">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<span data-ttu-id="d8eac-1065">类型：对象</span><span class="sxs-lookup"><span data-stu-id="d8eac-1065">Type: Object</span></span>

##### <a name="example"></a><span data-ttu-id="d8eac-1066">示例</span><span class="sxs-lookup"><span data-stu-id="d8eac-1066">Example</span></span>

<span data-ttu-id="d8eac-1067">以下示例显示了如何访问正则表达式规则元素 `fruits` 和 `veggies` 的匹配项的数组，这些元素在清单中指定。</span><span class="sxs-lookup"><span data-stu-id="d8eac-1067">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```js
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veggies = allMatches.veggies;
```

<br>

---
---

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="d8eac-1068">getRegExMatchesByName(name) → (nullable) {Array.<String>}</span><span class="sxs-lookup"><span data-stu-id="d8eac-1068">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="d8eac-1069">返回匹配在清单 XML 文件中定义的命名正则表达式的所选项目中的字符串值。</span><span class="sxs-lookup"><span data-stu-id="d8eac-1069">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="d8eac-1070">iOS 版 Outlook 或 Android 版 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="d8eac-1070">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="d8eac-1071">`getRegExMatchesByName` 方法返回匹配在具有指定 `RegExName` 元素值的清单 XML 文件中的 `ItemHasRegularExpressionMatch` 规则元素中定义的正则表达式的字符串。</span><span class="sxs-lookup"><span data-stu-id="d8eac-1071">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="d8eac-p171">如果在项目的正文属性上指定 `ItemHasRegularExpressionMatch` 规则，则正则表达式应进一步筛选正文，不应尝试返回该项目的整个正文。使用正则表达式（如 `.*`）获取项目的整个正文并不总是返回预期的结果。</span><span class="sxs-lookup"><span data-stu-id="d8eac-p171">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d8eac-1074">Parameters</span><span class="sxs-lookup"><span data-stu-id="d8eac-1074">Parameters</span></span>

|<span data-ttu-id="d8eac-1075">名称</span><span class="sxs-lookup"><span data-stu-id="d8eac-1075">Name</span></span>|<span data-ttu-id="d8eac-1076">类型</span><span class="sxs-lookup"><span data-stu-id="d8eac-1076">Type</span></span>|<span data-ttu-id="d8eac-1077">说明</span><span class="sxs-lookup"><span data-stu-id="d8eac-1077">Description</span></span>|
|---|---|---|
|`name`|<span data-ttu-id="d8eac-1078">字符串</span><span class="sxs-lookup"><span data-stu-id="d8eac-1078">String</span></span>|<span data-ttu-id="d8eac-1079">定义筛选器匹配的 `ItemHasRegularExpressionMatch` 规则元素的名称。</span><span class="sxs-lookup"><span data-stu-id="d8eac-1079">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d8eac-1080">Requirements</span><span class="sxs-lookup"><span data-stu-id="d8eac-1080">Requirements</span></span>

|<span data-ttu-id="d8eac-1081">要求</span><span class="sxs-lookup"><span data-stu-id="d8eac-1081">Requirement</span></span>|<span data-ttu-id="d8eac-1082">值</span><span class="sxs-lookup"><span data-stu-id="d8eac-1082">Value</span></span>|
|---|---|
|[<span data-ttu-id="d8eac-1083">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="d8eac-1083">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d8eac-1084">1.0</span><span class="sxs-lookup"><span data-stu-id="d8eac-1084">1.0</span></span>|
|[<span data-ttu-id="d8eac-1085">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="d8eac-1085">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d8eac-1086">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d8eac-1086">ReadItem</span></span>|
|[<span data-ttu-id="d8eac-1087">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="d8eac-1087">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="d8eac-1088">阅读</span><span class="sxs-lookup"><span data-stu-id="d8eac-1088">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="d8eac-1089">返回：</span><span class="sxs-lookup"><span data-stu-id="d8eac-1089">Returns:</span></span>

<span data-ttu-id="d8eac-1090">一个包含匹配在清单 XML 文件中定义的正则表达式的字符串的数组。</span><span class="sxs-lookup"><span data-stu-id="d8eac-1090">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<span data-ttu-id="d8eac-1091">类型：Array.< String ></span><span class="sxs-lookup"><span data-stu-id="d8eac-1091">Type: Array.< String ></span></span>

##### <a name="example"></a><span data-ttu-id="d8eac-1092">示例</span><span class="sxs-lookup"><span data-stu-id="d8eac-1092">Example</span></span>

```js
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

<br>

---
---

#### <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="d8eac-1093">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="d8eac-1093">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="d8eac-1094">以异步方式返回邮件的主题或正文中选定的数据。</span><span class="sxs-lookup"><span data-stu-id="d8eac-1094">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="d8eac-p172">如果没有选定内容，但光标位于正文或主题中，此方法将会为所选数据返回 null。如果选定的是字段，而不是正文或主题，则此方法返回 `InvalidSelection` 错误。</span><span class="sxs-lookup"><span data-stu-id="d8eac-p172">If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

> [!NOTE]
> <span data-ttu-id="d8eac-1097">在 web 上的 Outlook 中，如果未选择任何文本，但光标在正文中，则该方法将返回字符串 "null"。</span><span class="sxs-lookup"><span data-stu-id="d8eac-1097">In Outlook on the web, the method returns the string "null" if no text is selected but the cursor is in the body.</span></span> <span data-ttu-id="d8eac-1098">若要检查此情况，请包含与以下内容类似的代码：</span><span class="sxs-lookup"><span data-stu-id="d8eac-1098">To check for this situation, include code similar to the following:</span></span>
>
> `var selectedText = (asyncResult.value.endPosition === asyncResult.value.startPosition) ? "" : asyncResult.value.data;`

##### <a name="parameters"></a><span data-ttu-id="d8eac-1099">Parameters</span><span class="sxs-lookup"><span data-stu-id="d8eac-1099">Parameters</span></span>

|<span data-ttu-id="d8eac-1100">名称</span><span class="sxs-lookup"><span data-stu-id="d8eac-1100">Name</span></span>|<span data-ttu-id="d8eac-1101">类型</span><span class="sxs-lookup"><span data-stu-id="d8eac-1101">Type</span></span>|<span data-ttu-id="d8eac-1102">属性</span><span class="sxs-lookup"><span data-stu-id="d8eac-1102">Attributes</span></span>|<span data-ttu-id="d8eac-1103">说明</span><span class="sxs-lookup"><span data-stu-id="d8eac-1103">Description</span></span>|
|---|---|---|---|
|`coercionType`|[<span data-ttu-id="d8eac-1104">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="d8eac-1104">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="d8eac-p174">请求数据的格式。如果为文本，则此方法返回纯文本作为字符串，删除任何显示的 HTML 标记。如果为 HTML，则此方法返回所选文本，不论是纯文本还是 HTML。</span><span class="sxs-lookup"><span data-stu-id="d8eac-p174">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`|<span data-ttu-id="d8eac-1108">对象</span><span class="sxs-lookup"><span data-stu-id="d8eac-1108">Object</span></span>|<span data-ttu-id="d8eac-1109">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="d8eac-1109">&lt;optional&gt;</span></span>|<span data-ttu-id="d8eac-1110">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="d8eac-1110">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="d8eac-1111">对象</span><span class="sxs-lookup"><span data-stu-id="d8eac-1111">Object</span></span>|<span data-ttu-id="d8eac-1112">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="d8eac-1112">&lt;optional&gt;</span></span>|<span data-ttu-id="d8eac-1113">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="d8eac-1113">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="d8eac-1114">function</span><span class="sxs-lookup"><span data-stu-id="d8eac-1114">function</span></span>||<span data-ttu-id="d8eac-1115">方法完成后，使用单个参数 `callback`（一个 [`asyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `AsyncResult` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="d8eac-1115">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="d8eac-1116">若要从回调方法访问所选数据，请调用 `asyncResult.value.data`。</span><span class="sxs-lookup"><span data-stu-id="d8eac-1116">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="d8eac-1117">若要访问选定内容的源属性，请调用 `asyncResult.value.sourceProperty`，这将为 `body` 或 `subject`。</span><span class="sxs-lookup"><span data-stu-id="d8eac-1117">To access the source property that the selection comes from, call `asyncResult.value.sourceProperty`, which will be either `body` or `subject`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d8eac-1118">Requirements</span><span class="sxs-lookup"><span data-stu-id="d8eac-1118">Requirements</span></span>

|<span data-ttu-id="d8eac-1119">要求</span><span class="sxs-lookup"><span data-stu-id="d8eac-1119">Requirement</span></span>|<span data-ttu-id="d8eac-1120">值</span><span class="sxs-lookup"><span data-stu-id="d8eac-1120">Value</span></span>|
|---|---|
|[<span data-ttu-id="d8eac-1121">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="d8eac-1121">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d8eac-1122">1.2</span><span class="sxs-lookup"><span data-stu-id="d8eac-1122">1.2</span></span>|
|[<span data-ttu-id="d8eac-1123">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="d8eac-1123">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d8eac-1124">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d8eac-1124">ReadItem</span></span>|
|[<span data-ttu-id="d8eac-1125">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="d8eac-1125">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="d8eac-1126">撰写</span><span class="sxs-lookup"><span data-stu-id="d8eac-1126">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="d8eac-1127">返回：</span><span class="sxs-lookup"><span data-stu-id="d8eac-1127">Returns:</span></span>

<span data-ttu-id="d8eac-1128">作为字符串的所选数据的格式由 `coercionType` 确定。</span><span class="sxs-lookup"><span data-stu-id="d8eac-1128">The selected data as a string with format determined by `coercionType`.</span></span>

<span data-ttu-id="d8eac-1129">类型：字符串</span><span class="sxs-lookup"><span data-stu-id="d8eac-1129">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="d8eac-1130">示例</span><span class="sxs-lookup"><span data-stu-id="d8eac-1130">Example</span></span>

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

#### <a name="getselectedentities--entitiesjavascriptapioutlookofficeentitiesviewoutlook-js-17"></a><span data-ttu-id="d8eac-1131">getSelectedEntities() → {[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.7)}</span><span class="sxs-lookup"><span data-stu-id="d8eac-1131">getSelectedEntities() → {[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.7)}</span></span>

<span data-ttu-id="d8eac-1132">获取在用户已选择的突出显示匹配项中找到的实体。</span><span class="sxs-lookup"><span data-stu-id="d8eac-1132">Gets the entities found in a highlighted match a user has selected.</span></span> <span data-ttu-id="d8eac-1133">突出显示匹配项适用于[上下文外接程序](/outlook/add-ins/contextual-outlook-add-ins)。</span><span class="sxs-lookup"><span data-stu-id="d8eac-1133">Highlighted matches apply to [contextual add-ins](/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="d8eac-1134">iOS 版 Outlook 或 Android 版 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="d8eac-1134">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="d8eac-1135">要求</span><span class="sxs-lookup"><span data-stu-id="d8eac-1135">Requirements</span></span>

|<span data-ttu-id="d8eac-1136">要求</span><span class="sxs-lookup"><span data-stu-id="d8eac-1136">Requirement</span></span>|<span data-ttu-id="d8eac-1137">值</span><span class="sxs-lookup"><span data-stu-id="d8eac-1137">Value</span></span>|
|---|---|
|[<span data-ttu-id="d8eac-1138">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="d8eac-1138">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d8eac-1139">1.6</span><span class="sxs-lookup"><span data-stu-id="d8eac-1139">1.6</span></span>|
|[<span data-ttu-id="d8eac-1140">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="d8eac-1140">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d8eac-1141">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d8eac-1141">ReadItem</span></span>|
|[<span data-ttu-id="d8eac-1142">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="d8eac-1142">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="d8eac-1143">阅读</span><span class="sxs-lookup"><span data-stu-id="d8eac-1143">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="d8eac-1144">返回：</span><span class="sxs-lookup"><span data-stu-id="d8eac-1144">Returns:</span></span>

<span data-ttu-id="d8eac-1145">类型：[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="d8eac-1145">Type: [Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.7)</span></span>

##### <a name="example"></a><span data-ttu-id="d8eac-1146">示例</span><span class="sxs-lookup"><span data-stu-id="d8eac-1146">Example</span></span>

<span data-ttu-id="d8eac-1147">以下示例访问用户选择的突出显示匹配项中的地址实体。</span><span class="sxs-lookup"><span data-stu-id="d8eac-1147">The following example accesses the addresses entities in the highlighted match selected by the user.</span></span>

```js
var contacts = Office.context.mailbox.item.getSelectedEntities().addresses;
```

<br>

---
---

#### <a name="getselectedregexmatches--object"></a><span data-ttu-id="d8eac-1148">getSelectedRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="d8eac-1148">getSelectedRegExMatches() → {Object}</span></span>

<span data-ttu-id="d8eac-p177">返回突出显示匹配项中匹配在清单 XML 文件中定义的正则表达式的字符串值。突出显示匹配项适用于[上下文外接程序](/outlook/add-ins/contextual-outlook-add-ins)。</span><span class="sxs-lookup"><span data-stu-id="d8eac-p177">Returns string values in a highlighted match that match the regular expressions defined in the manifest XML file. Highlighted matches apply to [contextual add-ins](/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="d8eac-1151">iOS 版 Outlook 或 Android 版 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="d8eac-1151">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="d8eac-p178">`getSelectedRegExMatches` 方法返回匹配在清单 XML 文件中的每个 `ItemHasRegularExpressionMatch` 或 `ItemHasKnownEntity` 规则元素中定义的正则表达式的字符串。对于 `ItemHasRegularExpressionMatch` 规则，匹配字符串必须发生在该规则指定的项目的属性中。`PropertyName` 简单类型定义支持的属性。</span><span class="sxs-lookup"><span data-stu-id="d8eac-p178">The `getSelectedRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="d8eac-1155">例如，考虑一个外接程序清单具有以下 `Rule` 元素：</span><span class="sxs-lookup"><span data-stu-id="d8eac-1155">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="d8eac-1156">从 `getRegExMatches` 返回的对象应有两个属性：`fruits` 和 `veggies`。</span><span class="sxs-lookup"><span data-stu-id="d8eac-1156">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="d8eac-p179">如果在项目的正文属性上指定 `ItemHasRegularExpressionMatch` 规则，则正则表达式应进一步筛选正文，不应尝试返回该项目的整个正文。使用正则表达式（如 `.*`）获取项目的整个正文并不总是返回预期的结果。而是使用 [`Body.getAsync`](/javascript/api/outlook/office.body?view=outlook-js-1.7#getasync-coerciontype--options--callback-) 方法检索整个正文。</span><span class="sxs-lookup"><span data-stu-id="d8eac-p179">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook/office.body?view=outlook-js-1.7#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="d8eac-1160">Requirements</span><span class="sxs-lookup"><span data-stu-id="d8eac-1160">Requirements</span></span>

|<span data-ttu-id="d8eac-1161">要求</span><span class="sxs-lookup"><span data-stu-id="d8eac-1161">Requirement</span></span>|<span data-ttu-id="d8eac-1162">值</span><span class="sxs-lookup"><span data-stu-id="d8eac-1162">Value</span></span>|
|---|---|
|[<span data-ttu-id="d8eac-1163">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="d8eac-1163">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d8eac-1164">1.6</span><span class="sxs-lookup"><span data-stu-id="d8eac-1164">1.6</span></span>|
|[<span data-ttu-id="d8eac-1165">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="d8eac-1165">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d8eac-1166">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d8eac-1166">ReadItem</span></span>|
|[<span data-ttu-id="d8eac-1167">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="d8eac-1167">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="d8eac-1168">阅读</span><span class="sxs-lookup"><span data-stu-id="d8eac-1168">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="d8eac-1169">返回：</span><span class="sxs-lookup"><span data-stu-id="d8eac-1169">Returns:</span></span>

<span data-ttu-id="d8eac-p180">一个包含匹配在清单 XML 文件中定义的正则表达式的字符串数组的对象。每个数组的名称等于匹配 `ItemHasRegularExpressionMatch` 规则的 `RegExName` 属性或匹配 `ItemHasKnownEntity` 规则的 `FilterName` 属性的相应值。</span><span class="sxs-lookup"><span data-stu-id="d8eac-p180">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

##### <a name="example"></a><span data-ttu-id="d8eac-1172">示例</span><span class="sxs-lookup"><span data-stu-id="d8eac-1172">Example</span></span>

<span data-ttu-id="d8eac-1173">以下示例显示了如何访问正则表达式规则元素 `fruits` 和 `veggies` 的匹配项的数组，这些元素在清单中指定。</span><span class="sxs-lookup"><span data-stu-id="d8eac-1173">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```js
var selectedMatches = Office.context.mailbox.item.getSelectedRegExMatches();
var fruits = selectedMatches.fruits;
var veggies = selectedMatches.veggies;
```

<br>

---
---

#### <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="d8eac-1174">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="d8eac-1174">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="d8eac-1175">异步加载所选项目上此外接程序的自定义属性。</span><span class="sxs-lookup"><span data-stu-id="d8eac-1175">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="d8eac-p181">自定义属性基于每个应用、每个项目存储为键/值对。此方法在回调中返回 `CustomProperties` 对象，该回调提供访问特定于当前项目和当前外接程序的自定义属性的方法。自定义属性未在项目上加密，因此这不应用作安全存储。</span><span class="sxs-lookup"><span data-stu-id="d8eac-p181">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d8eac-1179">参数</span><span class="sxs-lookup"><span data-stu-id="d8eac-1179">Parameters</span></span>

|<span data-ttu-id="d8eac-1180">名称</span><span class="sxs-lookup"><span data-stu-id="d8eac-1180">Name</span></span>|<span data-ttu-id="d8eac-1181">类型</span><span class="sxs-lookup"><span data-stu-id="d8eac-1181">Type</span></span>|<span data-ttu-id="d8eac-1182">属性</span><span class="sxs-lookup"><span data-stu-id="d8eac-1182">Attributes</span></span>|<span data-ttu-id="d8eac-1183">说明</span><span class="sxs-lookup"><span data-stu-id="d8eac-1183">Description</span></span>|
|---|---|---|---|
|`callback`|<span data-ttu-id="d8eac-1184">函数</span><span class="sxs-lookup"><span data-stu-id="d8eac-1184">function</span></span>||<span data-ttu-id="d8eac-1185">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="d8eac-1185">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="d8eac-1186">自定义属性作为 `asyncResult.value` 属性中的 [`CustomProperties`](/javascript/api/outlook/office.customproperties?view=outlook-js-1.7) 对象提供。</span><span class="sxs-lookup"><span data-stu-id="d8eac-1186">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook/office.customproperties?view=outlook-js-1.7) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="d8eac-1187">此对象可用于获取、设置以及从项目中删除自定义属性，并将自定义属性集的更改重新保存到服务器。</span><span class="sxs-lookup"><span data-stu-id="d8eac-1187">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`|<span data-ttu-id="d8eac-1188">对象</span><span class="sxs-lookup"><span data-stu-id="d8eac-1188">Object</span></span>|<span data-ttu-id="d8eac-1189">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="d8eac-1189">&lt;optional&gt;</span></span>|<span data-ttu-id="d8eac-1190">开发人员可以提供他们想要在回调函数中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="d8eac-1190">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="d8eac-1191">此对象可以通过回调函数中的 `asyncResult.asyncContext` 属性进行访问。</span><span class="sxs-lookup"><span data-stu-id="d8eac-1191">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d8eac-1192">Requirements</span><span class="sxs-lookup"><span data-stu-id="d8eac-1192">Requirements</span></span>

|<span data-ttu-id="d8eac-1193">要求</span><span class="sxs-lookup"><span data-stu-id="d8eac-1193">Requirement</span></span>|<span data-ttu-id="d8eac-1194">值</span><span class="sxs-lookup"><span data-stu-id="d8eac-1194">Value</span></span>|
|---|---|
|[<span data-ttu-id="d8eac-1195">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="d8eac-1195">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d8eac-1196">1.0</span><span class="sxs-lookup"><span data-stu-id="d8eac-1196">1.0</span></span>|
|[<span data-ttu-id="d8eac-1197">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="d8eac-1197">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d8eac-1198">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d8eac-1198">ReadItem</span></span>|
|[<span data-ttu-id="d8eac-1199">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="d8eac-1199">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="d8eac-1200">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="d8eac-1200">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d8eac-1201">示例</span><span class="sxs-lookup"><span data-stu-id="d8eac-1201">Example</span></span>

<span data-ttu-id="d8eac-p184">以下代码示例显示了如何使用 `loadCustomPropertiesAsync` 方法异步加载特定于当前项目的自定义属性。该示例还显示了如何使用 `CustomProperties.saveAsync` 方法将这些属性重新保存到服务器。加载自定义属性后，该代码示例将使用 `CustomProperties.get` 方法读取自定义属性 `myProp`，使用 `CustomProperties.set` 方法写入自定义属性 `otherProp`，最后调用 `saveAsync` 方法保存这些自定义属性。</span><span class="sxs-lookup"><span data-stu-id="d8eac-p184">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

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

#### <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="d8eac-1205">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="d8eac-1205">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="d8eac-1206">将附件从邮件或约会中删除。</span><span class="sxs-lookup"><span data-stu-id="d8eac-1206">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="d8eac-1207">`removeAttachmentAsync` 方法删除项目中带指定标识符的附件。</span><span class="sxs-lookup"><span data-stu-id="d8eac-1207">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item.</span></span> <span data-ttu-id="d8eac-1208">最佳做法是，仅当同一个邮件应用程序在同一会话中添加了一个附件时，你才应使用该附件标识符来删除该附件。</span><span class="sxs-lookup"><span data-stu-id="d8eac-1208">As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session.</span></span> <span data-ttu-id="d8eac-1209">在 Outlook 网页版和移动设备上，附件标识符只在同一个会话中才有效。</span><span class="sxs-lookup"><span data-stu-id="d8eac-1209">In Outlook on the web and mobile devices, the attachment identifier is valid only within the same session.</span></span> <span data-ttu-id="d8eac-1210">当用户关闭应用程序，或者如果用户开始在内嵌窗体中撰写，并在随后弹出的内嵌窗体中继续在单独的窗口撰写时，会话即结束。</span><span class="sxs-lookup"><span data-stu-id="d8eac-1210">A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d8eac-1211">Parameters</span><span class="sxs-lookup"><span data-stu-id="d8eac-1211">Parameters</span></span>

|<span data-ttu-id="d8eac-1212">名称</span><span class="sxs-lookup"><span data-stu-id="d8eac-1212">Name</span></span>|<span data-ttu-id="d8eac-1213">类型</span><span class="sxs-lookup"><span data-stu-id="d8eac-1213">Type</span></span>|<span data-ttu-id="d8eac-1214">属性</span><span class="sxs-lookup"><span data-stu-id="d8eac-1214">Attributes</span></span>|<span data-ttu-id="d8eac-1215">说明</span><span class="sxs-lookup"><span data-stu-id="d8eac-1215">Description</span></span>|
|---|---|---|---|
|`attachmentId`|<span data-ttu-id="d8eac-1216">字符串</span><span class="sxs-lookup"><span data-stu-id="d8eac-1216">String</span></span>||<span data-ttu-id="d8eac-1217">要删除的附件的标识符。</span><span class="sxs-lookup"><span data-stu-id="d8eac-1217">The identifier of the attachment to remove.</span></span>|
|`options`|<span data-ttu-id="d8eac-1218">对象</span><span class="sxs-lookup"><span data-stu-id="d8eac-1218">Object</span></span>|<span data-ttu-id="d8eac-1219">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="d8eac-1219">&lt;optional&gt;</span></span>|<span data-ttu-id="d8eac-1220">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="d8eac-1220">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="d8eac-1221">对象</span><span class="sxs-lookup"><span data-stu-id="d8eac-1221">Object</span></span>|<span data-ttu-id="d8eac-1222">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="d8eac-1222">&lt;optional&gt;</span></span>|<span data-ttu-id="d8eac-1223">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="d8eac-1223">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="d8eac-1224">函数</span><span class="sxs-lookup"><span data-stu-id="d8eac-1224">function</span></span>|<span data-ttu-id="d8eac-1225">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="d8eac-1225">&lt;optional&gt;</span></span>|<span data-ttu-id="d8eac-1226">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="d8eac-1226">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="d8eac-1227">如果删除附件失败，`asyncResult.error` 属性将包含一个说明失败原因的错误代码。</span><span class="sxs-lookup"><span data-stu-id="d8eac-1227">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="d8eac-1228">错误</span><span class="sxs-lookup"><span data-stu-id="d8eac-1228">Errors</span></span>

|<span data-ttu-id="d8eac-1229">错误代码</span><span class="sxs-lookup"><span data-stu-id="d8eac-1229">Error code</span></span>|<span data-ttu-id="d8eac-1230">说明</span><span class="sxs-lookup"><span data-stu-id="d8eac-1230">Description</span></span>|
|------------|-------------|
|`InvalidAttachmentId`|<span data-ttu-id="d8eac-1231">附件标识符不存在。</span><span class="sxs-lookup"><span data-stu-id="d8eac-1231">The attachment identifier does not exist.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d8eac-1232">要求</span><span class="sxs-lookup"><span data-stu-id="d8eac-1232">Requirements</span></span>

|<span data-ttu-id="d8eac-1233">要求</span><span class="sxs-lookup"><span data-stu-id="d8eac-1233">Requirement</span></span>|<span data-ttu-id="d8eac-1234">值</span><span class="sxs-lookup"><span data-stu-id="d8eac-1234">Value</span></span>|
|---|---|
|[<span data-ttu-id="d8eac-1235">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="d8eac-1235">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d8eac-1236">1.1</span><span class="sxs-lookup"><span data-stu-id="d8eac-1236">1.1</span></span>|
|[<span data-ttu-id="d8eac-1237">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="d8eac-1237">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d8eac-1238">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="d8eac-1238">ReadWriteItem</span></span>|
|[<span data-ttu-id="d8eac-1239">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="d8eac-1239">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="d8eac-1240">撰写</span><span class="sxs-lookup"><span data-stu-id="d8eac-1240">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="d8eac-1241">示例</span><span class="sxs-lookup"><span data-stu-id="d8eac-1241">Example</span></span>

<span data-ttu-id="d8eac-1242">以下代码删除包含标识符 '0' 的附件。</span><span class="sxs-lookup"><span data-stu-id="d8eac-1242">The following code removes an attachment with an identifier of '0'.</span></span>

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

#### <a name="removehandlerasynceventtype-options-callback"></a><span data-ttu-id="d8eac-1243">removeHandlerAsync(eventType, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="d8eac-1243">removeHandlerAsync(eventType, [options], [callback])</span></span>

<span data-ttu-id="d8eac-1244">删除受支持事件类型的事件处理程序。</span><span class="sxs-lookup"><span data-stu-id="d8eac-1244">Removes the event handlers for a supported event type.</span></span>

<span data-ttu-id="d8eac-1245">目前，受支持的事件`Office.EventType.AppointmentTimeChanged`类型`Office.EventType.RecipientsChanged`是、和`Office.EventType.RecurrenceChanged`</span><span class="sxs-lookup"><span data-stu-id="d8eac-1245">Currently the supported event types are `Office.EventType.AppointmentTimeChanged`, `Office.EventType.RecipientsChanged`, and `Office.EventType.RecurrenceChanged`</span></span>

##### <a name="parameters"></a><span data-ttu-id="d8eac-1246">Parameters</span><span class="sxs-lookup"><span data-stu-id="d8eac-1246">Parameters</span></span>

| <span data-ttu-id="d8eac-1247">名称</span><span class="sxs-lookup"><span data-stu-id="d8eac-1247">Name</span></span> | <span data-ttu-id="d8eac-1248">类型</span><span class="sxs-lookup"><span data-stu-id="d8eac-1248">Type</span></span> | <span data-ttu-id="d8eac-1249">属性</span><span class="sxs-lookup"><span data-stu-id="d8eac-1249">Attributes</span></span> | <span data-ttu-id="d8eac-1250">说明</span><span class="sxs-lookup"><span data-stu-id="d8eac-1250">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="d8eac-1251">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="d8eac-1251">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="d8eac-1252">应调用处理程序的事件。</span><span class="sxs-lookup"><span data-stu-id="d8eac-1252">The event that should invoke the handler.</span></span> |
| `options` | <span data-ttu-id="d8eac-1253">对象</span><span class="sxs-lookup"><span data-stu-id="d8eac-1253">Object</span></span> | <span data-ttu-id="d8eac-1254">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="d8eac-1254">&lt;optional&gt;</span></span> | <span data-ttu-id="d8eac-1255">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="d8eac-1255">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="d8eac-1256">对象</span><span class="sxs-lookup"><span data-stu-id="d8eac-1256">Object</span></span> | <span data-ttu-id="d8eac-1257">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="d8eac-1257">&lt;optional&gt;</span></span> | <span data-ttu-id="d8eac-1258">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="d8eac-1258">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="d8eac-1259">函数</span><span class="sxs-lookup"><span data-stu-id="d8eac-1259">function</span></span>| <span data-ttu-id="d8eac-1260">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="d8eac-1260">&lt;optional&gt;</span></span>|<span data-ttu-id="d8eac-1261">方法完成后，使用单个参数 `callback`（一个 [`asyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `AsyncResult` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="d8eac-1261">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d8eac-1262">Requirements</span><span class="sxs-lookup"><span data-stu-id="d8eac-1262">Requirements</span></span>

|<span data-ttu-id="d8eac-1263">要求</span><span class="sxs-lookup"><span data-stu-id="d8eac-1263">Requirement</span></span>| <span data-ttu-id="d8eac-1264">值</span><span class="sxs-lookup"><span data-stu-id="d8eac-1264">Value</span></span>|
|---|---|
|[<span data-ttu-id="d8eac-1265">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="d8eac-1265">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d8eac-1266">1.7</span><span class="sxs-lookup"><span data-stu-id="d8eac-1266">1.7</span></span> |
|[<span data-ttu-id="d8eac-1267">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="d8eac-1267">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d8eac-1268">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d8eac-1268">ReadItem</span></span> |
|[<span data-ttu-id="d8eac-1269">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="d8eac-1269">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d8eac-1270">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="d8eac-1270">Compose or Read</span></span> |

##### <a name="example"></a><span data-ttu-id="d8eac-1271">示例</span><span class="sxs-lookup"><span data-stu-id="d8eac-1271">Example</span></span>

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

#### <a name="saveasyncoptions-callback"></a><span data-ttu-id="d8eac-1272">saveAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="d8eac-1272">saveAsync([options], callback)</span></span>

<span data-ttu-id="d8eac-1273">异步保存项目。</span><span class="sxs-lookup"><span data-stu-id="d8eac-1273">Asynchronously saves an item.</span></span>

<span data-ttu-id="d8eac-1274">调用时，此方法将当前邮件保存为草稿，并通过回调方法返回项目 ID。</span><span class="sxs-lookup"><span data-stu-id="d8eac-1274">When invoked, this method saves the current message as a draft and returns the item id via the callback method.</span></span> <span data-ttu-id="d8eac-1275">在 Outlook 网页版或 Outlook 联机模式下，该项目被保存到服务器中。</span><span class="sxs-lookup"><span data-stu-id="d8eac-1275">In Outlook on the web or Outlook in online mode, the item is saved to the server.</span></span> <span data-ttu-id="d8eac-1276">在 Outlook 缓存模式下，该项目被保存到本地缓存中。</span><span class="sxs-lookup"><span data-stu-id="d8eac-1276">In Outlook in cached mode, the item is saved to the local cache.</span></span>

> [!NOTE]
> <span data-ttu-id="d8eac-1277">如果加载项在撰写模式下对某个项目调用 `saveAsync` 来获得 `itemId`，以便与 EWS 或 REST API 一同使用，请注意，当 Outlook 处于高速缓存模式时，可能需要一段时间项目才能真正同步到服务器。</span><span class="sxs-lookup"><span data-stu-id="d8eac-1277">If your add-in calls `saveAsync` on an item in compose mode in order to get an `itemId` to use with EWS or the REST API, be aware that when Outlook is in cached mode, it may take some time before the item is actually synced to the server.</span></span> <span data-ttu-id="d8eac-1278">在项目同步前，使用 `itemId` 将返回一个错误。</span><span class="sxs-lookup"><span data-stu-id="d8eac-1278">Until the item is synced, using the `itemId` will return an error.</span></span>

<span data-ttu-id="d8eac-p188">由于约会没有草稿状态，如果以撰写模式在约会中调用 `saveAsync`，则该项将被保存为用户日历中的正常约会。对于之前未保存过的新约会，则不会发送邀请。保存现有约会将向添加或删除的与会者发送更新。</span><span class="sxs-lookup"><span data-stu-id="d8eac-p188">Since appointments have no draft state, if `saveAsync` is called on an appointment in compose mode, the item will be saved as a normal appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent. Saving an existing appointment will send an update to added or removed attendees.</span></span>

> [!NOTE]
> <span data-ttu-id="d8eac-1282">以下客户端在撰写模式下对约会上的 `saveAsync` 具有不同的行为：</span><span class="sxs-lookup"><span data-stu-id="d8eac-1282">The following clients have different behavior for `saveAsync` on appointments in compose mode:</span></span>
>
> - <span data-ttu-id="d8eac-1283">Mac 版 Outlook 不支持保存会议。</span><span class="sxs-lookup"><span data-stu-id="d8eac-1283">Outlook on Mac does not support saving a meeting.</span></span> <span data-ttu-id="d8eac-1284">在撰写模式下，无法从会议调用 `saveAsync` 方法。</span><span class="sxs-lookup"><span data-stu-id="d8eac-1284">The `saveAsync` method fails when called from a meeting in compose mode.</span></span> <span data-ttu-id="d8eac-1285">若需解决办法，请参阅[无法在 Outlook for Mac 中使用 Office JS API 将会议另存为草稿](https://support.microsoft.com/help/4505745)。</span><span class="sxs-lookup"><span data-stu-id="d8eac-1285">See [Cannot save a meeting as a draft in Outlook for Mac by using Office JS API](https://support.microsoft.com/help/4505745) for a workaround.</span></span>
> - <span data-ttu-id="d8eac-1286">在撰写模式下的约会上调用 `saveAsync` 时，Outlook 网页版始终发送邀请或更新。</span><span class="sxs-lookup"><span data-stu-id="d8eac-1286">Outlook on the web always sends an invitation or update when `saveAsync` is called on an appointment in compose mode.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d8eac-1287">参数</span><span class="sxs-lookup"><span data-stu-id="d8eac-1287">Parameters</span></span>

|<span data-ttu-id="d8eac-1288">名称</span><span class="sxs-lookup"><span data-stu-id="d8eac-1288">Name</span></span>|<span data-ttu-id="d8eac-1289">类型</span><span class="sxs-lookup"><span data-stu-id="d8eac-1289">Type</span></span>|<span data-ttu-id="d8eac-1290">属性</span><span class="sxs-lookup"><span data-stu-id="d8eac-1290">Attributes</span></span>|<span data-ttu-id="d8eac-1291">说明</span><span class="sxs-lookup"><span data-stu-id="d8eac-1291">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="d8eac-1292">对象</span><span class="sxs-lookup"><span data-stu-id="d8eac-1292">Object</span></span>|<span data-ttu-id="d8eac-1293">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="d8eac-1293">&lt;optional&gt;</span></span>|<span data-ttu-id="d8eac-1294">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="d8eac-1294">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="d8eac-1295">对象</span><span class="sxs-lookup"><span data-stu-id="d8eac-1295">Object</span></span>|<span data-ttu-id="d8eac-1296">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="d8eac-1296">&lt;optional&gt;</span></span>|<span data-ttu-id="d8eac-1297">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="d8eac-1297">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="d8eac-1298">函数</span><span class="sxs-lookup"><span data-stu-id="d8eac-1298">function</span></span>||<span data-ttu-id="d8eac-1299">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="d8eac-1299">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="d8eac-1300">如果成功，该项目标识符将在 `asyncResult.value` 属性中提供。</span><span class="sxs-lookup"><span data-stu-id="d8eac-1300">On success, the item identifier is provided in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d8eac-1301">要求</span><span class="sxs-lookup"><span data-stu-id="d8eac-1301">Requirements</span></span>

|<span data-ttu-id="d8eac-1302">要求</span><span class="sxs-lookup"><span data-stu-id="d8eac-1302">Requirement</span></span>|<span data-ttu-id="d8eac-1303">值</span><span class="sxs-lookup"><span data-stu-id="d8eac-1303">Value</span></span>|
|---|---|
|[<span data-ttu-id="d8eac-1304">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="d8eac-1304">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d8eac-1305">1.3</span><span class="sxs-lookup"><span data-stu-id="d8eac-1305">1.3</span></span>|
|[<span data-ttu-id="d8eac-1306">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="d8eac-1306">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d8eac-1307">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="d8eac-1307">ReadWriteItem</span></span>|
|[<span data-ttu-id="d8eac-1308">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="d8eac-1308">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="d8eac-1309">撰写</span><span class="sxs-lookup"><span data-stu-id="d8eac-1309">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="d8eac-1310">示例</span><span class="sxs-lookup"><span data-stu-id="d8eac-1310">Examples</span></span>

```js
Office.context.mailbox.item.saveAsync(
  function callback(result) {
    // Process the result.
  });
```

<span data-ttu-id="d8eac-p190">下面是传递给回调函数的 `result` 参数的示例。`value` 属性包含的项目的项目 ID。</span><span class="sxs-lookup"><span data-stu-id="d8eac-p190">The following is an example of the `result` parameter passed to the callback function. The `value` property contains the item ID of the item.</span></span>

```json
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

<br>

---
---

#### <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="d8eac-1313">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="d8eac-1313">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="d8eac-1314">以异步方式将数据插入到邮件的正文或主题中。</span><span class="sxs-lookup"><span data-stu-id="d8eac-1314">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="d8eac-p191">`setSelectedDataAsync` 方法将指定字符串插入到项目主题或正文的光标位置，或者，如果在编辑器中已选择文本，则该方法将替换选择的文本。如果光标不在正文或主题字段中，则返回错误。插入之后，光标会位于插入内容的末尾。</span><span class="sxs-lookup"><span data-stu-id="d8eac-p191">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d8eac-1318">参数</span><span class="sxs-lookup"><span data-stu-id="d8eac-1318">Parameters</span></span>

|<span data-ttu-id="d8eac-1319">名称</span><span class="sxs-lookup"><span data-stu-id="d8eac-1319">Name</span></span>|<span data-ttu-id="d8eac-1320">类型</span><span class="sxs-lookup"><span data-stu-id="d8eac-1320">Type</span></span>|<span data-ttu-id="d8eac-1321">属性</span><span class="sxs-lookup"><span data-stu-id="d8eac-1321">Attributes</span></span>|<span data-ttu-id="d8eac-1322">说明</span><span class="sxs-lookup"><span data-stu-id="d8eac-1322">Description</span></span>|
|---|---|---|---|
|`data`|<span data-ttu-id="d8eac-1323">字符串</span><span class="sxs-lookup"><span data-stu-id="d8eac-1323">String</span></span>||<span data-ttu-id="d8eac-p192">要插入的数据。数据不得超过 1,000,000 个字符。如果传入的数据超过 1,000,000 个字符，则会引发 `ArgumentOutOfRange` 异常。</span><span class="sxs-lookup"><span data-stu-id="d8eac-p192">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`|<span data-ttu-id="d8eac-1327">Object</span><span class="sxs-lookup"><span data-stu-id="d8eac-1327">Object</span></span>|<span data-ttu-id="d8eac-1328">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="d8eac-1328">&lt;optional&gt;</span></span>|<span data-ttu-id="d8eac-1329">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="d8eac-1329">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="d8eac-1330">对象</span><span class="sxs-lookup"><span data-stu-id="d8eac-1330">Object</span></span>|<span data-ttu-id="d8eac-1331">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="d8eac-1331">&lt;optional&gt;</span></span>|<span data-ttu-id="d8eac-1332">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="d8eac-1332">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`|[<span data-ttu-id="d8eac-1333">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="d8eac-1333">Office.CoercionType</span></span>](office.md#coerciontype-string)|<span data-ttu-id="d8eac-1334">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="d8eac-1334">&lt;optional&gt;</span></span>|<span data-ttu-id="d8eac-1335">如果为 `text`，系统在 Outlook 网页版和 Outlook 桌面版客户端中应用当前样式。</span><span class="sxs-lookup"><span data-stu-id="d8eac-1335">If `text`, the current style is applied in Outlook on the web and desktop clients.</span></span> <span data-ttu-id="d8eac-1336">如果字段是 HTML 编辑器，只会插入文本数据，即使数据为 HTML，也不例外。</span><span class="sxs-lookup"><span data-stu-id="d8eac-1336">If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="d8eac-1337">如果 `html` 和字段支持 HTML（主题不支持），系统在 Outlook 网页版中应用当前样式，而在 Outlook 桌面版客户端中则应用默认样式。</span><span class="sxs-lookup"><span data-stu-id="d8eac-1337">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook on the web and the default style is applied in Outlook desktop clients.</span></span> <span data-ttu-id="d8eac-1338">如果该字段是文本字段，则返回 `InvalidDataFormat` 错误。</span><span class="sxs-lookup"><span data-stu-id="d8eac-1338">If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="d8eac-1339">如果未设置 `coercionType`，则结果取决于该字段：如果该字段是 HTML，则使用 HTML；如果该字段是文本，则使用纯文本。</span><span class="sxs-lookup"><span data-stu-id="d8eac-1339">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`|<span data-ttu-id="d8eac-1340">function</span><span class="sxs-lookup"><span data-stu-id="d8eac-1340">function</span></span>||<span data-ttu-id="d8eac-1341">方法完成后，使用单个参数 `callback`（一个 [`asyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `AsyncResult` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="d8eac-1341">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d8eac-1342">Requirements</span><span class="sxs-lookup"><span data-stu-id="d8eac-1342">Requirements</span></span>

|<span data-ttu-id="d8eac-1343">要求</span><span class="sxs-lookup"><span data-stu-id="d8eac-1343">Requirement</span></span>|<span data-ttu-id="d8eac-1344">值</span><span class="sxs-lookup"><span data-stu-id="d8eac-1344">Value</span></span>|
|---|---|
|[<span data-ttu-id="d8eac-1345">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="d8eac-1345">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d8eac-1346">1.2</span><span class="sxs-lookup"><span data-stu-id="d8eac-1346">1.2</span></span>|
|[<span data-ttu-id="d8eac-1347">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="d8eac-1347">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d8eac-1348">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="d8eac-1348">ReadWriteItem</span></span>|
|[<span data-ttu-id="d8eac-1349">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="d8eac-1349">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)|<span data-ttu-id="d8eac-1350">撰写</span><span class="sxs-lookup"><span data-stu-id="d8eac-1350">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="d8eac-1351">示例</span><span class="sxs-lookup"><span data-stu-id="d8eac-1351">Example</span></span>

```js
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```
