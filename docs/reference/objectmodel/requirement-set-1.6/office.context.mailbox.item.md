---
title: Office.context.mailbox.item-要求设置 1.6
description: ''
ms.date: 01/30/2019
localization_priority: Normal
ms.openlocfilehash: 0c3eca68285e9d415954e6ce45d2a80508fa701b
ms.sourcegitcommit: bf5c56d9b8c573e42bf2268e10ca3fd4d2bb4ff9
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 02/01/2019
ms.locfileid: "29701895"
---
# <a name="item"></a><span data-ttu-id="09fac-102">item</span><span class="sxs-lookup"><span data-stu-id="09fac-102">item</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmditem"></a><span data-ttu-id="09fac-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span><span class="sxs-lookup"><span data-stu-id="09fac-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span></span>

<span data-ttu-id="09fac-p101">`item` 命名空间用于访问当前选定的邮件、会议请求或约会。可以通过使用 [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlook16officemailboxenumsitemtype) 属性确定 `item` 的类型。</span><span class="sxs-lookup"><span data-stu-id="09fac-p101">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlook16officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="09fac-106">Requirements</span><span class="sxs-lookup"><span data-stu-id="09fac-106">Requirements</span></span>

|<span data-ttu-id="09fac-107">要求</span><span class="sxs-lookup"><span data-stu-id="09fac-107">Requirement</span></span>| <span data-ttu-id="09fac-108">值</span><span class="sxs-lookup"><span data-stu-id="09fac-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="09fac-109">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="09fac-109">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="09fac-110">1.0</span><span class="sxs-lookup"><span data-stu-id="09fac-110">1.0</span></span>|
|[<span data-ttu-id="09fac-111">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="09fac-111">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="09fac-112">受限</span><span class="sxs-lookup"><span data-stu-id="09fac-112">Restricted</span></span>|
|[<span data-ttu-id="09fac-113">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="09fac-113">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="09fac-114">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="09fac-114">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="09fac-115">成员和方法</span><span class="sxs-lookup"><span data-stu-id="09fac-115">Members and methods</span></span>

| <span data-ttu-id="09fac-116">成员</span><span class="sxs-lookup"><span data-stu-id="09fac-116">Member</span></span> | <span data-ttu-id="09fac-117">类型</span><span class="sxs-lookup"><span data-stu-id="09fac-117">Type</span></span> |
|--------|------|
| [<span data-ttu-id="09fac-118">attachments</span><span class="sxs-lookup"><span data-stu-id="09fac-118">attachments</span></span>](#attachments-arrayattachmentdetailsjavascriptapioutlook16officeattachmentdetails) | <span data-ttu-id="09fac-119">成员</span><span class="sxs-lookup"><span data-stu-id="09fac-119">Member</span></span> |
| [<span data-ttu-id="09fac-120">bcc</span><span class="sxs-lookup"><span data-stu-id="09fac-120">bcc</span></span>](#bcc-recipientsjavascriptapioutlook16officerecipients) | <span data-ttu-id="09fac-121">成员</span><span class="sxs-lookup"><span data-stu-id="09fac-121">Member</span></span> |
| [<span data-ttu-id="09fac-122">body</span><span class="sxs-lookup"><span data-stu-id="09fac-122">body</span></span>](#body-bodyjavascriptapioutlook16officebody) | <span data-ttu-id="09fac-123">成员</span><span class="sxs-lookup"><span data-stu-id="09fac-123">Member</span></span> |
| [<span data-ttu-id="09fac-124">cc</span><span class="sxs-lookup"><span data-stu-id="09fac-124">cc</span></span>](#cc-arrayemailaddressdetailsjavascriptapioutlook16officeemailaddressdetailsrecipientsjavascriptapioutlook16officerecipients) | <span data-ttu-id="09fac-125">成员</span><span class="sxs-lookup"><span data-stu-id="09fac-125">Member</span></span> |
| [<span data-ttu-id="09fac-126">conversationId</span><span class="sxs-lookup"><span data-stu-id="09fac-126">conversationId</span></span>](#nullable-conversationid-string) | <span data-ttu-id="09fac-127">成员</span><span class="sxs-lookup"><span data-stu-id="09fac-127">Member</span></span> |
| [<span data-ttu-id="09fac-128">dateTimeCreated</span><span class="sxs-lookup"><span data-stu-id="09fac-128">dateTimeCreated</span></span>](#datetimecreated-date) | <span data-ttu-id="09fac-129">成员</span><span class="sxs-lookup"><span data-stu-id="09fac-129">Member</span></span> |
| [<span data-ttu-id="09fac-130">dateTimeModified</span><span class="sxs-lookup"><span data-stu-id="09fac-130">dateTimeModified</span></span>](#datetimemodified-date) | <span data-ttu-id="09fac-131">成员</span><span class="sxs-lookup"><span data-stu-id="09fac-131">Member</span></span> |
| [<span data-ttu-id="09fac-132">end</span><span class="sxs-lookup"><span data-stu-id="09fac-132">end</span></span>](#end-datetimejavascriptapioutlook16officetime) | <span data-ttu-id="09fac-133">成员</span><span class="sxs-lookup"><span data-stu-id="09fac-133">Member</span></span> |
| [<span data-ttu-id="09fac-134">from</span><span class="sxs-lookup"><span data-stu-id="09fac-134">from</span></span>](#from-emailaddressdetailsjavascriptapioutlook16officeemailaddressdetails) | <span data-ttu-id="09fac-135">成员</span><span class="sxs-lookup"><span data-stu-id="09fac-135">Member</span></span> |
| [<span data-ttu-id="09fac-136">internetMessageId</span><span class="sxs-lookup"><span data-stu-id="09fac-136">internetMessageId</span></span>](#internetmessageid-string) | <span data-ttu-id="09fac-137">成员</span><span class="sxs-lookup"><span data-stu-id="09fac-137">Member</span></span> |
| [<span data-ttu-id="09fac-138">itemClass</span><span class="sxs-lookup"><span data-stu-id="09fac-138">itemClass</span></span>](#itemclass-string) | <span data-ttu-id="09fac-139">成员</span><span class="sxs-lookup"><span data-stu-id="09fac-139">Member</span></span> |
| [<span data-ttu-id="09fac-140">itemId</span><span class="sxs-lookup"><span data-stu-id="09fac-140">itemId</span></span>](#nullable-itemid-string) | <span data-ttu-id="09fac-141">成员</span><span class="sxs-lookup"><span data-stu-id="09fac-141">Member</span></span> |
| [<span data-ttu-id="09fac-142">itemType</span><span class="sxs-lookup"><span data-stu-id="09fac-142">itemType</span></span>](#itemtype-officemailboxenumsitemtypejavascriptapioutlook16officemailboxenumsitemtype) | <span data-ttu-id="09fac-143">成员</span><span class="sxs-lookup"><span data-stu-id="09fac-143">Member</span></span> |
| [<span data-ttu-id="09fac-144">location</span><span class="sxs-lookup"><span data-stu-id="09fac-144">location</span></span>](#location-stringlocationjavascriptapioutlook16officelocation) | <span data-ttu-id="09fac-145">成员</span><span class="sxs-lookup"><span data-stu-id="09fac-145">Member</span></span> |
| [<span data-ttu-id="09fac-146">normalizedSubject</span><span class="sxs-lookup"><span data-stu-id="09fac-146">normalizedSubject</span></span>](#normalizedsubject-string) | <span data-ttu-id="09fac-147">成员</span><span class="sxs-lookup"><span data-stu-id="09fac-147">Member</span></span> |
| [<span data-ttu-id="09fac-148">notificationMessages</span><span class="sxs-lookup"><span data-stu-id="09fac-148">notificationMessages</span></span>](#notificationmessages-notificationmessagesjavascriptapioutlook16officenotificationmessages) | <span data-ttu-id="09fac-149">成员</span><span class="sxs-lookup"><span data-stu-id="09fac-149">Member</span></span> |
| [<span data-ttu-id="09fac-150">optionalAttendees</span><span class="sxs-lookup"><span data-stu-id="09fac-150">optionalAttendees</span></span>](#optionalattendees-arrayemailaddressdetailsjavascriptapioutlook16officeemailaddressdetailsrecipientsjavascriptapioutlook16officerecipients) | <span data-ttu-id="09fac-151">成员</span><span class="sxs-lookup"><span data-stu-id="09fac-151">Member</span></span> |
| [<span data-ttu-id="09fac-152">organizer</span><span class="sxs-lookup"><span data-stu-id="09fac-152">organizer</span></span>](#organizer-emailaddressdetailsjavascriptapioutlook16officeemailaddressdetails) | <span data-ttu-id="09fac-153">成员</span><span class="sxs-lookup"><span data-stu-id="09fac-153">Member</span></span> |
| [<span data-ttu-id="09fac-154">requiredAttendees</span><span class="sxs-lookup"><span data-stu-id="09fac-154">requiredAttendees</span></span>](#requiredattendees-arrayemailaddressdetailsjavascriptapioutlook16officeemailaddressdetailsrecipientsjavascriptapioutlook16officerecipients) | <span data-ttu-id="09fac-155">Member</span><span class="sxs-lookup"><span data-stu-id="09fac-155">Member</span></span> |
| [<span data-ttu-id="09fac-156">sender</span><span class="sxs-lookup"><span data-stu-id="09fac-156">sender</span></span>](#sender-emailaddressdetailsjavascriptapioutlook16officeemailaddressdetails) | <span data-ttu-id="09fac-157">成员</span><span class="sxs-lookup"><span data-stu-id="09fac-157">Member</span></span> |
| [<span data-ttu-id="09fac-158">start</span><span class="sxs-lookup"><span data-stu-id="09fac-158">start</span></span>](#start-datetimejavascriptapioutlook16officetime) | <span data-ttu-id="09fac-159">成员</span><span class="sxs-lookup"><span data-stu-id="09fac-159">Member</span></span> |
| [<span data-ttu-id="09fac-160">subject</span><span class="sxs-lookup"><span data-stu-id="09fac-160">subject</span></span>](#subject-stringsubjectjavascriptapioutlook16officesubject) | <span data-ttu-id="09fac-161">成员</span><span class="sxs-lookup"><span data-stu-id="09fac-161">Member</span></span> |
| [<span data-ttu-id="09fac-162">to</span><span class="sxs-lookup"><span data-stu-id="09fac-162">to</span></span>](#to-arrayemailaddressdetailsjavascriptapioutlook16officeemailaddressdetailsrecipientsjavascriptapioutlook16officerecipients) | <span data-ttu-id="09fac-163">成员</span><span class="sxs-lookup"><span data-stu-id="09fac-163">Member</span></span> |
| [<span data-ttu-id="09fac-164">addFileAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="09fac-164">addFileAttachmentAsync</span></span>](#addfileattachmentasyncuri-attachmentname-options-callback) | <span data-ttu-id="09fac-165">方法</span><span class="sxs-lookup"><span data-stu-id="09fac-165">Method</span></span> |
| [<span data-ttu-id="09fac-166">addItemAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="09fac-166">addItemAttachmentAsync</span></span>](#additemattachmentasyncitemid-attachmentname-options-callback) | <span data-ttu-id="09fac-167">方法</span><span class="sxs-lookup"><span data-stu-id="09fac-167">Method</span></span> |
| [<span data-ttu-id="09fac-168">close</span><span class="sxs-lookup"><span data-stu-id="09fac-168">close</span></span>](#close) | <span data-ttu-id="09fac-169">方法</span><span class="sxs-lookup"><span data-stu-id="09fac-169">Method</span></span> |
| [<span data-ttu-id="09fac-170">displayReplyAllForm</span><span class="sxs-lookup"><span data-stu-id="09fac-170">displayReplyAllForm</span></span>](#displayreplyallformformdata) | <span data-ttu-id="09fac-171">方法</span><span class="sxs-lookup"><span data-stu-id="09fac-171">Method</span></span> |
| [<span data-ttu-id="09fac-172">displayReplyForm</span><span class="sxs-lookup"><span data-stu-id="09fac-172">displayReplyForm</span></span>](#displayreplyformformdata) | <span data-ttu-id="09fac-173">方法</span><span class="sxs-lookup"><span data-stu-id="09fac-173">Method</span></span> |
| [<span data-ttu-id="09fac-174">getEntities</span><span class="sxs-lookup"><span data-stu-id="09fac-174">getEntities</span></span>](#getentities--entitiesjavascriptapioutlook16officeentities) | <span data-ttu-id="09fac-175">方法</span><span class="sxs-lookup"><span data-stu-id="09fac-175">Method</span></span> |
| [<span data-ttu-id="09fac-176">getEntitiesByType</span><span class="sxs-lookup"><span data-stu-id="09fac-176">getEntitiesByType</span></span>](#getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlook16officecontactmeetingsuggestionjavascriptapioutlook16officemeetingsuggestionphonenumberjavascriptapioutlook16officephonenumbertasksuggestionjavascriptapioutlook16officetasksuggestion) | <span data-ttu-id="09fac-177">方法</span><span class="sxs-lookup"><span data-stu-id="09fac-177">Method</span></span> |
| [<span data-ttu-id="09fac-178">getFilteredEntitiesByName</span><span class="sxs-lookup"><span data-stu-id="09fac-178">getFilteredEntitiesByName</span></span>](#getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlook16officecontactmeetingsuggestionjavascriptapioutlook16officemeetingsuggestionphonenumberjavascriptapioutlook16officephonenumbertasksuggestionjavascriptapioutlook16officetasksuggestion) | <span data-ttu-id="09fac-179">方法</span><span class="sxs-lookup"><span data-stu-id="09fac-179">Method</span></span> |
| [<span data-ttu-id="09fac-180">getRegExMatches</span><span class="sxs-lookup"><span data-stu-id="09fac-180">getRegExMatches</span></span>](#getregexmatches--object) | <span data-ttu-id="09fac-181">方法</span><span class="sxs-lookup"><span data-stu-id="09fac-181">Method</span></span> |
| [<span data-ttu-id="09fac-182">getRegExMatchesByName</span><span class="sxs-lookup"><span data-stu-id="09fac-182">getRegExMatchesByName</span></span>](#getregexmatchesbynamename--nullable-array-string-) | <span data-ttu-id="09fac-183">方法</span><span class="sxs-lookup"><span data-stu-id="09fac-183">Method</span></span> |
| [<span data-ttu-id="09fac-184">getSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="09fac-184">getSelectedDataAsync</span></span>](#getselecteddataasynccoerciontype-options-callback--string) | <span data-ttu-id="09fac-185">方法</span><span class="sxs-lookup"><span data-stu-id="09fac-185">Method</span></span> |
| [<span data-ttu-id="09fac-186">getSelectedEntities</span><span class="sxs-lookup"><span data-stu-id="09fac-186">getSelectedEntities</span></span>](#getselectedentities--entitiesjavascriptapioutlook16officeentities) | <span data-ttu-id="09fac-187">方法</span><span class="sxs-lookup"><span data-stu-id="09fac-187">Method</span></span> |
| [<span data-ttu-id="09fac-188">getSelectedRegExMatches</span><span class="sxs-lookup"><span data-stu-id="09fac-188">getSelectedRegExMatches</span></span>](#getselectedregexmatches--object) | <span data-ttu-id="09fac-189">方法</span><span class="sxs-lookup"><span data-stu-id="09fac-189">Method</span></span> |
| [<span data-ttu-id="09fac-190">loadCustomPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="09fac-190">loadCustomPropertiesAsync</span></span>](#loadcustompropertiesasynccallback-usercontext) | <span data-ttu-id="09fac-191">方法</span><span class="sxs-lookup"><span data-stu-id="09fac-191">Method</span></span> |
| [<span data-ttu-id="09fac-192">removeAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="09fac-192">removeAttachmentAsync</span></span>](#removeattachmentasyncattachmentid-options-callback) | <span data-ttu-id="09fac-193">方法</span><span class="sxs-lookup"><span data-stu-id="09fac-193">Method</span></span> |
| [<span data-ttu-id="09fac-194">saveAsync</span><span class="sxs-lookup"><span data-stu-id="09fac-194">saveAsync</span></span>](#saveasyncoptions-callback) | <span data-ttu-id="09fac-195">方法</span><span class="sxs-lookup"><span data-stu-id="09fac-195">Method</span></span> |
| [<span data-ttu-id="09fac-196">setSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="09fac-196">setSelectedDataAsync</span></span>](#setselecteddataasyncdata-options-callback) | <span data-ttu-id="09fac-197">方法</span><span class="sxs-lookup"><span data-stu-id="09fac-197">Method</span></span> |

### <a name="example"></a><span data-ttu-id="09fac-198">示例</span><span class="sxs-lookup"><span data-stu-id="09fac-198">Example</span></span>

<span data-ttu-id="09fac-199">以下 JavaScript 代码示例显示了如何访问 Outlook 中当前项目的 `subject` 属性。</span><span class="sxs-lookup"><span data-stu-id="09fac-199">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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

### <a name="members"></a><span data-ttu-id="09fac-200">成员</span><span class="sxs-lookup"><span data-stu-id="09fac-200">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlook16officeattachmentdetails"></a><span data-ttu-id="09fac-201">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_6/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="09fac-201">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_6/office.attachmentdetails)></span></span>

<span data-ttu-id="09fac-p102">获取项目的附件数组。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="09fac-p102">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="09fac-204">由于潜在的安全问题，某些类型的文件会受到 Outlook 阻止，并且不会返回。</span><span class="sxs-lookup"><span data-stu-id="09fac-204">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="09fac-205">如需了解更多信息，请参阅 [Outlook 中阻止的附件](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519)。</span><span class="sxs-lookup"><span data-stu-id="09fac-205">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="09fac-206">类型：</span><span class="sxs-lookup"><span data-stu-id="09fac-206">Type:</span></span>

*   <span data-ttu-id="09fac-207">Array.<[AttachmentDetails](/javascript/api/outlook_1_6/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="09fac-207">Array.<[AttachmentDetails](/javascript/api/outlook_1_6/office.attachmentdetails)></span></span>

##### <a name="requirements"></a><span data-ttu-id="09fac-208">Requirements</span><span class="sxs-lookup"><span data-stu-id="09fac-208">Requirements</span></span>

|<span data-ttu-id="09fac-209">要求</span><span class="sxs-lookup"><span data-stu-id="09fac-209">Requirement</span></span>| <span data-ttu-id="09fac-210">值</span><span class="sxs-lookup"><span data-stu-id="09fac-210">Value</span></span>|
|---|---|
|[<span data-ttu-id="09fac-211">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="09fac-211">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="09fac-212">1.0</span><span class="sxs-lookup"><span data-stu-id="09fac-212">1.0</span></span>|
|[<span data-ttu-id="09fac-213">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="09fac-213">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="09fac-214">ReadItem</span><span class="sxs-lookup"><span data-stu-id="09fac-214">ReadItem</span></span>|
|[<span data-ttu-id="09fac-215">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="09fac-215">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="09fac-216">阅读</span><span class="sxs-lookup"><span data-stu-id="09fac-216">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="09fac-217">示例</span><span class="sxs-lookup"><span data-stu-id="09fac-217">Example</span></span>

<span data-ttu-id="09fac-218">以下代码使用当前项目上所有附件的详细信息构成 HTML 字符串。</span><span class="sxs-lookup"><span data-stu-id="09fac-218">The following code builds an HTML string with details of all attachments on the current item.</span></span>

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

####  <a name="bcc-recipientsjavascriptapioutlook16officerecipients"></a><span data-ttu-id="09fac-219">bcc :[Recipients](/javascript/api/outlook_1_6/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="09fac-219">bcc :[Recipients](/javascript/api/outlook_1_6/office.recipients)</span></span>

<span data-ttu-id="09fac-220">获取对象，该对象提供用于获取或更新邮件的密件抄送 (Bcc) 行上的收件人的方法。</span><span class="sxs-lookup"><span data-stu-id="09fac-220">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="09fac-221">仅限撰写模式。</span><span class="sxs-lookup"><span data-stu-id="09fac-221">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="09fac-222">类型:</span><span class="sxs-lookup"><span data-stu-id="09fac-222">Type:</span></span>

*   [<span data-ttu-id="09fac-223">收件人</span><span class="sxs-lookup"><span data-stu-id="09fac-223">Recipients</span></span>](/javascript/api/outlook_1_6/office.recipients)

##### <a name="requirements"></a><span data-ttu-id="09fac-224">Requirements</span><span class="sxs-lookup"><span data-stu-id="09fac-224">Requirements</span></span>

|<span data-ttu-id="09fac-225">要求</span><span class="sxs-lookup"><span data-stu-id="09fac-225">Requirement</span></span>| <span data-ttu-id="09fac-226">值</span><span class="sxs-lookup"><span data-stu-id="09fac-226">Value</span></span>|
|---|---|
|[<span data-ttu-id="09fac-227">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="09fac-227">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="09fac-228">1.1</span><span class="sxs-lookup"><span data-stu-id="09fac-228">1.1</span></span>|
|[<span data-ttu-id="09fac-229">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="09fac-229">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="09fac-230">ReadItem</span><span class="sxs-lookup"><span data-stu-id="09fac-230">ReadItem</span></span>|
|[<span data-ttu-id="09fac-231">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="09fac-231">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="09fac-232">撰写</span><span class="sxs-lookup"><span data-stu-id="09fac-232">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="09fac-233">示例</span><span class="sxs-lookup"><span data-stu-id="09fac-233">Example</span></span>

```js
Office.context.mailbox.item.bcc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.bcc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.bcc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfBccRecipients = asyncResult.value;
}
```

####  <a name="body-bodyjavascriptapioutlook16officebody"></a><span data-ttu-id="09fac-234">body :[Body](/javascript/api/outlook_1_6/office.body)</span><span class="sxs-lookup"><span data-stu-id="09fac-234">body :[Body](/javascript/api/outlook_1_6/office.body)</span></span>

<span data-ttu-id="09fac-235">获取一个提供用于处理项目正文的方法的对象。</span><span class="sxs-lookup"><span data-stu-id="09fac-235">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="09fac-236">类型：</span><span class="sxs-lookup"><span data-stu-id="09fac-236">Type:</span></span>

*   [<span data-ttu-id="09fac-237">Body</span><span class="sxs-lookup"><span data-stu-id="09fac-237">Body</span></span>](/javascript/api/outlook_1_6/office.body)

##### <a name="requirements"></a><span data-ttu-id="09fac-238">要求</span><span class="sxs-lookup"><span data-stu-id="09fac-238">Requirements</span></span>

|<span data-ttu-id="09fac-239">要求</span><span class="sxs-lookup"><span data-stu-id="09fac-239">Requirement</span></span>| <span data-ttu-id="09fac-240">值</span><span class="sxs-lookup"><span data-stu-id="09fac-240">Value</span></span>|
|---|---|
|[<span data-ttu-id="09fac-241">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="09fac-241">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="09fac-242">1.1</span><span class="sxs-lookup"><span data-stu-id="09fac-242">1.1</span></span>|
|[<span data-ttu-id="09fac-243">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="09fac-243">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="09fac-244">ReadItem</span><span class="sxs-lookup"><span data-stu-id="09fac-244">ReadItem</span></span>|
|[<span data-ttu-id="09fac-245">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="09fac-245">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="09fac-246">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="09fac-246">Compose or read</span></span>|

####  <a name="cc-arrayemailaddressdetailsjavascriptapioutlook16officeemailaddressdetailsrecipientsjavascriptapioutlook16officerecipients"></a><span data-ttu-id="09fac-247">cc :Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_6/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="09fac-247">cc :Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_6/office.recipients)</span></span>

<span data-ttu-id="09fac-248">提供对邮件的抄送 (Cc) 收件人的访问权限。</span><span class="sxs-lookup"><span data-stu-id="09fac-248">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="09fac-249">对象的类型和访问级别取决于当前项目的模式。</span><span class="sxs-lookup"><span data-stu-id="09fac-249">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="09fac-250">阅读模式</span><span class="sxs-lookup"><span data-stu-id="09fac-250">Read mode</span></span>

<span data-ttu-id="09fac-p106">`cc` 属性返回包含邮件的**抄送**行上所列的每个收件人的 `EmailAddressDetails` 对象的数组。集合上限为 100 个成员。</span><span class="sxs-lookup"><span data-stu-id="09fac-p106">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="09fac-253">撰写模式</span><span class="sxs-lookup"><span data-stu-id="09fac-253">Compose mode</span></span>

<span data-ttu-id="09fac-254">`cc` 属性返回一个 `Recipients` 对象，该对象提供用于获取或更新邮件的“**抄送**”行上收件人的方法。</span><span class="sxs-lookup"><span data-stu-id="09fac-254">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="09fac-255">类型：</span><span class="sxs-lookup"><span data-stu-id="09fac-255">Type:</span></span>

*   <span data-ttu-id="09fac-256">Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_6/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="09fac-256">Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_6/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="09fac-257">要求</span><span class="sxs-lookup"><span data-stu-id="09fac-257">Requirements</span></span>

|<span data-ttu-id="09fac-258">要求</span><span class="sxs-lookup"><span data-stu-id="09fac-258">Requirement</span></span>| <span data-ttu-id="09fac-259">值</span><span class="sxs-lookup"><span data-stu-id="09fac-259">Value</span></span>|
|---|---|
|[<span data-ttu-id="09fac-260">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="09fac-260">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="09fac-261">1.0</span><span class="sxs-lookup"><span data-stu-id="09fac-261">1.0</span></span>|
|[<span data-ttu-id="09fac-262">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="09fac-262">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="09fac-263">ReadItem</span><span class="sxs-lookup"><span data-stu-id="09fac-263">ReadItem</span></span>|
|[<span data-ttu-id="09fac-264">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="09fac-264">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="09fac-265">Compose 或 Read</span><span class="sxs-lookup"><span data-stu-id="09fac-265">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="09fac-266">示例</span><span class="sxs-lookup"><span data-stu-id="09fac-266">Example</span></span>

```js
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

####  <a name="nullable-conversationid-string"></a><span data-ttu-id="09fac-267">(nullable) conversationId :String</span><span class="sxs-lookup"><span data-stu-id="09fac-267">(nullable) conversationId :String</span></span>

<span data-ttu-id="09fac-268">获取包含特定消息的电子邮件会话的标识符。</span><span class="sxs-lookup"><span data-stu-id="09fac-268">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="09fac-p107">如果在阅读窗体或撰写窗体的回复中激活邮件应用程序，则此属性可以获得一个整数值。如果用户随后更改了回复邮件的主题（若发送回复），则该邮件的对话 ID 将改变且之前获取的值将不适用。</span><span class="sxs-lookup"><span data-stu-id="09fac-p107">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="09fac-p108">对于撰写窗体的新项目，此属性获得一个 null 值。如果用户设置一个主题并保存该项目，`conversationId` 属性将返回一个值。</span><span class="sxs-lookup"><span data-stu-id="09fac-p108">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="09fac-273">类型：</span><span class="sxs-lookup"><span data-stu-id="09fac-273">Type:</span></span>

*   <span data-ttu-id="09fac-274">String</span><span class="sxs-lookup"><span data-stu-id="09fac-274">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="09fac-275">要求</span><span class="sxs-lookup"><span data-stu-id="09fac-275">Requirements</span></span>

|<span data-ttu-id="09fac-276">要求</span><span class="sxs-lookup"><span data-stu-id="09fac-276">Requirement</span></span>| <span data-ttu-id="09fac-277">值</span><span class="sxs-lookup"><span data-stu-id="09fac-277">Value</span></span>|
|---|---|
|[<span data-ttu-id="09fac-278">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="09fac-278">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="09fac-279">1.0</span><span class="sxs-lookup"><span data-stu-id="09fac-279">1.0</span></span>|
|[<span data-ttu-id="09fac-280">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="09fac-280">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="09fac-281">ReadItem</span><span class="sxs-lookup"><span data-stu-id="09fac-281">ReadItem</span></span>|
|[<span data-ttu-id="09fac-282">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="09fac-282">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="09fac-283">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="09fac-283">Compose or read</span></span>|

#### <a name="datetimecreated-date"></a><span data-ttu-id="09fac-284">dateTimeCreated :Date</span><span class="sxs-lookup"><span data-stu-id="09fac-284">dateTimeCreated :Date</span></span>

<span data-ttu-id="09fac-p109">获取项目创建的日期和时间。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="09fac-p109">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="09fac-287">类型：</span><span class="sxs-lookup"><span data-stu-id="09fac-287">Type:</span></span>

*   <span data-ttu-id="09fac-288">日期</span><span class="sxs-lookup"><span data-stu-id="09fac-288">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="09fac-289">Requirements</span><span class="sxs-lookup"><span data-stu-id="09fac-289">Requirements</span></span>

|<span data-ttu-id="09fac-290">要求</span><span class="sxs-lookup"><span data-stu-id="09fac-290">Requirement</span></span>| <span data-ttu-id="09fac-291">值</span><span class="sxs-lookup"><span data-stu-id="09fac-291">Value</span></span>|
|---|---|
|[<span data-ttu-id="09fac-292">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="09fac-292">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="09fac-293">1.0</span><span class="sxs-lookup"><span data-stu-id="09fac-293">1.0</span></span>|
|[<span data-ttu-id="09fac-294">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="09fac-294">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="09fac-295">ReadItem</span><span class="sxs-lookup"><span data-stu-id="09fac-295">ReadItem</span></span>|
|[<span data-ttu-id="09fac-296">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="09fac-296">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="09fac-297">阅读</span><span class="sxs-lookup"><span data-stu-id="09fac-297">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="09fac-298">示例</span><span class="sxs-lookup"><span data-stu-id="09fac-298">Example</span></span>

```js
var created = Office.context.mailbox.item.dateTimeCreated;
```

#### <a name="datetimemodified-date"></a><span data-ttu-id="09fac-299">dateTimeModified :Date</span><span class="sxs-lookup"><span data-stu-id="09fac-299">dateTimeModified :Date</span></span>

<span data-ttu-id="09fac-p110">获取项目最近一次修改的日期和时间。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="09fac-p110">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="09fac-302">在 Outlook for iOS 或 Outlook for Android 中不支持此成员。</span><span class="sxs-lookup"><span data-stu-id="09fac-302">This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="type"></a><span data-ttu-id="09fac-303">类型：</span><span class="sxs-lookup"><span data-stu-id="09fac-303">Type:</span></span>

*   <span data-ttu-id="09fac-304">日期</span><span class="sxs-lookup"><span data-stu-id="09fac-304">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="09fac-305">Requirements</span><span class="sxs-lookup"><span data-stu-id="09fac-305">Requirements</span></span>

|<span data-ttu-id="09fac-306">要求</span><span class="sxs-lookup"><span data-stu-id="09fac-306">Requirement</span></span>| <span data-ttu-id="09fac-307">值</span><span class="sxs-lookup"><span data-stu-id="09fac-307">Value</span></span>|
|---|---|
|[<span data-ttu-id="09fac-308">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="09fac-308">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="09fac-309">1.0</span><span class="sxs-lookup"><span data-stu-id="09fac-309">1.0</span></span>|
|[<span data-ttu-id="09fac-310">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="09fac-310">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="09fac-311">ReadItem</span><span class="sxs-lookup"><span data-stu-id="09fac-311">ReadItem</span></span>|
|[<span data-ttu-id="09fac-312">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="09fac-312">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="09fac-313">阅读</span><span class="sxs-lookup"><span data-stu-id="09fac-313">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="09fac-314">示例</span><span class="sxs-lookup"><span data-stu-id="09fac-314">Example</span></span>

```js
var modified = Office.context.mailbox.item.dateTimeModified;
```

####  <a name="end-datetimejavascriptapioutlook16officetime"></a><span data-ttu-id="09fac-315">end :Date|[Time](/javascript/api/outlook_1_6/office.time)</span><span class="sxs-lookup"><span data-stu-id="09fac-315">end :Date|[Time](/javascript/api/outlook_1_6/office.time)</span></span>

<span data-ttu-id="09fac-316">获取或设置约会结束的日期和时间。</span><span class="sxs-lookup"><span data-stu-id="09fac-316">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="09fac-p111">将 `end` 属性表示为协调世界时 (UTC) 的日期和时间值。可使用 [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook16officelocalclienttime) 方法将 end 属性值转换为客户端的本地日期和时间。</span><span class="sxs-lookup"><span data-stu-id="09fac-p111">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook16officelocalclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="09fac-319">阅读模式</span><span class="sxs-lookup"><span data-stu-id="09fac-319">Read mode</span></span>

<span data-ttu-id="09fac-320">`end` 属性返回 `Date` 对象。</span><span class="sxs-lookup"><span data-stu-id="09fac-320">The `end` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="09fac-321">撰写模式</span><span class="sxs-lookup"><span data-stu-id="09fac-321">Compose mode</span></span>

<span data-ttu-id="09fac-322">`end` 属性返回 `Time` 对象。</span><span class="sxs-lookup"><span data-stu-id="09fac-322">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="09fac-323">使用 [`Time.setAsync`](/javascript/api/outlook_1_6/office.time#setasync-datetime--options--callback-) 方法设置结束时间时，应使用 [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) 方法将客户端的本地时间转换为服务器的 UTC。</span><span class="sxs-lookup"><span data-stu-id="09fac-323">When you use the [`Time.setAsync`](/javascript/api/outlook_1_6/office.time#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="09fac-324">类型：</span><span class="sxs-lookup"><span data-stu-id="09fac-324">Type:</span></span>

*   <span data-ttu-id="09fac-325">Date | [Time](/javascript/api/outlook_1_6/office.time)</span><span class="sxs-lookup"><span data-stu-id="09fac-325">Date | [Time](/javascript/api/outlook_1_6/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="09fac-326">Requirements</span><span class="sxs-lookup"><span data-stu-id="09fac-326">Requirements</span></span>

|<span data-ttu-id="09fac-327">要求</span><span class="sxs-lookup"><span data-stu-id="09fac-327">Requirement</span></span>| <span data-ttu-id="09fac-328">值</span><span class="sxs-lookup"><span data-stu-id="09fac-328">Value</span></span>|
|---|---|
|[<span data-ttu-id="09fac-329">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="09fac-329">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="09fac-330">1.0</span><span class="sxs-lookup"><span data-stu-id="09fac-330">1.0</span></span>|
|[<span data-ttu-id="09fac-331">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="09fac-331">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="09fac-332">ReadItem</span><span class="sxs-lookup"><span data-stu-id="09fac-332">ReadItem</span></span>|
|[<span data-ttu-id="09fac-333">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="09fac-333">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="09fac-334">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="09fac-334">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="09fac-335">示例</span><span class="sxs-lookup"><span data-stu-id="09fac-335">Example</span></span>

<span data-ttu-id="09fac-336">以下示例通过使用 `Time` 对象的 [`setAsync`](/javascript/api/outlook_1_6/office.time#setasync-datetime--options--callback-) 方法，设置撰写模式下约会的结束时间。</span><span class="sxs-lookup"><span data-stu-id="09fac-336">The following example sets the end time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_6/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

#### <a name="from-emailaddressdetailsjavascriptapioutlook16officeemailaddressdetails"></a><span data-ttu-id="09fac-337">from :[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="09fac-337">from :[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)</span></span>

<span data-ttu-id="09fac-p112">获取邮件发件人的电子邮件地址。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="09fac-p112">Gets the email address of the sender of a message. Read mode only.</span></span>

<span data-ttu-id="09fac-p113">`from` 和 [`sender`](#sender-emailaddressdetailsjavascriptapioutlook16officeemailaddressdetails) 属性表示同一个人，邮件由代理人发送的除外。在此情况下，`from` 属性表示代理程序，而 sender 属性表示代理人。</span><span class="sxs-lookup"><span data-stu-id="09fac-p113">The `from` and [`sender`](#sender-emailaddressdetailsjavascriptapioutlook16officeemailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="09fac-342">`from` 属性中 `EmailAddressDetails` 对象的 `recipientType` 属性为 `undefined`。</span><span class="sxs-lookup"><span data-stu-id="09fac-342">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="09fac-343">类型：</span><span class="sxs-lookup"><span data-stu-id="09fac-343">Type:</span></span>

*   [<span data-ttu-id="09fac-344">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="09fac-344">EmailAddressDetails</span></span>](/javascript/api/outlook_1_6/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="09fac-345">要求</span><span class="sxs-lookup"><span data-stu-id="09fac-345">Requirements</span></span>

|<span data-ttu-id="09fac-346">要求</span><span class="sxs-lookup"><span data-stu-id="09fac-346">Requirement</span></span>| <span data-ttu-id="09fac-347">值</span><span class="sxs-lookup"><span data-stu-id="09fac-347">Value</span></span>|
|---|---|
|[<span data-ttu-id="09fac-348">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="09fac-348">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="09fac-349">1.0</span><span class="sxs-lookup"><span data-stu-id="09fac-349">1.0</span></span>|
|[<span data-ttu-id="09fac-350">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="09fac-350">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="09fac-351">ReadItem</span><span class="sxs-lookup"><span data-stu-id="09fac-351">ReadItem</span></span>|
|[<span data-ttu-id="09fac-352">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="09fac-352">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="09fac-353">阅读</span><span class="sxs-lookup"><span data-stu-id="09fac-353">Read</span></span>|

#### <a name="internetmessageid-string"></a><span data-ttu-id="09fac-354">internetMessageId :String</span><span class="sxs-lookup"><span data-stu-id="09fac-354">internetMessageId :String</span></span>

<span data-ttu-id="09fac-p114">获取电子邮件的 Internet 消息标识符。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="09fac-p114">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="09fac-357">类型：</span><span class="sxs-lookup"><span data-stu-id="09fac-357">Type:</span></span>

*   <span data-ttu-id="09fac-358">String</span><span class="sxs-lookup"><span data-stu-id="09fac-358">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="09fac-359">要求</span><span class="sxs-lookup"><span data-stu-id="09fac-359">Requirements</span></span>

|<span data-ttu-id="09fac-360">要求</span><span class="sxs-lookup"><span data-stu-id="09fac-360">Requirement</span></span>| <span data-ttu-id="09fac-361">值</span><span class="sxs-lookup"><span data-stu-id="09fac-361">Value</span></span>|
|---|---|
|[<span data-ttu-id="09fac-362">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="09fac-362">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="09fac-363">1.0</span><span class="sxs-lookup"><span data-stu-id="09fac-363">1.0</span></span>|
|[<span data-ttu-id="09fac-364">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="09fac-364">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="09fac-365">ReadItem</span><span class="sxs-lookup"><span data-stu-id="09fac-365">ReadItem</span></span>|
|[<span data-ttu-id="09fac-366">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="09fac-366">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="09fac-367">阅读</span><span class="sxs-lookup"><span data-stu-id="09fac-367">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="09fac-368">示例</span><span class="sxs-lookup"><span data-stu-id="09fac-368">Example</span></span>

```js
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

#### <a name="itemclass-string"></a><span data-ttu-id="09fac-369">itemClass :String</span><span class="sxs-lookup"><span data-stu-id="09fac-369">itemClass :String</span></span>

<span data-ttu-id="09fac-p115">获取选定项目的 Exchange Web 服务项目类。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="09fac-p115">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="09fac-p116">`itemClass` 属性指定所选项目的邮件类别。以下是邮件或约会项目的默认邮件类别。</span><span class="sxs-lookup"><span data-stu-id="09fac-p116">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

| <span data-ttu-id="09fac-374">类型</span><span class="sxs-lookup"><span data-stu-id="09fac-374">Type</span></span> | <span data-ttu-id="09fac-375">说明</span><span class="sxs-lookup"><span data-stu-id="09fac-375">Description</span></span> | <span data-ttu-id="09fac-376">项目类</span><span class="sxs-lookup"><span data-stu-id="09fac-376">item class</span></span> |
| --- | --- | --- |
| <span data-ttu-id="09fac-377">约会项目</span><span class="sxs-lookup"><span data-stu-id="09fac-377">Appointment items</span></span> | <span data-ttu-id="09fac-378">这些是项目类 `IPM.Appointment` 或 `IPM.Appointment.Occurrence` 的日历项目。</span><span class="sxs-lookup"><span data-stu-id="09fac-378">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurrence`.</span></span> | `IPM.Appointment`<br />`IPM.Appointment.Occurrence` |
| <span data-ttu-id="09fac-379">邮件项目</span><span class="sxs-lookup"><span data-stu-id="09fac-379">Message items</span></span> | <span data-ttu-id="09fac-380">这些项目包括具有默认邮件类别 `IPM.Note` 的电子邮件，以及使用 `IPM.Schedule.Meeting` 作为基础邮件类别的会议请求、响应和取消。</span><span class="sxs-lookup"><span data-stu-id="09fac-380">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span> | `IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled` |

<span data-ttu-id="09fac-381">你可以创建用于扩展默认邮件类别的自定义邮件类别，例如，自定义约会邮件类别 `IPM.Appointment.Contoso`。</span><span class="sxs-lookup"><span data-stu-id="09fac-381">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="09fac-382">类型：</span><span class="sxs-lookup"><span data-stu-id="09fac-382">Type:</span></span>

*   <span data-ttu-id="09fac-383">String</span><span class="sxs-lookup"><span data-stu-id="09fac-383">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="09fac-384">要求</span><span class="sxs-lookup"><span data-stu-id="09fac-384">Requirements</span></span>

|<span data-ttu-id="09fac-385">要求</span><span class="sxs-lookup"><span data-stu-id="09fac-385">Requirement</span></span>| <span data-ttu-id="09fac-386">值</span><span class="sxs-lookup"><span data-stu-id="09fac-386">Value</span></span>|
|---|---|
|[<span data-ttu-id="09fac-387">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="09fac-387">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="09fac-388">1.0</span><span class="sxs-lookup"><span data-stu-id="09fac-388">1.0</span></span>|
|[<span data-ttu-id="09fac-389">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="09fac-389">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="09fac-390">ReadItem</span><span class="sxs-lookup"><span data-stu-id="09fac-390">ReadItem</span></span>|
|[<span data-ttu-id="09fac-391">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="09fac-391">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="09fac-392">阅读</span><span class="sxs-lookup"><span data-stu-id="09fac-392">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="09fac-393">示例</span><span class="sxs-lookup"><span data-stu-id="09fac-393">Example</span></span>

```js
var itemClass = Office.context.mailbox.item.itemClass;
```

#### <a name="nullable-itemid-string"></a><span data-ttu-id="09fac-394">(nullable) itemId :String</span><span class="sxs-lookup"><span data-stu-id="09fac-394">(nullable) itemId :String</span></span>

<span data-ttu-id="09fac-p117">获取当前项目的 Exchange Web 服务项目标识符。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="09fac-p117">Gets the Exchange Web Services item identifier for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="09fac-397">`itemId` 属性返回的标识符与 Exchange Web 服务项目标识符相同。</span><span class="sxs-lookup"><span data-stu-id="09fac-397">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="09fac-398">`itemId` 属性与 Outlook 条目 ID 或 Outlook REST API 使用的 ID 不同。</span><span class="sxs-lookup"><span data-stu-id="09fac-398">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="09fac-399">使用此值进行 REST API 调用前，应使用 [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string) 对它进行转换。</span><span class="sxs-lookup"><span data-stu-id="09fac-399">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="09fac-400">有关详细信息，请参阅[从 Outlook 加载项使用 Outlook REST API](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id)。</span><span class="sxs-lookup"><span data-stu-id="09fac-400">For more details, see [Use the Outlook REST APIs from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

<span data-ttu-id="09fac-p119">`itemId` 属性在撰写模式下不可用。如果需要项目标识符，[`saveAsync`](#saveasyncoptions-callback) 方法可用于将项目保存到存储，这将在回调函数的 [`AsyncResult.value`](/javascript/api/office/office.asyncresult) 参数中返回项目标识符。</span><span class="sxs-lookup"><span data-stu-id="09fac-p119">The `itemId` property is not available in compose mode. If an item identifier is required, the [`saveAsync`](#saveasyncoptions-callback) method can be used to save the item to the store, which will return the item identifier in the [`AsyncResult.value`](/javascript/api/office/office.asyncresult) parameter in the callback function.</span></span>

##### <a name="type"></a><span data-ttu-id="09fac-403">类型：</span><span class="sxs-lookup"><span data-stu-id="09fac-403">Type:</span></span>

*   <span data-ttu-id="09fac-404">String</span><span class="sxs-lookup"><span data-stu-id="09fac-404">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="09fac-405">要求</span><span class="sxs-lookup"><span data-stu-id="09fac-405">Requirements</span></span>

|<span data-ttu-id="09fac-406">要求</span><span class="sxs-lookup"><span data-stu-id="09fac-406">Requirement</span></span>| <span data-ttu-id="09fac-407">值</span><span class="sxs-lookup"><span data-stu-id="09fac-407">Value</span></span>|
|---|---|
|[<span data-ttu-id="09fac-408">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="09fac-408">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="09fac-409">1.0</span><span class="sxs-lookup"><span data-stu-id="09fac-409">1.0</span></span>|
|[<span data-ttu-id="09fac-410">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="09fac-410">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="09fac-411">ReadItem</span><span class="sxs-lookup"><span data-stu-id="09fac-411">ReadItem</span></span>|
|[<span data-ttu-id="09fac-412">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="09fac-412">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="09fac-413">阅读</span><span class="sxs-lookup"><span data-stu-id="09fac-413">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="09fac-414">示例</span><span class="sxs-lookup"><span data-stu-id="09fac-414">Example</span></span>

<span data-ttu-id="09fac-p120">以下代码检查项目标识符是否存在。如果 `itemId` 属性返回 `null` 或 `undefined`，则将项目保存到存储，并从异步结果中获取项目标识符。</span><span class="sxs-lookup"><span data-stu-id="09fac-p120">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

```js
var itemId = Office.context.mailbox.item.itemId;
if (itemId === null || itemId == undefined) {
  Office.context.mailbox.item.saveAsync(function(result){
    itemId = result.value;
  });
}
```

####  <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlook16officemailboxenumsitemtype"></a><span data-ttu-id="09fac-417">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_6/office.mailboxenums.itemtype)</span><span class="sxs-lookup"><span data-stu-id="09fac-417">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_6/office.mailboxenums.itemtype)</span></span>

<span data-ttu-id="09fac-418">获取实例表示的项的类型。</span><span class="sxs-lookup"><span data-stu-id="09fac-418">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="09fac-419">`itemType` 属性返回其中一个 `ItemType` 枚举值，指示 `item` 对象实例是邮件还是约会。</span><span class="sxs-lookup"><span data-stu-id="09fac-419">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="09fac-420">类型：</span><span class="sxs-lookup"><span data-stu-id="09fac-420">Type:</span></span>

*   [<span data-ttu-id="09fac-421">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="09fac-421">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook_1_6/office.mailboxenums.itemtype)

##### <a name="requirements"></a><span data-ttu-id="09fac-422">Requirements</span><span class="sxs-lookup"><span data-stu-id="09fac-422">Requirements</span></span>

|<span data-ttu-id="09fac-423">要求</span><span class="sxs-lookup"><span data-stu-id="09fac-423">Requirement</span></span>| <span data-ttu-id="09fac-424">值</span><span class="sxs-lookup"><span data-stu-id="09fac-424">Value</span></span>|
|---|---|
|[<span data-ttu-id="09fac-425">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="09fac-425">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="09fac-426">1.0</span><span class="sxs-lookup"><span data-stu-id="09fac-426">1.0</span></span>|
|[<span data-ttu-id="09fac-427">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="09fac-427">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="09fac-428">ReadItem</span><span class="sxs-lookup"><span data-stu-id="09fac-428">ReadItem</span></span>|
|[<span data-ttu-id="09fac-429">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="09fac-429">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="09fac-430">Compose 或 Read</span><span class="sxs-lookup"><span data-stu-id="09fac-430">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="09fac-431">示例</span><span class="sxs-lookup"><span data-stu-id="09fac-431">Example</span></span>

```js
if (Office.context.mailbox.item.itemType == Office.MailboxEnums.ItemType.Message)
  // do something
else
  // do something else
```

####  <a name="location-stringlocationjavascriptapioutlook16officelocation"></a><span data-ttu-id="09fac-432">location :String|[Location](/javascript/api/outlook_1_6/office.location)</span><span class="sxs-lookup"><span data-stu-id="09fac-432">location :String|[Location](/javascript/api/outlook_1_6/office.location)</span></span>

<span data-ttu-id="09fac-433">获取或设置约会的位置。</span><span class="sxs-lookup"><span data-stu-id="09fac-433">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="09fac-434">阅读模式</span><span class="sxs-lookup"><span data-stu-id="09fac-434">Read mode</span></span>

<span data-ttu-id="09fac-435">`location` 属性返回一个包含约会位置的字符串。</span><span class="sxs-lookup"><span data-stu-id="09fac-435">The `location` property returns a string that contains the location of the appointment.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="09fac-436">撰写模式</span><span class="sxs-lookup"><span data-stu-id="09fac-436">Compose mode</span></span>

<span data-ttu-id="09fac-437">`location` 属性返回一个 `Location` 对象，该对象提供用于获取和设置约会位置的方法。</span><span class="sxs-lookup"><span data-stu-id="09fac-437">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="09fac-438">类型：</span><span class="sxs-lookup"><span data-stu-id="09fac-438">Type:</span></span>

*   <span data-ttu-id="09fac-439">String | [Location](/javascript/api/outlook_1_6/office.location)</span><span class="sxs-lookup"><span data-stu-id="09fac-439">String | [Location](/javascript/api/outlook_1_6/office.location)</span></span>

##### <a name="requirements"></a><span data-ttu-id="09fac-440">Requirements</span><span class="sxs-lookup"><span data-stu-id="09fac-440">Requirements</span></span>

|<span data-ttu-id="09fac-441">要求</span><span class="sxs-lookup"><span data-stu-id="09fac-441">Requirement</span></span>| <span data-ttu-id="09fac-442">值</span><span class="sxs-lookup"><span data-stu-id="09fac-442">Value</span></span>|
|---|---|
|[<span data-ttu-id="09fac-443">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="09fac-443">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="09fac-444">1.0</span><span class="sxs-lookup"><span data-stu-id="09fac-444">1.0</span></span>|
|[<span data-ttu-id="09fac-445">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="09fac-445">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="09fac-446">ReadItem</span><span class="sxs-lookup"><span data-stu-id="09fac-446">ReadItem</span></span>|
|[<span data-ttu-id="09fac-447">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="09fac-447">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="09fac-448">Compose 或 Read</span><span class="sxs-lookup"><span data-stu-id="09fac-448">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="09fac-449">示例</span><span class="sxs-lookup"><span data-stu-id="09fac-449">Example</span></span>

```js
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

#### <a name="normalizedsubject-string"></a><span data-ttu-id="09fac-450">normalizedSubject :String</span><span class="sxs-lookup"><span data-stu-id="09fac-450">normalizedSubject :String</span></span>

<span data-ttu-id="09fac-p121">获取删除了所有前缀（包括 `RE:` 和 `FWD:`）的项目主题。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="09fac-p121">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="09fac-p122">normalizedSubject 属性获取包含由电子邮件程序添加的任何标准前缀（如 `RE:` 和 `FW:`）的项目主题。若要获取包含完整前缀的项目主题，请使用 [`subject`](#subject-stringsubjectjavascriptapioutlook16officesubject) 属性。</span><span class="sxs-lookup"><span data-stu-id="09fac-p122">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubjectjavascriptapioutlook16officesubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="09fac-455">类型：</span><span class="sxs-lookup"><span data-stu-id="09fac-455">Type:</span></span>

*   <span data-ttu-id="09fac-456">String</span><span class="sxs-lookup"><span data-stu-id="09fac-456">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="09fac-457">要求</span><span class="sxs-lookup"><span data-stu-id="09fac-457">Requirements</span></span>

|<span data-ttu-id="09fac-458">要求</span><span class="sxs-lookup"><span data-stu-id="09fac-458">Requirement</span></span>| <span data-ttu-id="09fac-459">值</span><span class="sxs-lookup"><span data-stu-id="09fac-459">Value</span></span>|
|---|---|
|[<span data-ttu-id="09fac-460">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="09fac-460">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="09fac-461">1.0</span><span class="sxs-lookup"><span data-stu-id="09fac-461">1.0</span></span>|
|[<span data-ttu-id="09fac-462">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="09fac-462">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="09fac-463">ReadItem</span><span class="sxs-lookup"><span data-stu-id="09fac-463">ReadItem</span></span>|
|[<span data-ttu-id="09fac-464">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="09fac-464">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="09fac-465">阅读</span><span class="sxs-lookup"><span data-stu-id="09fac-465">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="09fac-466">示例</span><span class="sxs-lookup"><span data-stu-id="09fac-466">Example</span></span>

```js
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
```

####  <a name="notificationmessages-notificationmessagesjavascriptapioutlook16officenotificationmessages"></a><span data-ttu-id="09fac-467">notificationMessages :[NotificationMessages](/javascript/api/outlook_1_6/office.notificationmessages)</span><span class="sxs-lookup"><span data-stu-id="09fac-467">notificationMessages :[NotificationMessages](/javascript/api/outlook_1_6/office.notificationmessages)</span></span>

<span data-ttu-id="09fac-468">获取项目的通知邮件。</span><span class="sxs-lookup"><span data-stu-id="09fac-468">Gets the notification messages for an item.</span></span>

##### <a name="type"></a><span data-ttu-id="09fac-469">类型：</span><span class="sxs-lookup"><span data-stu-id="09fac-469">Type:</span></span>

*   [<span data-ttu-id="09fac-470">NotificationMessages</span><span class="sxs-lookup"><span data-stu-id="09fac-470">NotificationMessages</span></span>](/javascript/api/outlook_1_6/office.notificationmessages)

##### <a name="requirements"></a><span data-ttu-id="09fac-471">要求</span><span class="sxs-lookup"><span data-stu-id="09fac-471">Requirements</span></span>

|<span data-ttu-id="09fac-472">要求</span><span class="sxs-lookup"><span data-stu-id="09fac-472">Requirement</span></span>| <span data-ttu-id="09fac-473">值</span><span class="sxs-lookup"><span data-stu-id="09fac-473">Value</span></span>|
|---|---|
|[<span data-ttu-id="09fac-474">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="09fac-474">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="09fac-475">1.3</span><span class="sxs-lookup"><span data-stu-id="09fac-475">1.3</span></span>|
|[<span data-ttu-id="09fac-476">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="09fac-476">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="09fac-477">ReadItem</span><span class="sxs-lookup"><span data-stu-id="09fac-477">ReadItem</span></span>|
|[<span data-ttu-id="09fac-478">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="09fac-478">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="09fac-479">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="09fac-479">Compose or read</span></span>|

####  <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlook16officeemailaddressdetailsrecipientsjavascriptapioutlook16officerecipients"></a><span data-ttu-id="09fac-480">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_6/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="09fac-480">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_6/office.recipients)</span></span>

<span data-ttu-id="09fac-481">提供对事件的可选与会者的访问权限。</span><span class="sxs-lookup"><span data-stu-id="09fac-481">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="09fac-482">对象的类型和访问级别取决于当前项目的模式。</span><span class="sxs-lookup"><span data-stu-id="09fac-482">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="09fac-483">阅读模式</span><span class="sxs-lookup"><span data-stu-id="09fac-483">Read mode</span></span>

<span data-ttu-id="09fac-484">`optionalAttendees` 属性返回一个数组，其中包含每个可选与会者的 `EmailAddressDetails` 对象。</span><span class="sxs-lookup"><span data-stu-id="09fac-484">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="09fac-485">撰写模式</span><span class="sxs-lookup"><span data-stu-id="09fac-485">Compose mode</span></span>

<span data-ttu-id="09fac-486">`optionalAttendees` 属性返回一个 `Recipients` 对象，该对象提供用于获取或更新可选与会者的方法。</span><span class="sxs-lookup"><span data-stu-id="09fac-486">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="09fac-487">类型：</span><span class="sxs-lookup"><span data-stu-id="09fac-487">Type:</span></span>

*   <span data-ttu-id="09fac-488">Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="09fac-488">Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="09fac-489">Requirements</span><span class="sxs-lookup"><span data-stu-id="09fac-489">Requirements</span></span>

|<span data-ttu-id="09fac-490">要求</span><span class="sxs-lookup"><span data-stu-id="09fac-490">Requirement</span></span>| <span data-ttu-id="09fac-491">值</span><span class="sxs-lookup"><span data-stu-id="09fac-491">Value</span></span>|
|---|---|
|[<span data-ttu-id="09fac-492">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="09fac-492">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="09fac-493">1.0</span><span class="sxs-lookup"><span data-stu-id="09fac-493">1.0</span></span>|
|[<span data-ttu-id="09fac-494">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="09fac-494">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="09fac-495">ReadItem</span><span class="sxs-lookup"><span data-stu-id="09fac-495">ReadItem</span></span>|
|[<span data-ttu-id="09fac-496">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="09fac-496">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="09fac-497">Compose 或 Read</span><span class="sxs-lookup"><span data-stu-id="09fac-497">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="09fac-498">示例</span><span class="sxs-lookup"><span data-stu-id="09fac-498">Example</span></span>

```js
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

#### <a name="organizer-emailaddressdetailsjavascriptapioutlook16officeemailaddressdetails"></a><span data-ttu-id="09fac-499">organizer :[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="09fac-499">organizer :[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)</span></span>

<span data-ttu-id="09fac-p124">获取指定会议的会议组织者的电子邮件地址。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="09fac-p124">Gets the email address of the meeting organizer for a specified meeting. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="09fac-502">类型:</span><span class="sxs-lookup"><span data-stu-id="09fac-502">Type:</span></span>

*   [<span data-ttu-id="09fac-503">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="09fac-503">EmailAddressDetails</span></span>](/javascript/api/outlook_1_6/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="09fac-504">Requirements</span><span class="sxs-lookup"><span data-stu-id="09fac-504">Requirements</span></span>

|<span data-ttu-id="09fac-505">要求</span><span class="sxs-lookup"><span data-stu-id="09fac-505">Requirement</span></span>| <span data-ttu-id="09fac-506">值</span><span class="sxs-lookup"><span data-stu-id="09fac-506">Value</span></span>|
|---|---|
|[<span data-ttu-id="09fac-507">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="09fac-507">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="09fac-508">1.0</span><span class="sxs-lookup"><span data-stu-id="09fac-508">1.0</span></span>|
|[<span data-ttu-id="09fac-509">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="09fac-509">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="09fac-510">ReadItem</span><span class="sxs-lookup"><span data-stu-id="09fac-510">ReadItem</span></span>|
|[<span data-ttu-id="09fac-511">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="09fac-511">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="09fac-512">阅读</span><span class="sxs-lookup"><span data-stu-id="09fac-512">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="09fac-513">示例</span><span class="sxs-lookup"><span data-stu-id="09fac-513">Example</span></span>

```js
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
```

####  <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlook16officeemailaddressdetailsrecipientsjavascriptapioutlook16officerecipients"></a><span data-ttu-id="09fac-514">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_6/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="09fac-514">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_6/office.recipients)</span></span>

<span data-ttu-id="09fac-515">提供对事件的必需与会者的访问权限。</span><span class="sxs-lookup"><span data-stu-id="09fac-515">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="09fac-516">对象的类型和访问级别取决于当前项目的模式。</span><span class="sxs-lookup"><span data-stu-id="09fac-516">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="09fac-517">阅读模式</span><span class="sxs-lookup"><span data-stu-id="09fac-517">Read mode</span></span>

<span data-ttu-id="09fac-518">`requiredAttendees` 属性返回一个数组，其中包含每个必需与会者的 `EmailAddressDetails` 对象。</span><span class="sxs-lookup"><span data-stu-id="09fac-518">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="09fac-519">撰写模式</span><span class="sxs-lookup"><span data-stu-id="09fac-519">Compose mode</span></span>

<span data-ttu-id="09fac-520">`requiredAttendees` 属性返回一个 `Recipients` 对象，该对象提供用于获取或更新必需与会者的方法。</span><span class="sxs-lookup"><span data-stu-id="09fac-520">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="09fac-521">类型：</span><span class="sxs-lookup"><span data-stu-id="09fac-521">Type:</span></span>

*   <span data-ttu-id="09fac-522">Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_6/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="09fac-522">Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_6/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="09fac-523">Requirements</span><span class="sxs-lookup"><span data-stu-id="09fac-523">Requirements</span></span>

|<span data-ttu-id="09fac-524">要求</span><span class="sxs-lookup"><span data-stu-id="09fac-524">Requirement</span></span>| <span data-ttu-id="09fac-525">值</span><span class="sxs-lookup"><span data-stu-id="09fac-525">Value</span></span>|
|---|---|
|[<span data-ttu-id="09fac-526">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="09fac-526">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="09fac-527">1.0</span><span class="sxs-lookup"><span data-stu-id="09fac-527">1.0</span></span>|
|[<span data-ttu-id="09fac-528">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="09fac-528">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="09fac-529">ReadItem</span><span class="sxs-lookup"><span data-stu-id="09fac-529">ReadItem</span></span>|
|[<span data-ttu-id="09fac-530">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="09fac-530">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="09fac-531">Compose 或 Read</span><span class="sxs-lookup"><span data-stu-id="09fac-531">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="09fac-532">示例</span><span class="sxs-lookup"><span data-stu-id="09fac-532">Example</span></span>

```js
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
}
```

#### <a name="sender-emailaddressdetailsjavascriptapioutlook16officeemailaddressdetails"></a><span data-ttu-id="09fac-533">sender :[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="09fac-533">sender :[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)</span></span>

<span data-ttu-id="09fac-p126">获取电子邮件发件人的电子邮件地址。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="09fac-p126">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="09fac-p127">[`from`](#from-emailaddressdetailsjavascriptapioutlook16officeemailaddressdetails) 和 `sender` 属性表示同一个人，邮件由代理人发送的除外。在此情况下，`from` 属性表示代理程序，而 sender 属性表示代理人。</span><span class="sxs-lookup"><span data-stu-id="09fac-p127">The [`from`](#from-emailaddressdetailsjavascriptapioutlook16officeemailaddressdetails) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="09fac-538">`sender` 属性中 `EmailAddressDetails` 对象的 `recipientType` 属性为 `undefined`。</span><span class="sxs-lookup"><span data-stu-id="09fac-538">The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="09fac-539">类型:</span><span class="sxs-lookup"><span data-stu-id="09fac-539">Type:</span></span>

*   [<span data-ttu-id="09fac-540">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="09fac-540">EmailAddressDetails</span></span>](/javascript/api/outlook_1_6/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="09fac-541">Requirements</span><span class="sxs-lookup"><span data-stu-id="09fac-541">Requirements</span></span>

|<span data-ttu-id="09fac-542">要求</span><span class="sxs-lookup"><span data-stu-id="09fac-542">Requirement</span></span>| <span data-ttu-id="09fac-543">值</span><span class="sxs-lookup"><span data-stu-id="09fac-543">Value</span></span>|
|---|---|
|[<span data-ttu-id="09fac-544">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="09fac-544">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="09fac-545">1.0</span><span class="sxs-lookup"><span data-stu-id="09fac-545">1.0</span></span>|
|[<span data-ttu-id="09fac-546">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="09fac-546">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="09fac-547">ReadItem</span><span class="sxs-lookup"><span data-stu-id="09fac-547">ReadItem</span></span>|
|[<span data-ttu-id="09fac-548">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="09fac-548">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="09fac-549">阅读</span><span class="sxs-lookup"><span data-stu-id="09fac-549">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="09fac-550">示例</span><span class="sxs-lookup"><span data-stu-id="09fac-550">Example</span></span>

```js
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
```

####  <a name="start-datetimejavascriptapioutlook16officetime"></a><span data-ttu-id="09fac-551">start :Date|[Time](/javascript/api/outlook_1_6/office.time)</span><span class="sxs-lookup"><span data-stu-id="09fac-551">start :Date|[Time](/javascript/api/outlook_1_6/office.time)</span></span>

<span data-ttu-id="09fac-552">获取或设置约会开始的日期和时间。</span><span class="sxs-lookup"><span data-stu-id="09fac-552">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="09fac-p128">将 `start` 属性表示为协调世界时 (UTC) 的日期和时间值。可使用 [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook16officelocalclienttime) 方法将该值转换为客户端的本地日期和时间。</span><span class="sxs-lookup"><span data-stu-id="09fac-p128">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook16officelocalclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="09fac-555">阅读模式</span><span class="sxs-lookup"><span data-stu-id="09fac-555">Read mode</span></span>

<span data-ttu-id="09fac-556">`start` 属性返回 `Date` 对象。</span><span class="sxs-lookup"><span data-stu-id="09fac-556">The `start` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="09fac-557">撰写模式</span><span class="sxs-lookup"><span data-stu-id="09fac-557">Compose mode</span></span>

<span data-ttu-id="09fac-558">`start` 属性返回 `Time` 对象。</span><span class="sxs-lookup"><span data-stu-id="09fac-558">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="09fac-559">使用 [`Time.setAsync`](/javascript/api/outlook_1_6/office.time#setasync-datetime--options--callback-) 方法设置开始时间时，应使用 [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) 方法将客户端的本地时间转换为服务器的 UTC。</span><span class="sxs-lookup"><span data-stu-id="09fac-559">When you use the [`Time.setAsync`](/javascript/api/outlook_1_6/office.time#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="09fac-560">类型：</span><span class="sxs-lookup"><span data-stu-id="09fac-560">Type:</span></span>

*   <span data-ttu-id="09fac-561">Date | [Time](/javascript/api/outlook_1_6/office.time)</span><span class="sxs-lookup"><span data-stu-id="09fac-561">Date | [Time](/javascript/api/outlook_1_6/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="09fac-562">Requirements</span><span class="sxs-lookup"><span data-stu-id="09fac-562">Requirements</span></span>

|<span data-ttu-id="09fac-563">要求</span><span class="sxs-lookup"><span data-stu-id="09fac-563">Requirement</span></span>| <span data-ttu-id="09fac-564">值</span><span class="sxs-lookup"><span data-stu-id="09fac-564">Value</span></span>|
|---|---|
|[<span data-ttu-id="09fac-565">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="09fac-565">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="09fac-566">1.0</span><span class="sxs-lookup"><span data-stu-id="09fac-566">1.0</span></span>|
|[<span data-ttu-id="09fac-567">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="09fac-567">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="09fac-568">ReadItem</span><span class="sxs-lookup"><span data-stu-id="09fac-568">ReadItem</span></span>|
|[<span data-ttu-id="09fac-569">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="09fac-569">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="09fac-570">Compose 或 Read</span><span class="sxs-lookup"><span data-stu-id="09fac-570">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="09fac-571">示例</span><span class="sxs-lookup"><span data-stu-id="09fac-571">Example</span></span>

<span data-ttu-id="09fac-572">以下示例通过使用 `Time` 对象的 [`setAsync`](/javascript/api/outlook_1_6/office.time#setasync-datetime--options--callback-) 方法，设置撰写模式下约会的开始时间。</span><span class="sxs-lookup"><span data-stu-id="09fac-572">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_6/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

####  <a name="subject-stringsubjectjavascriptapioutlook16officesubject"></a><span data-ttu-id="09fac-573">subject :String|[Subject](/javascript/api/outlook_1_6/office.subject)</span><span class="sxs-lookup"><span data-stu-id="09fac-573">subject :String|[Subject](/javascript/api/outlook_1_6/office.subject)</span></span>

<span data-ttu-id="09fac-574">获取或设置显示在项目的主题字段中的说明。</span><span class="sxs-lookup"><span data-stu-id="09fac-574">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="09fac-575">`subject` 属性获取或设置由电子邮件服务器发送项目时的整个主题。</span><span class="sxs-lookup"><span data-stu-id="09fac-575">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="09fac-576">阅读模式</span><span class="sxs-lookup"><span data-stu-id="09fac-576">Read mode</span></span>

<span data-ttu-id="09fac-p129">`subject` 属性返回一个字符串。使用 [`normalizedSubject`](#normalizedsubject-string) 属性获取不带任何前导前缀（如 `RE:` 和 `FW:`）的主题。</span><span class="sxs-lookup"><span data-stu-id="09fac-p129">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

```js
var subject = Office.context.mailbox.item.subject;
```

##### <a name="compose-mode"></a><span data-ttu-id="09fac-579">撰写模式</span><span class="sxs-lookup"><span data-stu-id="09fac-579">Compose mode</span></span>

<span data-ttu-id="09fac-580">`subject` 属性返回一个 `Subject` 对象，该对象提供用于获取和设置主题的方法。</span><span class="sxs-lookup"><span data-stu-id="09fac-580">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```js
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="09fac-581">类型：</span><span class="sxs-lookup"><span data-stu-id="09fac-581">Type:</span></span>

*   <span data-ttu-id="09fac-582">String | [Subject](/javascript/api/outlook_1_6/office.subject)</span><span class="sxs-lookup"><span data-stu-id="09fac-582">String | [Subject](/javascript/api/outlook_1_6/office.subject)</span></span>

##### <a name="requirements"></a><span data-ttu-id="09fac-583">要求</span><span class="sxs-lookup"><span data-stu-id="09fac-583">Requirements</span></span>

|<span data-ttu-id="09fac-584">要求</span><span class="sxs-lookup"><span data-stu-id="09fac-584">Requirement</span></span>| <span data-ttu-id="09fac-585">值</span><span class="sxs-lookup"><span data-stu-id="09fac-585">Value</span></span>|
|---|---|
|[<span data-ttu-id="09fac-586">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="09fac-586">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="09fac-587">1.0</span><span class="sxs-lookup"><span data-stu-id="09fac-587">1.0</span></span>|
|[<span data-ttu-id="09fac-588">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="09fac-588">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="09fac-589">ReadItem</span><span class="sxs-lookup"><span data-stu-id="09fac-589">ReadItem</span></span>|
|[<span data-ttu-id="09fac-590">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="09fac-590">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="09fac-591">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="09fac-591">Compose or read</span></span>|

####  <a name="to-arrayemailaddressdetailsjavascriptapioutlook16officeemailaddressdetailsrecipientsjavascriptapioutlook16officerecipients"></a><span data-ttu-id="09fac-592">to :Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_6/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="09fac-592">to :Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_6/office.recipients)</span></span>

<span data-ttu-id="09fac-593">提供对邮件的“**收件人**”行上的收件人的访问权限。</span><span class="sxs-lookup"><span data-stu-id="09fac-593">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="09fac-594">对象的类型和访问级别取决于当前项目的模式。</span><span class="sxs-lookup"><span data-stu-id="09fac-594">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="09fac-595">阅读模式</span><span class="sxs-lookup"><span data-stu-id="09fac-595">Read mode</span></span>

<span data-ttu-id="09fac-p131">`to` 属性返回包含邮件的**收件人**行上所列的每个收件人的 `EmailAddressDetails` 对象的数组。集合上限为 100 个成员。</span><span class="sxs-lookup"><span data-stu-id="09fac-p131">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="09fac-598">撰写模式</span><span class="sxs-lookup"><span data-stu-id="09fac-598">Compose mode</span></span>

<span data-ttu-id="09fac-599">`to` 属性返回一个 `Recipients` 对象，该对象提供用于获取或更新邮件的“**收件人**”行上收件人的方法。</span><span class="sxs-lookup"><span data-stu-id="09fac-599">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="09fac-600">类型：</span><span class="sxs-lookup"><span data-stu-id="09fac-600">Type:</span></span>

*   <span data-ttu-id="09fac-601">Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_6/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="09fac-601">Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_6/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="09fac-602">Requirements</span><span class="sxs-lookup"><span data-stu-id="09fac-602">Requirements</span></span>

|<span data-ttu-id="09fac-603">要求</span><span class="sxs-lookup"><span data-stu-id="09fac-603">Requirement</span></span>| <span data-ttu-id="09fac-604">值</span><span class="sxs-lookup"><span data-stu-id="09fac-604">Value</span></span>|
|---|---|
|[<span data-ttu-id="09fac-605">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="09fac-605">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="09fac-606">1.0</span><span class="sxs-lookup"><span data-stu-id="09fac-606">1.0</span></span>|
|[<span data-ttu-id="09fac-607">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="09fac-607">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="09fac-608">ReadItem</span><span class="sxs-lookup"><span data-stu-id="09fac-608">ReadItem</span></span>|
|[<span data-ttu-id="09fac-609">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="09fac-609">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="09fac-610">Compose 或 Read</span><span class="sxs-lookup"><span data-stu-id="09fac-610">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="09fac-611">示例</span><span class="sxs-lookup"><span data-stu-id="09fac-611">Example</span></span>

```js
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

### <a name="methods"></a><span data-ttu-id="09fac-612">方法</span><span class="sxs-lookup"><span data-stu-id="09fac-612">Methods</span></span>

####  <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="09fac-613">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="09fac-613">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="09fac-614">将文件作为附件添加到邮件或约会。</span><span class="sxs-lookup"><span data-stu-id="09fac-614">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="09fac-615">`addFileAttachmentAsync` 方法在指定的 URI 上载文件并将其附加到撰写窗体中的项目。</span><span class="sxs-lookup"><span data-stu-id="09fac-615">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="09fac-616">你随后可以将该标识符与 [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) 方法一同使用，以删除同一个会话中的附件。</span><span class="sxs-lookup"><span data-stu-id="09fac-616">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="09fac-617">参数：</span><span class="sxs-lookup"><span data-stu-id="09fac-617">Parameters:</span></span>

|<span data-ttu-id="09fac-618">名称</span><span class="sxs-lookup"><span data-stu-id="09fac-618">Name</span></span>| <span data-ttu-id="09fac-619">类型</span><span class="sxs-lookup"><span data-stu-id="09fac-619">Type</span></span>| <span data-ttu-id="09fac-620">属性</span><span class="sxs-lookup"><span data-stu-id="09fac-620">Attributes</span></span>| <span data-ttu-id="09fac-621">说明</span><span class="sxs-lookup"><span data-stu-id="09fac-621">Description</span></span>|
|---|---|---|---|
|`uri`| <span data-ttu-id="09fac-622">String</span><span class="sxs-lookup"><span data-stu-id="09fac-622">String</span></span>||<span data-ttu-id="09fac-p132">提供附加到邮件或约会的文件的位置的 URI。最大长度为 2048 个字符。</span><span class="sxs-lookup"><span data-stu-id="09fac-p132">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="09fac-625">字符串</span><span class="sxs-lookup"><span data-stu-id="09fac-625">String</span></span>||<span data-ttu-id="09fac-p133">在附件上载过程中显示的附件名称。最大长度为 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="09fac-p133">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="09fac-628">Object</span><span class="sxs-lookup"><span data-stu-id="09fac-628">Object</span></span>| <span data-ttu-id="09fac-629">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="09fac-629">&lt;optional&gt;</span></span>|<span data-ttu-id="09fac-630">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="09fac-630">An object literal that contains one or more of the following properties.</span></span>|
| `options.asyncContext` | <span data-ttu-id="09fac-631">对象</span><span class="sxs-lookup"><span data-stu-id="09fac-631">Object</span></span> | <span data-ttu-id="09fac-632">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="09fac-632">&lt;optional&gt;</span></span> | <span data-ttu-id="09fac-633">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="09fac-633">Developers can provide any object they wish to access in the callback method.</span></span> |
| `options.isInline` | <span data-ttu-id="09fac-634">布尔值</span><span class="sxs-lookup"><span data-stu-id="09fac-634">Boolean</span></span> | <span data-ttu-id="09fac-635">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="09fac-635">&lt;optional&gt;</span></span> | <span data-ttu-id="09fac-636">如果为 `true`，则表示附件将在邮件正文中内联显示，并且不应显示在附件列表中。</span><span class="sxs-lookup"><span data-stu-id="09fac-636">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
|`callback`| <span data-ttu-id="09fac-637">函数</span><span class="sxs-lookup"><span data-stu-id="09fac-637">function</span></span>| <span data-ttu-id="09fac-638">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="09fac-638">&lt;optional&gt;</span></span>|<span data-ttu-id="09fac-639">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="09fac-639">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="09fac-640">如果成功，附件标识符将在 `asyncResult.value` 属性中提供。</span><span class="sxs-lookup"><span data-stu-id="09fac-640">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="09fac-641">如果上传附件失败，`asyncResult` 对象将包含一个提供错误说明的 `Error` 对象。</span><span class="sxs-lookup"><span data-stu-id="09fac-641">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="09fac-642">错误</span><span class="sxs-lookup"><span data-stu-id="09fac-642">Errors</span></span>

| <span data-ttu-id="09fac-643">错误代码</span><span class="sxs-lookup"><span data-stu-id="09fac-643">Error code</span></span> | <span data-ttu-id="09fac-644">说明</span><span class="sxs-lookup"><span data-stu-id="09fac-644">Description</span></span> |
|------------|-------------|
| `AttachmentSizeExceeded` | <span data-ttu-id="09fac-645">附件大小超过了允许的大小。</span><span class="sxs-lookup"><span data-stu-id="09fac-645">The attachment is larger than allowed.</span></span> |
| `FileTypeNotSupported` | <span data-ttu-id="09fac-646">该附件的扩展名不是允许的扩展名。</span><span class="sxs-lookup"><span data-stu-id="09fac-646">The attachment has an extension that is not allowed.</span></span> |
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="09fac-647">邮件或约会具有的附件过多。</span><span class="sxs-lookup"><span data-stu-id="09fac-647">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="09fac-648">Requirements</span><span class="sxs-lookup"><span data-stu-id="09fac-648">Requirements</span></span>

|<span data-ttu-id="09fac-649">要求</span><span class="sxs-lookup"><span data-stu-id="09fac-649">Requirement</span></span>| <span data-ttu-id="09fac-650">值</span><span class="sxs-lookup"><span data-stu-id="09fac-650">Value</span></span>|
|---|---|
|[<span data-ttu-id="09fac-651">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="09fac-651">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="09fac-652">1.1</span><span class="sxs-lookup"><span data-stu-id="09fac-652">1.1</span></span>|
|[<span data-ttu-id="09fac-653">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="09fac-653">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="09fac-654">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="09fac-654">ReadWriteItem</span></span>|
|[<span data-ttu-id="09fac-655">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="09fac-655">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="09fac-656">撰写</span><span class="sxs-lookup"><span data-stu-id="09fac-656">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="09fac-657">示例</span><span class="sxs-lookup"><span data-stu-id="09fac-657">Examples</span></span>

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

<span data-ttu-id="09fac-658">以下示例将图像文件添加为内联附件，并在邮件正文中引用该附件。</span><span class="sxs-lookup"><span data-stu-id="09fac-658">The following example adds an image file as an inline attachment and references the attachment in the message body.</span></span>

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

####  <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="09fac-659">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="09fac-659">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="09fac-660">将 Exchange 项目（如邮件）作为附件添加到邮件或约会。</span><span class="sxs-lookup"><span data-stu-id="09fac-660">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="09fac-p134">`addItemAttachmentAsync` 方法将包含指定 Exchange 标识符的项目附加到撰写窗体中的项目。如果指定一个回调方法，此方法使用 `asyncResult` 参数调用，该参数包含一个附件标识符或代码，指示附加项目过程中出现的任何错误。可以使用 `options` 参数将状态信息传递给回调方法（如果需要）。</span><span class="sxs-lookup"><span data-stu-id="09fac-p134">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="09fac-664">随后可以将该标识符与 [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) 方法一同使用，以删除同一个会话中的附件。</span><span class="sxs-lookup"><span data-stu-id="09fac-664">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="09fac-665">如果 Office 加载项在 Outlook Web App 中运行，则 `addItemAttachmentAsync` 方法可以将项目附加到项目（正在编辑的项目除外）中；然而，不支持也不建议这样做。</span><span class="sxs-lookup"><span data-stu-id="09fac-665">If your Office Add-in is running in Outlook Web App, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="09fac-666">参数：</span><span class="sxs-lookup"><span data-stu-id="09fac-666">Parameters:</span></span>

|<span data-ttu-id="09fac-667">名称</span><span class="sxs-lookup"><span data-stu-id="09fac-667">Name</span></span>| <span data-ttu-id="09fac-668">类型</span><span class="sxs-lookup"><span data-stu-id="09fac-668">Type</span></span>| <span data-ttu-id="09fac-669">属性</span><span class="sxs-lookup"><span data-stu-id="09fac-669">Attributes</span></span>| <span data-ttu-id="09fac-670">说明</span><span class="sxs-lookup"><span data-stu-id="09fac-670">Description</span></span>|
|---|---|---|---|
|`itemId`| <span data-ttu-id="09fac-671">字符串</span><span class="sxs-lookup"><span data-stu-id="09fac-671">String</span></span>||<span data-ttu-id="09fac-p135">要附加的项目的 Exchange 标识符。最大长度为 100 个字符。</span><span class="sxs-lookup"><span data-stu-id="09fac-p135">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="09fac-674">String</span><span class="sxs-lookup"><span data-stu-id="09fac-674">String</span></span>||<span data-ttu-id="09fac-p136">要附加的项目的主题。最大长度为 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="09fac-p136">The sujbect of the item to be attached. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="09fac-677">Object</span><span class="sxs-lookup"><span data-stu-id="09fac-677">Object</span></span>| <span data-ttu-id="09fac-678">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="09fac-678">&lt;optional&gt;</span></span>|<span data-ttu-id="09fac-679">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="09fac-679">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="09fac-680">对象</span><span class="sxs-lookup"><span data-stu-id="09fac-680">Object</span></span>| <span data-ttu-id="09fac-681">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="09fac-681">&lt;optional&gt;</span></span>|<span data-ttu-id="09fac-682">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="09fac-682">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="09fac-683">function</span><span class="sxs-lookup"><span data-stu-id="09fac-683">function</span></span>| <span data-ttu-id="09fac-684">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="09fac-684">&lt;optional&gt;</span></span>|<span data-ttu-id="09fac-685">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="09fac-685">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="09fac-686">如果成功，附件标识符将在 `asyncResult.value` 属性中提供。</span><span class="sxs-lookup"><span data-stu-id="09fac-686">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="09fac-687">如果添加附件失败，`asyncResult` 对象将包含一个提供错误说明的 `Error` 对象。</span><span class="sxs-lookup"><span data-stu-id="09fac-687">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="09fac-688">错误</span><span class="sxs-lookup"><span data-stu-id="09fac-688">Errors</span></span>

| <span data-ttu-id="09fac-689">错误代码</span><span class="sxs-lookup"><span data-stu-id="09fac-689">Error code</span></span> | <span data-ttu-id="09fac-690">说明</span><span class="sxs-lookup"><span data-stu-id="09fac-690">Description</span></span> |
|------------|-------------|
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="09fac-691">邮件或约会具有的附件过多。</span><span class="sxs-lookup"><span data-stu-id="09fac-691">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="09fac-692">Requirements</span><span class="sxs-lookup"><span data-stu-id="09fac-692">Requirements</span></span>

|<span data-ttu-id="09fac-693">要求</span><span class="sxs-lookup"><span data-stu-id="09fac-693">Requirement</span></span>| <span data-ttu-id="09fac-694">值</span><span class="sxs-lookup"><span data-stu-id="09fac-694">Value</span></span>|
|---|---|
|[<span data-ttu-id="09fac-695">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="09fac-695">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="09fac-696">1.1</span><span class="sxs-lookup"><span data-stu-id="09fac-696">1.1</span></span>|
|[<span data-ttu-id="09fac-697">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="09fac-697">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="09fac-698">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="09fac-698">ReadWriteItem</span></span>|
|[<span data-ttu-id="09fac-699">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="09fac-699">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="09fac-700">撰写</span><span class="sxs-lookup"><span data-stu-id="09fac-700">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="09fac-701">示例</span><span class="sxs-lookup"><span data-stu-id="09fac-701">Example</span></span>

<span data-ttu-id="09fac-702">以下示例将现有的 Outlook 项目添加为名为 `My Attachment` 的附件。</span><span class="sxs-lookup"><span data-stu-id="09fac-702">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

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

####  <a name="close"></a><span data-ttu-id="09fac-703">close()</span><span class="sxs-lookup"><span data-stu-id="09fac-703">close()</span></span>

<span data-ttu-id="09fac-704">关闭当前正在撰写的项目。</span><span class="sxs-lookup"><span data-stu-id="09fac-704">Closes the current item that is being composed.</span></span>

<span data-ttu-id="09fac-p137">
            \`close\` 方法的行为取决于要撰写的项目的当前状态。如果项目具有未保存的更改，客户端将提示用户保存、放弃或取消关闭操作。</span><span class="sxs-lookup"><span data-stu-id="09fac-p137">The behavior of the `close` method depends on the current state of the item being composed. If the item has unsaved changes, the client prompts the user to save, discard, or cancel the close action.</span></span>

> [!NOTE]
> <span data-ttu-id="09fac-707">在 Outlook 网页版中，如果该项目是约会并且之前已使用 `saveAsync` 保存，则即使自上次保存项目后未发生任何更改，也会提示用户保存、放弃或取消。</span><span class="sxs-lookup"><span data-stu-id="09fac-707">In Outlook on the web, if the item is an appointment and it has previously been saved using `saveAsync`, the user is prompted to save, discard, or cancel even if no changes have occurred since the item was last saved.</span></span>

<span data-ttu-id="09fac-708">在 Outlook 桌面客户端中，如果邮件是内联答复，`close` 方法不起作用。</span><span class="sxs-lookup"><span data-stu-id="09fac-708">In the Outlook desktop client, if the message is an inline reply, the `close` method has no effect.</span></span>

##### <a name="requirements"></a><span data-ttu-id="09fac-709">要求</span><span class="sxs-lookup"><span data-stu-id="09fac-709">Requirements</span></span>

|<span data-ttu-id="09fac-710">要求</span><span class="sxs-lookup"><span data-stu-id="09fac-710">Requirement</span></span>| <span data-ttu-id="09fac-711">值</span><span class="sxs-lookup"><span data-stu-id="09fac-711">Value</span></span>|
|---|---|
|[<span data-ttu-id="09fac-712">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="09fac-712">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="09fac-713">1.3</span><span class="sxs-lookup"><span data-stu-id="09fac-713">1.3</span></span>|
|[<span data-ttu-id="09fac-714">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="09fac-714">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="09fac-715">受限</span><span class="sxs-lookup"><span data-stu-id="09fac-715">Restricted</span></span>|
|[<span data-ttu-id="09fac-716">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="09fac-716">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="09fac-717">撰写</span><span class="sxs-lookup"><span data-stu-id="09fac-717">Compose</span></span>|

#### <a name="displayreplyallformformdata"></a><span data-ttu-id="09fac-718">displayReplyAllForm(formData)</span><span class="sxs-lookup"><span data-stu-id="09fac-718">displayReplyAllForm(formData)</span></span>

<span data-ttu-id="09fac-719">显示答复窗体，其中包括所选邮件的发件人和所有收件人或所选约会的组织者和所有与会者。</span><span class="sxs-lookup"><span data-stu-id="09fac-719">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="09fac-720">在 Outlook for iOS 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="09fac-720">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="09fac-721">在 Outlook Web App 中，答复窗体显示为包含 3 列视图的弹出式窗体以及包含 2 列或 1 列视图的弹出式窗体。</span><span class="sxs-lookup"><span data-stu-id="09fac-721">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="09fac-722">如果任意字符串参数超出其限制，`displayReplyAllForm` 将引发异常。</span><span class="sxs-lookup"><span data-stu-id="09fac-722">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="09fac-p138">当在 `formData.attachments` 参数中指定附件时，Outlook 和 Outlook Web App 尝试下载所有附件并将其附加到答复窗体。如果无法添加任何附件，则在窗体 UI 中显示错误。如果这不可能，则不引发错误消息。</span><span class="sxs-lookup"><span data-stu-id="09fac-p138">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="09fac-726">参数：</span><span class="sxs-lookup"><span data-stu-id="09fac-726">Parameters:</span></span>

| <span data-ttu-id="09fac-727">名称</span><span class="sxs-lookup"><span data-stu-id="09fac-727">Name</span></span> | <span data-ttu-id="09fac-728">类型</span><span class="sxs-lookup"><span data-stu-id="09fac-728">Type</span></span> | <span data-ttu-id="09fac-729">属性</span><span class="sxs-lookup"><span data-stu-id="09fac-729">Attributes</span></span> | <span data-ttu-id="09fac-730">说明</span><span class="sxs-lookup"><span data-stu-id="09fac-730">Description</span></span> |
|---|---|---|---|
|`formData`| <span data-ttu-id="09fac-731">字符串 &#124; 对象</span><span class="sxs-lookup"><span data-stu-id="09fac-731">String &#124; Object</span></span>| |<span data-ttu-id="09fac-p139">一个包含文本和 HTML 且表示答复窗体的正文的字符串。字符串限制为 32 KB。</span><span class="sxs-lookup"><span data-stu-id="09fac-p139">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="09fac-734">**或**</span><span class="sxs-lookup"><span data-stu-id="09fac-734">**OR**</span></span><br/><span data-ttu-id="09fac-p140">包含正文或附件数据和回调函数的对象。对象定义如下。</span><span class="sxs-lookup"><span data-stu-id="09fac-p140">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="09fac-737">String</span><span class="sxs-lookup"><span data-stu-id="09fac-737">String</span></span> | <span data-ttu-id="09fac-738">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="09fac-738">&lt;optional&gt;</span></span> | <span data-ttu-id="09fac-p141">一个包含文本和 HTML 且表示答复窗体的正文的字符串。字符串限制为 32 KB。</span><span class="sxs-lookup"><span data-stu-id="09fac-p141">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="09fac-741">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="09fac-741">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="09fac-742">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="09fac-742">&lt;optional&gt;</span></span> | <span data-ttu-id="09fac-743">JSON 对象（是文件或项目附件）的数组。</span><span class="sxs-lookup"><span data-stu-id="09fac-743">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="09fac-744">String</span><span class="sxs-lookup"><span data-stu-id="09fac-744">String</span></span> | | <span data-ttu-id="09fac-p142">指示附件的类型。必须是文件附件的 `file` 或项目附件的 `item`。</span><span class="sxs-lookup"><span data-stu-id="09fac-p142">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="09fac-747">字符串</span><span class="sxs-lookup"><span data-stu-id="09fac-747">String</span></span> | | <span data-ttu-id="09fac-748">一个包含附件的名称的字符串，最多包含 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="09fac-748">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="09fac-749">String</span><span class="sxs-lookup"><span data-stu-id="09fac-749">String</span></span> | | <span data-ttu-id="09fac-p143">仅在将 `type` 设置为 `file` 时使用。文件的位置的 URI。</span><span class="sxs-lookup"><span data-stu-id="09fac-p143">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.isInline` | <span data-ttu-id="09fac-752">Boolean</span><span class="sxs-lookup"><span data-stu-id="09fac-752">Boolean</span></span> | | <span data-ttu-id="09fac-p144">仅在将 `type` 设置为 `file` 时使用。如果为 `true`，则表示附件将在邮件正文中内联显示，并且不应显示在附件列表中。</span><span class="sxs-lookup"><span data-stu-id="09fac-p144">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="09fac-755">String</span><span class="sxs-lookup"><span data-stu-id="09fac-755">String</span></span> | | <span data-ttu-id="09fac-p145">仅在将 `type` 设置为 `item` 时使用。附件的 EWS 项目 ID。字符串最长为 100 个字符。</span><span class="sxs-lookup"><span data-stu-id="09fac-p145">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="09fac-759">函数</span><span class="sxs-lookup"><span data-stu-id="09fac-759">function</span></span> | <span data-ttu-id="09fac-760">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="09fac-760">&lt;optional&gt;</span></span> | <span data-ttu-id="09fac-761">方法完成后，使用单个参数 `asyncResult`（一个 [AsyncResult](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="09fac-761">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="09fac-762">要求</span><span class="sxs-lookup"><span data-stu-id="09fac-762">Requirements</span></span>

|<span data-ttu-id="09fac-763">要求</span><span class="sxs-lookup"><span data-stu-id="09fac-763">Requirement</span></span>| <span data-ttu-id="09fac-764">值</span><span class="sxs-lookup"><span data-stu-id="09fac-764">Value</span></span>|
|---|---|
|[<span data-ttu-id="09fac-765">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="09fac-765">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="09fac-766">1.0</span><span class="sxs-lookup"><span data-stu-id="09fac-766">1.0</span></span>|
|[<span data-ttu-id="09fac-767">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="09fac-767">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="09fac-768">ReadItem</span><span class="sxs-lookup"><span data-stu-id="09fac-768">ReadItem</span></span>|
|[<span data-ttu-id="09fac-769">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="09fac-769">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="09fac-770">阅读</span><span class="sxs-lookup"><span data-stu-id="09fac-770">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="09fac-771">示例</span><span class="sxs-lookup"><span data-stu-id="09fac-771">Examples</span></span>

<span data-ttu-id="09fac-772">以下代码将一个字符串传递到 `displayReplyAllForm` 函数。</span><span class="sxs-lookup"><span data-stu-id="09fac-772">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="09fac-773">使用空白正文答复。</span><span class="sxs-lookup"><span data-stu-id="09fac-773">Reply with an empty body.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="09fac-774">仅使用正文答复。</span><span class="sxs-lookup"><span data-stu-id="09fac-774">Reply with just a body.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="09fac-775">使用正文和文件附件答复。</span><span class="sxs-lookup"><span data-stu-id="09fac-775">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="09fac-776">使用正文和项目附件答复。</span><span class="sxs-lookup"><span data-stu-id="09fac-776">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="09fac-777">使用正文、文件附件、项目附件和回调答复。</span><span class="sxs-lookup"><span data-stu-id="09fac-777">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="displayreplyformformdata"></a><span data-ttu-id="09fac-778">displayReplyForm(formData)</span><span class="sxs-lookup"><span data-stu-id="09fac-778">displayReplyForm(formData)</span></span>

<span data-ttu-id="09fac-779">显示答复窗体，其中仅包括所选邮件的发件人或所选约会的组织者。</span><span class="sxs-lookup"><span data-stu-id="09fac-779">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="09fac-780">在 Outlook for iOS 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="09fac-780">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="09fac-781">在 Outlook Web App 中，答复窗体显示为包含 3 列视图的弹出式窗体以及包含 2 列或 1 列视图的弹出式窗体。</span><span class="sxs-lookup"><span data-stu-id="09fac-781">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="09fac-782">如果任意字符串参数超出其限制，`displayReplyForm` 将引发异常。</span><span class="sxs-lookup"><span data-stu-id="09fac-782">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="09fac-p146">当在 `formData.attachments` 参数中指定附件时，Outlook 和 Outlook Web App 尝试下载所有附件并将其附加到答复窗体。如果无法添加任何附件，则在窗体 UI 中显示错误。如果这不可能，则不引发错误消息。</span><span class="sxs-lookup"><span data-stu-id="09fac-p146">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="09fac-786">参数：</span><span class="sxs-lookup"><span data-stu-id="09fac-786">Parameters:</span></span>

| <span data-ttu-id="09fac-787">名称</span><span class="sxs-lookup"><span data-stu-id="09fac-787">Name</span></span> | <span data-ttu-id="09fac-788">类型</span><span class="sxs-lookup"><span data-stu-id="09fac-788">Type</span></span> | <span data-ttu-id="09fac-789">属性</span><span class="sxs-lookup"><span data-stu-id="09fac-789">Attributes</span></span> | <span data-ttu-id="09fac-790">说明</span><span class="sxs-lookup"><span data-stu-id="09fac-790">Description</span></span> |
|---|---|---|---|
|`formData`| <span data-ttu-id="09fac-791">字符串 &#124; 对象</span><span class="sxs-lookup"><span data-stu-id="09fac-791">String &#124; Object</span></span>| | <span data-ttu-id="09fac-p147">一个包含文本和 HTML 且表示答复窗体的正文的字符串。字符串限制为 32 KB。</span><span class="sxs-lookup"><span data-stu-id="09fac-p147">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="09fac-794">**或**</span><span class="sxs-lookup"><span data-stu-id="09fac-794">**OR**</span></span><br/><span data-ttu-id="09fac-p148">包含正文或附件数据和回调函数的对象。对象定义如下。</span><span class="sxs-lookup"><span data-stu-id="09fac-p148">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="09fac-797">String</span><span class="sxs-lookup"><span data-stu-id="09fac-797">String</span></span> | <span data-ttu-id="09fac-798">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="09fac-798">&lt;optional&gt;</span></span> | <span data-ttu-id="09fac-p149">一个包含文本和 HTML 且表示答复窗体的正文的字符串。字符串限制为 32 KB。</span><span class="sxs-lookup"><span data-stu-id="09fac-p149">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="09fac-801">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="09fac-801">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="09fac-802">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="09fac-802">&lt;optional&gt;</span></span> | <span data-ttu-id="09fac-803">JSON 对象（是文件或项目附件）的数组。</span><span class="sxs-lookup"><span data-stu-id="09fac-803">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="09fac-804">String</span><span class="sxs-lookup"><span data-stu-id="09fac-804">String</span></span> | | <span data-ttu-id="09fac-p150">指示附件的类型。必须是文件附件的 `file` 或项目附件的 `item`。</span><span class="sxs-lookup"><span data-stu-id="09fac-p150">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="09fac-807">字符串</span><span class="sxs-lookup"><span data-stu-id="09fac-807">String</span></span> | | <span data-ttu-id="09fac-808">一个包含附件的名称的字符串，最多包含 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="09fac-808">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="09fac-809">String</span><span class="sxs-lookup"><span data-stu-id="09fac-809">String</span></span> | | <span data-ttu-id="09fac-p151">仅在将 `type` 设置为 `file` 时使用。文件的位置的 URI。</span><span class="sxs-lookup"><span data-stu-id="09fac-p151">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.isInline` | <span data-ttu-id="09fac-812">Boolean</span><span class="sxs-lookup"><span data-stu-id="09fac-812">Boolean</span></span> | | <span data-ttu-id="09fac-p152">仅在将 `type` 设置为 `file` 时使用。如果为 `true`，则表示附件将在邮件正文中内联显示，并且不应显示在附件列表中。</span><span class="sxs-lookup"><span data-stu-id="09fac-p152">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="09fac-815">String</span><span class="sxs-lookup"><span data-stu-id="09fac-815">String</span></span> | | <span data-ttu-id="09fac-p153">仅在将 `type` 设置为 `item` 时使用。附件的 EWS 项目 ID。字符串最长为 100 个字符。</span><span class="sxs-lookup"><span data-stu-id="09fac-p153">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="09fac-819">函数</span><span class="sxs-lookup"><span data-stu-id="09fac-819">function</span></span> | <span data-ttu-id="09fac-820">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="09fac-820">&lt;optional&gt;</span></span> | <span data-ttu-id="09fac-821">方法完成后，使用单个参数 `asyncResult`（一个 [AsyncResult](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="09fac-821">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="09fac-822">Requirements</span><span class="sxs-lookup"><span data-stu-id="09fac-822">Requirements</span></span>

|<span data-ttu-id="09fac-823">要求</span><span class="sxs-lookup"><span data-stu-id="09fac-823">Requirement</span></span>| <span data-ttu-id="09fac-824">值</span><span class="sxs-lookup"><span data-stu-id="09fac-824">Value</span></span>|
|---|---|
|[<span data-ttu-id="09fac-825">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="09fac-825">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="09fac-826">1.0</span><span class="sxs-lookup"><span data-stu-id="09fac-826">1.0</span></span>|
|[<span data-ttu-id="09fac-827">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="09fac-827">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="09fac-828">ReadItem</span><span class="sxs-lookup"><span data-stu-id="09fac-828">ReadItem</span></span>|
|[<span data-ttu-id="09fac-829">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="09fac-829">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="09fac-830">阅读</span><span class="sxs-lookup"><span data-stu-id="09fac-830">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="09fac-831">示例</span><span class="sxs-lookup"><span data-stu-id="09fac-831">Examples</span></span>

<span data-ttu-id="09fac-832">以下代码将一个字符串传递到 `displayReplyForm` 函数。</span><span class="sxs-lookup"><span data-stu-id="09fac-832">The following code passes a string to the `displayReplyForm` function.</span></span>

```js
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="09fac-833">使用空白正文答复。</span><span class="sxs-lookup"><span data-stu-id="09fac-833">Reply with an empty body.</span></span>

```js
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="09fac-834">仅使用正文答复。</span><span class="sxs-lookup"><span data-stu-id="09fac-834">Reply with just a body.</span></span>

```js
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="09fac-835">使用正文和文件附件答复。</span><span class="sxs-lookup"><span data-stu-id="09fac-835">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="09fac-836">使用正文和项目附件答复。</span><span class="sxs-lookup"><span data-stu-id="09fac-836">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="09fac-837">使用正文、文件附件、项目附件和回调答复。</span><span class="sxs-lookup"><span data-stu-id="09fac-837">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="getentities--entitiesjavascriptapioutlook16officeentities"></a><span data-ttu-id="09fac-838">getEntities() → {[Entities](/javascript/api/outlook_1_6/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="09fac-838">getEntities() → {[Entities](/javascript/api/outlook_1_6/office.entities)}</span></span>

<span data-ttu-id="09fac-839">获取在所选项目的正文中找到的实体。</span><span class="sxs-lookup"><span data-stu-id="09fac-839">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="09fac-840">在 Outlook for iOS 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="09fac-840">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="09fac-841">要求</span><span class="sxs-lookup"><span data-stu-id="09fac-841">Requirements</span></span>

|<span data-ttu-id="09fac-842">要求</span><span class="sxs-lookup"><span data-stu-id="09fac-842">Requirement</span></span>| <span data-ttu-id="09fac-843">值</span><span class="sxs-lookup"><span data-stu-id="09fac-843">Value</span></span>|
|---|---|
|[<span data-ttu-id="09fac-844">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="09fac-844">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="09fac-845">1.0</span><span class="sxs-lookup"><span data-stu-id="09fac-845">1.0</span></span>|
|[<span data-ttu-id="09fac-846">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="09fac-846">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="09fac-847">ReadItem</span><span class="sxs-lookup"><span data-stu-id="09fac-847">ReadItem</span></span>|
|[<span data-ttu-id="09fac-848">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="09fac-848">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="09fac-849">阅读</span><span class="sxs-lookup"><span data-stu-id="09fac-849">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="09fac-850">返回：</span><span class="sxs-lookup"><span data-stu-id="09fac-850">Returns:</span></span>

<span data-ttu-id="09fac-851">类型：[Entities](/javascript/api/outlook_1_6/office.entities)</span><span class="sxs-lookup"><span data-stu-id="09fac-851">Type: [Entities](/javascript/api/outlook_1_6/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="09fac-852">示例</span><span class="sxs-lookup"><span data-stu-id="09fac-852">Example</span></span>

<span data-ttu-id="09fac-853">以下示例访问当前项目的正文中的联系人实体。</span><span class="sxs-lookup"><span data-stu-id="09fac-853">The following example accesses the contacts entities in the current item's body.</span></span>

```js
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlook16officecontactmeetingsuggestionjavascriptapioutlook16officemeetingsuggestionphonenumberjavascriptapioutlook16officephonenumbertasksuggestionjavascriptapioutlook16officetasksuggestion"></a><span data-ttu-id="09fac-854">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_6/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_6/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_6/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_6/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="09fac-854">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_6/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_6/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_6/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_6/office.tasksuggestion))>}</span></span>

<span data-ttu-id="09fac-855">获取所选项目的正文中找到的指定实体类型的所有实体的数组。</span><span class="sxs-lookup"><span data-stu-id="09fac-855">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="09fac-856">在 Outlook for iOS 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="09fac-856">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="09fac-857">参数：</span><span class="sxs-lookup"><span data-stu-id="09fac-857">Parameters:</span></span>

|<span data-ttu-id="09fac-858">名称</span><span class="sxs-lookup"><span data-stu-id="09fac-858">Name</span></span>| <span data-ttu-id="09fac-859">说明</span><span class="sxs-lookup"><span data-stu-id="09fac-859">Type</span></span>| <span data-ttu-id="09fac-860">说明</span><span class="sxs-lookup"><span data-stu-id="09fac-860">Description</span></span>|
|---|---|---|
|`entityType`| [<span data-ttu-id="09fac-861">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="09fac-861">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook_1_6/office.mailboxenums.entitytype)|<span data-ttu-id="09fac-862">EntityType 枚举值之一。</span><span class="sxs-lookup"><span data-stu-id="09fac-862">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="09fac-863">Requirements</span><span class="sxs-lookup"><span data-stu-id="09fac-863">Requirements</span></span>

|<span data-ttu-id="09fac-864">要求</span><span class="sxs-lookup"><span data-stu-id="09fac-864">Requirement</span></span>| <span data-ttu-id="09fac-865">值</span><span class="sxs-lookup"><span data-stu-id="09fac-865">Value</span></span>|
|---|---|
|[<span data-ttu-id="09fac-866">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="09fac-866">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="09fac-867">1.0</span><span class="sxs-lookup"><span data-stu-id="09fac-867">1.0</span></span>|
|[<span data-ttu-id="09fac-868">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="09fac-868">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="09fac-869">受限</span><span class="sxs-lookup"><span data-stu-id="09fac-869">Restricted</span></span>|
|[<span data-ttu-id="09fac-870">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="09fac-870">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="09fac-871">阅读</span><span class="sxs-lookup"><span data-stu-id="09fac-871">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="09fac-872">返回：</span><span class="sxs-lookup"><span data-stu-id="09fac-872">Returns:</span></span>

<span data-ttu-id="09fac-873">如果在 `entityType` 中传递的值不是 `EntityType` 枚举的有效成员，该方法返回 null。</span><span class="sxs-lookup"><span data-stu-id="09fac-873">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="09fac-874">如果指定类型的任何实体都不存在于该项目的正文中，该方法将返回空数组。</span><span class="sxs-lookup"><span data-stu-id="09fac-874">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="09fac-875">否则，返回的数组中对象的类型取决于 `entityType` 参数中请求实体的类型。</span><span class="sxs-lookup"><span data-stu-id="09fac-875">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="09fac-876">当使用此方法的最低权限级别**受限**时，某些实体类型需要 **ReadItem** 才能进行访问，如下表中所指定。</span><span class="sxs-lookup"><span data-stu-id="09fac-876">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

| <span data-ttu-id="09fac-877">`entityType` 的值</span><span class="sxs-lookup"><span data-stu-id="09fac-877">Value of `entityType`</span></span> | <span data-ttu-id="09fac-878">返回的数组中对象的类型</span><span class="sxs-lookup"><span data-stu-id="09fac-878">Type of objects in returned array</span></span> | <span data-ttu-id="09fac-879">所需权限级别</span><span class="sxs-lookup"><span data-stu-id="09fac-879">Required Permission Level</span></span> |
| --- | --- | --- |
| `Address` | <span data-ttu-id="09fac-880">String</span><span class="sxs-lookup"><span data-stu-id="09fac-880">String</span></span> | <span data-ttu-id="09fac-881">**受限**</span><span class="sxs-lookup"><span data-stu-id="09fac-881">**Restricted**</span></span> |
| `Contact` | <span data-ttu-id="09fac-882">Contact</span><span class="sxs-lookup"><span data-stu-id="09fac-882">Contact</span></span> | <span data-ttu-id="09fac-883">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="09fac-883">**ReadItem**</span></span> |
| `EmailAddress` | <span data-ttu-id="09fac-884">String</span><span class="sxs-lookup"><span data-stu-id="09fac-884">String</span></span> | <span data-ttu-id="09fac-885">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="09fac-885">**ReadItem**</span></span> |
| `MeetingSuggestion` | <span data-ttu-id="09fac-886">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="09fac-886">MeetingSuggestion</span></span> | <span data-ttu-id="09fac-887">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="09fac-887">**ReadItem**</span></span> |
| `PhoneNumber` | <span data-ttu-id="09fac-888">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="09fac-888">PhoneNumber</span></span> | <span data-ttu-id="09fac-889">**受限**</span><span class="sxs-lookup"><span data-stu-id="09fac-889">**Restricted**</span></span> |
| `TaskSuggestion` | <span data-ttu-id="09fac-890">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="09fac-890">TaskSuggestion</span></span> | <span data-ttu-id="09fac-891">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="09fac-891">**ReadItem**</span></span> |
| `URL` | <span data-ttu-id="09fac-892">String</span><span class="sxs-lookup"><span data-stu-id="09fac-892">String</span></span> | <span data-ttu-id="09fac-893">**受限**</span><span class="sxs-lookup"><span data-stu-id="09fac-893">**Restricted**</span></span> |

<span data-ttu-id="09fac-894">类型：Array.<(String|[Contact](/javascript/api/outlook_1_6/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_6/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_6/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_6/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="09fac-894">Type: Array.<(String|[Contact](/javascript/api/outlook_1_6/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_6/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_6/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_6/office.tasksuggestion))></span></span>

##### <a name="example"></a><span data-ttu-id="09fac-895">示例</span><span class="sxs-lookup"><span data-stu-id="09fac-895">Example</span></span>

<span data-ttu-id="09fac-896">以下示例显示了如何访问表示当前项目的正文中的邮政地址的字符串数组。</span><span class="sxs-lookup"><span data-stu-id="09fac-896">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlook16officecontactmeetingsuggestionjavascriptapioutlook16officemeetingsuggestionphonenumberjavascriptapioutlook16officephonenumbertasksuggestionjavascriptapioutlook16officetasksuggestion"></a><span data-ttu-id="09fac-897">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_6/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_6/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_6/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_6/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="09fac-897">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_6/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_6/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_6/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_6/office.tasksuggestion))>}</span></span>

<span data-ttu-id="09fac-898">返回传递清单 XML 文件中定义的命名筛选器的所选项目中的已知实体。</span><span class="sxs-lookup"><span data-stu-id="09fac-898">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="09fac-899">在 Outlook for iOS 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="09fac-899">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="09fac-900">`getFilteredEntitiesByName` 方法返回匹配在具有指定 `FilterName` 元素值的清单 XML 文件中的 [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) 规则元素中定义的正则表达式的实体。</span><span class="sxs-lookup"><span data-stu-id="09fac-900">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="09fac-901">参数：</span><span class="sxs-lookup"><span data-stu-id="09fac-901">Parameters:</span></span>

|<span data-ttu-id="09fac-902">名称</span><span class="sxs-lookup"><span data-stu-id="09fac-902">Name</span></span>| <span data-ttu-id="09fac-903">说明</span><span class="sxs-lookup"><span data-stu-id="09fac-903">Type</span></span>| <span data-ttu-id="09fac-904">描述</span><span class="sxs-lookup"><span data-stu-id="09fac-904">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="09fac-905">字符串</span><span class="sxs-lookup"><span data-stu-id="09fac-905">String</span></span>|<span data-ttu-id="09fac-906">定义筛选器匹配的 `ItemHasKnownEntity` 规则元素的名称。</span><span class="sxs-lookup"><span data-stu-id="09fac-906">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="09fac-907">Requirements</span><span class="sxs-lookup"><span data-stu-id="09fac-907">Requirements</span></span>

|<span data-ttu-id="09fac-908">要求</span><span class="sxs-lookup"><span data-stu-id="09fac-908">Requirement</span></span>| <span data-ttu-id="09fac-909">值</span><span class="sxs-lookup"><span data-stu-id="09fac-909">Value</span></span>|
|---|---|
|[<span data-ttu-id="09fac-910">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="09fac-910">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="09fac-911">1.0</span><span class="sxs-lookup"><span data-stu-id="09fac-911">1.0</span></span>|
|[<span data-ttu-id="09fac-912">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="09fac-912">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="09fac-913">ReadItem</span><span class="sxs-lookup"><span data-stu-id="09fac-913">ReadItem</span></span>|
|[<span data-ttu-id="09fac-914">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="09fac-914">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="09fac-915">阅读</span><span class="sxs-lookup"><span data-stu-id="09fac-915">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="09fac-916">返回：</span><span class="sxs-lookup"><span data-stu-id="09fac-916">Returns:</span></span>

<span data-ttu-id="09fac-p155">如果具有匹配 `name` 参数的 `FilterName` 元素值的清单中没有任何 `ItemHasKnownEntity` 元素，则该方法返回 `null`。如果 `name` 参数匹配清单中的 `ItemHasKnownEntity` 元素，但在匹配的当前项目中没有实体，则该方法返回一个空数组。</span><span class="sxs-lookup"><span data-stu-id="09fac-p155">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="09fac-919">类型：Array.<(String|[Contact](/javascript/api/outlook_1_6/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_6/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_6/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_6/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="09fac-919">Type: Array.<(String|[Contact](/javascript/api/outlook_1_6/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_6/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_6/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_6/office.tasksuggestion))></span></span>

#### <a name="getregexmatches--object"></a><span data-ttu-id="09fac-920">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="09fac-920">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="09fac-921">返回所选项目中匹配在清单 XML 文件中定义的正则表达式的字符串值。</span><span class="sxs-lookup"><span data-stu-id="09fac-921">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="09fac-922">在 Outlook for iOS 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="09fac-922">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="09fac-p156">`getRegExMatches` 方法返回匹配在清单 XML 文件中的每个 `ItemHasRegularExpressionMatch` 或 `ItemHasKnownEntity` 规则元素中定义的正则表达式的字符串。对于 `ItemHasRegularExpressionMatch` 规则，匹配字符串必须发生在该规则指定的项目的属性中。`PropertyName` 简单类型定义支持的属性。</span><span class="sxs-lookup"><span data-stu-id="09fac-p156">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="09fac-926">例如，考虑一个外接程序清单具有以下 `Rule` 元素：</span><span class="sxs-lookup"><span data-stu-id="09fac-926">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="09fac-927">从 `getRegExMatches` 返回的对象应有两个属性：`fruits` 和 `veggies`。</span><span class="sxs-lookup"><span data-stu-id="09fac-927">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="09fac-p157">如果在项目的正文属性上指定 `ItemHasRegularExpressionMatch` 规则，则正则表达式应进一步筛选正文，不应尝试返回该项目的整个正文。使用正则表达式（如 `.*`）获取项目的整个正文并不总是返回预期的结果。而是使用 [`Body.getAsync`](/javascript/api/outlook_1_6/office.body#getasync-coerciontype--options--callback-) 方法检索整个正文。</span><span class="sxs-lookup"><span data-stu-id="09fac-p157">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook_1_6/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="09fac-931">Requirements</span><span class="sxs-lookup"><span data-stu-id="09fac-931">Requirements</span></span>

|<span data-ttu-id="09fac-932">要求</span><span class="sxs-lookup"><span data-stu-id="09fac-932">Requirement</span></span>| <span data-ttu-id="09fac-933">值</span><span class="sxs-lookup"><span data-stu-id="09fac-933">Value</span></span>|
|---|---|
|[<span data-ttu-id="09fac-934">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="09fac-934">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="09fac-935">1.0</span><span class="sxs-lookup"><span data-stu-id="09fac-935">1.0</span></span>|
|[<span data-ttu-id="09fac-936">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="09fac-936">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="09fac-937">ReadItem</span><span class="sxs-lookup"><span data-stu-id="09fac-937">ReadItem</span></span>|
|[<span data-ttu-id="09fac-938">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="09fac-938">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="09fac-939">阅读</span><span class="sxs-lookup"><span data-stu-id="09fac-939">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="09fac-940">返回：</span><span class="sxs-lookup"><span data-stu-id="09fac-940">Returns:</span></span>

<span data-ttu-id="09fac-p158">一个包含匹配在清单 XML 文件中定义的正则表达式的字符串数组的对象。每个数组的名称等于匹配 `ItemHasRegularExpressionMatch` 规则的 `RegExName` 属性或匹配 `ItemHasKnownEntity` 规则的 `FilterName` 属性的相应值。</span><span class="sxs-lookup"><span data-stu-id="09fac-p158">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<dl class="param-type"><span data-ttu-id="09fac-943">

<dt>
类型</dt>


</span><span class="sxs-lookup"><span data-stu-id="09fac-943">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="09fac-944">对象</span><span class="sxs-lookup"><span data-stu-id="09fac-944">Object</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="09fac-945">示例</span><span class="sxs-lookup"><span data-stu-id="09fac-945">Example</span></span>

<span data-ttu-id="09fac-946">以下示例显示了如何访问正则表达式规则元素 `fruits` 和 `veggies` 的匹配项的数组，这些元素在清单中指定。</span><span class="sxs-lookup"><span data-stu-id="09fac-946">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```js
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veges = allMatches.veggies;
```

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="09fac-947">getRegExMatchesByName(name) → (nullable) {Array.<String>}</span><span class="sxs-lookup"><span data-stu-id="09fac-947">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="09fac-948">返回匹配在清单 XML 文件中定义的命名正则表达式的所选项目中的字符串值。</span><span class="sxs-lookup"><span data-stu-id="09fac-948">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="09fac-949">在 Outlook for iOS 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="09fac-949">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="09fac-950">`getRegExMatchesByName` 方法返回匹配在具有指定 `RegExName` 元素值的清单 XML 文件中的 `ItemHasRegularExpressionMatch` 规则元素中定义的正则表达式的字符串。</span><span class="sxs-lookup"><span data-stu-id="09fac-950">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="09fac-p159">如果在项目的正文属性上指定 `ItemHasRegularExpressionMatch` 规则，则正则表达式应进一步筛选正文，不应尝试返回该项目的整个正文。使用正则表达式（如 `.*`）获取项目的整个正文并不总是返回预期的结果。</span><span class="sxs-lookup"><span data-stu-id="09fac-p159">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="09fac-953">参数：</span><span class="sxs-lookup"><span data-stu-id="09fac-953">Parameters:</span></span>

|<span data-ttu-id="09fac-954">名称</span><span class="sxs-lookup"><span data-stu-id="09fac-954">Name</span></span>| <span data-ttu-id="09fac-955">说明</span><span class="sxs-lookup"><span data-stu-id="09fac-955">Type</span></span>| <span data-ttu-id="09fac-956">描述</span><span class="sxs-lookup"><span data-stu-id="09fac-956">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="09fac-957">字符串</span><span class="sxs-lookup"><span data-stu-id="09fac-957">String</span></span>|<span data-ttu-id="09fac-958">定义筛选器匹配的 `ItemHasRegularExpressionMatch` 规则元素的名称。</span><span class="sxs-lookup"><span data-stu-id="09fac-958">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="09fac-959">Requirements</span><span class="sxs-lookup"><span data-stu-id="09fac-959">Requirements</span></span>

|<span data-ttu-id="09fac-960">要求</span><span class="sxs-lookup"><span data-stu-id="09fac-960">Requirement</span></span>| <span data-ttu-id="09fac-961">值</span><span class="sxs-lookup"><span data-stu-id="09fac-961">Value</span></span>|
|---|---|
|[<span data-ttu-id="09fac-962">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="09fac-962">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="09fac-963">1.0</span><span class="sxs-lookup"><span data-stu-id="09fac-963">1.0</span></span>|
|[<span data-ttu-id="09fac-964">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="09fac-964">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="09fac-965">ReadItem</span><span class="sxs-lookup"><span data-stu-id="09fac-965">ReadItem</span></span>|
|[<span data-ttu-id="09fac-966">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="09fac-966">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="09fac-967">阅读</span><span class="sxs-lookup"><span data-stu-id="09fac-967">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="09fac-968">返回：</span><span class="sxs-lookup"><span data-stu-id="09fac-968">Returns:</span></span>

<span data-ttu-id="09fac-969">一个包含匹配在清单 XML 文件中定义的正则表达式的字符串的数组。</span><span class="sxs-lookup"><span data-stu-id="09fac-969">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<dl class="param-type"><span data-ttu-id="09fac-970">

<dt>类型</dt>

</span><span class="sxs-lookup"><span data-stu-id="09fac-970">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="09fac-971">Array.< String ></span><span class="sxs-lookup"><span data-stu-id="09fac-971">Array.< String ></span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="09fac-972">示例</span><span class="sxs-lookup"><span data-stu-id="09fac-972">Example</span></span>

```js
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

####  <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="09fac-973">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="09fac-973">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="09fac-974">以异步方式返回邮件的主题或正文中选定的数据。</span><span class="sxs-lookup"><span data-stu-id="09fac-974">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="09fac-p160">如果没有选定内容，但光标位于正文或主题中，此方法将会为所选数据返回 null。如果选定的是字段，而不是正文或主题，则此方法返回 `InvalidSelection` 错误。</span><span class="sxs-lookup"><span data-stu-id="09fac-p160">If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="09fac-977">参数：</span><span class="sxs-lookup"><span data-stu-id="09fac-977">Parameters:</span></span>

|<span data-ttu-id="09fac-978">名称</span><span class="sxs-lookup"><span data-stu-id="09fac-978">Name</span></span>| <span data-ttu-id="09fac-979">类型</span><span class="sxs-lookup"><span data-stu-id="09fac-979">Type</span></span>| <span data-ttu-id="09fac-980">属性</span><span class="sxs-lookup"><span data-stu-id="09fac-980">Attributes</span></span>| <span data-ttu-id="09fac-981">说明</span><span class="sxs-lookup"><span data-stu-id="09fac-981">Description</span></span>|
|---|---|---|---|
|`coercionType`| [<span data-ttu-id="09fac-982">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="09fac-982">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="09fac-p161">请求数据的格式。如果为文本，则此方法返回纯文本作为字符串，删除任何显示的 HTML 标记。如果为 HTML，则此方法返回所选文本，不论是纯文本还是 HTML。</span><span class="sxs-lookup"><span data-stu-id="09fac-p161">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`| <span data-ttu-id="09fac-986">对象</span><span class="sxs-lookup"><span data-stu-id="09fac-986">Object</span></span>| <span data-ttu-id="09fac-987">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="09fac-987">&lt;optional&gt;</span></span>|<span data-ttu-id="09fac-988">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="09fac-988">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="09fac-989">对象</span><span class="sxs-lookup"><span data-stu-id="09fac-989">Object</span></span>| <span data-ttu-id="09fac-990">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="09fac-990">&lt;optional&gt;</span></span>|<span data-ttu-id="09fac-991">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="09fac-991">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="09fac-992">函数</span><span class="sxs-lookup"><span data-stu-id="09fac-992">function</span></span>||<span data-ttu-id="09fac-993">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="09fac-993">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="09fac-994">若要从回调方法访问所选数据，请调用 `asyncResult.value.data`。</span><span class="sxs-lookup"><span data-stu-id="09fac-994">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="09fac-995">若要访问选定内容的源属性，请调用 `asyncResult.value.sourceProperty`，这将为 `body` 或 `subject`。</span><span class="sxs-lookup"><span data-stu-id="09fac-995">To access the source property that the selection comes from, call `asyncResult.value.sourceProperty`, which will be either `body` or `subject`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="09fac-996">要求</span><span class="sxs-lookup"><span data-stu-id="09fac-996">Requirements</span></span>

|<span data-ttu-id="09fac-997">要求</span><span class="sxs-lookup"><span data-stu-id="09fac-997">Requirement</span></span>| <span data-ttu-id="09fac-998">值</span><span class="sxs-lookup"><span data-stu-id="09fac-998">Value</span></span>|
|---|---|
|[<span data-ttu-id="09fac-999">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="09fac-999">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="09fac-1000">1.2</span><span class="sxs-lookup"><span data-stu-id="09fac-1000">1.2</span></span>|
|[<span data-ttu-id="09fac-1001">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="09fac-1001">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="09fac-1002">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="09fac-1002">ReadWriteItem</span></span>|
|[<span data-ttu-id="09fac-1003">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="09fac-1003">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="09fac-1004">撰写</span><span class="sxs-lookup"><span data-stu-id="09fac-1004">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="09fac-1005">返回：</span><span class="sxs-lookup"><span data-stu-id="09fac-1005">Returns:</span></span>

<span data-ttu-id="09fac-1006">作为字符串的所选数据的格式由 `coercionType` 确定。</span><span class="sxs-lookup"><span data-stu-id="09fac-1006">The selected data as a string with format determined by `coercionType`.</span></span>

<dl class="param-type"><span data-ttu-id="09fac-1007">

<dt>
类型</dt>


</span><span class="sxs-lookup"><span data-stu-id="09fac-1007">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="09fac-1008">String</span><span class="sxs-lookup"><span data-stu-id="09fac-1008">String</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="09fac-1009">示例</span><span class="sxs-lookup"><span data-stu-id="09fac-1009">Example</span></span>

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

#### <a name="getselectedentities--entitiesjavascriptapioutlook16officeentities"></a><span data-ttu-id="09fac-1010">getSelectedEntities() → {[Entities](/javascript/api/outlook_1_6/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="09fac-1010">getSelectedEntities() → {[Entities](/javascript/api/outlook_1_6/office.entities)}</span></span>

<span data-ttu-id="09fac-p163">获取在用户已选择的突出显示匹配项中找到的实体。突出显示匹配项适用于[上下文加载项](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins)。</span><span class="sxs-lookup"><span data-stu-id="09fac-p163">Gets the entities found in a highlighted match a user has selected. Highlighted matches apply to [contextual add-ins](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="09fac-1013">在 Outlook for iOS 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="09fac-1013">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="09fac-1014">Requirements</span><span class="sxs-lookup"><span data-stu-id="09fac-1014">Requirements</span></span>

|<span data-ttu-id="09fac-1015">要求</span><span class="sxs-lookup"><span data-stu-id="09fac-1015">Requirement</span></span>| <span data-ttu-id="09fac-1016">值</span><span class="sxs-lookup"><span data-stu-id="09fac-1016">Value</span></span>|
|---|---|
|[<span data-ttu-id="09fac-1017">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="09fac-1017">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="09fac-1018">1.6</span><span class="sxs-lookup"><span data-stu-id="09fac-1018">1.6</span></span> |
|[<span data-ttu-id="09fac-1019">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="09fac-1019">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="09fac-1020">ReadItem</span><span class="sxs-lookup"><span data-stu-id="09fac-1020">ReadItem</span></span>|
|[<span data-ttu-id="09fac-1021">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="09fac-1021">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="09fac-1022">阅读</span><span class="sxs-lookup"><span data-stu-id="09fac-1022">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="09fac-1023">返回：</span><span class="sxs-lookup"><span data-stu-id="09fac-1023">Returns:</span></span>

<span data-ttu-id="09fac-1024">类型：[Entities](/javascript/api/outlook_1_6/office.entities)</span><span class="sxs-lookup"><span data-stu-id="09fac-1024">Type: [Entities](/javascript/api/outlook_1_6/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="09fac-1025">示例</span><span class="sxs-lookup"><span data-stu-id="09fac-1025">Example</span></span>

<span data-ttu-id="09fac-1026">以下示例访问用户选择的突出显示匹配项中的地址实体。</span><span class="sxs-lookup"><span data-stu-id="09fac-1026">The following example accesses the addresses entities in the highlighted match selected by the user.</span></span>

```js
var contacts = Office.context.mailbox.item.getSelectedEntities().addresses;
```

#### <a name="getselectedregexmatches--object"></a><span data-ttu-id="09fac-1027">getSelectedRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="09fac-1027">getSelectedRegExMatches() → {Object}</span></span>

<span data-ttu-id="09fac-p164">返回突出显示匹配项中匹配在清单 XML 文件中定义的正则表达式的字符串值。突出显示匹配项适用于[上下文外接程序](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins)。</span><span class="sxs-lookup"><span data-stu-id="09fac-p164">Returns string values in a highlighted match that match the regular expressions defined in the manifest XML file. Highlighted matches apply to [contextual add-ins](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="09fac-1030">在 Outlook for iOS 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="09fac-1030">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="09fac-p165">`getSelectedRegExMatches` 方法返回匹配在清单 XML 文件中的每个 `ItemHasRegularExpressionMatch` 或 `ItemHasKnownEntity` 规则元素中定义的正则表达式的字符串。对于 `ItemHasRegularExpressionMatch` 规则，匹配字符串必须发生在该规则指定的项目的属性中。`PropertyName` 简单类型定义支持的属性。</span><span class="sxs-lookup"><span data-stu-id="09fac-p165">The `getSelectedRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="09fac-1034">例如，考虑一个外接程序清单具有以下 `Rule` 元素：</span><span class="sxs-lookup"><span data-stu-id="09fac-1034">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="09fac-1035">从 `getRegExMatches` 返回的对象应有两个属性：`fruits` 和 `veggies`。</span><span class="sxs-lookup"><span data-stu-id="09fac-1035">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="09fac-p166">如果在项目的正文属性上指定 `ItemHasRegularExpressionMatch` 规则，则正则表达式应进一步筛选正文，不应尝试返回该项目的整个正文。使用正则表达式（如 `.*`）获取项目的整个正文并不总是返回预期的结果。而是使用 [`Body.getAsync`](/javascript/api/outlook_1_6/office.body#getasync-coerciontype--options--callback-) 方法检索整个正文。</span><span class="sxs-lookup"><span data-stu-id="09fac-p166">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook_1_6/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="09fac-1039">Requirements</span><span class="sxs-lookup"><span data-stu-id="09fac-1039">Requirements</span></span>

|<span data-ttu-id="09fac-1040">要求</span><span class="sxs-lookup"><span data-stu-id="09fac-1040">Requirement</span></span>| <span data-ttu-id="09fac-1041">值</span><span class="sxs-lookup"><span data-stu-id="09fac-1041">Value</span></span>|
|---|---|
|[<span data-ttu-id="09fac-1042">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="09fac-1042">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="09fac-1043">1.6</span><span class="sxs-lookup"><span data-stu-id="09fac-1043">1.6</span></span> |
|[<span data-ttu-id="09fac-1044">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="09fac-1044">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="09fac-1045">ReadItem</span><span class="sxs-lookup"><span data-stu-id="09fac-1045">ReadItem</span></span>|
|[<span data-ttu-id="09fac-1046">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="09fac-1046">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="09fac-1047">阅读</span><span class="sxs-lookup"><span data-stu-id="09fac-1047">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="09fac-1048">返回：</span><span class="sxs-lookup"><span data-stu-id="09fac-1048">Returns:</span></span>

<span data-ttu-id="09fac-p167">一个包含匹配在清单 XML 文件中定义的正则表达式的字符串数组的对象。每个数组的名称等于匹配 `ItemHasRegularExpressionMatch` 规则的 `RegExName` 属性或匹配 `ItemHasKnownEntity` 规则的 `FilterName` 属性的相应值。</span><span class="sxs-lookup"><span data-stu-id="09fac-p167">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

##### <a name="example"></a><span data-ttu-id="09fac-1051">示例</span><span class="sxs-lookup"><span data-stu-id="09fac-1051">Example</span></span>

<span data-ttu-id="09fac-1052">以下示例显示了如何访问正则表达式规则元素 `fruits` 和 `veggies` 的匹配项的数组，这些元素在清单中指定。</span><span class="sxs-lookup"><span data-stu-id="09fac-1052">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```javascript
var selectedMatches = Office.context.mailbox.item.getSelectedRegExMatches();
var fruits = selectedMatches.fruits;
var veggies = selectedMatches.veggies;
```

####  <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="09fac-1053">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="09fac-1053">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="09fac-1054">异步加载所选项目上此外接程序的自定义属性。</span><span class="sxs-lookup"><span data-stu-id="09fac-1054">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="09fac-p168">自定义属性基于每个应用、每个项目存储为键/值对。此方法在回调中返回 `CustomProperties` 对象，该回调提供访问特定于当前项目和当前外接程序的自定义属性的方法。自定义属性未在项目上加密，因此这不应用作安全存储。</span><span class="sxs-lookup"><span data-stu-id="09fac-p168">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="09fac-1058">参数：</span><span class="sxs-lookup"><span data-stu-id="09fac-1058">Parameters:</span></span>

|<span data-ttu-id="09fac-1059">名称</span><span class="sxs-lookup"><span data-stu-id="09fac-1059">Name</span></span>| <span data-ttu-id="09fac-1060">类型</span><span class="sxs-lookup"><span data-stu-id="09fac-1060">Type</span></span>| <span data-ttu-id="09fac-1061">属性</span><span class="sxs-lookup"><span data-stu-id="09fac-1061">Attributes</span></span>| <span data-ttu-id="09fac-1062">说明</span><span class="sxs-lookup"><span data-stu-id="09fac-1062">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="09fac-1063">函数</span><span class="sxs-lookup"><span data-stu-id="09fac-1063">function</span></span>||<span data-ttu-id="09fac-1064">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="09fac-1064">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="09fac-1065">自定义属性作为 `asyncResult.value` 属性中的 [`CustomProperties`](/javascript/api/outlook_1_6/office.customproperties) 对象提供。</span><span class="sxs-lookup"><span data-stu-id="09fac-1065">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook_1_6/office.customproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="09fac-1066">此对象可用于获取、设置以及从项目中删除自定义属性，并将自定义属性集的更改重新保存到服务器。</span><span class="sxs-lookup"><span data-stu-id="09fac-1066">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`| <span data-ttu-id="09fac-1067">对象</span><span class="sxs-lookup"><span data-stu-id="09fac-1067">Object</span></span>| <span data-ttu-id="09fac-1068">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="09fac-1068">&lt;optional&gt;</span></span>|<span data-ttu-id="09fac-1069">开发人员可以提供他们想要在回调函数中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="09fac-1069">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="09fac-1070">此对象可以通过回调函数中的 `asyncResult.asyncContext` 属性进行访问。</span><span class="sxs-lookup"><span data-stu-id="09fac-1070">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="09fac-1071">要求</span><span class="sxs-lookup"><span data-stu-id="09fac-1071">Requirements</span></span>

|<span data-ttu-id="09fac-1072">要求</span><span class="sxs-lookup"><span data-stu-id="09fac-1072">Requirement</span></span>| <span data-ttu-id="09fac-1073">值</span><span class="sxs-lookup"><span data-stu-id="09fac-1073">Value</span></span>|
|---|---|
|[<span data-ttu-id="09fac-1074">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="09fac-1074">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="09fac-1075">1.0</span><span class="sxs-lookup"><span data-stu-id="09fac-1075">1.0</span></span>|
|[<span data-ttu-id="09fac-1076">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="09fac-1076">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="09fac-1077">ReadItem</span><span class="sxs-lookup"><span data-stu-id="09fac-1077">ReadItem</span></span>|
|[<span data-ttu-id="09fac-1078">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="09fac-1078">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="09fac-1079">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="09fac-1079">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="09fac-1080">示例</span><span class="sxs-lookup"><span data-stu-id="09fac-1080">Example</span></span>

<span data-ttu-id="09fac-p171">以下代码示例显示了如何使用 `loadCustomPropertiesAsync` 方法异步加载特定于当前项目的自定义属性。该示例还显示了如何使用 `CustomProperties.saveAsync` 方法将这些属性重新保存到服务器。加载自定义属性后，该代码示例将使用 `CustomProperties.get` 方法读取自定义属性 `myProp`，使用 `CustomProperties.set` 方法写入自定义属性 `otherProp`，最后调用 `saveAsync` 方法保存这些自定义属性。</span><span class="sxs-lookup"><span data-stu-id="09fac-p171">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

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

####  <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="09fac-1084">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="09fac-1084">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="09fac-1085">将附件从邮件或约会中删除。</span><span class="sxs-lookup"><span data-stu-id="09fac-1085">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="09fac-p172">`removeAttachmentAsync` 方法删除项目中带指定标识符的附件。最佳做法是，仅当同一个邮件应用程序在同一会话中添加了一个附件时，你才应使用该附件标识符来删除该附件。在 Outlook Web App 和适用于设备的 OWA 中，附件标识符只在同一个会话中才有效。当用户关闭应用程序，或者如果用户开始在内嵌窗体中撰写，并在随后弹出的内嵌窗体中继续在单独的窗口撰写时，会话即结束。</span><span class="sxs-lookup"><span data-stu-id="09fac-p172">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item. As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session. In Outlook Web App and OWA for Devices, the attachment identifier is valid only within the same session. A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="09fac-1090">参数：</span><span class="sxs-lookup"><span data-stu-id="09fac-1090">Parameters:</span></span>

|<span data-ttu-id="09fac-1091">名称</span><span class="sxs-lookup"><span data-stu-id="09fac-1091">Name</span></span>| <span data-ttu-id="09fac-1092">类型</span><span class="sxs-lookup"><span data-stu-id="09fac-1092">Type</span></span>| <span data-ttu-id="09fac-1093">属性</span><span class="sxs-lookup"><span data-stu-id="09fac-1093">Attributes</span></span>| <span data-ttu-id="09fac-1094">说明</span><span class="sxs-lookup"><span data-stu-id="09fac-1094">Description</span></span>|
|---|---|---|---|
|`attachmentId`| <span data-ttu-id="09fac-1095">字符串</span><span class="sxs-lookup"><span data-stu-id="09fac-1095">String</span></span>||<span data-ttu-id="09fac-1096">要删除的附件的标识符。</span><span class="sxs-lookup"><span data-stu-id="09fac-1096">The identifier of the attachment to remove.</span></span>|
|`options`| <span data-ttu-id="09fac-1097">对象</span><span class="sxs-lookup"><span data-stu-id="09fac-1097">Object</span></span>| <span data-ttu-id="09fac-1098">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="09fac-1098">&lt;optional&gt;</span></span>|<span data-ttu-id="09fac-1099">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="09fac-1099">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="09fac-1100">对象</span><span class="sxs-lookup"><span data-stu-id="09fac-1100">Object</span></span>| <span data-ttu-id="09fac-1101">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="09fac-1101">&lt;optional&gt;</span></span>|<span data-ttu-id="09fac-1102">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="09fac-1102">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="09fac-1103">function</span><span class="sxs-lookup"><span data-stu-id="09fac-1103">function</span></span>| <span data-ttu-id="09fac-1104">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="09fac-1104">&lt;optional&gt;</span></span>|<span data-ttu-id="09fac-1105">方法完成后，使用单个参数 `callback`（一个 [`asyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `AsyncResult` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="09fac-1105">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="09fac-1106">如果删除附件失败，`asyncResult.error` 属性将包含一个说明失败原因的错误代码。</span><span class="sxs-lookup"><span data-stu-id="09fac-1106">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="09fac-1107">错误</span><span class="sxs-lookup"><span data-stu-id="09fac-1107">Errors</span></span>

| <span data-ttu-id="09fac-1108">错误代码</span><span class="sxs-lookup"><span data-stu-id="09fac-1108">Error code</span></span> | <span data-ttu-id="09fac-1109">说明</span><span class="sxs-lookup"><span data-stu-id="09fac-1109">Description</span></span> |
|------------|-------------|
| `InvalidAttachmentId` | <span data-ttu-id="09fac-1110">附件标识符不存在。</span><span class="sxs-lookup"><span data-stu-id="09fac-1110">The attachment identifier does not exist.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="09fac-1111">Requirements</span><span class="sxs-lookup"><span data-stu-id="09fac-1111">Requirements</span></span>

|<span data-ttu-id="09fac-1112">要求</span><span class="sxs-lookup"><span data-stu-id="09fac-1112">Requirement</span></span>| <span data-ttu-id="09fac-1113">值</span><span class="sxs-lookup"><span data-stu-id="09fac-1113">Value</span></span>|
|---|---|
|[<span data-ttu-id="09fac-1114">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="09fac-1114">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="09fac-1115">1.1</span><span class="sxs-lookup"><span data-stu-id="09fac-1115">1.1</span></span>|
|[<span data-ttu-id="09fac-1116">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="09fac-1116">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="09fac-1117">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="09fac-1117">ReadWriteItem</span></span>|
|[<span data-ttu-id="09fac-1118">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="09fac-1118">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="09fac-1119">撰写</span><span class="sxs-lookup"><span data-stu-id="09fac-1119">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="09fac-1120">示例</span><span class="sxs-lookup"><span data-stu-id="09fac-1120">Example</span></span>

<span data-ttu-id="09fac-1121">以下代码删除包含标识符 '0' 的附件。</span><span class="sxs-lookup"><span data-stu-id="09fac-1121">The following code removes an attachment with an identifier of '0'.</span></span>

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

####  <a name="saveasyncoptions-callback"></a><span data-ttu-id="09fac-1122">saveAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="09fac-1122">saveAsync([options], callback)</span></span>

<span data-ttu-id="09fac-1123">异步保存项目。</span><span class="sxs-lookup"><span data-stu-id="09fac-1123">Asynchronously saves an item.</span></span>

<span data-ttu-id="09fac-p173">调用时，此方法将当前邮件保存为草稿，并通过回调方法返回项目 ID。在 Outlook Web App 或 Outlook 联机模式下，该项目被保存到服务器中。在 Outlook 缓存模式下，该项目被保存到本地缓存中。</span><span class="sxs-lookup"><span data-stu-id="09fac-p173">When invoked, this method saves the current message as a draft and returns the item id via the callback method. In Outlook Web App or Outlook in online mode, the item is saved to the server. In Outlook in cached mode, the item is saved to the local cache.</span></span>

> [!NOTE]
> <span data-ttu-id="09fac-1127">如果加载项在撰写模式下对某个项目调用 `saveAsync` 来获得 `itemId`，以便与 EWS 或 REST API 一同使用，请注意，当 Outlook 处于高速缓存模式时，可能需要一段时间项目才能真正同步到服务器。</span><span class="sxs-lookup"><span data-stu-id="09fac-1127">If your add-in calls `saveAsync` on an item in compose mode in order to get an `itemId` to use with EWS or the REST API, be aware that when Outlook is in cached mode, it may take some time before the item is actually synced to the server.</span></span> <span data-ttu-id="09fac-1128">在项目同步前，使用 `itemId` 将返回一个错误。</span><span class="sxs-lookup"><span data-stu-id="09fac-1128">Until the item is synced, using the `itemId` will return an error.</span></span>

<span data-ttu-id="09fac-p175">由于约会没有草稿状态，如果以撰写模式在约会中调用 `saveAsync`，则该项将被保存为用户日历中的正常约会。对于之前未保存过的新约会，则不会发送邀请。保存现有约会将向添加或删除的与会者发送更新。</span><span class="sxs-lookup"><span data-stu-id="09fac-p175">Since appointments have no draft state, if `saveAsync` is called on an appointment in compose mode, the item will be saved as a normal appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent. Saving an existing appointment will send an update to added or removed attendees.</span></span>

> [!NOTE]
> <span data-ttu-id="09fac-1132">以下客户端在撰写模式下对约会上的 `saveAsync` 具有不同的行为：</span><span class="sxs-lookup"><span data-stu-id="09fac-1132">The following clients have different behavior for `saveAsync` on appointments in compose mode:</span></span>
>
> - <span data-ttu-id="09fac-1133">Mac Outlook 不支持在撰写模式下对会议执行 `saveAsync` 操作。</span><span class="sxs-lookup"><span data-stu-id="09fac-1133">Mac Outlook does not support `saveAsync` on a meeting in compose mode.</span></span> <span data-ttu-id="09fac-1134">对 Mac Outlook 中的会议调用 `saveAsync` 将会返回一个错误。</span><span class="sxs-lookup"><span data-stu-id="09fac-1134">Calling `saveAsync` on a meeting in Mac Outlook will return an error.</span></span>
> - <span data-ttu-id="09fac-1135">在撰写模式下的约会上调用 `saveAsync` 时，Outlook 网页版始终发送邀请或更新。</span><span class="sxs-lookup"><span data-stu-id="09fac-1135">Outlook on the web always sends an invitation or update when `saveAsync` is called on an appointment in compose mode.</span></span>

##### <a name="parameters"></a><span data-ttu-id="09fac-1136">参数：</span><span class="sxs-lookup"><span data-stu-id="09fac-1136">Parameters:</span></span>

|<span data-ttu-id="09fac-1137">名称</span><span class="sxs-lookup"><span data-stu-id="09fac-1137">Name</span></span>| <span data-ttu-id="09fac-1138">类型</span><span class="sxs-lookup"><span data-stu-id="09fac-1138">Type</span></span>| <span data-ttu-id="09fac-1139">属性</span><span class="sxs-lookup"><span data-stu-id="09fac-1139">Attributes</span></span>| <span data-ttu-id="09fac-1140">说明</span><span class="sxs-lookup"><span data-stu-id="09fac-1140">Description</span></span>|
|---|---|---|---|
|`options`| <span data-ttu-id="09fac-1141">对象</span><span class="sxs-lookup"><span data-stu-id="09fac-1141">Object</span></span>| <span data-ttu-id="09fac-1142">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="09fac-1142">&lt;optional&gt;</span></span>|<span data-ttu-id="09fac-1143">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="09fac-1143">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="09fac-1144">对象</span><span class="sxs-lookup"><span data-stu-id="09fac-1144">Object</span></span>| <span data-ttu-id="09fac-1145">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="09fac-1145">&lt;optional&gt;</span></span>|<span data-ttu-id="09fac-1146">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="09fac-1146">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="09fac-1147">函数</span><span class="sxs-lookup"><span data-stu-id="09fac-1147">function</span></span>||<span data-ttu-id="09fac-1148">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="09fac-1148">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="09fac-1149">如果成功，该项目标识符将在 `asyncResult.value` 属性中提供。</span><span class="sxs-lookup"><span data-stu-id="09fac-1149">On success, the item identifier is provided in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="09fac-1150">要求</span><span class="sxs-lookup"><span data-stu-id="09fac-1150">Requirements</span></span>

|<span data-ttu-id="09fac-1151">要求</span><span class="sxs-lookup"><span data-stu-id="09fac-1151">Requirement</span></span>| <span data-ttu-id="09fac-1152">值</span><span class="sxs-lookup"><span data-stu-id="09fac-1152">Value</span></span>|
|---|---|
|[<span data-ttu-id="09fac-1153">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="09fac-1153">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="09fac-1154">1.3</span><span class="sxs-lookup"><span data-stu-id="09fac-1154">1.3</span></span>|
|[<span data-ttu-id="09fac-1155">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="09fac-1155">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="09fac-1156">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="09fac-1156">ReadWriteItem</span></span>|
|[<span data-ttu-id="09fac-1157">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="09fac-1157">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="09fac-1158">撰写</span><span class="sxs-lookup"><span data-stu-id="09fac-1158">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="09fac-1159">示例</span><span class="sxs-lookup"><span data-stu-id="09fac-1159">Examples</span></span>

```js
Office.context.mailbox.item.saveAsync(
  function callback(result) {
    // Process the result
  });
```

<span data-ttu-id="09fac-p177">下面是传递给回调函数的 `result` 参数的示例。`value` 属性包含的项目的项目 ID。</span><span class="sxs-lookup"><span data-stu-id="09fac-p177">The following is an example of the `result` parameter passed to the callback function. The `value` property contains the item ID of the item.</span></span>

```json
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

####  <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="09fac-1162">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="09fac-1162">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="09fac-1163">以异步方式将数据插入到邮件的正文或主题中。</span><span class="sxs-lookup"><span data-stu-id="09fac-1163">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="09fac-p178">`setSelectedDataAsync` 方法将指定字符串插入到项目主题或正文的光标位置，或者，如果在编辑器中已选择文本，则该方法将替换选择的文本。如果光标不在正文或主题字段中，则返回错误。插入之后，光标会位于插入内容的末尾。</span><span class="sxs-lookup"><span data-stu-id="09fac-p178">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="09fac-1167">参数：</span><span class="sxs-lookup"><span data-stu-id="09fac-1167">Parameters:</span></span>

|<span data-ttu-id="09fac-1168">名称</span><span class="sxs-lookup"><span data-stu-id="09fac-1168">Name</span></span>| <span data-ttu-id="09fac-1169">类型</span><span class="sxs-lookup"><span data-stu-id="09fac-1169">Type</span></span>| <span data-ttu-id="09fac-1170">属性</span><span class="sxs-lookup"><span data-stu-id="09fac-1170">Attributes</span></span>| <span data-ttu-id="09fac-1171">说明</span><span class="sxs-lookup"><span data-stu-id="09fac-1171">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="09fac-1172">字符串</span><span class="sxs-lookup"><span data-stu-id="09fac-1172">String</span></span>||<span data-ttu-id="09fac-p179">要插入的数据。数据不得超过 1,000,000 个字符。如果传入的数据超过 1,000,000 个字符，则会引发 `ArgumentOutOfRange` 异常。</span><span class="sxs-lookup"><span data-stu-id="09fac-p179">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`| <span data-ttu-id="09fac-1176">Object</span><span class="sxs-lookup"><span data-stu-id="09fac-1176">Object</span></span>| <span data-ttu-id="09fac-1177">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="09fac-1177">&lt;optional&gt;</span></span>|<span data-ttu-id="09fac-1178">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="09fac-1178">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="09fac-1179">对象</span><span class="sxs-lookup"><span data-stu-id="09fac-1179">Object</span></span>| <span data-ttu-id="09fac-1180">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="09fac-1180">&lt;optional&gt;</span></span>|<span data-ttu-id="09fac-1181">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="09fac-1181">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`| [<span data-ttu-id="09fac-1182">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="09fac-1182">Office.CoercionType</span></span>](office.md#coerciontype-string)| <span data-ttu-id="09fac-1183">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="09fac-1183">&lt;optional&gt;</span></span>|<span data-ttu-id="09fac-p180">如果为 `text`，则在 Outlook Web App 和 Outlook 中应用当前样式。如果该字段是 HTML 编辑器，则仅插入文本数据，即使数据为 HTML。</span><span class="sxs-lookup"><span data-stu-id="09fac-p180">If `text`, the current style is applied in Outlook Web App and Outlook. If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="09fac-p181">如果 `html` 和该字段支持 HTML（主题不支持），则在 Outlook Web App 中应用当前样式，而在 Outlook 中应用默认样式。如果该字段是文本字段，则返回 `InvalidDataFormat` 错误。</span><span class="sxs-lookup"><span data-stu-id="09fac-p181">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook Web App and the default style is applied in Outlook. If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="09fac-1188">如果未设置 `coercionType`，则结果取决于该字段：如果该字段是 HTML，则使用 HTML；如果该字段是文本，则使用纯文本。</span><span class="sxs-lookup"><span data-stu-id="09fac-1188">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`| <span data-ttu-id="09fac-1189">function</span><span class="sxs-lookup"><span data-stu-id="09fac-1189">function</span></span>||<span data-ttu-id="09fac-1190">方法完成后，使用单个参数 `callback`（一个 [`asyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `AsyncResult` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="09fac-1190">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="09fac-1191">Requirements</span><span class="sxs-lookup"><span data-stu-id="09fac-1191">Requirements</span></span>

|<span data-ttu-id="09fac-1192">要求</span><span class="sxs-lookup"><span data-stu-id="09fac-1192">Requirement</span></span>| <span data-ttu-id="09fac-1193">值</span><span class="sxs-lookup"><span data-stu-id="09fac-1193">Value</span></span>|
|---|---|
|[<span data-ttu-id="09fac-1194">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="09fac-1194">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="09fac-1195">1.2</span><span class="sxs-lookup"><span data-stu-id="09fac-1195">1.2</span></span>|
|[<span data-ttu-id="09fac-1196">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="09fac-1196">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="09fac-1197">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="09fac-1197">ReadWriteItem</span></span>|
|[<span data-ttu-id="09fac-1198">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="09fac-1198">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="09fac-1199">撰写</span><span class="sxs-lookup"><span data-stu-id="09fac-1199">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="09fac-1200">示例</span><span class="sxs-lookup"><span data-stu-id="09fac-1200">Example</span></span>

```js
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```
