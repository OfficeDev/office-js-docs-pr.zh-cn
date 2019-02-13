---
title: Office.context.mailbox.item-预览要求集
description: ''
ms.date: 01/30/2019
localization_priority: Normal
ms.openlocfilehash: a660f8bafdd2587f97d704e42c47abbe6c7d533d
ms.sourcegitcommit: a59f4e322238efa187f388a75b7709462c71e668
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 02/13/2019
ms.locfileid: "29982046"
---
# <a name="item"></a><span data-ttu-id="d7fe0-102">item</span><span class="sxs-lookup"><span data-stu-id="d7fe0-102">item</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmditem"></a><span data-ttu-id="d7fe0-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span><span class="sxs-lookup"><span data-stu-id="d7fe0-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span></span>

<span data-ttu-id="d7fe0-p101">`item` 命名空间用于访问当前选定的邮件、会议请求或约会。可以通过使用 [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtype) 属性确定 `item` 的类型。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-p101">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="d7fe0-106">Requirements</span><span class="sxs-lookup"><span data-stu-id="d7fe0-106">Requirements</span></span>

|<span data-ttu-id="d7fe0-107">要求</span><span class="sxs-lookup"><span data-stu-id="d7fe0-107">Requirement</span></span>|<span data-ttu-id="d7fe0-108">值</span><span class="sxs-lookup"><span data-stu-id="d7fe0-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="d7fe0-109">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="d7fe0-109">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d7fe0-110">1.0</span><span class="sxs-lookup"><span data-stu-id="d7fe0-110">1.0</span></span>|
|[<span data-ttu-id="d7fe0-111">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="d7fe0-111">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d7fe0-112">受限</span><span class="sxs-lookup"><span data-stu-id="d7fe0-112">Restricted</span></span>|
|[<span data-ttu-id="d7fe0-113">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="d7fe0-113">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="d7fe0-114">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="d7fe0-114">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="d7fe0-115">成员和方法</span><span class="sxs-lookup"><span data-stu-id="d7fe0-115">Members and methods</span></span>

| <span data-ttu-id="d7fe0-116">成员</span><span class="sxs-lookup"><span data-stu-id="d7fe0-116">Member</span></span> | <span data-ttu-id="d7fe0-117">类型</span><span class="sxs-lookup"><span data-stu-id="d7fe0-117">Type</span></span> |
|--------|------|
| [<span data-ttu-id="d7fe0-118">attachments</span><span class="sxs-lookup"><span data-stu-id="d7fe0-118">attachments</span></span>](#attachments-arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetails) | <span data-ttu-id="d7fe0-119">成员</span><span class="sxs-lookup"><span data-stu-id="d7fe0-119">Member</span></span> |
| [<span data-ttu-id="d7fe0-120">bcc</span><span class="sxs-lookup"><span data-stu-id="d7fe0-120">bcc</span></span>](#bcc-recipientsjavascriptapioutlookofficerecipients) | <span data-ttu-id="d7fe0-121">成员</span><span class="sxs-lookup"><span data-stu-id="d7fe0-121">Member</span></span> |
| [<span data-ttu-id="d7fe0-122">body</span><span class="sxs-lookup"><span data-stu-id="d7fe0-122">body</span></span>](#body-bodyjavascriptapioutlookofficebody) | <span data-ttu-id="d7fe0-123">成员</span><span class="sxs-lookup"><span data-stu-id="d7fe0-123">Member</span></span> |
| [<span data-ttu-id="d7fe0-124">cc</span><span class="sxs-lookup"><span data-stu-id="d7fe0-124">cc</span></span>](#cc-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients) | <span data-ttu-id="d7fe0-125">成员</span><span class="sxs-lookup"><span data-stu-id="d7fe0-125">Member</span></span> |
| [<span data-ttu-id="d7fe0-126">conversationId</span><span class="sxs-lookup"><span data-stu-id="d7fe0-126">conversationId</span></span>](#nullable-conversationid-string) | <span data-ttu-id="d7fe0-127">成员</span><span class="sxs-lookup"><span data-stu-id="d7fe0-127">Member</span></span> |
| [<span data-ttu-id="d7fe0-128">dateTimeCreated</span><span class="sxs-lookup"><span data-stu-id="d7fe0-128">dateTimeCreated</span></span>](#datetimecreated-date) | <span data-ttu-id="d7fe0-129">成员</span><span class="sxs-lookup"><span data-stu-id="d7fe0-129">Member</span></span> |
| [<span data-ttu-id="d7fe0-130">dateTimeModified</span><span class="sxs-lookup"><span data-stu-id="d7fe0-130">dateTimeModified</span></span>](#datetimemodified-date) | <span data-ttu-id="d7fe0-131">成员</span><span class="sxs-lookup"><span data-stu-id="d7fe0-131">Member</span></span> |
| [<span data-ttu-id="d7fe0-132">end</span><span class="sxs-lookup"><span data-stu-id="d7fe0-132">end</span></span>](#end-datetimejavascriptapioutlookofficetime) | <span data-ttu-id="d7fe0-133">成员</span><span class="sxs-lookup"><span data-stu-id="d7fe0-133">Member</span></span> |
| [<span data-ttu-id="d7fe0-134">enhancedLocation</span><span class="sxs-lookup"><span data-stu-id="d7fe0-134">enhancedLocation</span></span>](#enhancedlocation-enhancedlocationjavascriptapioutlookofficeenhancedlocation) | <span data-ttu-id="d7fe0-135">成员</span><span class="sxs-lookup"><span data-stu-id="d7fe0-135">Member</span></span> |
| [<span data-ttu-id="d7fe0-136">from</span><span class="sxs-lookup"><span data-stu-id="d7fe0-136">from</span></span>](#from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsfromjavascriptapioutlookofficefrom) | <span data-ttu-id="d7fe0-137">成员</span><span class="sxs-lookup"><span data-stu-id="d7fe0-137">Member</span></span> |
| [<span data-ttu-id="d7fe0-138">internetHeaders</span><span class="sxs-lookup"><span data-stu-id="d7fe0-138">internetHeaders</span></span>](#internetheaders-internetheadersjavascriptapioutlookofficeinternetheaders) | <span data-ttu-id="d7fe0-139">成员</span><span class="sxs-lookup"><span data-stu-id="d7fe0-139">Member</span></span> |
| [<span data-ttu-id="d7fe0-140">internetMessageId</span><span class="sxs-lookup"><span data-stu-id="d7fe0-140">internetMessageId</span></span>](#internetmessageid-string) | <span data-ttu-id="d7fe0-141">成员</span><span class="sxs-lookup"><span data-stu-id="d7fe0-141">Member</span></span> |
| [<span data-ttu-id="d7fe0-142">itemClass</span><span class="sxs-lookup"><span data-stu-id="d7fe0-142">itemClass</span></span>](#itemclass-string) | <span data-ttu-id="d7fe0-143">成员</span><span class="sxs-lookup"><span data-stu-id="d7fe0-143">Member</span></span> |
| [<span data-ttu-id="d7fe0-144">itemId</span><span class="sxs-lookup"><span data-stu-id="d7fe0-144">itemId</span></span>](#nullable-itemid-string) | <span data-ttu-id="d7fe0-145">成员</span><span class="sxs-lookup"><span data-stu-id="d7fe0-145">Member</span></span> |
| [<span data-ttu-id="d7fe0-146">itemType</span><span class="sxs-lookup"><span data-stu-id="d7fe0-146">itemType</span></span>](#itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtype) | <span data-ttu-id="d7fe0-147">成员</span><span class="sxs-lookup"><span data-stu-id="d7fe0-147">Member</span></span> |
| [<span data-ttu-id="d7fe0-148">location</span><span class="sxs-lookup"><span data-stu-id="d7fe0-148">location</span></span>](#location-stringlocationjavascriptapioutlookofficelocation) | <span data-ttu-id="d7fe0-149">成员</span><span class="sxs-lookup"><span data-stu-id="d7fe0-149">Member</span></span> |
| [<span data-ttu-id="d7fe0-150">normalizedSubject</span><span class="sxs-lookup"><span data-stu-id="d7fe0-150">normalizedSubject</span></span>](#normalizedsubject-string) | <span data-ttu-id="d7fe0-151">成员</span><span class="sxs-lookup"><span data-stu-id="d7fe0-151">Member</span></span> |
| [<span data-ttu-id="d7fe0-152">notificationMessages</span><span class="sxs-lookup"><span data-stu-id="d7fe0-152">notificationMessages</span></span>](#notificationmessages-notificationmessagesjavascriptapioutlookofficenotificationmessages) | <span data-ttu-id="d7fe0-153">成员</span><span class="sxs-lookup"><span data-stu-id="d7fe0-153">Member</span></span> |
| [<span data-ttu-id="d7fe0-154">optionalAttendees</span><span class="sxs-lookup"><span data-stu-id="d7fe0-154">optionalAttendees</span></span>](#optionalattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients) | <span data-ttu-id="d7fe0-155">Member</span><span class="sxs-lookup"><span data-stu-id="d7fe0-155">Member</span></span> |
| [<span data-ttu-id="d7fe0-156">organizer</span><span class="sxs-lookup"><span data-stu-id="d7fe0-156">organizer</span></span>](#organizer-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsorganizerjavascriptapioutlookofficeorganizer) | <span data-ttu-id="d7fe0-157">成员</span><span class="sxs-lookup"><span data-stu-id="d7fe0-157">Member</span></span> |
| [<span data-ttu-id="d7fe0-158">recurrence</span><span class="sxs-lookup"><span data-stu-id="d7fe0-158">recurrence</span></span>](#nullable-recurrence-recurrencejavascriptapioutlookofficerecurrence) | <span data-ttu-id="d7fe0-159">成员</span><span class="sxs-lookup"><span data-stu-id="d7fe0-159">Member</span></span> |
| [<span data-ttu-id="d7fe0-160">requiredAttendees</span><span class="sxs-lookup"><span data-stu-id="d7fe0-160">requiredAttendees</span></span>](#requiredattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients) | <span data-ttu-id="d7fe0-161">成员</span><span class="sxs-lookup"><span data-stu-id="d7fe0-161">Member</span></span> |
| [<span data-ttu-id="d7fe0-162">sender</span><span class="sxs-lookup"><span data-stu-id="d7fe0-162">sender</span></span>](#sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetails) | <span data-ttu-id="d7fe0-163">成员</span><span class="sxs-lookup"><span data-stu-id="d7fe0-163">Member</span></span> |
| [<span data-ttu-id="d7fe0-164">seriesId</span><span class="sxs-lookup"><span data-stu-id="d7fe0-164">seriesId</span></span>](#nullable-seriesid-string) | <span data-ttu-id="d7fe0-165">Member</span><span class="sxs-lookup"><span data-stu-id="d7fe0-165">Member</span></span> |
| [<span data-ttu-id="d7fe0-166">start</span><span class="sxs-lookup"><span data-stu-id="d7fe0-166">start</span></span>](#start-datetimejavascriptapioutlookofficetime) | <span data-ttu-id="d7fe0-167">Member</span><span class="sxs-lookup"><span data-stu-id="d7fe0-167">Member</span></span> |
| [<span data-ttu-id="d7fe0-168">subject</span><span class="sxs-lookup"><span data-stu-id="d7fe0-168">subject</span></span>](#subject-stringsubjectjavascriptapioutlookofficesubject) | <span data-ttu-id="d7fe0-169">Member</span><span class="sxs-lookup"><span data-stu-id="d7fe0-169">Member</span></span> |
| [<span data-ttu-id="d7fe0-170">to</span><span class="sxs-lookup"><span data-stu-id="d7fe0-170">to</span></span>](#to-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients) | <span data-ttu-id="d7fe0-171">成员</span><span class="sxs-lookup"><span data-stu-id="d7fe0-171">Member</span></span> |
| [<span data-ttu-id="d7fe0-172">addFileAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="d7fe0-172">addFileAttachmentAsync</span></span>](#addfileattachmentasyncuri-attachmentname-options-callback) | <span data-ttu-id="d7fe0-173">方法</span><span class="sxs-lookup"><span data-stu-id="d7fe0-173">Method</span></span> |
| [<span data-ttu-id="d7fe0-174">addFileAttachmentFromBase64Async</span><span class="sxs-lookup"><span data-stu-id="d7fe0-174">addFileAttachmentFromBase64Async</span></span>](#addfileattachmentfrombase64asyncbase64file-attachmentname-options-callback) | <span data-ttu-id="d7fe0-175">方法</span><span class="sxs-lookup"><span data-stu-id="d7fe0-175">Method</span></span> |
| [<span data-ttu-id="d7fe0-176">addHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="d7fe0-176">addHandlerAsync</span></span>](#addhandlerasynceventtype-handler-options-callback) | <span data-ttu-id="d7fe0-177">方法</span><span class="sxs-lookup"><span data-stu-id="d7fe0-177">Method</span></span> |
| [<span data-ttu-id="d7fe0-178">addItemAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="d7fe0-178">addItemAttachmentAsync</span></span>](#additemattachmentasyncitemid-attachmentname-options-callback) | <span data-ttu-id="d7fe0-179">方法</span><span class="sxs-lookup"><span data-stu-id="d7fe0-179">Method</span></span> |
| [<span data-ttu-id="d7fe0-180">close</span><span class="sxs-lookup"><span data-stu-id="d7fe0-180">close</span></span>](#close) | <span data-ttu-id="d7fe0-181">方法</span><span class="sxs-lookup"><span data-stu-id="d7fe0-181">Method</span></span> |
| [<span data-ttu-id="d7fe0-182">displayReplyAllForm</span><span class="sxs-lookup"><span data-stu-id="d7fe0-182">displayReplyAllForm</span></span>](#displayreplyallformformdata) | <span data-ttu-id="d7fe0-183">方法</span><span class="sxs-lookup"><span data-stu-id="d7fe0-183">Method</span></span> |
| [<span data-ttu-id="d7fe0-184">displayReplyForm</span><span class="sxs-lookup"><span data-stu-id="d7fe0-184">displayReplyForm</span></span>](#displayreplyformformdata) | <span data-ttu-id="d7fe0-185">方法</span><span class="sxs-lookup"><span data-stu-id="d7fe0-185">Method</span></span> |
| [<span data-ttu-id="d7fe0-186">getAttachmentContentAsync</span><span class="sxs-lookup"><span data-stu-id="d7fe0-186">getAttachmentContentAsync</span></span>](#getattachmentcontentasyncattachmentid-options-callback--attachmentcontentjavascriptapioutlookofficeattachmentcontent) | <span data-ttu-id="d7fe0-187">方法</span><span class="sxs-lookup"><span data-stu-id="d7fe0-187">Method</span></span> |
| [<span data-ttu-id="d7fe0-188">getAttachmentsAsync</span><span class="sxs-lookup"><span data-stu-id="d7fe0-188">getAttachmentsAsync</span></span>](#getattachmentsasyncoptions-callback--arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetails) | <span data-ttu-id="d7fe0-189">方法</span><span class="sxs-lookup"><span data-stu-id="d7fe0-189">Method</span></span> |
| [<span data-ttu-id="d7fe0-190">getEntities</span><span class="sxs-lookup"><span data-stu-id="d7fe0-190">getEntities</span></span>](#getentities--entitiesjavascriptapioutlookofficeentities) | <span data-ttu-id="d7fe0-191">方法</span><span class="sxs-lookup"><span data-stu-id="d7fe0-191">Method</span></span> |
| [<span data-ttu-id="d7fe0-192">getEntitiesByType</span><span class="sxs-lookup"><span data-stu-id="d7fe0-192">getEntitiesByType</span></span>](#getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlookofficecontactmeetingsuggestionjavascriptapioutlookofficemeetingsuggestionphonenumberjavascriptapioutlookofficephonenumbertasksuggestionjavascriptapioutlookofficetasksuggestion) | <span data-ttu-id="d7fe0-193">方法</span><span class="sxs-lookup"><span data-stu-id="d7fe0-193">Method</span></span> |
| [<span data-ttu-id="d7fe0-194">getFilteredEntitiesByName</span><span class="sxs-lookup"><span data-stu-id="d7fe0-194">getFilteredEntitiesByName</span></span>](#getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlookofficecontactmeetingsuggestionjavascriptapioutlookofficemeetingsuggestionphonenumberjavascriptapioutlookofficephonenumbertasksuggestionjavascriptapioutlookofficetasksuggestion) | <span data-ttu-id="d7fe0-195">方法</span><span class="sxs-lookup"><span data-stu-id="d7fe0-195">Method</span></span> |
| [<span data-ttu-id="d7fe0-196">getInitializationContextAsync</span><span class="sxs-lookup"><span data-stu-id="d7fe0-196">getInitializationContextAsync</span></span>](#getinitializationcontextasyncoptions-callback) | <span data-ttu-id="d7fe0-197">方法</span><span class="sxs-lookup"><span data-stu-id="d7fe0-197">Method</span></span> |
| [<span data-ttu-id="d7fe0-198">getRegExMatches</span><span class="sxs-lookup"><span data-stu-id="d7fe0-198">getRegExMatches</span></span>](#getregexmatches--object) | <span data-ttu-id="d7fe0-199">方法</span><span class="sxs-lookup"><span data-stu-id="d7fe0-199">Method</span></span> |
| [<span data-ttu-id="d7fe0-200">getRegExMatchesByName</span><span class="sxs-lookup"><span data-stu-id="d7fe0-200">getRegExMatchesByName</span></span>](#getregexmatchesbynamename--nullable-array-string-) | <span data-ttu-id="d7fe0-201">方法</span><span class="sxs-lookup"><span data-stu-id="d7fe0-201">Method</span></span> |
| [<span data-ttu-id="d7fe0-202">getSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="d7fe0-202">getSelectedDataAsync</span></span>](#getselecteddataasynccoerciontype-options-callback--string) | <span data-ttu-id="d7fe0-203">方法</span><span class="sxs-lookup"><span data-stu-id="d7fe0-203">Method</span></span> |
| [<span data-ttu-id="d7fe0-204">getSelectedEntities</span><span class="sxs-lookup"><span data-stu-id="d7fe0-204">getSelectedEntities</span></span>](#getselectedentities--entitiesjavascriptapioutlookofficeentities) | <span data-ttu-id="d7fe0-205">方法</span><span class="sxs-lookup"><span data-stu-id="d7fe0-205">Method</span></span> |
| [<span data-ttu-id="d7fe0-206">getSelectedRegExMatches</span><span class="sxs-lookup"><span data-stu-id="d7fe0-206">getSelectedRegExMatches</span></span>](#getselectedregexmatches--object) | <span data-ttu-id="d7fe0-207">方法</span><span class="sxs-lookup"><span data-stu-id="d7fe0-207">Method</span></span> |
| [<span data-ttu-id="d7fe0-208">getSharedPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="d7fe0-208">getSharedPropertiesAsync</span></span>](#getsharedpropertiesasyncoptions-callback) | <span data-ttu-id="d7fe0-209">方法</span><span class="sxs-lookup"><span data-stu-id="d7fe0-209">Method</span></span> |
| [<span data-ttu-id="d7fe0-210">loadCustomPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="d7fe0-210">loadCustomPropertiesAsync</span></span>](#loadcustompropertiesasynccallback-usercontext) | <span data-ttu-id="d7fe0-211">方法</span><span class="sxs-lookup"><span data-stu-id="d7fe0-211">Method</span></span> |
| [<span data-ttu-id="d7fe0-212">removeAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="d7fe0-212">removeAttachmentAsync</span></span>](#removeattachmentasyncattachmentid-options-callback) | <span data-ttu-id="d7fe0-213">方法</span><span class="sxs-lookup"><span data-stu-id="d7fe0-213">Method</span></span> |
| [<span data-ttu-id="d7fe0-214">removeHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="d7fe0-214">removeHandlerAsync</span></span>](#removehandlerasynceventtype-options-callback) | <span data-ttu-id="d7fe0-215">方法</span><span class="sxs-lookup"><span data-stu-id="d7fe0-215">Method</span></span> |
| [<span data-ttu-id="d7fe0-216">saveAsync</span><span class="sxs-lookup"><span data-stu-id="d7fe0-216">saveAsync</span></span>](#saveasyncoptions-callback) | <span data-ttu-id="d7fe0-217">方法</span><span class="sxs-lookup"><span data-stu-id="d7fe0-217">Method</span></span> |
| [<span data-ttu-id="d7fe0-218">setSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="d7fe0-218">setSelectedDataAsync</span></span>](#setselecteddataasyncdata-options-callback) | <span data-ttu-id="d7fe0-219">方法</span><span class="sxs-lookup"><span data-stu-id="d7fe0-219">Method</span></span> |

### <a name="example"></a><span data-ttu-id="d7fe0-220">示例</span><span class="sxs-lookup"><span data-stu-id="d7fe0-220">Example</span></span>

<span data-ttu-id="d7fe0-221">以下 JavaScript 代码示例显示了如何访问 Outlook 中当前项目的 `subject` 属性。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-221">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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

### <a name="members"></a><span data-ttu-id="d7fe0-222">成员</span><span class="sxs-lookup"><span data-stu-id="d7fe0-222">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetails"></a><span data-ttu-id="d7fe0-223">attachments :Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="d7fe0-223">attachments :Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span></span>

<span data-ttu-id="d7fe0-224">获取项目的附件作为数组。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-224">Gets the item's attachments as an array.</span></span> <span data-ttu-id="d7fe0-225">仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-225">Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="d7fe0-226">由于潜在的安全问题，某些类型的文件会受到 Outlook 阻止，并且不会返回。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-226">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="d7fe0-227">如需了解更多信息，请参阅 [Outlook 中阻止的附件](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519)。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-227">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="d7fe0-228">类型：</span><span class="sxs-lookup"><span data-stu-id="d7fe0-228">Type:</span></span>

*   <span data-ttu-id="d7fe0-229">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="d7fe0-229">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span></span>

##### <a name="requirements"></a><span data-ttu-id="d7fe0-230">要求</span><span class="sxs-lookup"><span data-stu-id="d7fe0-230">Requirements</span></span>

|<span data-ttu-id="d7fe0-231">要求</span><span class="sxs-lookup"><span data-stu-id="d7fe0-231">Requirement</span></span>|<span data-ttu-id="d7fe0-232">值</span><span class="sxs-lookup"><span data-stu-id="d7fe0-232">Value</span></span>|
|---|---|
|[<span data-ttu-id="d7fe0-233">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="d7fe0-233">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d7fe0-234">1.0</span><span class="sxs-lookup"><span data-stu-id="d7fe0-234">1.0</span></span>|
|[<span data-ttu-id="d7fe0-235">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="d7fe0-235">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d7fe0-236">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d7fe0-236">ReadItem</span></span>|
|[<span data-ttu-id="d7fe0-237">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="d7fe0-237">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="d7fe0-238">阅读</span><span class="sxs-lookup"><span data-stu-id="d7fe0-238">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d7fe0-239">示例</span><span class="sxs-lookup"><span data-stu-id="d7fe0-239">Example</span></span>

<span data-ttu-id="d7fe0-240">以下代码使用当前项目上所有附件的详细信息构成 HTML 字符串。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-240">The following code builds an HTML string with details of all attachments on the current item.</span></span>

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

####  <a name="bcc-recipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="d7fe0-241">bcc :[Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="d7fe0-241">bcc :[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="d7fe0-242">获取对象，该对象提供用于获取或更新邮件的密件抄送 (Bcc) 行上的收件人的方法。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-242">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="d7fe0-243">仅限撰写模式。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-243">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="d7fe0-244">类型:</span><span class="sxs-lookup"><span data-stu-id="d7fe0-244">Type:</span></span>

*   [<span data-ttu-id="d7fe0-245">收件人</span><span class="sxs-lookup"><span data-stu-id="d7fe0-245">Recipients</span></span>](/javascript/api/outlook/office.recipients)

##### <a name="requirements"></a><span data-ttu-id="d7fe0-246">要求</span><span class="sxs-lookup"><span data-stu-id="d7fe0-246">Requirements</span></span>

|<span data-ttu-id="d7fe0-247">要求</span><span class="sxs-lookup"><span data-stu-id="d7fe0-247">Requirement</span></span>|<span data-ttu-id="d7fe0-248">值</span><span class="sxs-lookup"><span data-stu-id="d7fe0-248">Value</span></span>|
|---|---|
|[<span data-ttu-id="d7fe0-249">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="d7fe0-249">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d7fe0-250">1.1</span><span class="sxs-lookup"><span data-stu-id="d7fe0-250">1.1</span></span>|
|[<span data-ttu-id="d7fe0-251">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="d7fe0-251">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d7fe0-252">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d7fe0-252">ReadItem</span></span>|
|[<span data-ttu-id="d7fe0-253">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="d7fe0-253">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="d7fe0-254">撰写</span><span class="sxs-lookup"><span data-stu-id="d7fe0-254">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="d7fe0-255">示例</span><span class="sxs-lookup"><span data-stu-id="d7fe0-255">Example</span></span>

```javascript
Office.context.mailbox.item.bcc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.bcc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.bcc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfBccRecipients = asyncResult.value;
}
```

####  <a name="body-bodyjavascriptapioutlookofficebody"></a><span data-ttu-id="d7fe0-256">body :[Body](/javascript/api/outlook/office.body)</span><span class="sxs-lookup"><span data-stu-id="d7fe0-256">body :[Body](/javascript/api/outlook/office.body)</span></span>

<span data-ttu-id="d7fe0-257">获取一个提供用于处理项目正文的方法的对象。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-257">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="d7fe0-258">类型：</span><span class="sxs-lookup"><span data-stu-id="d7fe0-258">Type:</span></span>

*   [<span data-ttu-id="d7fe0-259">Body</span><span class="sxs-lookup"><span data-stu-id="d7fe0-259">Body</span></span>](/javascript/api/outlook/office.body)

##### <a name="requirements"></a><span data-ttu-id="d7fe0-260">要求</span><span class="sxs-lookup"><span data-stu-id="d7fe0-260">Requirements</span></span>

|<span data-ttu-id="d7fe0-261">要求</span><span class="sxs-lookup"><span data-stu-id="d7fe0-261">Requirement</span></span>|<span data-ttu-id="d7fe0-262">值</span><span class="sxs-lookup"><span data-stu-id="d7fe0-262">Value</span></span>|
|---|---|
|[<span data-ttu-id="d7fe0-263">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="d7fe0-263">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d7fe0-264">1.1</span><span class="sxs-lookup"><span data-stu-id="d7fe0-264">1.1</span></span>|
|[<span data-ttu-id="d7fe0-265">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="d7fe0-265">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d7fe0-266">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d7fe0-266">ReadItem</span></span>|
|[<span data-ttu-id="d7fe0-267">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="d7fe0-267">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="d7fe0-268">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="d7fe0-268">Compose or read</span></span>|

####  <a name="cc-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="d7fe0-269">cc :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="d7fe0-269">cc :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="d7fe0-270">提供对邮件的抄送 (Cc) 收件人的访问权限。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-270">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="d7fe0-271">对象的类型和访问级别取决于当前项目的模式。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-271">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="d7fe0-272">读取模式</span><span class="sxs-lookup"><span data-stu-id="d7fe0-272">Read mode</span></span>

<span data-ttu-id="d7fe0-p106">`cc` 属性返回包含邮件的**抄送**行上所列的每个收件人的 `EmailAddressDetails` 对象的数组。集合上限为 100 个成员。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-p106">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="d7fe0-275">撰写模式</span><span class="sxs-lookup"><span data-stu-id="d7fe0-275">Compose mode</span></span>

<span data-ttu-id="d7fe0-276">`cc` 属性返回一个 `Recipients` 对象，该对象提供用于获取或更新邮件的“**抄送**”行上收件人的方法。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-276">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="d7fe0-277">类型：</span><span class="sxs-lookup"><span data-stu-id="d7fe0-277">Type:</span></span>

*   <span data-ttu-id="d7fe0-278">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="d7fe0-278">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="d7fe0-279">Requirements</span><span class="sxs-lookup"><span data-stu-id="d7fe0-279">Requirements</span></span>

|<span data-ttu-id="d7fe0-280">要求</span><span class="sxs-lookup"><span data-stu-id="d7fe0-280">Requirement</span></span>|<span data-ttu-id="d7fe0-281">值</span><span class="sxs-lookup"><span data-stu-id="d7fe0-281">Value</span></span>|
|---|---|
|[<span data-ttu-id="d7fe0-282">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="d7fe0-282">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d7fe0-283">1.0</span><span class="sxs-lookup"><span data-stu-id="d7fe0-283">1.0</span></span>|
|[<span data-ttu-id="d7fe0-284">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="d7fe0-284">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d7fe0-285">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d7fe0-285">ReadItem</span></span>|
|[<span data-ttu-id="d7fe0-286">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="d7fe0-286">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="d7fe0-287">Compose 或 Read</span><span class="sxs-lookup"><span data-stu-id="d7fe0-287">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="d7fe0-288">示例</span><span class="sxs-lookup"><span data-stu-id="d7fe0-288">Example</span></span>

```javascript
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

####  <a name="nullable-conversationid-string"></a><span data-ttu-id="d7fe0-289">(nullable) conversationId :String</span><span class="sxs-lookup"><span data-stu-id="d7fe0-289">(nullable) conversationId :String</span></span>

<span data-ttu-id="d7fe0-290">获取包含特定消息的电子邮件会话的标识符。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-290">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="d7fe0-p107">如果在阅读窗体或撰写窗体的回复中激活邮件应用程序，则此属性可以获得一个整数值。如果用户随后更改了回复邮件的主题（若发送回复），则该邮件的对话 ID 将改变且之前获取的值将不适用。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-p107">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="d7fe0-p108">对于撰写窗体的新项目，此属性获得一个 null 值。如果用户设置一个主题并保存该项目，`conversationId` 属性将返回一个值。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-p108">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="d7fe0-295">类型：</span><span class="sxs-lookup"><span data-stu-id="d7fe0-295">Type:</span></span>

*   <span data-ttu-id="d7fe0-296">String</span><span class="sxs-lookup"><span data-stu-id="d7fe0-296">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="d7fe0-297">要求</span><span class="sxs-lookup"><span data-stu-id="d7fe0-297">Requirements</span></span>

|<span data-ttu-id="d7fe0-298">要求</span><span class="sxs-lookup"><span data-stu-id="d7fe0-298">Requirement</span></span>|<span data-ttu-id="d7fe0-299">值</span><span class="sxs-lookup"><span data-stu-id="d7fe0-299">Value</span></span>|
|---|---|
|[<span data-ttu-id="d7fe0-300">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="d7fe0-300">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d7fe0-301">1.0</span><span class="sxs-lookup"><span data-stu-id="d7fe0-301">1.0</span></span>|
|[<span data-ttu-id="d7fe0-302">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="d7fe0-302">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d7fe0-303">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d7fe0-303">ReadItem</span></span>|
|[<span data-ttu-id="d7fe0-304">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="d7fe0-304">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="d7fe0-305">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="d7fe0-305">Compose or read</span></span>|

#### <a name="datetimecreated-date"></a><span data-ttu-id="d7fe0-306">dateTimeCreated :Date</span><span class="sxs-lookup"><span data-stu-id="d7fe0-306">dateTimeCreated :Date</span></span>

<span data-ttu-id="d7fe0-p109">获取项目创建的日期和时间。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-p109">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="d7fe0-309">类型：</span><span class="sxs-lookup"><span data-stu-id="d7fe0-309">Type:</span></span>

*   <span data-ttu-id="d7fe0-310">日期</span><span class="sxs-lookup"><span data-stu-id="d7fe0-310">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="d7fe0-311">Requirements</span><span class="sxs-lookup"><span data-stu-id="d7fe0-311">Requirements</span></span>

|<span data-ttu-id="d7fe0-312">要求</span><span class="sxs-lookup"><span data-stu-id="d7fe0-312">Requirement</span></span>|<span data-ttu-id="d7fe0-313">值</span><span class="sxs-lookup"><span data-stu-id="d7fe0-313">Value</span></span>|
|---|---|
|[<span data-ttu-id="d7fe0-314">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="d7fe0-314">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d7fe0-315">1.0</span><span class="sxs-lookup"><span data-stu-id="d7fe0-315">1.0</span></span>|
|[<span data-ttu-id="d7fe0-316">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="d7fe0-316">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d7fe0-317">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d7fe0-317">ReadItem</span></span>|
|[<span data-ttu-id="d7fe0-318">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="d7fe0-318">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="d7fe0-319">阅读</span><span class="sxs-lookup"><span data-stu-id="d7fe0-319">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d7fe0-320">示例</span><span class="sxs-lookup"><span data-stu-id="d7fe0-320">Example</span></span>

```javascript
var created = Office.context.mailbox.item.dateTimeCreated;
```

#### <a name="datetimemodified-date"></a><span data-ttu-id="d7fe0-321">dateTimeModified :Date</span><span class="sxs-lookup"><span data-stu-id="d7fe0-321">dateTimeModified :Date</span></span>

<span data-ttu-id="d7fe0-p110">获取项目最近一次修改的日期和时间。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-p110">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="d7fe0-324">在 Outlook for iOS 或 Outlook for Android 中不支持此成员。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-324">This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="type"></a><span data-ttu-id="d7fe0-325">类型:</span><span class="sxs-lookup"><span data-stu-id="d7fe0-325">Type:</span></span>

*   <span data-ttu-id="d7fe0-326">日期</span><span class="sxs-lookup"><span data-stu-id="d7fe0-326">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="d7fe0-327">Requirements</span><span class="sxs-lookup"><span data-stu-id="d7fe0-327">Requirements</span></span>

|<span data-ttu-id="d7fe0-328">要求</span><span class="sxs-lookup"><span data-stu-id="d7fe0-328">Requirement</span></span>|<span data-ttu-id="d7fe0-329">值</span><span class="sxs-lookup"><span data-stu-id="d7fe0-329">Value</span></span>|
|---|---|
|[<span data-ttu-id="d7fe0-330">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="d7fe0-330">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d7fe0-331">1.0</span><span class="sxs-lookup"><span data-stu-id="d7fe0-331">1.0</span></span>|
|[<span data-ttu-id="d7fe0-332">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="d7fe0-332">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d7fe0-333">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d7fe0-333">ReadItem</span></span>|
|[<span data-ttu-id="d7fe0-334">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="d7fe0-334">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="d7fe0-335">阅读</span><span class="sxs-lookup"><span data-stu-id="d7fe0-335">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d7fe0-336">示例</span><span class="sxs-lookup"><span data-stu-id="d7fe0-336">Example</span></span>

```javascript
var modified = Office.context.mailbox.item.dateTimeModified;
```

####  <a name="end-datetimejavascriptapioutlookofficetime"></a><span data-ttu-id="d7fe0-337">end :Date|[Time](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="d7fe0-337">end :Date|[Time](/javascript/api/outlook/office.time)</span></span>

<span data-ttu-id="d7fe0-338">获取或设置约会结束的日期和时间。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-338">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="d7fe0-p111">将 `end` 属性表示为协调世界时 (UTC) 的日期和时间值。可使用 [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttime) 方法将 end 属性值转换为客户端的本地日期和时间。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-p111">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="d7fe0-341">阅读模式</span><span class="sxs-lookup"><span data-stu-id="d7fe0-341">Read mode</span></span>

<span data-ttu-id="d7fe0-342">`end` 属性返回 `Date` 对象。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-342">The `end` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="d7fe0-343">撰写模式</span><span class="sxs-lookup"><span data-stu-id="d7fe0-343">Compose mode</span></span>

<span data-ttu-id="d7fe0-344">`end` 属性返回 `Time` 对象。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-344">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="d7fe0-345">使用 [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) 方法设置结束时间时，应使用 [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) 方法将客户端的本地时间转换为服务器的 UTC。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-345">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="d7fe0-346">类型：</span><span class="sxs-lookup"><span data-stu-id="d7fe0-346">Type:</span></span>

*   <span data-ttu-id="d7fe0-347">Date | [Time](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="d7fe0-347">Date | [Time](/javascript/api/outlook/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="d7fe0-348">要求</span><span class="sxs-lookup"><span data-stu-id="d7fe0-348">Requirements</span></span>

|<span data-ttu-id="d7fe0-349">要求</span><span class="sxs-lookup"><span data-stu-id="d7fe0-349">Requirement</span></span>|<span data-ttu-id="d7fe0-350">值</span><span class="sxs-lookup"><span data-stu-id="d7fe0-350">Value</span></span>|
|---|---|
|[<span data-ttu-id="d7fe0-351">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="d7fe0-351">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d7fe0-352">1.0</span><span class="sxs-lookup"><span data-stu-id="d7fe0-352">1.0</span></span>|
|[<span data-ttu-id="d7fe0-353">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="d7fe0-353">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d7fe0-354">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d7fe0-354">ReadItem</span></span>|
|[<span data-ttu-id="d7fe0-355">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="d7fe0-355">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="d7fe0-356">Compose 或 Read</span><span class="sxs-lookup"><span data-stu-id="d7fe0-356">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="d7fe0-357">示例</span><span class="sxs-lookup"><span data-stu-id="d7fe0-357">Example</span></span>

<span data-ttu-id="d7fe0-358">以下示例通过使用 `Time` 对象的 [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) 方法，设置撰写模式下约会的结束时间。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-358">The following example sets the end time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

#### <a name="enhancedlocation-enhancedlocationjavascriptapioutlookofficeenhancedlocation"></a><span data-ttu-id="d7fe0-359">enhancedLocation:[EnhancedLocation](/javascript/api/outlook/office.enhancedlocation)</span><span class="sxs-lookup"><span data-stu-id="d7fe0-359">enhancedLocation :[EnhancedLocation](/javascript/api/outlook/office.enhancedlocation)</span></span>

<span data-ttu-id="d7fe0-360">获取或设置约会的位置。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-360">Gets or sets the locations of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="d7fe0-361">阅读模式</span><span class="sxs-lookup"><span data-stu-id="d7fe0-361">Read mode</span></span>

<span data-ttu-id="d7fe0-362">`enhancedLocation`属性返回允许您获取与约会关联 （每个由表示[LocationDetails](/javascript/api/outlook/office.locationdetails)对象） 的位置套[EnhancedLocation](/javascript/api/outlook/office.enhancedlocation)对象。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-362">The `enhancedLocation` property returns an [EnhancedLocation](/javascript/api/outlook/office.enhancedlocation) object that allows you to get the set of locations (each represented by a [LocationDetails](/javascript/api/outlook/office.locationdetails) object) associated with the appointment.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="d7fe0-363">撰写模式</span><span class="sxs-lookup"><span data-stu-id="d7fe0-363">Compose mode</span></span>

<span data-ttu-id="d7fe0-364">`enhancedLocation`属性返回[EnhancedLocation](/javascript/api/outlook/office.enhancedlocation)对象，提供用于获取、 删除或添加对约会位置的方法。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-364">The `enhancedLocation` property returns an [EnhancedLocation](/javascript/api/outlook/office.enhancedlocation) object that provides methods to get, remove, or add locations on an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="d7fe0-365">类型：</span><span class="sxs-lookup"><span data-stu-id="d7fe0-365">Type:</span></span>

*   [<span data-ttu-id="d7fe0-366">EnhancedLocation</span><span class="sxs-lookup"><span data-stu-id="d7fe0-366">EnhancedLocation</span></span>](/javascript/api/outlook/office.enhancedlocation)

##### <a name="requirements"></a><span data-ttu-id="d7fe0-367">Requirements</span><span class="sxs-lookup"><span data-stu-id="d7fe0-367">Requirements</span></span>

|<span data-ttu-id="d7fe0-368">要求</span><span class="sxs-lookup"><span data-stu-id="d7fe0-368">Requirement</span></span>|<span data-ttu-id="d7fe0-369">值</span><span class="sxs-lookup"><span data-stu-id="d7fe0-369">Value</span></span>|
|---|---|
|[<span data-ttu-id="d7fe0-370">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="d7fe0-370">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d7fe0-371">预览</span><span class="sxs-lookup"><span data-stu-id="d7fe0-371">Preview</span></span>|
|[<span data-ttu-id="d7fe0-372">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="d7fe0-372">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d7fe0-373">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d7fe0-373">ReadItem</span></span>|
|[<span data-ttu-id="d7fe0-374">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="d7fe0-374">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="d7fe0-375">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="d7fe0-375">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="d7fe0-376">示例</span><span class="sxs-lookup"><span data-stu-id="d7fe0-376">Example</span></span>

<span data-ttu-id="d7fe0-377">下面的示例获取当前的位置相关联的约会。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-377">The following example gets the current locations associated with the appointment.</span></span>

```javascript
Office.context.mailbox.item.enhancedLocation.getAsync(callbackFunction);

function callbackFunction(asyncResult) {
  asyncResult.value.forEach(function (place) {
    console.log("Display name: " + place.displayName);
    console.log("Type: " + place.locationIdentifier.type);
    if (place.locationIdentifier.type == Office.MailboxEnums.LocationType.Room) {
      console.log("Email address: " + place.emailAddress);
    }
  });
}

// Sample output:
// Display name: Conf Room 14
// Type: room
// Email address: cr14@contoso.com
// Display name: Paris
// Type: custom
```

#### <a name="from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsfromjavascriptapioutlookofficefrom"></a><span data-ttu-id="d7fe0-378">from :[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[From](/javascript/api/outlook/office.from)</span><span class="sxs-lookup"><span data-stu-id="d7fe0-378">from :[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[From](/javascript/api/outlook/office.from)</span></span>

<span data-ttu-id="d7fe0-379">获取邮件发件人的电子邮件地址。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-379">Gets the email address of the sender of a message.</span></span>

<span data-ttu-id="d7fe0-p112">`from` 和 [`sender`](#sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetails) 属性表示同一个人，邮件由代理人发送的除外。在此情况下，`from` 属性表示代理程序，而 sender 属性表示代理人。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-p112">The `from` and [`sender`](#sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="d7fe0-382">`from` 属性中 `EmailAddressDetails` 对象的 `recipientType` 属性为 `undefined`。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-382">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="d7fe0-383">阅读模式</span><span class="sxs-lookup"><span data-stu-id="d7fe0-383">Read mode</span></span>

<span data-ttu-id="d7fe0-384">`from` 属性返回一个 `EmailAddressDetails` 对象。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-384">The `from` property returns an `EmailAddressDetails` object.</span></span>

```javascript
var from = Office.context.mailbox.item.from;
```

##### <a name="compose-mode"></a><span data-ttu-id="d7fe0-385">撰写模式</span><span class="sxs-lookup"><span data-stu-id="d7fe0-385">Compose mode</span></span>

<span data-ttu-id="d7fe0-386">`from` 属性返回一个 `From` 对象，该对象提供从值中进行获取的方法。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-386">The `from` property returns a `From` object that provides a method to get the from value.</span></span>

```javascript
Office.context.mailbox.item.from.getAsync(callback);

function callback(asyncResult) {
  var from = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="d7fe0-387">类型：</span><span class="sxs-lookup"><span data-stu-id="d7fe0-387">Type:</span></span>

*   <span data-ttu-id="d7fe0-388">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [From](/javascript/api/outlook/office.from)</span><span class="sxs-lookup"><span data-stu-id="d7fe0-388">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [From](/javascript/api/outlook/office.from)</span></span>

##### <a name="requirements"></a><span data-ttu-id="d7fe0-389">要求</span><span class="sxs-lookup"><span data-stu-id="d7fe0-389">Requirements</span></span>

|<span data-ttu-id="d7fe0-390">要求</span><span class="sxs-lookup"><span data-stu-id="d7fe0-390">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="d7fe0-391">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="d7fe0-391">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d7fe0-392">1.0</span><span class="sxs-lookup"><span data-stu-id="d7fe0-392">1.0</span></span>|<span data-ttu-id="d7fe0-393">1.7</span><span class="sxs-lookup"><span data-stu-id="d7fe0-393">1.7</span></span>|
|[<span data-ttu-id="d7fe0-394">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="d7fe0-394">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d7fe0-395">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d7fe0-395">ReadItem</span></span>|<span data-ttu-id="d7fe0-396">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="d7fe0-396">ReadWriteItem</span></span>|
|[<span data-ttu-id="d7fe0-397">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="d7fe0-397">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="d7fe0-398">阅读</span><span class="sxs-lookup"><span data-stu-id="d7fe0-398">Read</span></span>|<span data-ttu-id="d7fe0-399">撰写</span><span class="sxs-lookup"><span data-stu-id="d7fe0-399">Compose</span></span>|

#### <a name="internetheaders-internetheadersjavascriptapioutlookofficeinternetheaders"></a><span data-ttu-id="d7fe0-400">internetHeaders :[InternetHeaders](/javascript/api/outlook/office.internetheaders)</span><span class="sxs-lookup"><span data-stu-id="d7fe0-400">internetHeaders :[InternetHeaders](/javascript/api/outlook/office.internetheaders)</span></span>

<span data-ttu-id="d7fe0-401">获取或设置消息的 Internet 标头。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-401">Gets or sets the internet headers of a message.</span></span>

##### <a name="type"></a><span data-ttu-id="d7fe0-402">类型：</span><span class="sxs-lookup"><span data-stu-id="d7fe0-402">Type:</span></span>

*   [<span data-ttu-id="d7fe0-403">InternetHeaders</span><span class="sxs-lookup"><span data-stu-id="d7fe0-403">InternetHeaders</span></span>](/javascript/api/outlook/office.internetheaders)

##### <a name="requirements"></a><span data-ttu-id="d7fe0-404">要求</span><span class="sxs-lookup"><span data-stu-id="d7fe0-404">Requirements</span></span>

|<span data-ttu-id="d7fe0-405">要求</span><span class="sxs-lookup"><span data-stu-id="d7fe0-405">Requirement</span></span>|<span data-ttu-id="d7fe0-406">值</span><span class="sxs-lookup"><span data-stu-id="d7fe0-406">Value</span></span>|
|---|---|
|[<span data-ttu-id="d7fe0-407">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="d7fe0-407">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d7fe0-408">预览</span><span class="sxs-lookup"><span data-stu-id="d7fe0-408">Preview</span></span>|
|[<span data-ttu-id="d7fe0-409">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="d7fe0-409">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d7fe0-410">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d7fe0-410">ReadItem</span></span>|
|[<span data-ttu-id="d7fe0-411">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="d7fe0-411">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="d7fe0-412">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="d7fe0-412">Compose or read</span></span>|

#### <a name="internetmessageid-string"></a><span data-ttu-id="d7fe0-413">internetMessageId :String</span><span class="sxs-lookup"><span data-stu-id="d7fe0-413">internetMessageId :String</span></span>

<span data-ttu-id="d7fe0-p113">获取电子邮件的 Internet 消息标识符。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-p113">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="d7fe0-416">类型：</span><span class="sxs-lookup"><span data-stu-id="d7fe0-416">Type:</span></span>

*   <span data-ttu-id="d7fe0-417">String</span><span class="sxs-lookup"><span data-stu-id="d7fe0-417">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="d7fe0-418">要求</span><span class="sxs-lookup"><span data-stu-id="d7fe0-418">Requirements</span></span>

|<span data-ttu-id="d7fe0-419">要求</span><span class="sxs-lookup"><span data-stu-id="d7fe0-419">Requirement</span></span>|<span data-ttu-id="d7fe0-420">值</span><span class="sxs-lookup"><span data-stu-id="d7fe0-420">Value</span></span>|
|---|---|
|[<span data-ttu-id="d7fe0-421">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="d7fe0-421">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d7fe0-422">1.0</span><span class="sxs-lookup"><span data-stu-id="d7fe0-422">1.0</span></span>|
|[<span data-ttu-id="d7fe0-423">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="d7fe0-423">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d7fe0-424">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d7fe0-424">ReadItem</span></span>|
|[<span data-ttu-id="d7fe0-425">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="d7fe0-425">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="d7fe0-426">阅读</span><span class="sxs-lookup"><span data-stu-id="d7fe0-426">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d7fe0-427">示例</span><span class="sxs-lookup"><span data-stu-id="d7fe0-427">Example</span></span>

```javascript
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

#### <a name="itemclass-string"></a><span data-ttu-id="d7fe0-428">itemClass :String</span><span class="sxs-lookup"><span data-stu-id="d7fe0-428">itemClass :String</span></span>

<span data-ttu-id="d7fe0-p114">获取选定项目的 Exchange Web 服务项目类。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-p114">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="d7fe0-p115">`itemClass` 属性指定所选项目的邮件类别。以下是邮件或约会项目的默认邮件类别。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-p115">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

|<span data-ttu-id="d7fe0-433">类型</span><span class="sxs-lookup"><span data-stu-id="d7fe0-433">Type</span></span>|<span data-ttu-id="d7fe0-434">说明</span><span class="sxs-lookup"><span data-stu-id="d7fe0-434">Description</span></span>|<span data-ttu-id="d7fe0-435">项目类</span><span class="sxs-lookup"><span data-stu-id="d7fe0-435">item class</span></span>|
|---|---|---|
|<span data-ttu-id="d7fe0-436">约会项目</span><span class="sxs-lookup"><span data-stu-id="d7fe0-436">Appointment items</span></span>|<span data-ttu-id="d7fe0-437">这些是项目类 `IPM.Appointment` 或 `IPM.Appointment.Occurrence` 的日历项目。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-437">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurrence`.</span></span>|`IPM.Appointment`<br />`IPM.Appointment.Occurrence`|
|<span data-ttu-id="d7fe0-438">邮件项目</span><span class="sxs-lookup"><span data-stu-id="d7fe0-438">Message items</span></span>|<span data-ttu-id="d7fe0-439">这些项目包括具有默认邮件类别 `IPM.Note` 的电子邮件，以及使用 `IPM.Schedule.Meeting` 作为基础邮件类别的会议请求、响应和取消。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-439">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span>|`IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled`|

<span data-ttu-id="d7fe0-440">你可以创建用于扩展默认邮件类别的自定义邮件类别，例如，自定义约会邮件类别 `IPM.Appointment.Contoso`。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-440">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="d7fe0-441">类型：</span><span class="sxs-lookup"><span data-stu-id="d7fe0-441">Type:</span></span>

*   <span data-ttu-id="d7fe0-442">String</span><span class="sxs-lookup"><span data-stu-id="d7fe0-442">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="d7fe0-443">要求</span><span class="sxs-lookup"><span data-stu-id="d7fe0-443">Requirements</span></span>

|<span data-ttu-id="d7fe0-444">要求</span><span class="sxs-lookup"><span data-stu-id="d7fe0-444">Requirement</span></span>|<span data-ttu-id="d7fe0-445">值</span><span class="sxs-lookup"><span data-stu-id="d7fe0-445">Value</span></span>|
|---|---|
|[<span data-ttu-id="d7fe0-446">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="d7fe0-446">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d7fe0-447">1.0</span><span class="sxs-lookup"><span data-stu-id="d7fe0-447">1.0</span></span>|
|[<span data-ttu-id="d7fe0-448">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="d7fe0-448">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d7fe0-449">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d7fe0-449">ReadItem</span></span>|
|[<span data-ttu-id="d7fe0-450">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="d7fe0-450">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="d7fe0-451">阅读</span><span class="sxs-lookup"><span data-stu-id="d7fe0-451">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d7fe0-452">示例</span><span class="sxs-lookup"><span data-stu-id="d7fe0-452">Example</span></span>

```javascript
var itemClass = Office.context.mailbox.item.itemClass;
```

#### <a name="nullable-itemid-string"></a><span data-ttu-id="d7fe0-453">(nullable) itemId :String</span><span class="sxs-lookup"><span data-stu-id="d7fe0-453">(nullable) itemId :String</span></span>

<span data-ttu-id="d7fe0-p116">获取当前项目的 Exchange Web 服务项目标识符。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-p116">Gets the Exchange Web Services item identifier for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="d7fe0-456">`itemId` 属性返回的标识符与 Exchange Web 服务项目标识符相同。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-456">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="d7fe0-457">`itemId` 属性与 Outlook 条目 ID 或 Outlook REST API 使用的 ID 不同。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-457">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="d7fe0-458">使用此值进行 REST API 调用前，应使用 [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string) 对它进行转换。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-458">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="d7fe0-459">有关详细信息，请参阅[从 Outlook 加载项使用 Outlook REST API](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id)。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-459">For more details, see [Use the Outlook REST APIs from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

<span data-ttu-id="d7fe0-p118">`itemId` 属性在撰写模式下不可用。如果需要项目标识符，[`saveAsync`](#saveasyncoptions-callback) 方法可用于将项目保存到存储，这将在回调函数的 [`AsyncResult.value`](/javascript/api/office/office.asyncresult) 参数中返回项目标识符。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-p118">The `itemId` property is not available in compose mode. If an item identifier is required, the [`saveAsync`](#saveasyncoptions-callback) method can be used to save the item to the store, which will return the item identifier in the [`AsyncResult.value`](/javascript/api/office/office.asyncresult) parameter in the callback function.</span></span>

##### <a name="type"></a><span data-ttu-id="d7fe0-462">类型:</span><span class="sxs-lookup"><span data-stu-id="d7fe0-462">Type:</span></span>

*   <span data-ttu-id="d7fe0-463">String</span><span class="sxs-lookup"><span data-stu-id="d7fe0-463">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="d7fe0-464">要求</span><span class="sxs-lookup"><span data-stu-id="d7fe0-464">Requirements</span></span>

|<span data-ttu-id="d7fe0-465">要求</span><span class="sxs-lookup"><span data-stu-id="d7fe0-465">Requirement</span></span>|<span data-ttu-id="d7fe0-466">值</span><span class="sxs-lookup"><span data-stu-id="d7fe0-466">Value</span></span>|
|---|---|
|[<span data-ttu-id="d7fe0-467">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="d7fe0-467">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d7fe0-468">1.0</span><span class="sxs-lookup"><span data-stu-id="d7fe0-468">1.0</span></span>|
|[<span data-ttu-id="d7fe0-469">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="d7fe0-469">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d7fe0-470">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d7fe0-470">ReadItem</span></span>|
|[<span data-ttu-id="d7fe0-471">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="d7fe0-471">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="d7fe0-472">阅读</span><span class="sxs-lookup"><span data-stu-id="d7fe0-472">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d7fe0-473">示例</span><span class="sxs-lookup"><span data-stu-id="d7fe0-473">Example</span></span>

<span data-ttu-id="d7fe0-p119">以下代码检查项目标识符是否存在。如果 `itemId` 属性返回 `null` 或 `undefined`，则将项目保存到存储，并从异步结果中获取项目标识符。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-p119">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

```javascript
var itemId = Office.context.mailbox.item.itemId;
if (itemId === null || itemId == undefined) {
  Office.context.mailbox.item.saveAsync(function(result){
    itemId = result.value;
  });
}
```

####  <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtype"></a><span data-ttu-id="d7fe0-476">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype)</span><span class="sxs-lookup"><span data-stu-id="d7fe0-476">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype)</span></span>

<span data-ttu-id="d7fe0-477">获取实例表示的项的类型。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-477">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="d7fe0-478">`itemType` 属性返回其中一个 `ItemType` 枚举值，指示 `item` 对象实例是邮件还是约会。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-478">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="d7fe0-479">类型：</span><span class="sxs-lookup"><span data-stu-id="d7fe0-479">Type:</span></span>

*   [<span data-ttu-id="d7fe0-480">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="d7fe0-480">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook/office.mailboxenums.itemtype)

##### <a name="requirements"></a><span data-ttu-id="d7fe0-481">Requirements</span><span class="sxs-lookup"><span data-stu-id="d7fe0-481">Requirements</span></span>

|<span data-ttu-id="d7fe0-482">要求</span><span class="sxs-lookup"><span data-stu-id="d7fe0-482">Requirement</span></span>|<span data-ttu-id="d7fe0-483">值</span><span class="sxs-lookup"><span data-stu-id="d7fe0-483">Value</span></span>|
|---|---|
|[<span data-ttu-id="d7fe0-484">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="d7fe0-484">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d7fe0-485">1.0</span><span class="sxs-lookup"><span data-stu-id="d7fe0-485">1.0</span></span>|
|[<span data-ttu-id="d7fe0-486">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="d7fe0-486">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d7fe0-487">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d7fe0-487">ReadItem</span></span>|
|[<span data-ttu-id="d7fe0-488">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="d7fe0-488">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="d7fe0-489">Compose 或 Read</span><span class="sxs-lookup"><span data-stu-id="d7fe0-489">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="d7fe0-490">示例</span><span class="sxs-lookup"><span data-stu-id="d7fe0-490">Example</span></span>

```javascript
if (Office.context.mailbox.item.itemType == Office.MailboxEnums.ItemType.Message)
  // do something
else
  // do something else
```

####  <a name="location-stringlocationjavascriptapioutlookofficelocation"></a><span data-ttu-id="d7fe0-491">location :String|[Location](/javascript/api/outlook/office.location)</span><span class="sxs-lookup"><span data-stu-id="d7fe0-491">location :String|[Location](/javascript/api/outlook/office.location)</span></span>

<span data-ttu-id="d7fe0-492">获取或设置约会的位置。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-492">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="d7fe0-493">阅读模式</span><span class="sxs-lookup"><span data-stu-id="d7fe0-493">Read mode</span></span>

<span data-ttu-id="d7fe0-494">`location` 属性返回一个包含约会位置的字符串。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-494">The `location` property returns a string that contains the location of the appointment.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="d7fe0-495">撰写模式</span><span class="sxs-lookup"><span data-stu-id="d7fe0-495">Compose mode</span></span>

<span data-ttu-id="d7fe0-496">`location` 属性返回一个 `Location` 对象，该对象提供用于获取和设置约会位置的方法。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-496">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="d7fe0-497">类型：</span><span class="sxs-lookup"><span data-stu-id="d7fe0-497">Type:</span></span>

*   <span data-ttu-id="d7fe0-498">String | [Location](/javascript/api/outlook/office.location)</span><span class="sxs-lookup"><span data-stu-id="d7fe0-498">String | [Location](/javascript/api/outlook/office.location)</span></span>

##### <a name="requirements"></a><span data-ttu-id="d7fe0-499">Requirements</span><span class="sxs-lookup"><span data-stu-id="d7fe0-499">Requirements</span></span>

|<span data-ttu-id="d7fe0-500">要求</span><span class="sxs-lookup"><span data-stu-id="d7fe0-500">Requirement</span></span>|<span data-ttu-id="d7fe0-501">值</span><span class="sxs-lookup"><span data-stu-id="d7fe0-501">Value</span></span>|
|---|---|
|[<span data-ttu-id="d7fe0-502">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="d7fe0-502">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d7fe0-503">1.0</span><span class="sxs-lookup"><span data-stu-id="d7fe0-503">1.0</span></span>|
|[<span data-ttu-id="d7fe0-504">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="d7fe0-504">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d7fe0-505">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d7fe0-505">ReadItem</span></span>|
|[<span data-ttu-id="d7fe0-506">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="d7fe0-506">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="d7fe0-507">Compose 或 Read</span><span class="sxs-lookup"><span data-stu-id="d7fe0-507">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="d7fe0-508">示例</span><span class="sxs-lookup"><span data-stu-id="d7fe0-508">Example</span></span>

```javascript
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

#### <a name="normalizedsubject-string"></a><span data-ttu-id="d7fe0-509">normalizedSubject :String</span><span class="sxs-lookup"><span data-stu-id="d7fe0-509">normalizedSubject :String</span></span>

<span data-ttu-id="d7fe0-p120">获取删除了所有前缀（包括 `RE:` 和 `FWD:`）的项目主题。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-p120">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="d7fe0-p121">normalizedSubject 属性获取包含由电子邮件程序添加的任何标准前缀（如 `RE:` 和 `FW:`）的项目主题。若要获取包含完整前缀的项目主题，请使用 [`subject`](#subject-stringsubjectjavascriptapioutlookofficesubject) 属性。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-p121">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubjectjavascriptapioutlookofficesubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="d7fe0-514">类型:</span><span class="sxs-lookup"><span data-stu-id="d7fe0-514">Type:</span></span>

*   <span data-ttu-id="d7fe0-515">String</span><span class="sxs-lookup"><span data-stu-id="d7fe0-515">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="d7fe0-516">要求</span><span class="sxs-lookup"><span data-stu-id="d7fe0-516">Requirements</span></span>

|<span data-ttu-id="d7fe0-517">要求</span><span class="sxs-lookup"><span data-stu-id="d7fe0-517">Requirement</span></span>|<span data-ttu-id="d7fe0-518">值</span><span class="sxs-lookup"><span data-stu-id="d7fe0-518">Value</span></span>|
|---|---|
|[<span data-ttu-id="d7fe0-519">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="d7fe0-519">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d7fe0-520">1.0</span><span class="sxs-lookup"><span data-stu-id="d7fe0-520">1.0</span></span>|
|[<span data-ttu-id="d7fe0-521">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="d7fe0-521">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d7fe0-522">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d7fe0-522">ReadItem</span></span>|
|[<span data-ttu-id="d7fe0-523">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="d7fe0-523">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="d7fe0-524">阅读</span><span class="sxs-lookup"><span data-stu-id="d7fe0-524">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d7fe0-525">示例</span><span class="sxs-lookup"><span data-stu-id="d7fe0-525">Example</span></span>

```javascript
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
```

####  <a name="notificationmessages-notificationmessagesjavascriptapioutlookofficenotificationmessages"></a><span data-ttu-id="d7fe0-526">notificationMessages :[NotificationMessages](/javascript/api/outlook/office.notificationmessages)</span><span class="sxs-lookup"><span data-stu-id="d7fe0-526">notificationMessages :[NotificationMessages](/javascript/api/outlook/office.notificationmessages)</span></span>

<span data-ttu-id="d7fe0-527">获取项目的通知邮件。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-527">Gets the notification messages for an item.</span></span>

##### <a name="type"></a><span data-ttu-id="d7fe0-528">类型：</span><span class="sxs-lookup"><span data-stu-id="d7fe0-528">Type:</span></span>

*   [<span data-ttu-id="d7fe0-529">NotificationMessages</span><span class="sxs-lookup"><span data-stu-id="d7fe0-529">NotificationMessages</span></span>](/javascript/api/outlook/office.notificationmessages)

##### <a name="requirements"></a><span data-ttu-id="d7fe0-530">Requirements</span><span class="sxs-lookup"><span data-stu-id="d7fe0-530">Requirements</span></span>

|<span data-ttu-id="d7fe0-531">要求</span><span class="sxs-lookup"><span data-stu-id="d7fe0-531">Requirement</span></span>|<span data-ttu-id="d7fe0-532">值</span><span class="sxs-lookup"><span data-stu-id="d7fe0-532">Value</span></span>|
|---|---|
|[<span data-ttu-id="d7fe0-533">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="d7fe0-533">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d7fe0-534">1.3</span><span class="sxs-lookup"><span data-stu-id="d7fe0-534">1.3</span></span>|
|[<span data-ttu-id="d7fe0-535">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="d7fe0-535">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d7fe0-536">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d7fe0-536">ReadItem</span></span>|
|[<span data-ttu-id="d7fe0-537">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="d7fe0-537">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="d7fe0-538">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="d7fe0-538">Compose or read</span></span>|

####  <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="d7fe0-539">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="d7fe0-539">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="d7fe0-540">提供对事件的可选与会者的访问权限。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-540">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="d7fe0-541">对象的类型和访问级别取决于当前项目的模式。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-541">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="d7fe0-542">读取模式</span><span class="sxs-lookup"><span data-stu-id="d7fe0-542">Read mode</span></span>

<span data-ttu-id="d7fe0-543">`optionalAttendees` 属性返回一个数组，其中包含每个可选与会者的 `EmailAddressDetails` 对象。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-543">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="d7fe0-544">撰写模式</span><span class="sxs-lookup"><span data-stu-id="d7fe0-544">Compose mode</span></span>

<span data-ttu-id="d7fe0-545">`optionalAttendees` 属性返回一个 `Recipients` 对象，该对象提供用于获取或更新可选与会者的方法。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-545">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="d7fe0-546">类型：</span><span class="sxs-lookup"><span data-stu-id="d7fe0-546">Type:</span></span>

*   <span data-ttu-id="d7fe0-547">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="d7fe0-547">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="d7fe0-548">Requirements</span><span class="sxs-lookup"><span data-stu-id="d7fe0-548">Requirements</span></span>

|<span data-ttu-id="d7fe0-549">要求</span><span class="sxs-lookup"><span data-stu-id="d7fe0-549">Requirement</span></span>|<span data-ttu-id="d7fe0-550">值</span><span class="sxs-lookup"><span data-stu-id="d7fe0-550">Value</span></span>|
|---|---|
|[<span data-ttu-id="d7fe0-551">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="d7fe0-551">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d7fe0-552">1.0</span><span class="sxs-lookup"><span data-stu-id="d7fe0-552">1.0</span></span>|
|[<span data-ttu-id="d7fe0-553">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="d7fe0-553">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d7fe0-554">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d7fe0-554">ReadItem</span></span>|
|[<span data-ttu-id="d7fe0-555">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="d7fe0-555">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="d7fe0-556">Compose 或 Read</span><span class="sxs-lookup"><span data-stu-id="d7fe0-556">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="d7fe0-557">示例</span><span class="sxs-lookup"><span data-stu-id="d7fe0-557">Example</span></span>

```javascript
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

#### <a name="organizer-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsorganizerjavascriptapioutlookofficeorganizer"></a><span data-ttu-id="d7fe0-558">organizer :[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[Organizer](/javascript/api/outlook/office.organizer)</span><span class="sxs-lookup"><span data-stu-id="d7fe0-558">organizer :[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[Organizer](/javascript/api/outlook/office.organizer)</span></span>

<span data-ttu-id="d7fe0-559">获取指定会议的组织者的电子邮件地址。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-559">Gets the email address of the organizer for a specified meeting.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="d7fe0-560">阅读模式</span><span class="sxs-lookup"><span data-stu-id="d7fe0-560">Read mode</span></span>

<span data-ttu-id="d7fe0-561">`organizer` 属性返回 [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) 对象，它表示会议组织者。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-561">The `organizer` property returns an [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) object that represents the meeting organizer.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="d7fe0-562">撰写模式</span><span class="sxs-lookup"><span data-stu-id="d7fe0-562">Compose mode</span></span>

<span data-ttu-id="d7fe0-563">`organizer` 属性返回 [Organizer](/javascript/api/outlook/office.organizer) 对象，该对象提供获取组织者值的方法。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-563">The `organizer` property returns an [Organizer](/javascript/api/outlook/office.organizer) object that provides a method to get the organizer value.</span></span>

##### <a name="type"></a><span data-ttu-id="d7fe0-564">类型：</span><span class="sxs-lookup"><span data-stu-id="d7fe0-564">Type:</span></span>

*   <span data-ttu-id="d7fe0-565">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [Organizer](/javascript/api/outlook/office.organizer)</span><span class="sxs-lookup"><span data-stu-id="d7fe0-565">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [Organizer](/javascript/api/outlook/office.organizer)</span></span>

##### <a name="requirements"></a><span data-ttu-id="d7fe0-566">要求</span><span class="sxs-lookup"><span data-stu-id="d7fe0-566">Requirements</span></span>

|<span data-ttu-id="d7fe0-567">要求</span><span class="sxs-lookup"><span data-stu-id="d7fe0-567">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="d7fe0-568">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="d7fe0-568">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d7fe0-569">1.0</span><span class="sxs-lookup"><span data-stu-id="d7fe0-569">1.0</span></span>|<span data-ttu-id="d7fe0-570">1.7</span><span class="sxs-lookup"><span data-stu-id="d7fe0-570">1.7</span></span>|
|[<span data-ttu-id="d7fe0-571">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="d7fe0-571">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d7fe0-572">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d7fe0-572">ReadItem</span></span>|<span data-ttu-id="d7fe0-573">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="d7fe0-573">ReadWriteItem</span></span>|
|[<span data-ttu-id="d7fe0-574">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="d7fe0-574">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="d7fe0-575">阅读</span><span class="sxs-lookup"><span data-stu-id="d7fe0-575">Read</span></span>|<span data-ttu-id="d7fe0-576">撰写</span><span class="sxs-lookup"><span data-stu-id="d7fe0-576">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="d7fe0-577">示例</span><span class="sxs-lookup"><span data-stu-id="d7fe0-577">Example</span></span>

```javascript
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
```

#### <a name="nullable-recurrence-recurrencejavascriptapioutlookofficerecurrence"></a><span data-ttu-id="d7fe0-578">(nullable) recurrence :[Recurrence](/javascript/api/outlook/office.recurrence)</span><span class="sxs-lookup"><span data-stu-id="d7fe0-578">(nullable) recurrence :[Recurrence](/javascript/api/outlook/office.recurrence)</span></span>

<span data-ttu-id="d7fe0-579">获取或设置约会的定期模式。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-579">Gets or sets the recurrence pattern of an appointment.</span></span> <span data-ttu-id="d7fe0-580">获取或设置会议请求的定期模式。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-580">Gets the recurrence pattern of a meeting request.</span></span> <span data-ttu-id="d7fe0-581">阅读撰写约会项目的模式。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-581">Read and compose modes for appointment items.</span></span> <span data-ttu-id="d7fe0-582">会议请求项目的阅读模式。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-582">Read mode for meeting request items.</span></span>

<span data-ttu-id="d7fe0-583">如果项目是一个系列或系列中的一个实例，则 `recurrence` 属性将返回定期约会的 [recurrence](/javascript/api/outlook/office.recurrence) 对象或会议请求。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-583">The `recurrence` property returns a [recurrence](/javascript/api/outlook/office.recurrence) object for recurring appointments or meetings requests if an item is a series or an instance in a series.</span></span> <span data-ttu-id="d7fe0-584">针对单个约会和单个约会的会议请求返回 `null`。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-584">`null` is returned for single appointments and meeting requests of single appointments.</span></span> <span data-ttu-id="d7fe0-585">针对非会议请求的邮件返回 `undefined`。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-585">`undefined` is returned for messages that are not meeting requests.</span></span>

> <span data-ttu-id="d7fe0-586">注意：会议请求的 `itemClass` 值为 IPM.Schedule.Meeting.Request。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-586">Note: Meeting requests have an `itemClass` value of IPM.Schedule.Meeting.Request.</span></span>

> <span data-ttu-id="d7fe0-587">注意：如果 recurrence 对象为 `null`，则这表示对象是单个约会或单个约会的会议请求，而不是系列的一部分。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-587">Note: If the recurrence object is `null`, this indicates that the object is a single appointment or a meeting request of a single appointment and NOT a part of a series.</span></span>

##### <a name="type"></a><span data-ttu-id="d7fe0-588">类型：</span><span class="sxs-lookup"><span data-stu-id="d7fe0-588">Type:</span></span>

* [<span data-ttu-id="d7fe0-589">Recurrence</span><span class="sxs-lookup"><span data-stu-id="d7fe0-589">Recurrence</span></span>](/javascript/api/outlook/office.recurrence)

|<span data-ttu-id="d7fe0-590">要求</span><span class="sxs-lookup"><span data-stu-id="d7fe0-590">Requirement</span></span>|<span data-ttu-id="d7fe0-591">值</span><span class="sxs-lookup"><span data-stu-id="d7fe0-591">Value</span></span>|
|---|---|
|[<span data-ttu-id="d7fe0-592">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="d7fe0-592">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d7fe0-593">1.7</span><span class="sxs-lookup"><span data-stu-id="d7fe0-593">1.7</span></span>|
|[<span data-ttu-id="d7fe0-594">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="d7fe0-594">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d7fe0-595">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d7fe0-595">ReadItem</span></span>|
|[<span data-ttu-id="d7fe0-596">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="d7fe0-596">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="d7fe0-597">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="d7fe0-597">Compose or read</span></span>|

####  <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="d7fe0-598">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="d7fe0-598">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="d7fe0-599">提供对事件的必需与会者的访问权限。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-599">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="d7fe0-600">对象的类型和访问级别取决于当前项目的模式。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-600">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="d7fe0-601">读取模式</span><span class="sxs-lookup"><span data-stu-id="d7fe0-601">Read mode</span></span>

<span data-ttu-id="d7fe0-602">`requiredAttendees` 属性返回一个数组，其中包含每个必需与会者的 `EmailAddressDetails` 对象。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-602">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="d7fe0-603">撰写模式</span><span class="sxs-lookup"><span data-stu-id="d7fe0-603">Compose mode</span></span>

<span data-ttu-id="d7fe0-604">`requiredAttendees` 属性返回一个 `Recipients` 对象，该对象提供用于获取或更新必需与会者的方法。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-604">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="d7fe0-605">类型：</span><span class="sxs-lookup"><span data-stu-id="d7fe0-605">Type:</span></span>

*   <span data-ttu-id="d7fe0-606">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="d7fe0-606">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="d7fe0-607">Requirements</span><span class="sxs-lookup"><span data-stu-id="d7fe0-607">Requirements</span></span>

|<span data-ttu-id="d7fe0-608">要求</span><span class="sxs-lookup"><span data-stu-id="d7fe0-608">Requirement</span></span>|<span data-ttu-id="d7fe0-609">值</span><span class="sxs-lookup"><span data-stu-id="d7fe0-609">Value</span></span>|
|---|---|
|[<span data-ttu-id="d7fe0-610">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="d7fe0-610">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d7fe0-611">1.0</span><span class="sxs-lookup"><span data-stu-id="d7fe0-611">1.0</span></span>|
|[<span data-ttu-id="d7fe0-612">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="d7fe0-612">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d7fe0-613">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d7fe0-613">ReadItem</span></span>|
|[<span data-ttu-id="d7fe0-614">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="d7fe0-614">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="d7fe0-615">Compose 或 Read</span><span class="sxs-lookup"><span data-stu-id="d7fe0-615">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="d7fe0-616">示例</span><span class="sxs-lookup"><span data-stu-id="d7fe0-616">Example</span></span>

```javascript
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
}
```

#### <a name="sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetails"></a><span data-ttu-id="d7fe0-617">sender :[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="d7fe0-617">sender :[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)</span></span>

<span data-ttu-id="d7fe0-p126">获取电子邮件发件人的电子邮件地址。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-p126">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="d7fe0-p127">[`from`](#from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsfromjavascriptapioutlookofficefrom) 和 `sender` 属性表示同一个人，邮件由代理人发送的除外。在此情况下，`from` 属性表示代理程序，而 sender 属性表示代理人。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-p127">The [`from`](#from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsfromjavascriptapioutlookofficefrom) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="d7fe0-622">`sender` 属性中 `EmailAddressDetails` 对象的 `recipientType` 属性为 `undefined`。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-622">The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="d7fe0-623">类型:</span><span class="sxs-lookup"><span data-stu-id="d7fe0-623">Type:</span></span>

*   [<span data-ttu-id="d7fe0-624">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="d7fe0-624">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="d7fe0-625">Requirements</span><span class="sxs-lookup"><span data-stu-id="d7fe0-625">Requirements</span></span>

|<span data-ttu-id="d7fe0-626">要求</span><span class="sxs-lookup"><span data-stu-id="d7fe0-626">Requirement</span></span>|<span data-ttu-id="d7fe0-627">值</span><span class="sxs-lookup"><span data-stu-id="d7fe0-627">Value</span></span>|
|---|---|
|[<span data-ttu-id="d7fe0-628">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="d7fe0-628">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d7fe0-629">1.0</span><span class="sxs-lookup"><span data-stu-id="d7fe0-629">1.0</span></span>|
|[<span data-ttu-id="d7fe0-630">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="d7fe0-630">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d7fe0-631">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d7fe0-631">ReadItem</span></span>|
|[<span data-ttu-id="d7fe0-632">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="d7fe0-632">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="d7fe0-633">阅读</span><span class="sxs-lookup"><span data-stu-id="d7fe0-633">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d7fe0-634">示例</span><span class="sxs-lookup"><span data-stu-id="d7fe0-634">Example</span></span>

```javascript
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
```

#### <a name="nullable-seriesid-string"></a><span data-ttu-id="d7fe0-635">(nullable) seriesId :String</span><span class="sxs-lookup"><span data-stu-id="d7fe0-635">(nullable) seriesId :String</span></span>

<span data-ttu-id="d7fe0-636">获取实例所属的系列的 ID。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-636">Gets the id of the series that an instance belongs to.</span></span>

<span data-ttu-id="d7fe0-637">在 OWA 和 Outlook 中，`seriesId` 返回此项目所属的父（系列）项目的 Exchange Web 服务 (EWS) ID。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-637">In OWA and Outlook, the `seriesId` returns the Exchange Web Services (EWS) ID of the parent (series) item that this item belongs to.</span></span> <span data-ttu-id="d7fe0-638">但是，在 iOS 和 Android 中，`seriesId` 返回父项目的其余部分 ID。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-638">However, in iOS and Android, the `seriesId` returns the REST ID of the parent item.</span></span>

> [!NOTE]
> <span data-ttu-id="d7fe0-639">`seriesId` 属性返回的标识符与 Exchange Web 服务项目标识符相同。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-639">The identifier returned by the `seriesId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="d7fe0-640">`seriesId` 属性与 Outlook REST API 使用的 Outlook ID 不同。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-640">The `seriesId` property is not identical to the Outlook IDs used by the Outlook REST API.</span></span> <span data-ttu-id="d7fe0-641">使用此值进行 REST API 调用前，应使用 [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string) 对它进行转换。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-641">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="d7fe0-642">有关详细信息，请参阅[从 Outlook 加载项使用 Outlook REST API](https://docs.microsoft.com/outlook/add-ins/use-rest-api)。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-642">For more details, see [Use the Outlook REST APIs from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/use-rest-api).</span></span>

<span data-ttu-id="d7fe0-643">`seriesId` 属性对于没有父项目（如单个约会、系列项目或会议请求）的项目返回 `null`，对于非会议请求的任何其他项目，返回 `undefined`。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-643">The `seriesId` property returns `null` for items that do not have parent items such as single appointments, series items, or meeting requests and returns `undefined` for any other items that are not meeting requests.</span></span>

##### <a name="type"></a><span data-ttu-id="d7fe0-644">类型：</span><span class="sxs-lookup"><span data-stu-id="d7fe0-644">Type:</span></span>

* <span data-ttu-id="d7fe0-645">String</span><span class="sxs-lookup"><span data-stu-id="d7fe0-645">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="d7fe0-646">要求</span><span class="sxs-lookup"><span data-stu-id="d7fe0-646">Requirements</span></span>

|<span data-ttu-id="d7fe0-647">要求</span><span class="sxs-lookup"><span data-stu-id="d7fe0-647">Requirement</span></span>|<span data-ttu-id="d7fe0-648">值</span><span class="sxs-lookup"><span data-stu-id="d7fe0-648">Value</span></span>|
|---|---|
|[<span data-ttu-id="d7fe0-649">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="d7fe0-649">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d7fe0-650">1.7</span><span class="sxs-lookup"><span data-stu-id="d7fe0-650">1.7</span></span>|
|[<span data-ttu-id="d7fe0-651">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="d7fe0-651">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d7fe0-652">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d7fe0-652">ReadItem</span></span>|
|[<span data-ttu-id="d7fe0-653">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="d7fe0-653">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="d7fe0-654">Compose 或 Read</span><span class="sxs-lookup"><span data-stu-id="d7fe0-654">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="d7fe0-655">示例</span><span class="sxs-lookup"><span data-stu-id="d7fe0-655">Example</span></span>

```javascript
var seriesId = Office.context.mailbox.item.seriesId;
var isSeries = (seriesId == null);
```

####  <a name="start-datetimejavascriptapioutlookofficetime"></a><span data-ttu-id="d7fe0-656">start :Date|[Time](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="d7fe0-656">start :Date|[Time](/javascript/api/outlook/office.time)</span></span>

<span data-ttu-id="d7fe0-657">获取或设置约会开始的日期和时间。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-657">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="d7fe0-p130">将 `start` 属性表示为协调世界时 (UTC) 的日期和时间值。可使用 [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttime) 方法将该值转换为客户端的本地日期和时间。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-p130">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="d7fe0-660">阅读模式</span><span class="sxs-lookup"><span data-stu-id="d7fe0-660">Read mode</span></span>

<span data-ttu-id="d7fe0-661">`start` 属性返回 `Date` 对象。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-661">The `start` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="d7fe0-662">撰写模式</span><span class="sxs-lookup"><span data-stu-id="d7fe0-662">Compose mode</span></span>

<span data-ttu-id="d7fe0-663">`start` 属性返回 `Time` 对象。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-663">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="d7fe0-664">使用 [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) 方法设置开始时间时，应使用 [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) 方法将客户端的本地时间转换为服务器的 UTC。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-664">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="d7fe0-665">类型：</span><span class="sxs-lookup"><span data-stu-id="d7fe0-665">Type:</span></span>

*   <span data-ttu-id="d7fe0-666">Date | [Time](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="d7fe0-666">Date | [Time](/javascript/api/outlook/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="d7fe0-667">Requirements</span><span class="sxs-lookup"><span data-stu-id="d7fe0-667">Requirements</span></span>

|<span data-ttu-id="d7fe0-668">要求</span><span class="sxs-lookup"><span data-stu-id="d7fe0-668">Requirement</span></span>|<span data-ttu-id="d7fe0-669">值</span><span class="sxs-lookup"><span data-stu-id="d7fe0-669">Value</span></span>|
|---|---|
|[<span data-ttu-id="d7fe0-670">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="d7fe0-670">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d7fe0-671">1.0</span><span class="sxs-lookup"><span data-stu-id="d7fe0-671">1.0</span></span>|
|[<span data-ttu-id="d7fe0-672">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="d7fe0-672">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d7fe0-673">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d7fe0-673">ReadItem</span></span>|
|[<span data-ttu-id="d7fe0-674">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="d7fe0-674">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="d7fe0-675">Compose 或 Read</span><span class="sxs-lookup"><span data-stu-id="d7fe0-675">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="d7fe0-676">示例</span><span class="sxs-lookup"><span data-stu-id="d7fe0-676">Example</span></span>

<span data-ttu-id="d7fe0-677">以下示例通过使用 `Time` 对象的 [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) 方法，设置撰写模式下约会的开始时间。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-677">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

####  <a name="subject-stringsubjectjavascriptapioutlookofficesubject"></a><span data-ttu-id="d7fe0-678">subject :String|[Subject](/javascript/api/outlook/office.subject)</span><span class="sxs-lookup"><span data-stu-id="d7fe0-678">subject :String|[Subject](/javascript/api/outlook/office.subject)</span></span>

<span data-ttu-id="d7fe0-679">获取或设置显示在项目的主题字段中的说明。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-679">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="d7fe0-680">`subject` 属性获取或设置由电子邮件服务器发送项目时的整个主题。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-680">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="d7fe0-681">阅读模式</span><span class="sxs-lookup"><span data-stu-id="d7fe0-681">Read mode</span></span>

<span data-ttu-id="d7fe0-p131">`subject` 属性返回一个字符串。使用 [`normalizedSubject`](#normalizedsubject-string) 属性获取不带任何前导前缀（如 `RE:` 和 `FW:`）的主题。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-p131">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

```javascript
var subject = Office.context.mailbox.item.subject;
```

##### <a name="compose-mode"></a><span data-ttu-id="d7fe0-684">撰写模式</span><span class="sxs-lookup"><span data-stu-id="d7fe0-684">Compose mode</span></span>

<span data-ttu-id="d7fe0-685">`subject` 属性返回一个 `Subject` 对象，该对象提供用于获取和设置主题的方法。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-685">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```javascript
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="d7fe0-686">类型：</span><span class="sxs-lookup"><span data-stu-id="d7fe0-686">Type:</span></span>

*   <span data-ttu-id="d7fe0-687">String | [Subject](/javascript/api/outlook/office.subject)</span><span class="sxs-lookup"><span data-stu-id="d7fe0-687">String | [Subject](/javascript/api/outlook/office.subject)</span></span>

##### <a name="requirements"></a><span data-ttu-id="d7fe0-688">要求</span><span class="sxs-lookup"><span data-stu-id="d7fe0-688">Requirements</span></span>

|<span data-ttu-id="d7fe0-689">要求</span><span class="sxs-lookup"><span data-stu-id="d7fe0-689">Requirement</span></span>|<span data-ttu-id="d7fe0-690">值</span><span class="sxs-lookup"><span data-stu-id="d7fe0-690">Value</span></span>|
|---|---|
|[<span data-ttu-id="d7fe0-691">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="d7fe0-691">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d7fe0-692">1.0</span><span class="sxs-lookup"><span data-stu-id="d7fe0-692">1.0</span></span>|
|[<span data-ttu-id="d7fe0-693">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="d7fe0-693">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d7fe0-694">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d7fe0-694">ReadItem</span></span>|
|[<span data-ttu-id="d7fe0-695">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="d7fe0-695">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="d7fe0-696">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="d7fe0-696">Compose or read</span></span>|

####  <a name="to-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="d7fe0-697">to :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="d7fe0-697">to :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="d7fe0-698">提供对邮件的“**收件人**”行上的收件人的访问权限。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-698">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="d7fe0-699">对象的类型和访问级别取决于当前项目的模式。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-699">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="d7fe0-700">阅读模式</span><span class="sxs-lookup"><span data-stu-id="d7fe0-700">Read mode</span></span>

<span data-ttu-id="d7fe0-p133">`to` 属性返回包含邮件的**收件人**行上所列的每个收件人的 `EmailAddressDetails` 对象的数组。集合上限为 100 个成员。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-p133">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="d7fe0-703">撰写模式</span><span class="sxs-lookup"><span data-stu-id="d7fe0-703">Compose mode</span></span>

<span data-ttu-id="d7fe0-704">`to` 属性返回一个 `Recipients` 对象，该对象提供用于获取或更新邮件的“**收件人**”行上收件人的方法。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-704">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="d7fe0-705">类型：</span><span class="sxs-lookup"><span data-stu-id="d7fe0-705">Type:</span></span>

*   <span data-ttu-id="d7fe0-706">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="d7fe0-706">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="d7fe0-707">要求</span><span class="sxs-lookup"><span data-stu-id="d7fe0-707">Requirements</span></span>

|<span data-ttu-id="d7fe0-708">要求</span><span class="sxs-lookup"><span data-stu-id="d7fe0-708">Requirement</span></span>|<span data-ttu-id="d7fe0-709">值</span><span class="sxs-lookup"><span data-stu-id="d7fe0-709">Value</span></span>|
|---|---|
|[<span data-ttu-id="d7fe0-710">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="d7fe0-710">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d7fe0-711">1.0</span><span class="sxs-lookup"><span data-stu-id="d7fe0-711">1.0</span></span>|
|[<span data-ttu-id="d7fe0-712">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="d7fe0-712">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d7fe0-713">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d7fe0-713">ReadItem</span></span>|
|[<span data-ttu-id="d7fe0-714">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="d7fe0-714">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="d7fe0-715">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="d7fe0-715">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="d7fe0-716">示例</span><span class="sxs-lookup"><span data-stu-id="d7fe0-716">Example</span></span>

```javascript
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

### <a name="methods"></a><span data-ttu-id="d7fe0-717">方法</span><span class="sxs-lookup"><span data-stu-id="d7fe0-717">Methods</span></span>

####  <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="d7fe0-718">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="d7fe0-718">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="d7fe0-719">将文件作为附件添加到邮件或约会。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-719">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="d7fe0-720">`addFileAttachmentAsync` 方法在指定的 URI 上载文件并将其附加到撰写窗体中的项目。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-720">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="d7fe0-721">你随后可以将该标识符与 [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) 方法一同使用，以删除同一个会话中的附件。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-721">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d7fe0-722">参数：</span><span class="sxs-lookup"><span data-stu-id="d7fe0-722">Parameters:</span></span>
|<span data-ttu-id="d7fe0-723">名称</span><span class="sxs-lookup"><span data-stu-id="d7fe0-723">Name</span></span>|<span data-ttu-id="d7fe0-724">类型</span><span class="sxs-lookup"><span data-stu-id="d7fe0-724">Type</span></span>|<span data-ttu-id="d7fe0-725">属性</span><span class="sxs-lookup"><span data-stu-id="d7fe0-725">Attributes</span></span>|<span data-ttu-id="d7fe0-726">说明</span><span class="sxs-lookup"><span data-stu-id="d7fe0-726">Description</span></span>|
|---|---|---|---|
|`uri`|<span data-ttu-id="d7fe0-727">String</span><span class="sxs-lookup"><span data-stu-id="d7fe0-727">String</span></span>||<span data-ttu-id="d7fe0-p134">提供附加到邮件或约会的文件的位置的 URI。最大长度为 2048 个字符。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-p134">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`|<span data-ttu-id="d7fe0-730">字符串</span><span class="sxs-lookup"><span data-stu-id="d7fe0-730">String</span></span>||<span data-ttu-id="d7fe0-p135">在附件上载过程中显示的附件名称。最大长度为 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-p135">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="d7fe0-733">Object</span><span class="sxs-lookup"><span data-stu-id="d7fe0-733">Object</span></span>|<span data-ttu-id="d7fe0-734">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="d7fe0-734">&lt;optional&gt;</span></span>|<span data-ttu-id="d7fe0-735">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-735">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="d7fe0-736">对象</span><span class="sxs-lookup"><span data-stu-id="d7fe0-736">Object</span></span>|<span data-ttu-id="d7fe0-737">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="d7fe0-737">&lt;optional&gt;</span></span>|<span data-ttu-id="d7fe0-738">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-738">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.isInline`|<span data-ttu-id="d7fe0-739">布尔值</span><span class="sxs-lookup"><span data-stu-id="d7fe0-739">Boolean</span></span>|<span data-ttu-id="d7fe0-740">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="d7fe0-740">&lt;optional&gt;</span></span>|<span data-ttu-id="d7fe0-741">如果为 `true`，则表示附件将在邮件正文中内联显示，并且不应显示在附件列表中。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-741">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`callback`|<span data-ttu-id="d7fe0-742">函数</span><span class="sxs-lookup"><span data-stu-id="d7fe0-742">function</span></span>|<span data-ttu-id="d7fe0-743">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="d7fe0-743">&lt;optional&gt;</span></span>|<span data-ttu-id="d7fe0-744">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-744">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="d7fe0-745">如果成功，附件标识符将在 `asyncResult.value` 属性中提供。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-745">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="d7fe0-746">如果上传附件失败，`asyncResult` 对象将包含一个提供错误说明的 `Error` 对象。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-746">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="d7fe0-747">错误</span><span class="sxs-lookup"><span data-stu-id="d7fe0-747">Errors</span></span>

|<span data-ttu-id="d7fe0-748">错误代码</span><span class="sxs-lookup"><span data-stu-id="d7fe0-748">Error code</span></span>|<span data-ttu-id="d7fe0-749">说明</span><span class="sxs-lookup"><span data-stu-id="d7fe0-749">Description</span></span>|
|------------|-------------|
|`AttachmentSizeExceeded`|<span data-ttu-id="d7fe0-750">附件大小超过了允许的大小。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-750">The attachment is larger than allowed.</span></span>|
|`FileTypeNotSupported`|<span data-ttu-id="d7fe0-751">该附件的扩展名不是允许的扩展名。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-751">The attachment has an extension that is not allowed.</span></span>|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="d7fe0-752">邮件或约会具有的附件过多。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-752">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d7fe0-753">要求</span><span class="sxs-lookup"><span data-stu-id="d7fe0-753">Requirements</span></span>

|<span data-ttu-id="d7fe0-754">要求</span><span class="sxs-lookup"><span data-stu-id="d7fe0-754">Requirement</span></span>|<span data-ttu-id="d7fe0-755">值</span><span class="sxs-lookup"><span data-stu-id="d7fe0-755">Value</span></span>|
|---|---|
|[<span data-ttu-id="d7fe0-756">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="d7fe0-756">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d7fe0-757">1.1</span><span class="sxs-lookup"><span data-stu-id="d7fe0-757">1.1</span></span>|
|[<span data-ttu-id="d7fe0-758">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="d7fe0-758">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d7fe0-759">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="d7fe0-759">ReadWriteItem</span></span>|
|[<span data-ttu-id="d7fe0-760">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="d7fe0-760">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="d7fe0-761">撰写</span><span class="sxs-lookup"><span data-stu-id="d7fe0-761">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="d7fe0-762">示例</span><span class="sxs-lookup"><span data-stu-id="d7fe0-762">Examples</span></span>

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

<span data-ttu-id="d7fe0-763">以下示例将图像文件添加为内联附件，并在邮件正文中引用该附件。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-763">The following example adds an image file as an inline attachment and references the attachment in the message body.</span></span>

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

#### <a name="addfileattachmentfrombase64asyncbase64file-attachmentname-options-callback"></a><span data-ttu-id="d7fe0-764">addFileAttachmentFromBase64Async(base64File, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="d7fe0-764">addFileAttachmentFromBase64Async(base64File, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="d7fe0-765">将 base64 编码中的文件作为附件添加到消息或约会。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-765">Adds a file from the base64 encoding to a message or appointment as an attachment.</span></span>

<span data-ttu-id="d7fe0-766">`addFileAttachmentFromBase64Async` 方法从 base64 编码上传文件，并将其附加到撰写表单中的项目。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-766">The `addFileAttachmentFromBase64Async` method uploads the file from the base64 encoding and attaches it to the item in the compose form.</span></span> <span data-ttu-id="d7fe0-767">此方法返回 AsyncResult.value 对象中的附件标识符。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-767">This method returns the attachment identifier in the AsyncResult.value object.</span></span>

<span data-ttu-id="d7fe0-768">随后可以将该标识符与 [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) 方法一同使用，以删除同一个会话中的附件。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-768">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d7fe0-769">参数：</span><span class="sxs-lookup"><span data-stu-id="d7fe0-769">Parameters:</span></span>
|<span data-ttu-id="d7fe0-770">名称</span><span class="sxs-lookup"><span data-stu-id="d7fe0-770">Name</span></span>|<span data-ttu-id="d7fe0-771">类型</span><span class="sxs-lookup"><span data-stu-id="d7fe0-771">Type</span></span>|<span data-ttu-id="d7fe0-772">属性</span><span class="sxs-lookup"><span data-stu-id="d7fe0-772">Attributes</span></span>|<span data-ttu-id="d7fe0-773">说明</span><span class="sxs-lookup"><span data-stu-id="d7fe0-773">Description</span></span>|
|---|---|---|---|
|`base64File`|<span data-ttu-id="d7fe0-774">String</span><span class="sxs-lookup"><span data-stu-id="d7fe0-774">String</span></span>||<span data-ttu-id="d7fe0-775">要添加到电子邮件或事件的图像或文件的 base64 编码内容。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-775">The base64 encoded content of an image or file to be added to an email or event.</span></span>|
|`attachmentName`|<span data-ttu-id="d7fe0-776">字符串</span><span class="sxs-lookup"><span data-stu-id="d7fe0-776">String</span></span>||<span data-ttu-id="d7fe0-p137">在附件上载过程中显示的附件名称。最大长度为 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-p137">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="d7fe0-779">Object</span><span class="sxs-lookup"><span data-stu-id="d7fe0-779">Object</span></span>|<span data-ttu-id="d7fe0-780">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="d7fe0-780">&lt;optional&gt;</span></span>|<span data-ttu-id="d7fe0-781">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-781">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="d7fe0-782">对象</span><span class="sxs-lookup"><span data-stu-id="d7fe0-782">Object</span></span>|<span data-ttu-id="d7fe0-783">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="d7fe0-783">&lt;optional&gt;</span></span>|<span data-ttu-id="d7fe0-784">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-784">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.isInline`|<span data-ttu-id="d7fe0-785">布尔值</span><span class="sxs-lookup"><span data-stu-id="d7fe0-785">Boolean</span></span>|<span data-ttu-id="d7fe0-786">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="d7fe0-786">&lt;optional&gt;</span></span>|<span data-ttu-id="d7fe0-787">如果为 `true`，则表示附件将在邮件正文中内联显示，并且不应显示在附件列表中。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-787">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`callback`|<span data-ttu-id="d7fe0-788">函数</span><span class="sxs-lookup"><span data-stu-id="d7fe0-788">function</span></span>|<span data-ttu-id="d7fe0-789">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="d7fe0-789">&lt;optional&gt;</span></span>|<span data-ttu-id="d7fe0-790">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-790">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="d7fe0-791">如果成功，附件标识符将在 `asyncResult.value` 属性中提供。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-791">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="d7fe0-792">如果上传附件失败，`asyncResult` 对象将包含一个提供错误说明的 `Error` 对象。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-792">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="d7fe0-793">错误</span><span class="sxs-lookup"><span data-stu-id="d7fe0-793">Errors</span></span>

|<span data-ttu-id="d7fe0-794">错误代码</span><span class="sxs-lookup"><span data-stu-id="d7fe0-794">Error code</span></span>|<span data-ttu-id="d7fe0-795">说明</span><span class="sxs-lookup"><span data-stu-id="d7fe0-795">Description</span></span>|
|------------|-------------|
|`AttachmentSizeExceeded`|<span data-ttu-id="d7fe0-796">附件大小超过了允许的大小。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-796">The attachment is larger than allowed.</span></span>|
|`FileTypeNotSupported`|<span data-ttu-id="d7fe0-797">该附件的扩展名不是允许的扩展名。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-797">The attachment has an extension that is not allowed.</span></span>|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="d7fe0-798">邮件或约会具有的附件过多。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-798">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d7fe0-799">要求</span><span class="sxs-lookup"><span data-stu-id="d7fe0-799">Requirements</span></span>

|<span data-ttu-id="d7fe0-800">要求</span><span class="sxs-lookup"><span data-stu-id="d7fe0-800">Requirement</span></span>|<span data-ttu-id="d7fe0-801">值</span><span class="sxs-lookup"><span data-stu-id="d7fe0-801">Value</span></span>|
|---|---|
|[<span data-ttu-id="d7fe0-802">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="d7fe0-802">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d7fe0-803">预览</span><span class="sxs-lookup"><span data-stu-id="d7fe0-803">Preview</span></span>|
|[<span data-ttu-id="d7fe0-804">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="d7fe0-804">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d7fe0-805">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="d7fe0-805">ReadWriteItem</span></span>|
|[<span data-ttu-id="d7fe0-806">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="d7fe0-806">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="d7fe0-807">撰写</span><span class="sxs-lookup"><span data-stu-id="d7fe0-807">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="d7fe0-808">示例</span><span class="sxs-lookup"><span data-stu-id="d7fe0-808">Examples</span></span>

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

####  <a name="addhandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="d7fe0-809">addHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="d7fe0-809">addHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="d7fe0-810">添加支持事件的事件处理程序。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-810">Adds an event handler for a supported event.</span></span>

<span data-ttu-id="d7fe0-811">当前，支持的事件类型是 `Office.EventType.AttachmentsChanged`、`Office.EventType.AppointmentTimeChanged`、`Office.EventType.RecipientsChanged` 和 `Office.EventType.RecurrenceChanged`</span><span class="sxs-lookup"><span data-stu-id="d7fe0-811">Currently the supported event types are `Office.EventType.AttachmentsChanged`, `Office.EventType.AppointmentTimeChanged`, `Office.EventType.RecipientsChanged`, and `Office.EventType.RecurrenceChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d7fe0-812">参数：</span><span class="sxs-lookup"><span data-stu-id="d7fe0-812">Parameters:</span></span>

| <span data-ttu-id="d7fe0-813">名称</span><span class="sxs-lookup"><span data-stu-id="d7fe0-813">Name</span></span> | <span data-ttu-id="d7fe0-814">类型</span><span class="sxs-lookup"><span data-stu-id="d7fe0-814">Type</span></span> | <span data-ttu-id="d7fe0-815">属性</span><span class="sxs-lookup"><span data-stu-id="d7fe0-815">Attributes</span></span> | <span data-ttu-id="d7fe0-816">说明</span><span class="sxs-lookup"><span data-stu-id="d7fe0-816">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="d7fe0-817">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="d7fe0-817">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="d7fe0-818">应调用处理程序的事件。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-818">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="d7fe0-819">函数</span><span class="sxs-lookup"><span data-stu-id="d7fe0-819">Function</span></span> || <span data-ttu-id="d7fe0-p138">用于处理事件的函数。此函数必须接受一个参数，即对象文本。参数上的 `type` 属性将匹配传递给 `addHandlerAsync` 的 `eventType` 参数。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-p138">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `addHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="d7fe0-823">Object</span><span class="sxs-lookup"><span data-stu-id="d7fe0-823">Object</span></span> | <span data-ttu-id="d7fe0-824">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="d7fe0-824">&lt;optional&gt;</span></span> | <span data-ttu-id="d7fe0-825">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-825">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="d7fe0-826">对象</span><span class="sxs-lookup"><span data-stu-id="d7fe0-826">Object</span></span> | <span data-ttu-id="d7fe0-827">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="d7fe0-827">&lt;optional&gt;</span></span> | <span data-ttu-id="d7fe0-828">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-828">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="d7fe0-829">函数</span><span class="sxs-lookup"><span data-stu-id="d7fe0-829">function</span></span>| <span data-ttu-id="d7fe0-830">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="d7fe0-830">&lt;optional&gt;</span></span>|<span data-ttu-id="d7fe0-831">方法完成后，使用单个参数 `callback`（一个 [`asyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `AsyncResult` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-831">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d7fe0-832">Requirements</span><span class="sxs-lookup"><span data-stu-id="d7fe0-832">Requirements</span></span>

|<span data-ttu-id="d7fe0-833">要求</span><span class="sxs-lookup"><span data-stu-id="d7fe0-833">Requirement</span></span>| <span data-ttu-id="d7fe0-834">值</span><span class="sxs-lookup"><span data-stu-id="d7fe0-834">Value</span></span>|
|---|---|
|[<span data-ttu-id="d7fe0-835">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="d7fe0-835">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d7fe0-836">1.7</span><span class="sxs-lookup"><span data-stu-id="d7fe0-836">1.7</span></span> |
|[<span data-ttu-id="d7fe0-837">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="d7fe0-837">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d7fe0-838">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d7fe0-838">ReadItem</span></span> |
|[<span data-ttu-id="d7fe0-839">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="d7fe0-839">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="d7fe0-840">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="d7fe0-840">Compose or read</span></span> |

####  <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="d7fe0-841">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="d7fe0-841">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="d7fe0-842">将 Exchange 项目（如邮件）作为附件添加到邮件或约会。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-842">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="d7fe0-p139">`addItemAttachmentAsync` 方法将包含指定 Exchange 标识符的项目附加到撰写窗体中的项目。如果指定一个回调方法，此方法使用 `asyncResult` 参数调用，该参数包含一个附件标识符或代码，指示附加项目过程中出现的任何错误。可以使用 `options` 参数将状态信息传递给回调方法（如果需要）。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-p139">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="d7fe0-846">随后可以将该标识符与 [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) 方法一同使用，以删除同一个会话中的附件。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-846">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="d7fe0-847">如果 Office 加载项在 Outlook Web App 中运行，则 `addItemAttachmentAsync` 方法可以将项目附加到项目（正在编辑的项目除外）中；然而，不支持也不建议这样做。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-847">If your Office Add-in is running in Outlook Web App, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d7fe0-848">参数：</span><span class="sxs-lookup"><span data-stu-id="d7fe0-848">Parameters:</span></span>

|<span data-ttu-id="d7fe0-849">名称</span><span class="sxs-lookup"><span data-stu-id="d7fe0-849">Name</span></span>|<span data-ttu-id="d7fe0-850">类型</span><span class="sxs-lookup"><span data-stu-id="d7fe0-850">Type</span></span>|<span data-ttu-id="d7fe0-851">属性</span><span class="sxs-lookup"><span data-stu-id="d7fe0-851">Attributes</span></span>|<span data-ttu-id="d7fe0-852">说明</span><span class="sxs-lookup"><span data-stu-id="d7fe0-852">Description</span></span>|
|---|---|---|---|
|`itemId`|<span data-ttu-id="d7fe0-853">String</span><span class="sxs-lookup"><span data-stu-id="d7fe0-853">String</span></span>||<span data-ttu-id="d7fe0-p140">要附加的项目的 Exchange 标识符。最大长度为 100 个字符。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-p140">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`|<span data-ttu-id="d7fe0-856">String</span><span class="sxs-lookup"><span data-stu-id="d7fe0-856">String</span></span>||<span data-ttu-id="d7fe0-p141">要附加的项目的主题。最大长度为 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-p141">The sujbect of the item to be attached. The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="d7fe0-859">对象</span><span class="sxs-lookup"><span data-stu-id="d7fe0-859">Object</span></span>|<span data-ttu-id="d7fe0-860">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="d7fe0-860">&lt;optional&gt;</span></span>|<span data-ttu-id="d7fe0-861">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-861">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="d7fe0-862">对象</span><span class="sxs-lookup"><span data-stu-id="d7fe0-862">Object</span></span>|<span data-ttu-id="d7fe0-863">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="d7fe0-863">&lt;optional&gt;</span></span>|<span data-ttu-id="d7fe0-864">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-864">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="d7fe0-865">函数</span><span class="sxs-lookup"><span data-stu-id="d7fe0-865">function</span></span>|<span data-ttu-id="d7fe0-866">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="d7fe0-866">&lt;optional&gt;</span></span>|<span data-ttu-id="d7fe0-867">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-867">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="d7fe0-868">如果成功，附件标识符将在 `asyncResult.value` 属性中提供。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-868">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="d7fe0-869">如果添加附件失败，`asyncResult` 对象将包含一个提供错误说明的 `Error` 对象。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-869">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="d7fe0-870">错误</span><span class="sxs-lookup"><span data-stu-id="d7fe0-870">Errors</span></span>

|<span data-ttu-id="d7fe0-871">错误代码</span><span class="sxs-lookup"><span data-stu-id="d7fe0-871">Error code</span></span>|<span data-ttu-id="d7fe0-872">说明</span><span class="sxs-lookup"><span data-stu-id="d7fe0-872">Description</span></span>|
|------------|-------------|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="d7fe0-873">邮件或约会具有的附件过多。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-873">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d7fe0-874">Requirements</span><span class="sxs-lookup"><span data-stu-id="d7fe0-874">Requirements</span></span>

|<span data-ttu-id="d7fe0-875">要求</span><span class="sxs-lookup"><span data-stu-id="d7fe0-875">Requirement</span></span>|<span data-ttu-id="d7fe0-876">值</span><span class="sxs-lookup"><span data-stu-id="d7fe0-876">Value</span></span>|
|---|---|
|[<span data-ttu-id="d7fe0-877">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="d7fe0-877">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d7fe0-878">1.1</span><span class="sxs-lookup"><span data-stu-id="d7fe0-878">1.1</span></span>|
|[<span data-ttu-id="d7fe0-879">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="d7fe0-879">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d7fe0-880">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="d7fe0-880">ReadWriteItem</span></span>|
|[<span data-ttu-id="d7fe0-881">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="d7fe0-881">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="d7fe0-882">撰写</span><span class="sxs-lookup"><span data-stu-id="d7fe0-882">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="d7fe0-883">示例</span><span class="sxs-lookup"><span data-stu-id="d7fe0-883">Example</span></span>

<span data-ttu-id="d7fe0-884">以下示例将现有的 Outlook 项目添加为名为 `My Attachment` 的附件。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-884">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

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

####  <a name="close"></a><span data-ttu-id="d7fe0-885">close()</span><span class="sxs-lookup"><span data-stu-id="d7fe0-885">close()</span></span>

<span data-ttu-id="d7fe0-886">关闭当前正在撰写的项目。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-886">Closes the current item that is being composed.</span></span>

<span data-ttu-id="d7fe0-p142">
            \`close\` 方法的行为取决于要撰写的项目的当前状态。如果项目具有未保存的更改，客户端将提示用户保存、放弃或取消关闭操作。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-p142">The behavior of the `close` method depends on the current state of the item being composed. If the item has unsaved changes, the client prompts the user to save, discard, or cancel the close action.</span></span>

> [!NOTE]
> <span data-ttu-id="d7fe0-889">在 Outlook 网页版中，如果该项目是约会并且之前已使用 `saveAsync` 保存，则即使自上次保存项目后未发生任何更改，也会提示用户保存、放弃或取消。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-889">In Outlook on the web, if the item is an appointment and it has previously been saved using `saveAsync`, the user is prompted to save, discard, or cancel even if no changes have occurred since the item was last saved.</span></span>

<span data-ttu-id="d7fe0-890">在 Outlook 桌面客户端中，如果邮件是内联答复，`close` 方法不起作用。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-890">In the Outlook desktop client, if the message is an inline reply, the `close` method has no effect.</span></span>

##### <a name="requirements"></a><span data-ttu-id="d7fe0-891">要求</span><span class="sxs-lookup"><span data-stu-id="d7fe0-891">Requirements</span></span>

|<span data-ttu-id="d7fe0-892">要求</span><span class="sxs-lookup"><span data-stu-id="d7fe0-892">Requirement</span></span>|<span data-ttu-id="d7fe0-893">值</span><span class="sxs-lookup"><span data-stu-id="d7fe0-893">Value</span></span>|
|---|---|
|[<span data-ttu-id="d7fe0-894">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="d7fe0-894">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d7fe0-895">1.3</span><span class="sxs-lookup"><span data-stu-id="d7fe0-895">1.3</span></span>|
|[<span data-ttu-id="d7fe0-896">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="d7fe0-896">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d7fe0-897">受限</span><span class="sxs-lookup"><span data-stu-id="d7fe0-897">Restricted</span></span>|
|[<span data-ttu-id="d7fe0-898">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="d7fe0-898">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="d7fe0-899">撰写</span><span class="sxs-lookup"><span data-stu-id="d7fe0-899">Compose</span></span>|

#### <a name="displayreplyallformformdata"></a><span data-ttu-id="d7fe0-900">displayReplyAllForm(formData)</span><span class="sxs-lookup"><span data-stu-id="d7fe0-900">displayReplyAllForm(formData)</span></span>

<span data-ttu-id="d7fe0-901">显示答复窗体，其中包括所选邮件的发件人和所有收件人或所选约会的组织者和所有与会者。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-901">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="d7fe0-902">在 Outlook for iOS 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-902">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="d7fe0-903">在 Outlook Web App 中，答复窗体显示为包含 3 列视图的弹出式窗体以及包含 2 列或 1 列视图的弹出式窗体。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-903">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="d7fe0-904">如果任意字符串参数超出其限制，`displayReplyAllForm` 将引发异常。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-904">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="d7fe0-p143">当在 `formData.attachments` 参数中指定附件时，Outlook 和 Outlook Web App 尝试下载所有附件并将其附加到答复窗体。如果无法添加任何附件，则在窗体 UI 中显示错误。如果这不可能，则不引发错误消息。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-p143">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d7fe0-908">参数：</span><span class="sxs-lookup"><span data-stu-id="d7fe0-908">Parameters:</span></span>

|<span data-ttu-id="d7fe0-909">名称</span><span class="sxs-lookup"><span data-stu-id="d7fe0-909">Name</span></span>|<span data-ttu-id="d7fe0-910">类型</span><span class="sxs-lookup"><span data-stu-id="d7fe0-910">Type</span></span>|<span data-ttu-id="d7fe0-911">属性</span><span class="sxs-lookup"><span data-stu-id="d7fe0-911">Attributes</span></span>|<span data-ttu-id="d7fe0-912">说明</span><span class="sxs-lookup"><span data-stu-id="d7fe0-912">Description</span></span>|
|---|---|---|---|
|`formData`|<span data-ttu-id="d7fe0-913">字符串 &#124; 对象</span><span class="sxs-lookup"><span data-stu-id="d7fe0-913">String &#124; Object</span></span>||<span data-ttu-id="d7fe0-p144">一个包含文本和 HTML 且表示答复窗体的正文的字符串。字符串限制为 32 KB。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-p144">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="d7fe0-916">**或**</span><span class="sxs-lookup"><span data-stu-id="d7fe0-916">**OR**</span></span><br/><span data-ttu-id="d7fe0-p145">包含正文或附件数据和回调函数的对象。对象定义如下。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-p145">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span>|
|`formData.htmlBody`|<span data-ttu-id="d7fe0-919">字符串</span><span class="sxs-lookup"><span data-stu-id="d7fe0-919">String</span></span>|<span data-ttu-id="d7fe0-920">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="d7fe0-920">&lt;optional&gt;</span></span>|<span data-ttu-id="d7fe0-p146">一个包含文本和 HTML 且表示答复窗体的正文的字符串。字符串限制为 32 KB。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-p146">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
|`formData.attachments`|<span data-ttu-id="d7fe0-923">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="d7fe0-923">Array.&lt;Object&gt;</span></span>|<span data-ttu-id="d7fe0-924">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="d7fe0-924">&lt;optional&gt;</span></span>|<span data-ttu-id="d7fe0-925">JSON 对象（是文件或项目附件）的数组。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-925">An array of JSON objects that are either file or item attachments.</span></span>|
|`formData.attachments.type`|<span data-ttu-id="d7fe0-926">String</span><span class="sxs-lookup"><span data-stu-id="d7fe0-926">String</span></span>||<span data-ttu-id="d7fe0-p147">指示附件的类型。必须是文件附件的 `file` 或项目附件的 `item`。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-p147">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span>|
|`formData.attachments.name`|<span data-ttu-id="d7fe0-929">字符串</span><span class="sxs-lookup"><span data-stu-id="d7fe0-929">String</span></span>||<span data-ttu-id="d7fe0-930">一个包含附件的名称的字符串，最多包含 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-930">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
|`formData.attachments.url`|<span data-ttu-id="d7fe0-931">String</span><span class="sxs-lookup"><span data-stu-id="d7fe0-931">String</span></span>||<span data-ttu-id="d7fe0-p148">仅在将 `type` 设置为 `file` 时使用。文件的位置的 URI。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-p148">Only used if `type` is set to `file`. The URI of the location for the file.</span></span>|
|`formData.attachments.isInline`|<span data-ttu-id="d7fe0-934">Boolean</span><span class="sxs-lookup"><span data-stu-id="d7fe0-934">Boolean</span></span>||<span data-ttu-id="d7fe0-p149">仅在将 `type` 设置为 `file` 时使用。如果为 `true`，则表示附件将在邮件正文中内联显示，并且不应显示在附件列表中。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-p149">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`formData.attachments.itemId`|<span data-ttu-id="d7fe0-937">String</span><span class="sxs-lookup"><span data-stu-id="d7fe0-937">String</span></span>||<span data-ttu-id="d7fe0-p150">仅在将 `type` 设置为 `item` 时使用。附件的 EWS 项目 ID。字符串最长为 100 个字符。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-p150">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span>|
|`callback`|<span data-ttu-id="d7fe0-941">函数</span><span class="sxs-lookup"><span data-stu-id="d7fe0-941">function</span></span>|<span data-ttu-id="d7fe0-942">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="d7fe0-942">&lt;optional&gt;</span></span>|<span data-ttu-id="d7fe0-943">方法完成后，使用单个参数 `asyncResult`（一个 [AsyncResult](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-943">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d7fe0-944">要求</span><span class="sxs-lookup"><span data-stu-id="d7fe0-944">Requirements</span></span>

|<span data-ttu-id="d7fe0-945">要求</span><span class="sxs-lookup"><span data-stu-id="d7fe0-945">Requirement</span></span>|<span data-ttu-id="d7fe0-946">值</span><span class="sxs-lookup"><span data-stu-id="d7fe0-946">Value</span></span>|
|---|---|
|[<span data-ttu-id="d7fe0-947">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="d7fe0-947">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d7fe0-948">1.0</span><span class="sxs-lookup"><span data-stu-id="d7fe0-948">1.0</span></span>|
|[<span data-ttu-id="d7fe0-949">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="d7fe0-949">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d7fe0-950">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d7fe0-950">ReadItem</span></span>|
|[<span data-ttu-id="d7fe0-951">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="d7fe0-951">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="d7fe0-952">阅读</span><span class="sxs-lookup"><span data-stu-id="d7fe0-952">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="d7fe0-953">示例</span><span class="sxs-lookup"><span data-stu-id="d7fe0-953">Examples</span></span>

<span data-ttu-id="d7fe0-954">以下代码将一个字符串传递到 `displayReplyAllForm` 函数。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-954">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="d7fe0-955">使用空白正文答复。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-955">Reply with an empty body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="d7fe0-956">仅使用正文答复。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-956">Reply with just a body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="d7fe0-957">使用正文和文件附件答复。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-957">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="d7fe0-958">使用正文和项目附件答复。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-958">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="d7fe0-959">使用正文、文件附件、项目附件和回调答复。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-959">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="displayreplyformformdata"></a><span data-ttu-id="d7fe0-960">displayReplyForm(formData)</span><span class="sxs-lookup"><span data-stu-id="d7fe0-960">displayReplyForm(formData)</span></span>

<span data-ttu-id="d7fe0-961">显示答复窗体，其中仅包括所选邮件的发件人或所选约会的组织者。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-961">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="d7fe0-962">在 Outlook for iOS 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-962">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="d7fe0-963">在 Outlook Web App 中，答复窗体显示为包含 3 列视图的弹出式窗体以及包含 2 列或 1 列视图的弹出式窗体。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-963">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="d7fe0-964">如果任意字符串参数超出其限制，`displayReplyForm` 将引发异常。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-964">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="d7fe0-p151">当在 `formData.attachments` 参数中指定附件时，Outlook 和 Outlook Web App 尝试下载所有附件并将其附加到答复窗体。如果无法添加任何附件，则在窗体 UI 中显示错误。如果这不可能，则不引发错误消息。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-p151">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d7fe0-968">参数：</span><span class="sxs-lookup"><span data-stu-id="d7fe0-968">Parameters:</span></span>

|<span data-ttu-id="d7fe0-969">名称</span><span class="sxs-lookup"><span data-stu-id="d7fe0-969">Name</span></span>|<span data-ttu-id="d7fe0-970">类型</span><span class="sxs-lookup"><span data-stu-id="d7fe0-970">Type</span></span>|<span data-ttu-id="d7fe0-971">属性</span><span class="sxs-lookup"><span data-stu-id="d7fe0-971">Attributes</span></span>|<span data-ttu-id="d7fe0-972">说明</span><span class="sxs-lookup"><span data-stu-id="d7fe0-972">Description</span></span>|
|---|---|---|---|
|`formData`|<span data-ttu-id="d7fe0-973">字符串 &#124; 对象</span><span class="sxs-lookup"><span data-stu-id="d7fe0-973">String &#124; Object</span></span>||<span data-ttu-id="d7fe0-p152">一个包含文本和 HTML 且表示答复窗体的正文的字符串。字符串限制为 32 KB。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-p152">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="d7fe0-976">**或**</span><span class="sxs-lookup"><span data-stu-id="d7fe0-976">**OR**</span></span><br/><span data-ttu-id="d7fe0-p153">包含正文或附件数据和回调函数的对象。对象定义如下。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-p153">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span>|
|`formData.htmlBody`|<span data-ttu-id="d7fe0-979">String</span><span class="sxs-lookup"><span data-stu-id="d7fe0-979">String</span></span>|<span data-ttu-id="d7fe0-980">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="d7fe0-980">&lt;optional&gt;</span></span>|<span data-ttu-id="d7fe0-p154">一个包含文本和 HTML 且表示答复窗体的正文的字符串。字符串限制为 32 KB。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-p154">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
|`formData.attachments`|<span data-ttu-id="d7fe0-983">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="d7fe0-983">Array.&lt;Object&gt;</span></span>|<span data-ttu-id="d7fe0-984">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="d7fe0-984">&lt;optional&gt;</span></span>|<span data-ttu-id="d7fe0-985">JSON 对象（是文件或项目附件）的数组。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-985">An array of JSON objects that are either file or item attachments.</span></span>|
|`formData.attachments.type`|<span data-ttu-id="d7fe0-986">String</span><span class="sxs-lookup"><span data-stu-id="d7fe0-986">String</span></span>||<span data-ttu-id="d7fe0-p155">指示附件的类型。必须是文件附件的 `file` 或项目附件的 `item`。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-p155">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span>|
|`formData.attachments.name`|<span data-ttu-id="d7fe0-989">字符串</span><span class="sxs-lookup"><span data-stu-id="d7fe0-989">String</span></span>||<span data-ttu-id="d7fe0-990">一个包含附件的名称的字符串，最多包含 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-990">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
|`formData.attachments.url`|<span data-ttu-id="d7fe0-991">String</span><span class="sxs-lookup"><span data-stu-id="d7fe0-991">String</span></span>||<span data-ttu-id="d7fe0-p156">仅在将 `type` 设置为 `file` 时使用。文件的位置的 URI。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-p156">Only used if `type` is set to `file`. The URI of the location for the file.</span></span>|
|`formData.attachments.isInline`|<span data-ttu-id="d7fe0-994">Boolean</span><span class="sxs-lookup"><span data-stu-id="d7fe0-994">Boolean</span></span>||<span data-ttu-id="d7fe0-p157">仅在将 `type` 设置为 `file` 时使用。如果为 `true`，则表示附件将在邮件正文中内联显示，并且不应显示在附件列表中。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-p157">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`formData.attachments.itemId`|<span data-ttu-id="d7fe0-997">String</span><span class="sxs-lookup"><span data-stu-id="d7fe0-997">String</span></span>||<span data-ttu-id="d7fe0-p158">仅在将 `type` 设置为 `item` 时使用。附件的 EWS 项目 ID。字符串最长为 100 个字符。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-p158">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span>|
|`callback`|<span data-ttu-id="d7fe0-1001">函数</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1001">function</span></span>|<span data-ttu-id="d7fe0-1002">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1002">&lt;optional&gt;</span></span>|<span data-ttu-id="d7fe0-1003">方法完成后，使用单个参数 `asyncResult`（一个 [AsyncResult](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1003">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d7fe0-1004">要求</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1004">Requirements</span></span>

|<span data-ttu-id="d7fe0-1005">要求</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1005">Requirement</span></span>|<span data-ttu-id="d7fe0-1006">值</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1006">Value</span></span>|
|---|---|
|[<span data-ttu-id="d7fe0-1007">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1007">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d7fe0-1008">1.0</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1008">1.0</span></span>|
|[<span data-ttu-id="d7fe0-1009">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1009">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d7fe0-1010">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1010">ReadItem</span></span>|
|[<span data-ttu-id="d7fe0-1011">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1011">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="d7fe0-1012">阅读</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1012">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="d7fe0-1013">示例</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1013">Examples</span></span>

<span data-ttu-id="d7fe0-1014">以下代码将一个字符串传递到 `displayReplyForm` 函数。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1014">The following code passes a string to the `displayReplyForm` function.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="d7fe0-1015">使用空白正文答复。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1015">Reply with an empty body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="d7fe0-1016">仅使用正文答复。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1016">Reply with just a body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="d7fe0-1017">使用正文和文件附件答复。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1017">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="d7fe0-1018">使用正文和项目附件答复。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1018">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="d7fe0-1019">使用正文、文件附件、项目附件和回调答复。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1019">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="getattachmentcontentasyncattachmentid-options-callback--attachmentcontentjavascriptapioutlookofficeattachmentcontent"></a><span data-ttu-id="d7fe0-1020">getAttachmentContentAsync(attachmentId, [options], callback) → [AttachmentContent](/javascript/api/outlook/office.attachmentcontent)</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1020">getAttachmentContentAsync(attachmentId, [options], callback) → [AttachmentContent](/javascript/api/outlook/office.attachmentcontent)</span></span>

<span data-ttu-id="d7fe0-1021">从消息或约会中获取指定的附件，并将其作为 `AttachmentContent` 对象返回。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1021">Gets the specified attachment from a message or appointment and returns it as an `AttachmentContent` object.</span></span>

<span data-ttu-id="d7fe0-1022">`getAttachmentContentAsync` 方法获取项目中带指定标识符的附件。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1022">The `getAttachmentContentAsync` method gets the attachment with the specified identifier from the item.</span></span> <span data-ttu-id="d7fe0-1023">作为最佳做法，应使用标识符检索同一会话中的附件，在该会话中，使用 `getAttachmentsAsync` 或 `item.attachments` 调用检索附件 ID。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1023">As a best practice, you should use the identifier to retrieve an attachment in the same session that the attachmentIds were retrieved with the `getAttachmentsAsync` or `item.attachments` call.</span></span> <span data-ttu-id="d7fe0-1024">在 Outlook Web App 和适用于设备的 OWA 中，附件标识符只在同一个会话中才有效。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1024">In Outlook Web App and OWA for Devices, the attachment identifier is valid only within the same session.</span></span> <span data-ttu-id="d7fe0-1025">当用户关闭应用，或者如果用户开始在内嵌窗体中撰写，则随后弹出的窗体中继续在单独的窗口撰写时，会话即结束。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1025">A session is over when the user closes the app, or if the user starts composing an inline form then subsequently pops out the form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d7fe0-1026">参数：</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1026">Parameters:</span></span>

|<span data-ttu-id="d7fe0-1027">名称</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1027">Name</span></span>|<span data-ttu-id="d7fe0-1028">类型</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1028">Type</span></span>|<span data-ttu-id="d7fe0-1029">属性</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1029">Attributes</span></span>|<span data-ttu-id="d7fe0-1030">说明</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1030">Description</span></span>|
|---|---|---|---|
|`attachmentId`|<span data-ttu-id="d7fe0-1031">字符串</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1031">String</span></span>||<span data-ttu-id="d7fe0-1032">要获取的附件的标识符。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1032">The identifier of the attachment you want to get.</span></span>|
|`options`|<span data-ttu-id="d7fe0-1033">对象</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1033">Object</span></span>|<span data-ttu-id="d7fe0-1034">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1034">&lt;optional&gt;</span></span>|<span data-ttu-id="d7fe0-1035">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1035">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="d7fe0-1036">对象</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1036">Object</span></span>|<span data-ttu-id="d7fe0-1037">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1037">&lt;optional&gt;</span></span>|<span data-ttu-id="d7fe0-1038">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1038">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="d7fe0-1039">function</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1039">function</span></span>|<span data-ttu-id="d7fe0-1040">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1040">&lt;optional&gt;</span></span>|<span data-ttu-id="d7fe0-1041">方法完成后，使用单个参数 `asyncResult`（一个 [AsyncResult](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1041">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d7fe0-1042">要求</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1042">Requirements</span></span>

|<span data-ttu-id="d7fe0-1043">要求</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1043">Requirement</span></span>|<span data-ttu-id="d7fe0-1044">值</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1044">Value</span></span>|
|---|---|
|[<span data-ttu-id="d7fe0-1045">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1045">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d7fe0-1046">预览</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1046">Preview</span></span>|
|[<span data-ttu-id="d7fe0-1047">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1047">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d7fe0-1048">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1048">ReadItem</span></span>|
|[<span data-ttu-id="d7fe0-1049">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1049">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="d7fe0-1050">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1050">Compose or read</span></span>|

##### <a name="returns"></a><span data-ttu-id="d7fe0-1051">返回：</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1051">Returns:</span></span>

<span data-ttu-id="d7fe0-1052">类型：[AttachmentContent](/javascript/api/outlook/office.attachmentcontent)</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1052">Type: [AttachmentContent](/javascript/api/outlook/office.attachmentcontent)</span></span>

##### <a name="example"></a><span data-ttu-id="d7fe0-1053">示例</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1053">Example</span></span>

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

#### <a name="getattachmentsasyncoptions-callback--arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetails"></a><span data-ttu-id="d7fe0-1054">getAttachmentsAsync([options], callback) → Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="d7fe0-1054">getAttachmentsAsync([options], callback) → Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span></span>

<span data-ttu-id="d7fe0-1055">获取项目的附件作为数组。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1055">Gets the item's attachments as an array.</span></span> <span data-ttu-id="d7fe0-1056">仅限撰写模式。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1056">Compose mode only.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d7fe0-1057">参数：</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1057">Parameters:</span></span>

|<span data-ttu-id="d7fe0-1058">名称</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1058">Name</span></span>|<span data-ttu-id="d7fe0-1059">类型</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1059">Type</span></span>|<span data-ttu-id="d7fe0-1060">属性</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1060">Attributes</span></span>|<span data-ttu-id="d7fe0-1061">说明</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1061">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="d7fe0-1062">Object</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1062">Object</span></span>|<span data-ttu-id="d7fe0-1063">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1063">&lt;optional&gt;</span></span>|<span data-ttu-id="d7fe0-1064">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1064">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="d7fe0-1065">对象</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1065">Object</span></span>|<span data-ttu-id="d7fe0-1066">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1066">&lt;optional&gt;</span></span>|<span data-ttu-id="d7fe0-1067">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1067">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="d7fe0-1068">函数</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1068">function</span></span>|<span data-ttu-id="d7fe0-1069">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1069">&lt;optional&gt;</span></span>|<span data-ttu-id="d7fe0-1070">方法完成后，使用单个参数 `asyncResult`（一个 [AsyncResult](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1070">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d7fe0-1071">Requirements</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1071">Requirements</span></span>

|<span data-ttu-id="d7fe0-1072">要求</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1072">Requirement</span></span>|<span data-ttu-id="d7fe0-1073">值</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1073">Value</span></span>|
|---|---|
|[<span data-ttu-id="d7fe0-1074">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1074">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d7fe0-1075">预览</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1075">Preview</span></span>|
|[<span data-ttu-id="d7fe0-1076">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1076">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d7fe0-1077">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1077">ReadItem</span></span>|
|[<span data-ttu-id="d7fe0-1078">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1078">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="d7fe0-1079">撰写</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1079">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="d7fe0-1080">返回：</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1080">Returns:</span></span>

<span data-ttu-id="d7fe0-1081">类型：Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="d7fe0-1081">Type: Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span></span>

##### <a name="example"></a><span data-ttu-id="d7fe0-1082">示例</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1082">Example</span></span>

<span data-ttu-id="d7fe0-1083">以下示例使用当前项目上所有附件的详细信息构成 HTML 字符串。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1083">The following example builds an HTML string with details of all attachments on the current item.</span></span>

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

#### <a name="getentities--entitiesjavascriptapioutlookofficeentities"></a><span data-ttu-id="d7fe0-1084">getEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1084">getEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span></span>

<span data-ttu-id="d7fe0-1085">获取在所选项目的正文中找到的实体。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1085">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="d7fe0-1086">在 Outlook for iOS 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1086">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="d7fe0-1087">要求</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1087">Requirements</span></span>

|<span data-ttu-id="d7fe0-1088">要求</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1088">Requirement</span></span>|<span data-ttu-id="d7fe0-1089">值</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1089">Value</span></span>|
|---|---|
|[<span data-ttu-id="d7fe0-1090">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1090">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d7fe0-1091">1.0</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1091">1.0</span></span>|
|[<span data-ttu-id="d7fe0-1092">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1092">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d7fe0-1093">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1093">ReadItem</span></span>|
|[<span data-ttu-id="d7fe0-1094">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1094">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="d7fe0-1095">阅读</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1095">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="d7fe0-1096">返回：</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1096">Returns:</span></span>

<span data-ttu-id="d7fe0-1097">类型：[Entities](/javascript/api/outlook/office.entities)</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1097">Type: [Entities](/javascript/api/outlook/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="d7fe0-1098">示例</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1098">Example</span></span>

<span data-ttu-id="d7fe0-1099">以下示例访问当前项目的正文中的联系人实体。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1099">The following example accesses the contacts entities in the current item's body.</span></span>

```javascript
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlookofficecontactmeetingsuggestionjavascriptapioutlookofficemeetingsuggestionphonenumberjavascriptapioutlookofficephonenumbertasksuggestionjavascriptapioutlookofficetasksuggestion"></a><span data-ttu-id="d7fe0-1100">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1100">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span></span>

<span data-ttu-id="d7fe0-1101">获取所选项目的正文中找到的指定实体类型的所有实体的数组。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1101">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="d7fe0-1102">在 Outlook for iOS 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1102">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d7fe0-1103">参数：</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1103">Parameters:</span></span>

|<span data-ttu-id="d7fe0-1104">名称</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1104">Name</span></span>|<span data-ttu-id="d7fe0-1105">说明</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1105">Type</span></span>|<span data-ttu-id="d7fe0-1106">说明</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1106">Description</span></span>|
|---|---|---|
|`entityType`|[<span data-ttu-id="d7fe0-1107">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1107">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook/office.mailboxenums.entitytype)|<span data-ttu-id="d7fe0-1108">EntityType 枚举值之一。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1108">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d7fe0-1109">Requirements</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1109">Requirements</span></span>

|<span data-ttu-id="d7fe0-1110">要求</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1110">Requirement</span></span>|<span data-ttu-id="d7fe0-1111">值</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1111">Value</span></span>|
|---|---|
|[<span data-ttu-id="d7fe0-1112">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1112">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d7fe0-1113">1.0</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1113">1.0</span></span>|
|[<span data-ttu-id="d7fe0-1114">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1114">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d7fe0-1115">受限</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1115">Restricted</span></span>|
|[<span data-ttu-id="d7fe0-1116">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1116">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="d7fe0-1117">阅读</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1117">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="d7fe0-1118">返回：</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1118">Returns:</span></span>

<span data-ttu-id="d7fe0-1119">如果在 `entityType` 中传递的值不是 `EntityType` 枚举的有效成员，该方法返回 null。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1119">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="d7fe0-1120">如果指定类型的任何实体都不存在于该项目的正文中，该方法将返回空数组。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1120">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="d7fe0-1121">否则，返回的数组中对象的类型取决于 `entityType` 参数中请求实体的类型。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1121">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="d7fe0-1122">当使用此方法的最低权限级别**受限**时，某些实体类型需要 **ReadItem** 才能进行访问，如下表中所指定。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1122">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

|<span data-ttu-id="d7fe0-1123">`entityType` 的值</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1123">Value of `entityType`</span></span>|<span data-ttu-id="d7fe0-1124">返回的数组中对象的类型</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1124">Type of objects in returned array</span></span>|<span data-ttu-id="d7fe0-1125">所需权限级别</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1125">Required Permission Level</span></span>|
|---|---|---|
|`Address`|<span data-ttu-id="d7fe0-1126">字符串</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1126">String</span></span>|<span data-ttu-id="d7fe0-1127">**受限**</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1127">**Restricted**</span></span>|
|`Contact`|<span data-ttu-id="d7fe0-1128">Contact</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1128">Contact</span></span>|<span data-ttu-id="d7fe0-1129">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1129">**ReadItem**</span></span>|
|`EmailAddress`|<span data-ttu-id="d7fe0-1130">String</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1130">String</span></span>|<span data-ttu-id="d7fe0-1131">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1131">**ReadItem**</span></span>|
|`MeetingSuggestion`|<span data-ttu-id="d7fe0-1132">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1132">MeetingSuggestion</span></span>|<span data-ttu-id="d7fe0-1133">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1133">**ReadItem**</span></span>|
|`PhoneNumber`|<span data-ttu-id="d7fe0-1134">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1134">PhoneNumber</span></span>|<span data-ttu-id="d7fe0-1135">**受限**</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1135">**Restricted**</span></span>|
|`TaskSuggestion`|<span data-ttu-id="d7fe0-1136">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1136">TaskSuggestion</span></span>|<span data-ttu-id="d7fe0-1137">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1137">**ReadItem**</span></span>|
|`URL`|<span data-ttu-id="d7fe0-1138">String</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1138">String</span></span>|<span data-ttu-id="d7fe0-1139">**受限**</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1139">**Restricted**</span></span>|

<span data-ttu-id="d7fe0-1140">类型：Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="d7fe0-1140">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span></span>

##### <a name="example"></a><span data-ttu-id="d7fe0-1141">示例</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1141">Example</span></span>

<span data-ttu-id="d7fe0-1142">以下示例显示了如何访问表示当前项目的正文中的邮政地址的字符串数组。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1142">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlookofficecontactmeetingsuggestionjavascriptapioutlookofficemeetingsuggestionphonenumberjavascriptapioutlookofficephonenumbertasksuggestionjavascriptapioutlookofficetasksuggestion"></a><span data-ttu-id="d7fe0-1143">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1143">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span></span>

<span data-ttu-id="d7fe0-1144">返回传递清单 XML 文件中定义的命名筛选器的所选项目中的已知实体。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1144">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="d7fe0-1145">在 Outlook for iOS 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1145">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="d7fe0-1146">`getFilteredEntitiesByName` 方法返回匹配在具有指定 `FilterName` 元素值的清单 XML 文件中的 [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) 规则元素中定义的正则表达式的实体。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1146">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d7fe0-1147">参数：</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1147">Parameters:</span></span>

|<span data-ttu-id="d7fe0-1148">名称</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1148">Name</span></span>|<span data-ttu-id="d7fe0-1149">说明</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1149">Type</span></span>|<span data-ttu-id="d7fe0-1150">说明</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1150">Description</span></span>|
|---|---|---|
|`name`|<span data-ttu-id="d7fe0-1151">字符串</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1151">String</span></span>|<span data-ttu-id="d7fe0-1152">定义筛选器匹配的 `ItemHasKnownEntity` 规则元素的名称。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1152">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d7fe0-1153">Requirements</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1153">Requirements</span></span>

|<span data-ttu-id="d7fe0-1154">要求</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1154">Requirement</span></span>|<span data-ttu-id="d7fe0-1155">值</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1155">Value</span></span>|
|---|---|
|[<span data-ttu-id="d7fe0-1156">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1156">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d7fe0-1157">1.0</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1157">1.0</span></span>|
|[<span data-ttu-id="d7fe0-1158">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1158">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d7fe0-1159">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1159">ReadItem</span></span>|
|[<span data-ttu-id="d7fe0-1160">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1160">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="d7fe0-1161">阅读</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1161">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="d7fe0-1162">返回：</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1162">Returns:</span></span>

<span data-ttu-id="d7fe0-p162">如果具有匹配 `name` 参数的 `FilterName` 元素值的清单中没有任何 `ItemHasKnownEntity` 元素，则该方法返回 `null`。如果 `name` 参数匹配清单中的 `ItemHasKnownEntity` 元素，但在匹配的当前项目中没有实体，则该方法返回一个空数组。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-p162">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="d7fe0-1165">类型：Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="d7fe0-1165">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span></span>

#### <a name="getinitializationcontextasyncoptions-callback"></a><span data-ttu-id="d7fe0-1166">getInitializationContextAsync([options], [callback])</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1166">getInitializationContextAsync([options], [callback])</span></span>

<span data-ttu-id="d7fe0-1167">当[可操作邮件激活](https://docs.microsoft.com/outlook/actionable-messages/invoke-add-in-from-actionable-message)加载项时，获取传递的初始化数据。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1167">Gets initialization data passed when the add-in is [activated by an actionable message](https://docs.microsoft.com/outlook/actionable-messages/invoke-add-in-from-actionable-message).</span></span>

> [!NOTE]
> <span data-ttu-id="d7fe0-1168">仅 Outlook 2016 for Windows 或更高版本（高于 16.0.8413.1000 的即点即用版本）和适用于 Office 365 的 Outlook 网页版支持此方法。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1168">This method is only supported by Outlook 2016 or later for Windows (Click-to-Run versions later than 16.0.8413.1000) and Outlook on the web for Office 365.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d7fe0-1169">参数：</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1169">Parameters:</span></span>
|<span data-ttu-id="d7fe0-1170">名称</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1170">Name</span></span>|<span data-ttu-id="d7fe0-1171">类型</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1171">Type</span></span>|<span data-ttu-id="d7fe0-1172">属性</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1172">Attributes</span></span>|<span data-ttu-id="d7fe0-1173">说明</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1173">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="d7fe0-1174">Object</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1174">Object</span></span>|<span data-ttu-id="d7fe0-1175">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1175">&lt;optional&gt;</span></span>|<span data-ttu-id="d7fe0-1176">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1176">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="d7fe0-1177">对象</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1177">Object</span></span>|<span data-ttu-id="d7fe0-1178">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1178">&lt;optional&gt;</span></span>|<span data-ttu-id="d7fe0-1179">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1179">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="d7fe0-1180">函数</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1180">function</span></span>|<span data-ttu-id="d7fe0-1181">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1181">&lt;optional&gt;</span></span>|<span data-ttu-id="d7fe0-1182">方法完成后，使用单个参数 `callback`（一个 [`asyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `AsyncResult` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1182">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="d7fe0-1183">成功后，`asyncResult.value` 属性便以字符串形式提供初始化数据。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1183">On success, the initialization data is provided in the `asyncResult.value` property as a string.</span></span><br/><span data-ttu-id="d7fe0-1184">如果没有初始化上下文，`asyncResult` 对象包含 `Error` 对象，并将它的 `code` 和 `name` 属性分别设置为 `9020` 和 `GenericResponseError`。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1184">If there is no initialization context, the `asyncResult` object will contain an `Error` object with its `code` property set to `9020` and its `name` property set to `GenericResponseError`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d7fe0-1185">要求</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1185">Requirements</span></span>

|<span data-ttu-id="d7fe0-1186">要求</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1186">Requirement</span></span>|<span data-ttu-id="d7fe0-1187">值</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1187">Value</span></span>|
|---|---|
|[<span data-ttu-id="d7fe0-1188">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1188">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d7fe0-1189">预览</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1189">Preview</span></span>|
|[<span data-ttu-id="d7fe0-1190">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1190">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d7fe0-1191">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1191">ReadItem</span></span>|
|[<span data-ttu-id="d7fe0-1192">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1192">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="d7fe0-1193">阅读</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1193">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d7fe0-1194">示例</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1194">Example</span></span>

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

#### <a name="getregexmatches--object"></a><span data-ttu-id="d7fe0-1195">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1195">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="d7fe0-1196">返回所选项目中匹配在清单 XML 文件中定义的正则表达式的字符串值。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1196">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="d7fe0-1197">在 Outlook for iOS 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1197">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="d7fe0-p163">`getRegExMatches` 方法返回匹配在清单 XML 文件中的每个 `ItemHasRegularExpressionMatch` 或 `ItemHasKnownEntity` 规则元素中定义的正则表达式的字符串。对于 `ItemHasRegularExpressionMatch` 规则，匹配字符串必须发生在该规则指定的项目的属性中。`PropertyName` 简单类型定义支持的属性。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-p163">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="d7fe0-1201">例如，考虑一个外接程序清单具有以下 `Rule` 元素：</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1201">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="d7fe0-1202">从 `getRegExMatches` 返回的对象应有两个属性：`fruits` 和 `veggies`。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1202">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="d7fe0-p164">如果在项目的正文属性上指定 `ItemHasRegularExpressionMatch` 规则，则正则表达式应进一步筛选正文，不应尝试返回该项目的整个正文。使用正则表达式（如 `.*`）获取项目的整个正文并不总是返回预期的结果。而是使用 [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) 方法检索整个正文。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-p164">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="d7fe0-1206">Requirements</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1206">Requirements</span></span>

|<span data-ttu-id="d7fe0-1207">要求</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1207">Requirement</span></span>|<span data-ttu-id="d7fe0-1208">值</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1208">Value</span></span>|
|---|---|
|[<span data-ttu-id="d7fe0-1209">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1209">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d7fe0-1210">1.0</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1210">1.0</span></span>|
|[<span data-ttu-id="d7fe0-1211">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1211">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d7fe0-1212">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1212">ReadItem</span></span>|
|[<span data-ttu-id="d7fe0-1213">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1213">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="d7fe0-1214">阅读</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1214">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="d7fe0-1215">返回：</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1215">Returns:</span></span>

<span data-ttu-id="d7fe0-p165">一个包含匹配在清单 XML 文件中定义的正则表达式的字符串数组的对象。每个数组的名称等于匹配 `ItemHasRegularExpressionMatch` 规则的 `RegExName` 属性或匹配 `ItemHasKnownEntity` 规则的 `FilterName` 属性的相应值。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-p165">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<dl class="param-type"><span data-ttu-id="d7fe0-1218">

<dt>
类型</dt>


</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1218">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="d7fe0-1219">对象</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1219">Object</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="d7fe0-1220">示例</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1220">Example</span></span>

<span data-ttu-id="d7fe0-1221">以下示例显示了如何访问正则表达式规则元素 `fruits` 和 `veggies` 的匹配项的数组，这些元素在清单中指定。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1221">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```javascript
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veges = allMatches.veggies;
```

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="d7fe0-1222">getRegExMatchesByName(name) → (nullable) {Array.<String>}</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1222">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="d7fe0-1223">返回匹配在清单 XML 文件中定义的命名正则表达式的所选项目中的字符串值。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1223">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="d7fe0-1224">在 Outlook for iOS 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1224">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="d7fe0-1225">`getRegExMatchesByName` 方法返回匹配在具有指定 `RegExName` 元素值的清单 XML 文件中的 `ItemHasRegularExpressionMatch` 规则元素中定义的正则表达式的字符串。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1225">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="d7fe0-p166">如果在项目的正文属性上指定 `ItemHasRegularExpressionMatch` 规则，则正则表达式应进一步筛选正文，不应尝试返回该项目的整个正文。使用正则表达式（如 `.*`）获取项目的整个正文并不总是返回预期的结果。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-p166">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d7fe0-1228">参数：</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1228">Parameters:</span></span>

|<span data-ttu-id="d7fe0-1229">名称</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1229">Name</span></span>|<span data-ttu-id="d7fe0-1230">说明</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1230">Type</span></span>|<span data-ttu-id="d7fe0-1231">说明</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1231">Description</span></span>|
|---|---|---|
|`name`|<span data-ttu-id="d7fe0-1232">字符串</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1232">String</span></span>|<span data-ttu-id="d7fe0-1233">定义筛选器匹配的 `ItemHasRegularExpressionMatch` 规则元素的名称。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1233">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d7fe0-1234">Requirements</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1234">Requirements</span></span>

|<span data-ttu-id="d7fe0-1235">要求</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1235">Requirement</span></span>|<span data-ttu-id="d7fe0-1236">值</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1236">Value</span></span>|
|---|---|
|[<span data-ttu-id="d7fe0-1237">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1237">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d7fe0-1238">1.0</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1238">1.0</span></span>|
|[<span data-ttu-id="d7fe0-1239">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1239">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d7fe0-1240">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1240">ReadItem</span></span>|
|[<span data-ttu-id="d7fe0-1241">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1241">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="d7fe0-1242">阅读</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1242">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="d7fe0-1243">返回：</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1243">Returns:</span></span>

<span data-ttu-id="d7fe0-1244">一个包含匹配在清单 XML 文件中定义的正则表达式的字符串的数组。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1244">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<dl class="param-type"><span data-ttu-id="d7fe0-1245">

<dt>类型</dt>

</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1245">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="d7fe0-1246">Array.< String ></span><span class="sxs-lookup"><span data-stu-id="d7fe0-1246">Array.< String ></span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="d7fe0-1247">示例</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1247">Example</span></span>

```javascript
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

####  <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="d7fe0-1248">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1248">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="d7fe0-1249">以异步方式返回邮件的主题或正文中选定的数据。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1249">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="d7fe0-p167">如果没有选定内容，但光标位于正文或主题中，此方法将会为所选数据返回 null。如果选定的是字段，而不是正文或主题，则此方法返回 `InvalidSelection` 错误。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-p167">If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d7fe0-1252">参数：</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1252">Parameters:</span></span>

|<span data-ttu-id="d7fe0-1253">名称</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1253">Name</span></span>|<span data-ttu-id="d7fe0-1254">类型</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1254">Type</span></span>|<span data-ttu-id="d7fe0-1255">属性</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1255">Attributes</span></span>|<span data-ttu-id="d7fe0-1256">说明</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1256">Description</span></span>|
|---|---|---|---|
|`coercionType`|[<span data-ttu-id="d7fe0-1257">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1257">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="d7fe0-p168">请求数据的格式。如果为文本，则此方法返回纯文本作为字符串，删除任何显示的 HTML 标记。如果为 HTML，则此方法返回所选文本，不论是纯文本还是 HTML。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-p168">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`|<span data-ttu-id="d7fe0-1261">对象</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1261">Object</span></span>|<span data-ttu-id="d7fe0-1262">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1262">&lt;optional&gt;</span></span>|<span data-ttu-id="d7fe0-1263">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1263">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="d7fe0-1264">对象</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1264">Object</span></span>|<span data-ttu-id="d7fe0-1265">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1265">&lt;optional&gt;</span></span>|<span data-ttu-id="d7fe0-1266">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1266">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="d7fe0-1267">函数</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1267">function</span></span>||<span data-ttu-id="d7fe0-1268">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1268">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="d7fe0-1269">若要从回调方法访问所选数据，请调用 `asyncResult.value.data`。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1269">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="d7fe0-1270">若要访问选定内容的源属性，请调用 `asyncResult.value.sourceProperty`，这将为 `body` 或 `subject`。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1270">To access the source property that the selection comes from, call `asyncResult.value.sourceProperty`, which will be either `body` or `subject`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d7fe0-1271">Requirements</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1271">Requirements</span></span>

|<span data-ttu-id="d7fe0-1272">要求</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1272">Requirement</span></span>|<span data-ttu-id="d7fe0-1273">值</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1273">Value</span></span>|
|---|---|
|[<span data-ttu-id="d7fe0-1274">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1274">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d7fe0-1275">1.2</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1275">1.2</span></span>|
|[<span data-ttu-id="d7fe0-1276">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1276">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d7fe0-1277">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1277">ReadWriteItem</span></span>|
|[<span data-ttu-id="d7fe0-1278">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1278">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="d7fe0-1279">撰写</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1279">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="d7fe0-1280">返回：</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1280">Returns:</span></span>

<span data-ttu-id="d7fe0-1281">作为字符串的所选数据的格式由 `coercionType` 确定。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1281">The selected data as a string with format determined by `coercionType`.</span></span>

<dl class="param-type"><span data-ttu-id="d7fe0-1282">

<dt>
类型</dt>


</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1282">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="d7fe0-1283">字符串</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1283">String</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="d7fe0-1284">示例</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1284">Example</span></span>

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

#### <a name="getselectedentities--entitiesjavascriptapioutlookofficeentities"></a><span data-ttu-id="d7fe0-1285">getSelectedEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1285">getSelectedEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span></span>

<span data-ttu-id="d7fe0-p170">获取在用户已选择的突出显示匹配项中找到的实体。突出显示匹配项适用于[上下文加载项](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins)。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-p170">Gets the entities found in a highlighted match a user has selected. Highlighted matches apply to [contextual add-ins](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="d7fe0-1288">在 Outlook for iOS 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1288">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="d7fe0-1289">Requirements</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1289">Requirements</span></span>

|<span data-ttu-id="d7fe0-1290">要求</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1290">Requirement</span></span>|<span data-ttu-id="d7fe0-1291">值</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1291">Value</span></span>|
|---|---|
|[<span data-ttu-id="d7fe0-1292">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1292">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d7fe0-1293">1.6</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1293">1.6</span></span>|
|[<span data-ttu-id="d7fe0-1294">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1294">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d7fe0-1295">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1295">ReadItem</span></span>|
|[<span data-ttu-id="d7fe0-1296">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1296">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="d7fe0-1297">阅读</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1297">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="d7fe0-1298">返回：</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1298">Returns:</span></span>

<span data-ttu-id="d7fe0-1299">类型：[Entities](/javascript/api/outlook/office.entities)</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1299">Type: [Entities](/javascript/api/outlook/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="d7fe0-1300">示例</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1300">Example</span></span>

<span data-ttu-id="d7fe0-1301">以下示例访问用户选择的突出显示匹配项中的地址实体。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1301">The following example accesses the addresses entities in the highlighted match selected by the user.</span></span>

```javascript
var contacts = Office.context.mailbox.item.getSelectedEntities().addresses;
```

#### <a name="getselectedregexmatches--object"></a><span data-ttu-id="d7fe0-1302">getSelectedRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1302">getSelectedRegExMatches() → {Object}</span></span>

<span data-ttu-id="d7fe0-p171">返回突出显示匹配项中匹配在清单 XML 文件中定义的正则表达式的字符串值。突出显示匹配项适用于[上下文外接程序](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins)。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-p171">Returns string values in a highlighted match that match the regular expressions defined in the manifest XML file. Highlighted matches apply to [contextual add-ins](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="d7fe0-1305">在 Outlook for iOS 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1305">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="d7fe0-p172">`getSelectedRegExMatches` 方法返回匹配在清单 XML 文件中的每个 `ItemHasRegularExpressionMatch` 或 `ItemHasKnownEntity` 规则元素中定义的正则表达式的字符串。对于 `ItemHasRegularExpressionMatch` 规则，匹配字符串必须发生在该规则指定的项目的属性中。`PropertyName` 简单类型定义支持的属性。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-p172">The `getSelectedRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="d7fe0-1309">例如，考虑一个外接程序清单具有以下 `Rule` 元素：</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1309">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="d7fe0-1310">从 `getRegExMatches` 返回的对象应有两个属性：`fruits` 和 `veggies`。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1310">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="d7fe0-p173">如果在项目的正文属性上指定 `ItemHasRegularExpressionMatch` 规则，则正则表达式应进一步筛选正文，不应尝试返回该项目的整个正文。使用正则表达式（如 `.*`）获取项目的整个正文并不总是返回预期的结果。而是使用 [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) 方法检索整个正文。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-p173">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="d7fe0-1314">Requirements</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1314">Requirements</span></span>

|<span data-ttu-id="d7fe0-1315">要求</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1315">Requirement</span></span>|<span data-ttu-id="d7fe0-1316">值</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1316">Value</span></span>|
|---|---|
|[<span data-ttu-id="d7fe0-1317">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1317">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d7fe0-1318">1.6</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1318">1.6</span></span>|
|[<span data-ttu-id="d7fe0-1319">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1319">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d7fe0-1320">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1320">ReadItem</span></span>|
|[<span data-ttu-id="d7fe0-1321">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1321">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="d7fe0-1322">阅读</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1322">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="d7fe0-1323">返回：</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1323">Returns:</span></span>

<span data-ttu-id="d7fe0-p174">一个包含匹配在清单 XML 文件中定义的正则表达式的字符串数组的对象。每个数组的名称等于匹配 `ItemHasRegularExpressionMatch` 规则的 `RegExName` 属性或匹配 `ItemHasKnownEntity` 规则的 `FilterName` 属性的相应值。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-p174">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

##### <a name="example"></a><span data-ttu-id="d7fe0-1326">示例</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1326">Example</span></span>

<span data-ttu-id="d7fe0-1327">以下示例显示了如何访问正则表达式规则元素 `fruits` 和 `veggies` 的匹配项的数组，这些元素在清单中指定。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1327">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```javascript
var selectedMatches = Office.context.mailbox.item.getSelectedRegExMatches();
var fruits = selectedMatches.fruits;
var veggies = selectedMatches.veggies;
```

#### <a name="getsharedpropertiesasyncoptions-callback"></a><span data-ttu-id="d7fe0-1328">getSharedPropertiesAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1328">getSharedPropertiesAsync([options], callback)</span></span>

<span data-ttu-id="d7fe0-1329">获取共享文件夹、日历或邮箱中所选约会或邮件的属性。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1329">Gets the properties of the selected appointment or message in a shared folder, calendar, or mailbox.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d7fe0-1330">参数：</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1330">Parameters:</span></span>

|<span data-ttu-id="d7fe0-1331">名称</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1331">Name</span></span>|<span data-ttu-id="d7fe0-1332">类型</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1332">Type</span></span>|<span data-ttu-id="d7fe0-1333">属性</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1333">Attributes</span></span>|<span data-ttu-id="d7fe0-1334">说明</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1334">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="d7fe0-1335">Object</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1335">Object</span></span>|<span data-ttu-id="d7fe0-1336">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1336">&lt;optional&gt;</span></span>|<span data-ttu-id="d7fe0-1337">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1337">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="d7fe0-1338">对象</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1338">Object</span></span>|<span data-ttu-id="d7fe0-1339">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1339">&lt;optional&gt;</span></span>|<span data-ttu-id="d7fe0-1340">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1340">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="d7fe0-1341">function</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1341">function</span></span>||<span data-ttu-id="d7fe0-1342">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1342">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="d7fe0-1343">共享属性作为 `asyncResult.value` 属性中的 [`SharedProperties`](/javascript/api/outlook/office.sharedproperties) 对象提供。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1343">The shared properties are provided as a [`SharedProperties`](/javascript/api/outlook/office.sharedproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="d7fe0-1344">此对象可用于获取项目的共享属性。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1344">This object can be used to get the item's shared properties.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d7fe0-1345">要求</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1345">Requirements</span></span>

|<span data-ttu-id="d7fe0-1346">要求</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1346">Requirement</span></span>|<span data-ttu-id="d7fe0-1347">值</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1347">Value</span></span>|
|---|---|
|[<span data-ttu-id="d7fe0-1348">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1348">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d7fe0-1349">预览</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1349">Preview</span></span>|
|[<span data-ttu-id="d7fe0-1350">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1350">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d7fe0-1351">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1351">ReadItem</span></span>|
|[<span data-ttu-id="d7fe0-1352">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1352">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="d7fe0-1353">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1353">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="d7fe0-1354">示例</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1354">Example</span></span>

```js
Office.context.mailbox.item.getSharedPropertiesAsync(callback);
function callback (asyncResult) {
  var context=asyncResult.context;
  var sharedProperties = asyncResult.value;
}
```

####  <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="d7fe0-1355">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1355">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="d7fe0-1356">异步加载所选项目上此外接程序的自定义属性。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1356">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="d7fe0-p176">自定义属性基于每个应用、每个项目存储为键/值对。此方法在回调中返回 `CustomProperties` 对象，该回调提供访问特定于当前项目和当前外接程序的自定义属性的方法。自定义属性未在项目上加密，因此这不应用作安全存储。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-p176">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d7fe0-1360">参数：</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1360">Parameters:</span></span>

|<span data-ttu-id="d7fe0-1361">名称</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1361">Name</span></span>|<span data-ttu-id="d7fe0-1362">类型</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1362">Type</span></span>|<span data-ttu-id="d7fe0-1363">属性</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1363">Attributes</span></span>|<span data-ttu-id="d7fe0-1364">说明</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1364">Description</span></span>|
|---|---|---|---|
|`callback`|<span data-ttu-id="d7fe0-1365">函数</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1365">function</span></span>||<span data-ttu-id="d7fe0-1366">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1366">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="d7fe0-1367">自定义属性作为 `asyncResult.value` 属性中的 [`CustomProperties`](/javascript/api/outlook/office.customproperties) 对象提供。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1367">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook/office.customproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="d7fe0-1368">此对象可用于获取、设置以及从项目中删除自定义属性，并将自定义属性集的更改重新保存到服务器。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1368">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`|<span data-ttu-id="d7fe0-1369">对象</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1369">Object</span></span>|<span data-ttu-id="d7fe0-1370">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1370">&lt;optional&gt;</span></span>|<span data-ttu-id="d7fe0-1371">开发人员可以提供他们想要在回调函数中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1371">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="d7fe0-1372">此对象可以通过回调函数中的 `asyncResult.asyncContext` 属性进行访问。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1372">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d7fe0-1373">要求</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1373">Requirements</span></span>

|<span data-ttu-id="d7fe0-1374">要求</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1374">Requirement</span></span>|<span data-ttu-id="d7fe0-1375">值</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1375">Value</span></span>|
|---|---|
|[<span data-ttu-id="d7fe0-1376">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1376">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d7fe0-1377">1.0</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1377">1.0</span></span>|
|[<span data-ttu-id="d7fe0-1378">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1378">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d7fe0-1379">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1379">ReadItem</span></span>|
|[<span data-ttu-id="d7fe0-1380">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1380">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="d7fe0-1381">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1381">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="d7fe0-1382">示例</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1382">Example</span></span>

<span data-ttu-id="d7fe0-p179">以下代码示例显示了如何使用 `loadCustomPropertiesAsync` 方法异步加载特定于当前项目的自定义属性。该示例还显示了如何使用 `CustomProperties.saveAsync` 方法将这些属性重新保存到服务器。加载自定义属性后，该代码示例将使用 `CustomProperties.get` 方法读取自定义属性 `myProp`，使用 `CustomProperties.set` 方法写入自定义属性 `otherProp`，最后调用 `saveAsync` 方法保存这些自定义属性。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-p179">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

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

####  <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="d7fe0-1386">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1386">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="d7fe0-1387">将附件从邮件或约会中删除。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1387">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="d7fe0-1388">`removeAttachmentAsync` 方法删除项目中带指定标识符的附件。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1388">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item.</span></span> <span data-ttu-id="d7fe0-1389">最佳做法是，仅当同一个邮件应用程序在同一会话中添加了一个附件时，你才应使用该附件标识符来删除该附件。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1389">As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session.</span></span> <span data-ttu-id="d7fe0-1390">在 Outlook Web App 和适用于设备的 OWA 中，附件标识符只在同一个会话中才有效。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1390">In Outlook Web App and OWA for Devices, the attachment identifier is valid only within the same session.</span></span> <span data-ttu-id="d7fe0-1391">当用户关闭应用，或者如果用户开始在内嵌窗体中撰写，则随后弹出的窗体中继续在单独的窗口撰写时，会话即结束。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1391">A session is over when the user closes the app, or if the user starts composing an inline form then subsequently pops out the form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d7fe0-1392">参数：</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1392">Parameters:</span></span>

|<span data-ttu-id="d7fe0-1393">名称</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1393">Name</span></span>|<span data-ttu-id="d7fe0-1394">类型</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1394">Type</span></span>|<span data-ttu-id="d7fe0-1395">属性</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1395">Attributes</span></span>|<span data-ttu-id="d7fe0-1396">说明</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1396">Description</span></span>|
|---|---|---|---|
|`attachmentId`|<span data-ttu-id="d7fe0-1397">字符串</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1397">String</span></span>||<span data-ttu-id="d7fe0-1398">要删除的附件的标识符。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1398">The identifier of the attachment to remove.</span></span>|
|`options`|<span data-ttu-id="d7fe0-1399">对象</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1399">Object</span></span>|<span data-ttu-id="d7fe0-1400">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1400">&lt;optional&gt;</span></span>|<span data-ttu-id="d7fe0-1401">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1401">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="d7fe0-1402">对象</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1402">Object</span></span>|<span data-ttu-id="d7fe0-1403">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1403">&lt;optional&gt;</span></span>|<span data-ttu-id="d7fe0-1404">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1404">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="d7fe0-1405">函数</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1405">function</span></span>|<span data-ttu-id="d7fe0-1406">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1406">&lt;optional&gt;</span></span>|<span data-ttu-id="d7fe0-1407">方法完成后，使用单个参数 `callback`（一个 [`asyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `AsyncResult` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1407">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="d7fe0-1408">如果删除附件失败，`asyncResult.error` 属性将包含一个说明失败原因的错误代码。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1408">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="d7fe0-1409">错误</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1409">Errors</span></span>

|<span data-ttu-id="d7fe0-1410">错误代码</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1410">Error code</span></span>|<span data-ttu-id="d7fe0-1411">说明</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1411">Description</span></span>|
|------------|-------------|
|`InvalidAttachmentId`|<span data-ttu-id="d7fe0-1412">附件标识符不存在。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1412">The attachment identifier does not exist.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d7fe0-1413">Requirements</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1413">Requirements</span></span>

|<span data-ttu-id="d7fe0-1414">要求</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1414">Requirement</span></span>|<span data-ttu-id="d7fe0-1415">值</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1415">Value</span></span>|
|---|---|
|[<span data-ttu-id="d7fe0-1416">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1416">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d7fe0-1417">1.1</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1417">1.1</span></span>|
|[<span data-ttu-id="d7fe0-1418">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1418">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d7fe0-1419">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1419">ReadWriteItem</span></span>|
|[<span data-ttu-id="d7fe0-1420">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1420">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="d7fe0-1421">撰写</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1421">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="d7fe0-1422">示例</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1422">Example</span></span>

<span data-ttu-id="d7fe0-1423">以下代码删除包含标识符 '0' 的附件。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1423">The following code removes an attachment with an identifier of '0'.</span></span>

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

####  <a name="removehandlerasynceventtype-options-callback"></a><span data-ttu-id="d7fe0-1424">removeHandlerAsync(eventType, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1424">removeHandlerAsync(eventType, [options], [callback])</span></span>

<span data-ttu-id="d7fe0-1425">删除受支持事件类型的事件处理程序。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1425">Removes the event handlers for a supported event type.</span></span>

<span data-ttu-id="d7fe0-1426">当前，支持的事件类型是 `Office.EventType.AttachmentsChanged`、`Office.EventType.AppointmentTimeChanged`、`Office.EventType.RecipientsChanged` 和 `Office.EventType.RecurrenceChanged`</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1426">Currently the supported event types are `Office.EventType.AttachmentsChanged`, `Office.EventType.AppointmentTimeChanged`, `Office.EventType.RecipientsChanged`, and `Office.EventType.RecurrenceChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d7fe0-1427">参数：</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1427">Parameters:</span></span>

| <span data-ttu-id="d7fe0-1428">名称</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1428">Name</span></span> | <span data-ttu-id="d7fe0-1429">类型</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1429">Type</span></span> | <span data-ttu-id="d7fe0-1430">属性</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1430">Attributes</span></span> | <span data-ttu-id="d7fe0-1431">说明</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1431">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="d7fe0-1432">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1432">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="d7fe0-1433">应撤销处理程序的事件。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1433">The event that should revoke the handler.</span></span> |
| `options` | <span data-ttu-id="d7fe0-1434">对象</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1434">Object</span></span> | <span data-ttu-id="d7fe0-1435">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1435">&lt;optional&gt;</span></span> | <span data-ttu-id="d7fe0-1436">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1436">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="d7fe0-1437">对象</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1437">Object</span></span> | <span data-ttu-id="d7fe0-1438">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1438">&lt;optional&gt;</span></span> | <span data-ttu-id="d7fe0-1439">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1439">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="d7fe0-1440">函数</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1440">function</span></span>| <span data-ttu-id="d7fe0-1441">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1441">&lt;optional&gt;</span></span>|<span data-ttu-id="d7fe0-1442">方法完成后，使用单个参数 `callback`（一个 [`asyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `AsyncResult` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1442">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d7fe0-1443">Requirements</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1443">Requirements</span></span>

|<span data-ttu-id="d7fe0-1444">要求</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1444">Requirement</span></span>| <span data-ttu-id="d7fe0-1445">值</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1445">Value</span></span>|
|---|---|
|[<span data-ttu-id="d7fe0-1446">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1446">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d7fe0-1447">1.7</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1447">1.7</span></span> |
|[<span data-ttu-id="d7fe0-1448">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1448">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d7fe0-1449">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1449">ReadItem</span></span> |
|[<span data-ttu-id="d7fe0-1450">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1450">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="d7fe0-1451">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1451">Compose or read</span></span> |

####  <a name="saveasyncoptions-callback"></a><span data-ttu-id="d7fe0-1452">saveAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1452">saveAsync([options], callback)</span></span>

<span data-ttu-id="d7fe0-1453">异步保存项目。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1453">Asynchronously saves an item.</span></span>

<span data-ttu-id="d7fe0-p181">调用时，此方法将当前邮件保存为草稿，并通过回调方法返回项目 ID。在 Outlook Web App 或 Outlook 联机模式下，该项目被保存到服务器中。在 Outlook 缓存模式下，该项目被保存到本地缓存中。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-p181">When invoked, this method saves the current message as a draft and returns the item id via the callback method. In Outlook Web App or Outlook in online mode, the item is saved to the server. In Outlook in cached mode, the item is saved to the local cache.</span></span>

> [!NOTE]
> <span data-ttu-id="d7fe0-1457">如果加载项在撰写模式下对某个项目调用 `saveAsync` 来获得 `itemId`，以便与 EWS 或 REST API 一同使用，请注意，当 Outlook 处于高速缓存模式时，可能需要一段时间项目才能真正同步到服务器。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1457">If your add-in calls `saveAsync` on an item in compose mode in order to get an `itemId` to use with EWS or the REST API, be aware that when Outlook is in cached mode, it may take some time before the item is actually synced to the server.</span></span> <span data-ttu-id="d7fe0-1458">在项目同步前，使用 `itemId` 将返回一个错误。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1458">Until the item is synced, using the `itemId` will return an error.</span></span>

<span data-ttu-id="d7fe0-p183">由于约会没有草稿状态，如果以撰写模式在约会中调用 `saveAsync`，则该项将被保存为用户日历中的正常约会。对于之前未保存过的新约会，则不会发送邀请。保存现有约会将向添加或删除的与会者发送更新。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-p183">Since appointments have no draft state, if `saveAsync` is called on an appointment in compose mode, the item will be saved as a normal appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent. Saving an existing appointment will send an update to added or removed attendees.</span></span>

> [!NOTE]
> <span data-ttu-id="d7fe0-1462">以下客户端在撰写模式下对约会上的 `saveAsync` 具有不同的行为：</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1462">The following clients have different behavior for `saveAsync` on appointments in compose mode:</span></span>
>
> - <span data-ttu-id="d7fe0-1463">Mac Outlook 不支持在撰写模式下对会议执行 `saveAsync` 操作。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1463">Mac Outlook does not support `saveAsync` on a meeting in compose mode.</span></span> <span data-ttu-id="d7fe0-1464">对 Mac Outlook 中的会议调用 `saveAsync` 将会返回一个错误。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1464">Calling `saveAsync` on a meeting in Mac Outlook will return an error.</span></span>
> - <span data-ttu-id="d7fe0-1465">在撰写模式下的约会上调用 `saveAsync` 时，Outlook 网页版始终发送邀请或更新。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1465">Outlook on the web always sends an invitation or update when `saveAsync` is called on an appointment in compose mode.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d7fe0-1466">参数：</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1466">Parameters:</span></span>

|<span data-ttu-id="d7fe0-1467">名称</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1467">Name</span></span>|<span data-ttu-id="d7fe0-1468">类型</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1468">Type</span></span>|<span data-ttu-id="d7fe0-1469">属性</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1469">Attributes</span></span>|<span data-ttu-id="d7fe0-1470">说明</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1470">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="d7fe0-1471">Object</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1471">Object</span></span>|<span data-ttu-id="d7fe0-1472">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1472">&lt;optional&gt;</span></span>|<span data-ttu-id="d7fe0-1473">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1473">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="d7fe0-1474">对象</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1474">Object</span></span>|<span data-ttu-id="d7fe0-1475">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1475">&lt;optional&gt;</span></span>|<span data-ttu-id="d7fe0-1476">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1476">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="d7fe0-1477">函数</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1477">function</span></span>||<span data-ttu-id="d7fe0-1478">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1478">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="d7fe0-1479">如果成功，该项目标识符将在 `asyncResult.value` 属性中提供。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1479">On success, the item identifier is provided in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d7fe0-1480">要求</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1480">Requirements</span></span>

|<span data-ttu-id="d7fe0-1481">要求</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1481">Requirement</span></span>|<span data-ttu-id="d7fe0-1482">值</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1482">Value</span></span>|
|---|---|
|[<span data-ttu-id="d7fe0-1483">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1483">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d7fe0-1484">1.3</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1484">1.3</span></span>|
|[<span data-ttu-id="d7fe0-1485">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1485">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d7fe0-1486">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1486">ReadWriteItem</span></span>|
|[<span data-ttu-id="d7fe0-1487">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1487">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="d7fe0-1488">撰写</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1488">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="d7fe0-1489">示例</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1489">Examples</span></span>

```javascript
Office.context.mailbox.item.saveAsync(
  function callback(result) {
    // Process the result
  });
```

<span data-ttu-id="d7fe0-p185">下面是传递给回调函数的 `result` 参数的示例。`value` 属性包含的项目的项目 ID。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-p185">The following is an example of the `result` parameter passed to the callback function. The `value` property contains the item ID of the item.</span></span>

```json
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

####  <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="d7fe0-1492">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1492">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="d7fe0-1493">以异步方式将数据插入到邮件的正文或主题中。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1493">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="d7fe0-p186">`setSelectedDataAsync` 方法将指定字符串插入到项目主题或正文的光标位置，或者，如果在编辑器中已选择文本，则该方法将替换选择的文本。如果光标不在正文或主题字段中，则返回错误。插入之后，光标会位于插入内容的末尾。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-p186">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d7fe0-1497">参数：</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1497">Parameters:</span></span>

|<span data-ttu-id="d7fe0-1498">名称</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1498">Name</span></span>|<span data-ttu-id="d7fe0-1499">类型</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1499">Type</span></span>|<span data-ttu-id="d7fe0-1500">属性</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1500">Attributes</span></span>|<span data-ttu-id="d7fe0-1501">说明</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1501">Description</span></span>|
|---|---|---|---|
|`data`|<span data-ttu-id="d7fe0-1502">字符串</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1502">String</span></span>||<span data-ttu-id="d7fe0-p187">要插入的数据。数据不得超过 1,000,000 个字符。如果传入的数据超过 1,000,000 个字符，则会引发 `ArgumentOutOfRange` 异常。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-p187">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`|<span data-ttu-id="d7fe0-1506">Object</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1506">Object</span></span>|<span data-ttu-id="d7fe0-1507">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1507">&lt;optional&gt;</span></span>|<span data-ttu-id="d7fe0-1508">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1508">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="d7fe0-1509">对象</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1509">Object</span></span>|<span data-ttu-id="d7fe0-1510">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1510">&lt;optional&gt;</span></span>|<span data-ttu-id="d7fe0-1511">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1511">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`|[<span data-ttu-id="d7fe0-1512">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1512">Office.CoercionType</span></span>](office.md#coerciontype-string)|<span data-ttu-id="d7fe0-1513">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1513">&lt;optional&gt;</span></span>|<span data-ttu-id="d7fe0-p188">如果为 `text`，则在 Outlook Web App 和 Outlook 中应用当前样式。如果该字段是 HTML 编辑器，则仅插入文本数据，即使数据为 HTML。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-p188">If `text`, the current style is applied in Outlook Web App and Outlook. If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="d7fe0-p189">如果 `html` 和该字段支持 HTML（主题不支持），则在 Outlook Web App 中应用当前样式，而在 Outlook 中应用默认样式。如果该字段是文本字段，则返回 `InvalidDataFormat` 错误。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-p189">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook Web App and the default style is applied in Outlook. If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="d7fe0-1518">如果未设置 `coercionType`，则结果取决于该字段：如果该字段是 HTML，则使用 HTML；如果该字段是文本，则使用纯文本。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1518">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`|<span data-ttu-id="d7fe0-1519">function</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1519">function</span></span>||<span data-ttu-id="d7fe0-1520">方法完成后，使用单个参数 `callback`（一个 [`asyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `AsyncResult` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1520">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d7fe0-1521">Requirements</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1521">Requirements</span></span>

|<span data-ttu-id="d7fe0-1522">要求</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1522">Requirement</span></span>|<span data-ttu-id="d7fe0-1523">值</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1523">Value</span></span>|
|---|---|
|[<span data-ttu-id="d7fe0-1524">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1524">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="d7fe0-1525">1.2</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1525">1.2</span></span>|
|[<span data-ttu-id="d7fe0-1526">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1526">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="d7fe0-1527">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1527">ReadWriteItem</span></span>|
|[<span data-ttu-id="d7fe0-1528">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1528">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="d7fe0-1529">撰写</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1529">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="d7fe0-1530">示例</span><span class="sxs-lookup"><span data-stu-id="d7fe0-1530">Example</span></span>

```javascript
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```
