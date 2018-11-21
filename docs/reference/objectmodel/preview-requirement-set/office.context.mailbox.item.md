
# <a name="item"></a><span data-ttu-id="e64f0-101">item</span><span class="sxs-lookup"><span data-stu-id="e64f0-101">item</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmditem"></a><span data-ttu-id="e64f0-102">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span><span class="sxs-lookup"><span data-stu-id="e64f0-102">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span></span>

<span data-ttu-id="e64f0-p101">`item` 命名空间用于访问当前选定的邮件、会议请求或约会。可以通过使用 [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtype) 属性确定 `item` 的类型。</span><span class="sxs-lookup"><span data-stu-id="e64f0-p101">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="e64f0-105">Requirements</span><span class="sxs-lookup"><span data-stu-id="e64f0-105">Requirements</span></span>

|<span data-ttu-id="e64f0-106">要求</span><span class="sxs-lookup"><span data-stu-id="e64f0-106">Requirement</span></span>|<span data-ttu-id="e64f0-107">值</span><span class="sxs-lookup"><span data-stu-id="e64f0-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="e64f0-108">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="e64f0-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e64f0-109">1.0</span><span class="sxs-lookup"><span data-stu-id="e64f0-109">1.0</span></span>|
|[<span data-ttu-id="e64f0-110">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="e64f0-110">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e64f0-111">受限</span><span class="sxs-lookup"><span data-stu-id="e64f0-111">Restricted</span></span>|
|[<span data-ttu-id="e64f0-112">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="e64f0-112">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="e64f0-113">Compose 或 Read</span><span class="sxs-lookup"><span data-stu-id="e64f0-113">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="e64f0-114">成员和方法</span><span class="sxs-lookup"><span data-stu-id="e64f0-114">Members and methods</span></span>

| <span data-ttu-id="e64f0-115">成员</span><span class="sxs-lookup"><span data-stu-id="e64f0-115">Member</span></span> | <span data-ttu-id="e64f0-116">类型</span><span class="sxs-lookup"><span data-stu-id="e64f0-116">Type</span></span> |
|--------|------|
| [<span data-ttu-id="e64f0-117">attachments</span><span class="sxs-lookup"><span data-stu-id="e64f0-117">attachments</span></span>](#attachments-arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetails) | <span data-ttu-id="e64f0-118">成员</span><span class="sxs-lookup"><span data-stu-id="e64f0-118">Member</span></span> |
| [<span data-ttu-id="e64f0-119">bcc</span><span class="sxs-lookup"><span data-stu-id="e64f0-119">bcc</span></span>](#bcc-recipientsjavascriptapioutlookofficerecipients) | <span data-ttu-id="e64f0-120">成员</span><span class="sxs-lookup"><span data-stu-id="e64f0-120">Member</span></span> |
| [<span data-ttu-id="e64f0-121">body</span><span class="sxs-lookup"><span data-stu-id="e64f0-121">body</span></span>](#body-bodyjavascriptapioutlookofficebody) | <span data-ttu-id="e64f0-122">成员</span><span class="sxs-lookup"><span data-stu-id="e64f0-122">Member</span></span> |
| [<span data-ttu-id="e64f0-123">cc</span><span class="sxs-lookup"><span data-stu-id="e64f0-123">cc</span></span>](#cc-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients) | <span data-ttu-id="e64f0-124">成员</span><span class="sxs-lookup"><span data-stu-id="e64f0-124">Member</span></span> |
| [<span data-ttu-id="e64f0-125">conversationId</span><span class="sxs-lookup"><span data-stu-id="e64f0-125">conversationId</span></span>](#nullable-conversationid-string) | <span data-ttu-id="e64f0-126">成员</span><span class="sxs-lookup"><span data-stu-id="e64f0-126">Member</span></span> |
| [<span data-ttu-id="e64f0-127">dateTimeCreated</span><span class="sxs-lookup"><span data-stu-id="e64f0-127">dateTimeCreated</span></span>](#datetimecreated-date) | <span data-ttu-id="e64f0-128">成员</span><span class="sxs-lookup"><span data-stu-id="e64f0-128">Member</span></span> |
| [<span data-ttu-id="e64f0-129">dateTimeModified</span><span class="sxs-lookup"><span data-stu-id="e64f0-129">dateTimeModified</span></span>](#datetimemodified-date) | <span data-ttu-id="e64f0-130">成员</span><span class="sxs-lookup"><span data-stu-id="e64f0-130">Member</span></span> |
| [<span data-ttu-id="e64f0-131">end</span><span class="sxs-lookup"><span data-stu-id="e64f0-131">end</span></span>](#end-datetimejavascriptapioutlookofficetime) | <span data-ttu-id="e64f0-132">成员</span><span class="sxs-lookup"><span data-stu-id="e64f0-132">Member</span></span> |
| [<span data-ttu-id="e64f0-133">from</span><span class="sxs-lookup"><span data-stu-id="e64f0-133">from</span></span>](#from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsfromjavascriptapioutlookofficefrom) | <span data-ttu-id="e64f0-134">Member</span><span class="sxs-lookup"><span data-stu-id="e64f0-134">Member</span></span> |
| [<span data-ttu-id="e64f0-135">internetHeaders</span><span class="sxs-lookup"><span data-stu-id="e64f0-135">internetHeaders</span></span>](#internetheaders-internetheadersjavascriptapioutlookofficeinternetheaders) | <span data-ttu-id="e64f0-136">Member</span><span class="sxs-lookup"><span data-stu-id="e64f0-136">Member</span></span> |
| [<span data-ttu-id="e64f0-137">internetMessageId</span><span class="sxs-lookup"><span data-stu-id="e64f0-137">internetMessageId</span></span>](#internetmessageid-string) | <span data-ttu-id="e64f0-138">成员</span><span class="sxs-lookup"><span data-stu-id="e64f0-138">Member</span></span> |
| [<span data-ttu-id="e64f0-139">itemClass</span><span class="sxs-lookup"><span data-stu-id="e64f0-139">itemClass</span></span>](#itemclass-string) | <span data-ttu-id="e64f0-140">成员</span><span class="sxs-lookup"><span data-stu-id="e64f0-140">Member</span></span> |
| [<span data-ttu-id="e64f0-141">itemId</span><span class="sxs-lookup"><span data-stu-id="e64f0-141">itemId</span></span>](#nullable-itemid-string) | <span data-ttu-id="e64f0-142">成员</span><span class="sxs-lookup"><span data-stu-id="e64f0-142">Member</span></span> |
| [<span data-ttu-id="e64f0-143">itemType</span><span class="sxs-lookup"><span data-stu-id="e64f0-143">itemType</span></span>](#itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtype) | <span data-ttu-id="e64f0-144">成员</span><span class="sxs-lookup"><span data-stu-id="e64f0-144">Member</span></span> |
| [<span data-ttu-id="e64f0-145">location</span><span class="sxs-lookup"><span data-stu-id="e64f0-145">location</span></span>](#location-stringlocationjavascriptapioutlookofficelocation) | <span data-ttu-id="e64f0-146">成员</span><span class="sxs-lookup"><span data-stu-id="e64f0-146">Member</span></span> |
| [<span data-ttu-id="e64f0-147">normalizedSubject</span><span class="sxs-lookup"><span data-stu-id="e64f0-147">normalizedSubject</span></span>](#normalizedsubject-string) | <span data-ttu-id="e64f0-148">成员</span><span class="sxs-lookup"><span data-stu-id="e64f0-148">Member</span></span> |
| [<span data-ttu-id="e64f0-149">notificationMessages</span><span class="sxs-lookup"><span data-stu-id="e64f0-149">notificationMessages</span></span>](#notificationmessages-notificationmessagesjavascriptapioutlookofficenotificationmessages) | <span data-ttu-id="e64f0-150">成员</span><span class="sxs-lookup"><span data-stu-id="e64f0-150">Member</span></span> |
| [<span data-ttu-id="e64f0-151">optionalAttendees</span><span class="sxs-lookup"><span data-stu-id="e64f0-151">optionalAttendees</span></span>](#optionalattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients) | <span data-ttu-id="e64f0-152">成员</span><span class="sxs-lookup"><span data-stu-id="e64f0-152">Member</span></span> |
| [<span data-ttu-id="e64f0-153">organizer</span><span class="sxs-lookup"><span data-stu-id="e64f0-153">organizer</span></span>](#organizer-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsorganizerjavascriptapioutlookofficeorganizer) | <span data-ttu-id="e64f0-154">Member</span><span class="sxs-lookup"><span data-stu-id="e64f0-154">Member</span></span> |
| [<span data-ttu-id="e64f0-155">recurrence</span><span class="sxs-lookup"><span data-stu-id="e64f0-155">recurrence</span></span>](#nullable-recurrence-recurrencejavascriptapioutlookofficerecurrence) | <span data-ttu-id="e64f0-156">Member</span><span class="sxs-lookup"><span data-stu-id="e64f0-156">Member</span></span> |
| [<span data-ttu-id="e64f0-157">requiredAttendees</span><span class="sxs-lookup"><span data-stu-id="e64f0-157">requiredAttendees</span></span>](#requiredattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients) | <span data-ttu-id="e64f0-158">Member</span><span class="sxs-lookup"><span data-stu-id="e64f0-158">Member</span></span> |
| [<span data-ttu-id="e64f0-159">sender</span><span class="sxs-lookup"><span data-stu-id="e64f0-159">sender</span></span>](#sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetails) | <span data-ttu-id="e64f0-160">Member</span><span class="sxs-lookup"><span data-stu-id="e64f0-160">Member</span></span> |
| [<span data-ttu-id="e64f0-161">seriesId</span><span class="sxs-lookup"><span data-stu-id="e64f0-161">seriesId</span></span>](#nullable-seriesid-string) | <span data-ttu-id="e64f0-162">Member</span><span class="sxs-lookup"><span data-stu-id="e64f0-162">Member</span></span> |
| [<span data-ttu-id="e64f0-163">start</span><span class="sxs-lookup"><span data-stu-id="e64f0-163">start</span></span>](#start-datetimejavascriptapioutlookofficetime) | <span data-ttu-id="e64f0-164">成员</span><span class="sxs-lookup"><span data-stu-id="e64f0-164">Member</span></span> |
| [<span data-ttu-id="e64f0-165">subject</span><span class="sxs-lookup"><span data-stu-id="e64f0-165">subject</span></span>](#subject-stringsubjectjavascriptapioutlookofficesubject) | <span data-ttu-id="e64f0-166">成员</span><span class="sxs-lookup"><span data-stu-id="e64f0-166">Member</span></span> |
| [<span data-ttu-id="e64f0-167">to</span><span class="sxs-lookup"><span data-stu-id="e64f0-167">to</span></span>](#to-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients) | <span data-ttu-id="e64f0-168">成员</span><span class="sxs-lookup"><span data-stu-id="e64f0-168">Member</span></span> |
| [<span data-ttu-id="e64f0-169">addFileAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="e64f0-169">addFileAttachmentAsync</span></span>](#addfileattachmentasyncuri-attachmentname-options-callback) | <span data-ttu-id="e64f0-170">方法</span><span class="sxs-lookup"><span data-stu-id="e64f0-170">Method</span></span> |
| [<span data-ttu-id="e64f0-171">addFileAttachmentFromBase64Async</span><span class="sxs-lookup"><span data-stu-id="e64f0-171">addFileAttachmentFromBase64Async</span></span>](#addfileattachmentfrombase64asyncbase64file-attachmentname-options-callback) | <span data-ttu-id="e64f0-172">方法</span><span class="sxs-lookup"><span data-stu-id="e64f0-172">Method</span></span> |
| [<span data-ttu-id="e64f0-173">addHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="e64f0-173">addHandlerAsync</span></span>](#addhandlerasynceventtype-handler-options-callback) | <span data-ttu-id="e64f0-174">方法</span><span class="sxs-lookup"><span data-stu-id="e64f0-174">Method</span></span> |
| [<span data-ttu-id="e64f0-175">addItemAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="e64f0-175">addItemAttachmentAsync</span></span>](#additemattachmentasyncitemid-attachmentname-options-callback) | <span data-ttu-id="e64f0-176">方法</span><span class="sxs-lookup"><span data-stu-id="e64f0-176">Method</span></span> |
| [<span data-ttu-id="e64f0-177">close</span><span class="sxs-lookup"><span data-stu-id="e64f0-177">close</span></span>](#close) | <span data-ttu-id="e64f0-178">方法</span><span class="sxs-lookup"><span data-stu-id="e64f0-178">Method</span></span> |
| [<span data-ttu-id="e64f0-179">displayReplyAllForm</span><span class="sxs-lookup"><span data-stu-id="e64f0-179">displayReplyAllForm</span></span>](#displayreplyallformformdata) | <span data-ttu-id="e64f0-180">方法</span><span class="sxs-lookup"><span data-stu-id="e64f0-180">Method</span></span> |
| [<span data-ttu-id="e64f0-181">displayReplyForm</span><span class="sxs-lookup"><span data-stu-id="e64f0-181">displayReplyForm</span></span>](#displayreplyformformdata) | <span data-ttu-id="e64f0-182">方法</span><span class="sxs-lookup"><span data-stu-id="e64f0-182">Method</span></span> |
| [<span data-ttu-id="e64f0-183">getAttachmentContentAsync</span><span class="sxs-lookup"><span data-stu-id="e64f0-183">getAttachmentContentAsync</span></span>](#getattachmentcontentasyncattachmentid-options-callback--attachmentcontentjavascriptapioutlookofficeattachmentcontent) | <span data-ttu-id="e64f0-184">方法</span><span class="sxs-lookup"><span data-stu-id="e64f0-184">Method</span></span> |
| [<span data-ttu-id="e64f0-185">getAttachmentsAsync</span><span class="sxs-lookup"><span data-stu-id="e64f0-185">getAttachmentsAsync</span></span>](#getattachmentsasyncoptions-callback--arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetails) | <span data-ttu-id="e64f0-186">方法</span><span class="sxs-lookup"><span data-stu-id="e64f0-186">Method</span></span> |
| [<span data-ttu-id="e64f0-187">getEntities</span><span class="sxs-lookup"><span data-stu-id="e64f0-187">getEntities</span></span>](#getentities--entitiesjavascriptapioutlookofficeentities) | <span data-ttu-id="e64f0-188">方法</span><span class="sxs-lookup"><span data-stu-id="e64f0-188">Method</span></span> |
| [<span data-ttu-id="e64f0-189">getEntitiesByType</span><span class="sxs-lookup"><span data-stu-id="e64f0-189">getEntitiesByType</span></span>](#getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlookofficecontactmeetingsuggestionjavascriptapioutlookofficemeetingsuggestionphonenumberjavascriptapioutlookofficephonenumbertasksuggestionjavascriptapioutlookofficetasksuggestion) | <span data-ttu-id="e64f0-190">方法</span><span class="sxs-lookup"><span data-stu-id="e64f0-190">Method</span></span> |
| [<span data-ttu-id="e64f0-191">getFilteredEntitiesByName</span><span class="sxs-lookup"><span data-stu-id="e64f0-191">getFilteredEntitiesByName</span></span>](#getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlookofficecontactmeetingsuggestionjavascriptapioutlookofficemeetingsuggestionphonenumberjavascriptapioutlookofficephonenumbertasksuggestionjavascriptapioutlookofficetasksuggestion) | <span data-ttu-id="e64f0-192">方法</span><span class="sxs-lookup"><span data-stu-id="e64f0-192">Method</span></span> |
| [<span data-ttu-id="e64f0-193">getInitializationContextAsync</span><span class="sxs-lookup"><span data-stu-id="e64f0-193">getInitializationContextAsync</span></span>](#getinitializationcontextasyncoptions-callback) | <span data-ttu-id="e64f0-194">方法</span><span class="sxs-lookup"><span data-stu-id="e64f0-194">Method</span></span> |
| [<span data-ttu-id="e64f0-195">getRegExMatches</span><span class="sxs-lookup"><span data-stu-id="e64f0-195">getRegExMatches</span></span>](#getregexmatches--object) | <span data-ttu-id="e64f0-196">方法</span><span class="sxs-lookup"><span data-stu-id="e64f0-196">Method</span></span> |
| [<span data-ttu-id="e64f0-197">getRegExMatchesByName</span><span class="sxs-lookup"><span data-stu-id="e64f0-197">getRegExMatchesByName</span></span>](#getregexmatchesbynamename--nullable-array-string-) | <span data-ttu-id="e64f0-198">方法</span><span class="sxs-lookup"><span data-stu-id="e64f0-198">Method</span></span> |
| [<span data-ttu-id="e64f0-199">getSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="e64f0-199">getSelectedDataAsync</span></span>](#getselecteddataasynccoerciontype-options-callback--string) | <span data-ttu-id="e64f0-200">方法</span><span class="sxs-lookup"><span data-stu-id="e64f0-200">Method</span></span> |
| [<span data-ttu-id="e64f0-201">getSelectedEntities</span><span class="sxs-lookup"><span data-stu-id="e64f0-201">getSelectedEntities</span></span>](#getselectedentities--entitiesjavascriptapioutlookofficeentities) | <span data-ttu-id="e64f0-202">方法</span><span class="sxs-lookup"><span data-stu-id="e64f0-202">Method</span></span> |
| [<span data-ttu-id="e64f0-203">getSelectedRegExMatches</span><span class="sxs-lookup"><span data-stu-id="e64f0-203">getSelectedRegExMatches</span></span>](#getselectedregexmatches--object) | <span data-ttu-id="e64f0-204">方法</span><span class="sxs-lookup"><span data-stu-id="e64f0-204">Method</span></span> |
| [<span data-ttu-id="e64f0-205">getSharedPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="e64f0-205">getSharedPropertiesAsync</span></span>](#getsharedpropertiesasyncoptions-callback) | <span data-ttu-id="e64f0-206">方法</span><span class="sxs-lookup"><span data-stu-id="e64f0-206">Method</span></span> |
| [<span data-ttu-id="e64f0-207">loadCustomPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="e64f0-207">loadCustomPropertiesAsync</span></span>](#loadcustompropertiesasynccallback-usercontext) | <span data-ttu-id="e64f0-208">方法</span><span class="sxs-lookup"><span data-stu-id="e64f0-208">Method</span></span> |
| [<span data-ttu-id="e64f0-209">removeAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="e64f0-209">removeAttachmentAsync</span></span>](#removeattachmentasyncattachmentid-options-callback) | <span data-ttu-id="e64f0-210">方法</span><span class="sxs-lookup"><span data-stu-id="e64f0-210">Method</span></span> |
| [<span data-ttu-id="e64f0-211">removeHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="e64f0-211">removeHandlerAsync</span></span>](#removehandlerasynceventtype-handler-options-callback) | <span data-ttu-id="e64f0-212">方法</span><span class="sxs-lookup"><span data-stu-id="e64f0-212">Method</span></span> |
| [<span data-ttu-id="e64f0-213">saveAsync</span><span class="sxs-lookup"><span data-stu-id="e64f0-213">saveAsync</span></span>](#saveasyncoptions-callback) | <span data-ttu-id="e64f0-214">方法</span><span class="sxs-lookup"><span data-stu-id="e64f0-214">Method</span></span> |
| [<span data-ttu-id="e64f0-215">setSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="e64f0-215">setSelectedDataAsync</span></span>](#setselecteddataasyncdata-options-callback) | <span data-ttu-id="e64f0-216">方法</span><span class="sxs-lookup"><span data-stu-id="e64f0-216">Method</span></span> |

### <a name="example"></a><span data-ttu-id="e64f0-217">示例</span><span class="sxs-lookup"><span data-stu-id="e64f0-217">Example</span></span>

<span data-ttu-id="e64f0-218">以下 JavaScript 代码示例显示了如何访问 Outlook 中当前项目的 `subject` 属性。</span><span class="sxs-lookup"><span data-stu-id="e64f0-218">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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

### <a name="members"></a><span data-ttu-id="e64f0-219">成员</span><span class="sxs-lookup"><span data-stu-id="e64f0-219">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetails"></a><span data-ttu-id="e64f0-220">attachments :Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="e64f0-220">attachments :Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span></span>

<span data-ttu-id="e64f0-221">获取项目的附件作为数组。</span><span class="sxs-lookup"><span data-stu-id="e64f0-221">Gets the item's attachments as an array.</span></span> <span data-ttu-id="e64f0-222">仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="e64f0-222">Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="e64f0-223">由于潜在的安全问题，某些类型的文件会受到 Outlook 阻止，并且不会返回。</span><span class="sxs-lookup"><span data-stu-id="e64f0-223">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="e64f0-224">如需了解更多信息，请参阅 [Outlook 中阻止的附件](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519)。</span><span class="sxs-lookup"><span data-stu-id="e64f0-224">For more information see [Payments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="e64f0-225">类型：</span><span class="sxs-lookup"><span data-stu-id="e64f0-225">Type:</span></span>

*   <span data-ttu-id="e64f0-226">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="e64f0-226">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span></span>

##### <a name="requirements"></a><span data-ttu-id="e64f0-227">要求</span><span class="sxs-lookup"><span data-stu-id="e64f0-227">Requirements</span></span>

|<span data-ttu-id="e64f0-228">要求</span><span class="sxs-lookup"><span data-stu-id="e64f0-228">Requirement</span></span>|<span data-ttu-id="e64f0-229">值</span><span class="sxs-lookup"><span data-stu-id="e64f0-229">Value</span></span>|
|---|---|
|[<span data-ttu-id="e64f0-230">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="e64f0-230">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e64f0-231">1.0</span><span class="sxs-lookup"><span data-stu-id="e64f0-231">1.0</span></span>|
|[<span data-ttu-id="e64f0-232">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="e64f0-232">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e64f0-233">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e64f0-233">ReadItem</span></span>|
|[<span data-ttu-id="e64f0-234">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="e64f0-234">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="e64f0-235">阅读</span><span class="sxs-lookup"><span data-stu-id="e64f0-235">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="e64f0-236">示例</span><span class="sxs-lookup"><span data-stu-id="e64f0-236">Example</span></span>

<span data-ttu-id="e64f0-237">以下代码使用当前项目上所有附件的详细信息构成 HTML 字符串。</span><span class="sxs-lookup"><span data-stu-id="e64f0-237">The following code builds an HTML string with details of all attachments on the current item.</span></span>

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

####  <a name="bcc-recipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="e64f0-238">bcc :[Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="e64f0-238">bcc :[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="e64f0-239">获取对象，该对象提供用于获取或更新邮件的密件抄送 (Bcc) 行上的收件人的方法。</span><span class="sxs-lookup"><span data-stu-id="e64f0-239">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="e64f0-240">仅限撰写模式。</span><span class="sxs-lookup"><span data-stu-id="e64f0-240">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="e64f0-241">类型：</span><span class="sxs-lookup"><span data-stu-id="e64f0-241">Type:</span></span>

*   [<span data-ttu-id="e64f0-242">收件人</span><span class="sxs-lookup"><span data-stu-id="e64f0-242">Recipients</span></span>](/javascript/api/outlook/office.recipients)

##### <a name="requirements"></a><span data-ttu-id="e64f0-243">要求</span><span class="sxs-lookup"><span data-stu-id="e64f0-243">Requirements</span></span>

|<span data-ttu-id="e64f0-244">要求</span><span class="sxs-lookup"><span data-stu-id="e64f0-244">Requirement</span></span>|<span data-ttu-id="e64f0-245">值</span><span class="sxs-lookup"><span data-stu-id="e64f0-245">Value</span></span>|
|---|---|
|[<span data-ttu-id="e64f0-246">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="e64f0-246">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e64f0-247">1.1</span><span class="sxs-lookup"><span data-stu-id="e64f0-247">1.1</span></span>|
|[<span data-ttu-id="e64f0-248">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="e64f0-248">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e64f0-249">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e64f0-249">ReadItem</span></span>|
|[<span data-ttu-id="e64f0-250">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="e64f0-250">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="e64f0-251">撰写</span><span class="sxs-lookup"><span data-stu-id="e64f0-251">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="e64f0-252">示例</span><span class="sxs-lookup"><span data-stu-id="e64f0-252">Example</span></span>

```javascript
Office.context.mailbox.item.bcc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.bcc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.bcc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfBccRecipients = asyncResult.value;
}
```

####  <a name="body-bodyjavascriptapioutlookofficebody"></a><span data-ttu-id="e64f0-253">body :[Body](/javascript/api/outlook/office.body)</span><span class="sxs-lookup"><span data-stu-id="e64f0-253">body :[Body](/javascript/api/outlook/office.body)</span></span>

<span data-ttu-id="e64f0-254">获取一个提供用于处理项目正文的方法的对象。</span><span class="sxs-lookup"><span data-stu-id="e64f0-254">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="e64f0-255">类型：</span><span class="sxs-lookup"><span data-stu-id="e64f0-255">Type:</span></span>

*   [<span data-ttu-id="e64f0-256">Body</span><span class="sxs-lookup"><span data-stu-id="e64f0-256">Body</span></span>](/javascript/api/outlook/office.body)

##### <a name="requirements"></a><span data-ttu-id="e64f0-257">要求</span><span class="sxs-lookup"><span data-stu-id="e64f0-257">Requirements</span></span>

|<span data-ttu-id="e64f0-258">要求</span><span class="sxs-lookup"><span data-stu-id="e64f0-258">Requirement</span></span>|<span data-ttu-id="e64f0-259">值</span><span class="sxs-lookup"><span data-stu-id="e64f0-259">Value</span></span>|
|---|---|
|[<span data-ttu-id="e64f0-260">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="e64f0-260">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e64f0-261">1.1</span><span class="sxs-lookup"><span data-stu-id="e64f0-261">1.1</span></span>|
|[<span data-ttu-id="e64f0-262">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="e64f0-262">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e64f0-263">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e64f0-263">ReadItem</span></span>|
|[<span data-ttu-id="e64f0-264">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="e64f0-264">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="e64f0-265">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="e64f0-265">Compose or read</span></span>|

####  <a name="cc-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="e64f0-266">cc :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="e64f0-266">cc :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="e64f0-267">提供对邮件的抄送 (Cc) 收件人的访问权限。</span><span class="sxs-lookup"><span data-stu-id="e64f0-267">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="e64f0-268">对象的类型和访问级别取决于当前项目的模式。</span><span class="sxs-lookup"><span data-stu-id="e64f0-268">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="e64f0-269">阅读模式</span><span class="sxs-lookup"><span data-stu-id="e64f0-269">Read mode</span></span>

<span data-ttu-id="e64f0-p106">`cc` 属性返回包含邮件的**抄送**行上所列的每个收件人的 `EmailAddressDetails` 对象的数组。集合上限为 100 个成员。</span><span class="sxs-lookup"><span data-stu-id="e64f0-p106">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="e64f0-272">撰写模式</span><span class="sxs-lookup"><span data-stu-id="e64f0-272">Compose mode</span></span>

<span data-ttu-id="e64f0-273">`cc` 属性返回一个 `Recipients` 对象，该对象提供用于获取或更新邮件的“**抄送**”行上收件人的方法。</span><span class="sxs-lookup"><span data-stu-id="e64f0-273">The `cc` property returns a `Recipients` object that provides methods for manipulating the recipients on the **Cc** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="e64f0-274">类型：</span><span class="sxs-lookup"><span data-stu-id="e64f0-274">Type:</span></span>

*   <span data-ttu-id="e64f0-275">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="e64f0-275">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="e64f0-276">要求</span><span class="sxs-lookup"><span data-stu-id="e64f0-276">Requirements</span></span>

|<span data-ttu-id="e64f0-277">要求</span><span class="sxs-lookup"><span data-stu-id="e64f0-277">Requirement</span></span>|<span data-ttu-id="e64f0-278">值</span><span class="sxs-lookup"><span data-stu-id="e64f0-278">Value</span></span>|
|---|---|
|[<span data-ttu-id="e64f0-279">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="e64f0-279">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e64f0-280">1.0</span><span class="sxs-lookup"><span data-stu-id="e64f0-280">1.0</span></span>|
|[<span data-ttu-id="e64f0-281">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="e64f0-281">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e64f0-282">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e64f0-282">ReadItem</span></span>|
|[<span data-ttu-id="e64f0-283">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="e64f0-283">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="e64f0-284">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="e64f0-284">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="e64f0-285">示例</span><span class="sxs-lookup"><span data-stu-id="e64f0-285">Example</span></span>

```javascript
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

####  <a name="nullable-conversationid-string"></a><span data-ttu-id="e64f0-286">(nullable) conversationId :String</span><span class="sxs-lookup"><span data-stu-id="e64f0-286">(nullable) conversationId :String</span></span>

<span data-ttu-id="e64f0-287">获取包含特定消息的电子邮件会话的标识符。</span><span class="sxs-lookup"><span data-stu-id="e64f0-287">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="e64f0-p107">如果在阅读窗体或撰写窗体的回复中激活邮件应用程序，则此属性可以获得一个整数值。如果用户随后更改了回复邮件的主题（若发送回复），则该邮件的对话 ID 将改变且之前获取的值将不适用。</span><span class="sxs-lookup"><span data-stu-id="e64f0-p107">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="e64f0-p108">对于撰写窗体的新项目，此属性获得一个 null 值。如果用户设置一个主题并保存该项目，`conversationId` 属性将返回一个值。</span><span class="sxs-lookup"><span data-stu-id="e64f0-p108">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="e64f0-292">类型:</span><span class="sxs-lookup"><span data-stu-id="e64f0-292">Type:</span></span>

*   <span data-ttu-id="e64f0-293">String</span><span class="sxs-lookup"><span data-stu-id="e64f0-293">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="e64f0-294">要求</span><span class="sxs-lookup"><span data-stu-id="e64f0-294">Requirements</span></span>

|<span data-ttu-id="e64f0-295">要求</span><span class="sxs-lookup"><span data-stu-id="e64f0-295">Requirement</span></span>|<span data-ttu-id="e64f0-296">值</span><span class="sxs-lookup"><span data-stu-id="e64f0-296">Value</span></span>|
|---|---|
|[<span data-ttu-id="e64f0-297">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="e64f0-297">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e64f0-298">1.0</span><span class="sxs-lookup"><span data-stu-id="e64f0-298">1.0</span></span>|
|[<span data-ttu-id="e64f0-299">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="e64f0-299">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e64f0-300">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e64f0-300">ReadItem</span></span>|
|[<span data-ttu-id="e64f0-301">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="e64f0-301">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="e64f0-302">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="e64f0-302">Compose or read</span></span>|

#### <a name="datetimecreated-date"></a><span data-ttu-id="e64f0-303">dateTimeCreated :Date</span><span class="sxs-lookup"><span data-stu-id="e64f0-303">dateTimeCreated :Date</span></span>

<span data-ttu-id="e64f0-p109">获取项目创建的日期和时间。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="e64f0-p109">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="e64f0-306">类型：</span><span class="sxs-lookup"><span data-stu-id="e64f0-306">Type:</span></span>

*   <span data-ttu-id="e64f0-307">日期</span><span class="sxs-lookup"><span data-stu-id="e64f0-307">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="e64f0-308">要求</span><span class="sxs-lookup"><span data-stu-id="e64f0-308">Requirements</span></span>

|<span data-ttu-id="e64f0-309">要求</span><span class="sxs-lookup"><span data-stu-id="e64f0-309">Requirement</span></span>|<span data-ttu-id="e64f0-310">值</span><span class="sxs-lookup"><span data-stu-id="e64f0-310">Value</span></span>|
|---|---|
|[<span data-ttu-id="e64f0-311">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="e64f0-311">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e64f0-312">1.0</span><span class="sxs-lookup"><span data-stu-id="e64f0-312">1.0</span></span>|
|[<span data-ttu-id="e64f0-313">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="e64f0-313">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e64f0-314">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e64f0-314">ReadItem</span></span>|
|[<span data-ttu-id="e64f0-315">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="e64f0-315">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="e64f0-316">阅读</span><span class="sxs-lookup"><span data-stu-id="e64f0-316">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="e64f0-317">示例</span><span class="sxs-lookup"><span data-stu-id="e64f0-317">Example</span></span>

```javascript
var created = Office.context.mailbox.item.dateTimeCreated;
```

#### <a name="datetimemodified-date"></a><span data-ttu-id="e64f0-318">dateTimeModified :Date</span><span class="sxs-lookup"><span data-stu-id="e64f0-318">dateTimeModified :Date</span></span>

<span data-ttu-id="e64f0-p110">获取项目最近一次修改的日期和时间。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="e64f0-p110">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="e64f0-321">在 Outlook for iOS 或 Outlook for Android 中不支持此成员。</span><span class="sxs-lookup"><span data-stu-id="e64f0-321">Note: This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="type"></a><span data-ttu-id="e64f0-322">类型:</span><span class="sxs-lookup"><span data-stu-id="e64f0-322">Type:</span></span>

*   <span data-ttu-id="e64f0-323">日期</span><span class="sxs-lookup"><span data-stu-id="e64f0-323">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="e64f0-324">要求</span><span class="sxs-lookup"><span data-stu-id="e64f0-324">Requirements</span></span>

|<span data-ttu-id="e64f0-325">要求</span><span class="sxs-lookup"><span data-stu-id="e64f0-325">Requirement</span></span>|<span data-ttu-id="e64f0-326">值</span><span class="sxs-lookup"><span data-stu-id="e64f0-326">Value</span></span>|
|---|---|
|[<span data-ttu-id="e64f0-327">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="e64f0-327">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e64f0-328">1.0</span><span class="sxs-lookup"><span data-stu-id="e64f0-328">1.0</span></span>|
|[<span data-ttu-id="e64f0-329">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="e64f0-329">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e64f0-330">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e64f0-330">ReadItem</span></span>|
|[<span data-ttu-id="e64f0-331">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="e64f0-331">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="e64f0-332">阅读</span><span class="sxs-lookup"><span data-stu-id="e64f0-332">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="e64f0-333">示例</span><span class="sxs-lookup"><span data-stu-id="e64f0-333">Example</span></span>

```javascript
var modified = Office.context.mailbox.item.dateTimeModified;
```

####  <a name="end-datetimejavascriptapioutlookofficetime"></a><span data-ttu-id="e64f0-334">end :Date|[Time](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="e64f0-334">end :Date|[Time](/javascript/api/outlook/office.time)</span></span>

<span data-ttu-id="e64f0-335">获取或设置约会结束的日期和时间。</span><span class="sxs-lookup"><span data-stu-id="e64f0-335">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="e64f0-p111">将 `end` 属性表示为协调世界时 (UTC) 的日期和时间值。可使用 [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttime) 方法将 end 属性值转换为客户端的本地日期和时间。</span><span class="sxs-lookup"><span data-stu-id="e64f0-p111">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="e64f0-338">阅读模式</span><span class="sxs-lookup"><span data-stu-id="e64f0-338">Read mode</span></span>

<span data-ttu-id="e64f0-339">`end` 属性返回 `Date` 对象。</span><span class="sxs-lookup"><span data-stu-id="e64f0-339">The `end` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="e64f0-340">撰写模式</span><span class="sxs-lookup"><span data-stu-id="e64f0-340">Compose mode</span></span>

<span data-ttu-id="e64f0-341">`end` 属性返回 `Time` 对象。</span><span class="sxs-lookup"><span data-stu-id="e64f0-341">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="e64f0-342">使用 [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) 方法设置结束时间时，应使用 [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) 方法将客户端的本地时间转换为服务器的 UTC。</span><span class="sxs-lookup"><span data-stu-id="e64f0-342">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="e64f0-343">类型：</span><span class="sxs-lookup"><span data-stu-id="e64f0-343">Type:</span></span>

*   <span data-ttu-id="e64f0-344">Date | [Time](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="e64f0-344">Date | [Time](/javascript/api/outlook/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="e64f0-345">要求</span><span class="sxs-lookup"><span data-stu-id="e64f0-345">Requirements</span></span>

|<span data-ttu-id="e64f0-346">要求</span><span class="sxs-lookup"><span data-stu-id="e64f0-346">Requirement</span></span>|<span data-ttu-id="e64f0-347">值</span><span class="sxs-lookup"><span data-stu-id="e64f0-347">Value</span></span>|
|---|---|
|[<span data-ttu-id="e64f0-348">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="e64f0-348">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e64f0-349">1.0</span><span class="sxs-lookup"><span data-stu-id="e64f0-349">1.0</span></span>|
|[<span data-ttu-id="e64f0-350">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="e64f0-350">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e64f0-351">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e64f0-351">ReadItem</span></span>|
|[<span data-ttu-id="e64f0-352">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="e64f0-352">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="e64f0-353">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="e64f0-353">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="e64f0-354">示例</span><span class="sxs-lookup"><span data-stu-id="e64f0-354">Example</span></span>

<span data-ttu-id="e64f0-355">以下示例通过使用 `Time` 对象的 [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) 方法，设置撰写模式下约会的结束时间。</span><span class="sxs-lookup"><span data-stu-id="e64f0-355">The following example sets the end time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

#### <a name="from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsfromjavascriptapioutlookofficefrom"></a><span data-ttu-id="e64f0-356">from :[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[From](/javascript/api/outlook/office.from)</span><span class="sxs-lookup"><span data-stu-id="e64f0-356">from :[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[From](/javascript/api/outlook/office.from)</span></span>

<span data-ttu-id="e64f0-357">获取邮件发件人的电子邮件地址。</span><span class="sxs-lookup"><span data-stu-id="e64f0-357">Gets the email address of the sender of a message.</span></span>

<span data-ttu-id="e64f0-p112">`from` 和 [`sender`](#sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetails) 属性表示同一个人，邮件由代理人发送的除外。在此情况下，`from` 属性表示代理程序，而 sender 属性表示代理人。</span><span class="sxs-lookup"><span data-stu-id="e64f0-p112">The `from` and [`sender`](#sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="e64f0-360">`from` 属性中 `EmailAddressDetails` 对象的 `recipientType` 属性为 `undefined`。</span><span class="sxs-lookup"><span data-stu-id="e64f0-360">Note: The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="e64f0-361">阅读模式</span><span class="sxs-lookup"><span data-stu-id="e64f0-361">Read mode</span></span>

<span data-ttu-id="e64f0-362">`from` 属性返回一个 `EmailAddressDetails` 对象。</span><span class="sxs-lookup"><span data-stu-id="e64f0-362">The AssignToCategory`from` property always returns an AssignToCategoryRuleAction`EmailAddressDetails` object.</span></span>

```javascript
var subject = Office.context.mailbox.item.from;
```

##### <a name="compose-mode"></a><span data-ttu-id="e64f0-363">撰写模式</span><span class="sxs-lookup"><span data-stu-id="e64f0-363">Compose mode</span></span>

<span data-ttu-id="e64f0-364">`from` 属性返回一个 `From` 对象，该对象提供从值中进行获取的方法。</span><span class="sxs-lookup"><span data-stu-id="e64f0-364">The `from` property returns a `From` object that provides a method to get the from value.</span></span>

```javascript
Office.context.mailbox.item.from.getAsync(callback);

function callback(asyncResult) {
  var from = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="e64f0-365">类型：</span><span class="sxs-lookup"><span data-stu-id="e64f0-365">Type:</span></span>

*   <span data-ttu-id="e64f0-366">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [From](/javascript/api/outlook/office.from)</span><span class="sxs-lookup"><span data-stu-id="e64f0-366">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [From](/javascript/api/outlook/office.from)</span></span>

##### <a name="requirements"></a><span data-ttu-id="e64f0-367">要求</span><span class="sxs-lookup"><span data-stu-id="e64f0-367">Requirements</span></span>

|<span data-ttu-id="e64f0-368">要求</span><span class="sxs-lookup"><span data-stu-id="e64f0-368">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="e64f0-369">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="e64f0-369">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e64f0-370">1.0</span><span class="sxs-lookup"><span data-stu-id="e64f0-370">1.0</span></span>|<span data-ttu-id="e64f0-371">1.7</span><span class="sxs-lookup"><span data-stu-id="e64f0-371">-17</span></span>|
|[<span data-ttu-id="e64f0-372">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="e64f0-372">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e64f0-373">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e64f0-373">ReadItem</span></span>|<span data-ttu-id="e64f0-374">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="e64f0-374">ReadWriteItem</span></span>|
|[<span data-ttu-id="e64f0-375">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="e64f0-375">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="e64f0-376">阅读</span><span class="sxs-lookup"><span data-stu-id="e64f0-376">Read</span></span>|<span data-ttu-id="e64f0-377">撰写</span><span class="sxs-lookup"><span data-stu-id="e64f0-377">Compose</span></span>|

#### <a name="internetheaders-internetheadersjavascriptapioutlookofficeinternetheaders"></a><span data-ttu-id="e64f0-378">internetHeaders :[InternetHeaders](/javascript/api/outlook/office.internetheaders)</span><span class="sxs-lookup"><span data-stu-id="e64f0-378">internetHeaders :[InternetHeaders](/javascript/api/outlook/office.internetheaders)</span></span>

<span data-ttu-id="e64f0-379">获取或设置消息的 Internet 标头。</span><span class="sxs-lookup"><span data-stu-id="e64f0-379">Gets or sets the internet headers of a message.</span></span>

##### <a name="type"></a><span data-ttu-id="e64f0-380">类型：</span><span class="sxs-lookup"><span data-stu-id="e64f0-380">Type:</span></span>

*   [<span data-ttu-id="e64f0-381">InternetHeaders</span><span class="sxs-lookup"><span data-stu-id="e64f0-381">InternetHeaders</span></span>](/javascript/api/outlook/office.internetheaders)

##### <a name="requirements"></a><span data-ttu-id="e64f0-382">要求</span><span class="sxs-lookup"><span data-stu-id="e64f0-382">Requirements</span></span>

|<span data-ttu-id="e64f0-383">要求</span><span class="sxs-lookup"><span data-stu-id="e64f0-383">Requirement</span></span>|<span data-ttu-id="e64f0-384">值</span><span class="sxs-lookup"><span data-stu-id="e64f0-384">Value</span></span>|
|---|---|
|[<span data-ttu-id="e64f0-385">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="e64f0-385">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e64f0-386">预览</span><span class="sxs-lookup"><span data-stu-id="e64f0-386">Preview</span></span>|
|[<span data-ttu-id="e64f0-387">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="e64f0-387">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e64f0-388">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e64f0-388">ReadItem</span></span>|
|[<span data-ttu-id="e64f0-389">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="e64f0-389">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="e64f0-390">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="e64f0-390">Compose or read</span></span>|

#### <a name="internetmessageid-string"></a><span data-ttu-id="e64f0-391">internetMessageId :String</span><span class="sxs-lookup"><span data-stu-id="e64f0-391">internetMessageId :String</span></span>

<span data-ttu-id="e64f0-p113">获取电子邮件的 Internet 消息标识符。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="e64f0-p113">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="e64f0-394">类型：</span><span class="sxs-lookup"><span data-stu-id="e64f0-394">Type:</span></span>

*   <span data-ttu-id="e64f0-395">String</span><span class="sxs-lookup"><span data-stu-id="e64f0-395">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="e64f0-396">要求</span><span class="sxs-lookup"><span data-stu-id="e64f0-396">Requirements</span></span>

|<span data-ttu-id="e64f0-397">要求</span><span class="sxs-lookup"><span data-stu-id="e64f0-397">Requirement</span></span>|<span data-ttu-id="e64f0-398">值</span><span class="sxs-lookup"><span data-stu-id="e64f0-398">Value</span></span>|
|---|---|
|[<span data-ttu-id="e64f0-399">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="e64f0-399">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e64f0-400">1.0</span><span class="sxs-lookup"><span data-stu-id="e64f0-400">1.0</span></span>|
|[<span data-ttu-id="e64f0-401">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="e64f0-401">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e64f0-402">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e64f0-402">ReadItem</span></span>|
|[<span data-ttu-id="e64f0-403">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="e64f0-403">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="e64f0-404">阅读</span><span class="sxs-lookup"><span data-stu-id="e64f0-404">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="e64f0-405">示例</span><span class="sxs-lookup"><span data-stu-id="e64f0-405">Example</span></span>

```javascript
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

#### <a name="itemclass-string"></a><span data-ttu-id="e64f0-406">itemClass :String</span><span class="sxs-lookup"><span data-stu-id="e64f0-406">itemClass :String</span></span>

<span data-ttu-id="e64f0-p114">获取选定项目的 Exchange Web 服务项目类。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="e64f0-p114">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="e64f0-p115">`itemClass` 属性指定所选项目的邮件类别。以下是邮件或约会项目的默认邮件类别。</span><span class="sxs-lookup"><span data-stu-id="e64f0-p115">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

|<span data-ttu-id="e64f0-411">类型</span><span class="sxs-lookup"><span data-stu-id="e64f0-411">Type</span></span>|<span data-ttu-id="e64f0-412">说明</span><span class="sxs-lookup"><span data-stu-id="e64f0-412">Description</span></span>|<span data-ttu-id="e64f0-413">项目类</span><span class="sxs-lookup"><span data-stu-id="e64f0-413">item class</span></span>|
|---|---|---|
|<span data-ttu-id="e64f0-414">约会项目</span><span class="sxs-lookup"><span data-stu-id="e64f0-414">Appointment items</span></span>|<span data-ttu-id="e64f0-415">这些是项目类 `IPM.Appointment` 或 `IPM.Appointment.Occurence` 的日历项目。</span><span class="sxs-lookup"><span data-stu-id="e64f0-415">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurence`.</span></span>|`IPM.Appointment`<br />`IPM.Appointment.Occurence`|
|<span data-ttu-id="e64f0-416">邮件项目</span><span class="sxs-lookup"><span data-stu-id="e64f0-416">Message items</span></span>|<span data-ttu-id="e64f0-417">这些项目包括具有默认邮件类别 `IPM.Note` 的电子邮件，以及使用 `IPM.Schedule.Meeting` 作为基础邮件类别的会议请求、响应和取消。</span><span class="sxs-lookup"><span data-stu-id="e64f0-417">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span>|`IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled`|

<span data-ttu-id="e64f0-418">你可以创建用于扩展默认邮件类别的自定义邮件类别，例如，自定义约会邮件类别 `IPM.Appointment.Contoso`。</span><span class="sxs-lookup"><span data-stu-id="e64f0-418">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="e64f0-419">类型：</span><span class="sxs-lookup"><span data-stu-id="e64f0-419">Type:</span></span>

*   <span data-ttu-id="e64f0-420">String</span><span class="sxs-lookup"><span data-stu-id="e64f0-420">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="e64f0-421">要求</span><span class="sxs-lookup"><span data-stu-id="e64f0-421">Requirements</span></span>

|<span data-ttu-id="e64f0-422">要求</span><span class="sxs-lookup"><span data-stu-id="e64f0-422">Requirement</span></span>|<span data-ttu-id="e64f0-423">值</span><span class="sxs-lookup"><span data-stu-id="e64f0-423">Value</span></span>|
|---|---|
|[<span data-ttu-id="e64f0-424">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="e64f0-424">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e64f0-425">1.0</span><span class="sxs-lookup"><span data-stu-id="e64f0-425">1.0</span></span>|
|[<span data-ttu-id="e64f0-426">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="e64f0-426">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e64f0-427">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e64f0-427">ReadItem</span></span>|
|[<span data-ttu-id="e64f0-428">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="e64f0-428">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="e64f0-429">阅读</span><span class="sxs-lookup"><span data-stu-id="e64f0-429">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="e64f0-430">示例</span><span class="sxs-lookup"><span data-stu-id="e64f0-430">Example</span></span>

```javascript
var itemClass = Office.context.mailbox.item.itemClass;
```

#### <a name="nullable-itemid-string"></a><span data-ttu-id="e64f0-431">(nullable) itemId :String</span><span class="sxs-lookup"><span data-stu-id="e64f0-431">(nullable) itemId :String</span></span>

<span data-ttu-id="e64f0-p116">获取当前项目的 Exchange Web 服务项目标识符。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="e64f0-p116">Gets the Exchange Web Services item identifier for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="e64f0-434">`itemId` 属性返回的标识符与 Exchange Web 服务项目标识符相同。</span><span class="sxs-lookup"><span data-stu-id="e64f0-434">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="e64f0-435">`itemId` 属性与 Outlook 条目 ID 或 Outlook REST API 使用的 ID 不同。</span><span class="sxs-lookup"><span data-stu-id="e64f0-435">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="e64f0-436">使用此值进行 REST API 调用前，应使用 [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string) 对它进行转换。</span><span class="sxs-lookup"><span data-stu-id="e64f0-436">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="e64f0-437">有关详细信息，请参阅[从 Outlook 加载项使用 Outlook REST API](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id)。</span><span class="sxs-lookup"><span data-stu-id="e64f0-437">For more details, see [Use the Outlook REST APIs from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

<span data-ttu-id="e64f0-p118">`itemId` 属性在撰写模式下不可用。如果需要项目标识符，[`saveAsync`](#saveasyncoptions-callback) 方法可用于将项目保存到存储，这将在回调函数的 [`AsyncResult.value`](/javascript/api/office/office.asyncresult) 参数中返回项目标识符。</span><span class="sxs-lookup"><span data-stu-id="e64f0-p118">The `itemId` property is not available in compose mode. If an item identifier is required, the [`saveAsync`](#saveasyncoptions-callback) method can be used to save the item to the store, which will return the item identifier in the [`AsyncResult.value`](/javascript/api/office/office.asyncresult) parameter in the callback function.</span></span>

##### <a name="type"></a><span data-ttu-id="e64f0-440">类型:</span><span class="sxs-lookup"><span data-stu-id="e64f0-440">Type:</span></span>

*   <span data-ttu-id="e64f0-441">String</span><span class="sxs-lookup"><span data-stu-id="e64f0-441">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="e64f0-442">要求</span><span class="sxs-lookup"><span data-stu-id="e64f0-442">Requirements</span></span>

|<span data-ttu-id="e64f0-443">要求</span><span class="sxs-lookup"><span data-stu-id="e64f0-443">Requirement</span></span>|<span data-ttu-id="e64f0-444">值</span><span class="sxs-lookup"><span data-stu-id="e64f0-444">Value</span></span>|
|---|---|
|[<span data-ttu-id="e64f0-445">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="e64f0-445">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e64f0-446">1.0</span><span class="sxs-lookup"><span data-stu-id="e64f0-446">1.0</span></span>|
|[<span data-ttu-id="e64f0-447">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="e64f0-447">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e64f0-448">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e64f0-448">ReadItem</span></span>|
|[<span data-ttu-id="e64f0-449">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="e64f0-449">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="e64f0-450">阅读</span><span class="sxs-lookup"><span data-stu-id="e64f0-450">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="e64f0-451">示例</span><span class="sxs-lookup"><span data-stu-id="e64f0-451">Example</span></span>

<span data-ttu-id="e64f0-p119">以下代码检查项目标识符是否存在。如果 `itemId` 属性返回 `null` 或 `undefined`，则将项目保存到存储，并从异步结果中获取项目标识符。</span><span class="sxs-lookup"><span data-stu-id="e64f0-p119">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

```javascript
var itemId = Office.context.mailbox.item.itemId;
if (itemId === null || itemId == undefined) {
  Office.context.mailbox.item.saveAsync(function(result){
    itemId = result.value;
  });
}
```

####  <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtype"></a><span data-ttu-id="e64f0-454">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype)</span><span class="sxs-lookup"><span data-stu-id="e64f0-454">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype)</span></span>

<span data-ttu-id="e64f0-455">获取实例表示的项的类型。</span><span class="sxs-lookup"><span data-stu-id="e64f0-455">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="e64f0-456">`itemType` 属性返回其中一个 `ItemType` 枚举值，指示 `item` 对象实例是邮件还是约会。</span><span class="sxs-lookup"><span data-stu-id="e64f0-456">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="e64f0-457">类型：</span><span class="sxs-lookup"><span data-stu-id="e64f0-457">Type:</span></span>

*   [<span data-ttu-id="e64f0-458">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="e64f0-458">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook/office.mailboxenums.itemtype)

##### <a name="requirements"></a><span data-ttu-id="e64f0-459">要求</span><span class="sxs-lookup"><span data-stu-id="e64f0-459">Requirements</span></span>

|<span data-ttu-id="e64f0-460">要求</span><span class="sxs-lookup"><span data-stu-id="e64f0-460">Requirement</span></span>|<span data-ttu-id="e64f0-461">值</span><span class="sxs-lookup"><span data-stu-id="e64f0-461">Value</span></span>|
|---|---|
|[<span data-ttu-id="e64f0-462">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="e64f0-462">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e64f0-463">1.0</span><span class="sxs-lookup"><span data-stu-id="e64f0-463">1.0</span></span>|
|[<span data-ttu-id="e64f0-464">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="e64f0-464">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e64f0-465">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e64f0-465">ReadItem</span></span>|
|[<span data-ttu-id="e64f0-466">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="e64f0-466">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="e64f0-467">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="e64f0-467">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="e64f0-468">示例</span><span class="sxs-lookup"><span data-stu-id="e64f0-468">Example</span></span>

```javascript
if (Office.context.mailbox.item.itemType == Office.MailboxEnums.ItemType.Message)
  // do something
else
  // do something else
```

####  <a name="location-stringlocationjavascriptapioutlookofficelocation"></a><span data-ttu-id="e64f0-469">location :String|[Location](/javascript/api/outlook/office.location)</span><span class="sxs-lookup"><span data-stu-id="e64f0-469">location :String|[Location](/javascript/api/outlook/office.location)</span></span>

<span data-ttu-id="e64f0-470">获取或设置约会的位置。</span><span class="sxs-lookup"><span data-stu-id="e64f0-470">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="e64f0-471">阅读模式</span><span class="sxs-lookup"><span data-stu-id="e64f0-471">Read mode</span></span>

<span data-ttu-id="e64f0-472">`location` 属性返回一个包含约会位置的字符串。</span><span class="sxs-lookup"><span data-stu-id="e64f0-472">The `location` property returns a string that contains the location of the appointment.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="e64f0-473">撰写模式</span><span class="sxs-lookup"><span data-stu-id="e64f0-473">Compose mode</span></span>

<span data-ttu-id="e64f0-474">`location` 属性返回一个 `Location` 对象，该对象提供用于获取和设置约会位置的方法。</span><span class="sxs-lookup"><span data-stu-id="e64f0-474">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="e64f0-475">类型：</span><span class="sxs-lookup"><span data-stu-id="e64f0-475">Type:</span></span>

*   <span data-ttu-id="e64f0-476">String | [Location](/javascript/api/outlook/office.location)</span><span class="sxs-lookup"><span data-stu-id="e64f0-476">String | [Location](/javascript/api/outlook/office.location)</span></span>

##### <a name="requirements"></a><span data-ttu-id="e64f0-477">要求</span><span class="sxs-lookup"><span data-stu-id="e64f0-477">Requirements</span></span>

|<span data-ttu-id="e64f0-478">要求</span><span class="sxs-lookup"><span data-stu-id="e64f0-478">Requirement</span></span>|<span data-ttu-id="e64f0-479">值</span><span class="sxs-lookup"><span data-stu-id="e64f0-479">Value</span></span>|
|---|---|
|[<span data-ttu-id="e64f0-480">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="e64f0-480">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e64f0-481">1.0</span><span class="sxs-lookup"><span data-stu-id="e64f0-481">1.0</span></span>|
|[<span data-ttu-id="e64f0-482">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="e64f0-482">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e64f0-483">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e64f0-483">ReadItem</span></span>|
|[<span data-ttu-id="e64f0-484">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="e64f0-484">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="e64f0-485">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="e64f0-485">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="e64f0-486">示例</span><span class="sxs-lookup"><span data-stu-id="e64f0-486">Example</span></span>

```javascript
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

#### <a name="normalizedsubject-string"></a><span data-ttu-id="e64f0-487">normalizedSubject :String</span><span class="sxs-lookup"><span data-stu-id="e64f0-487">normalizedSubject :String</span></span>

<span data-ttu-id="e64f0-p120">获取删除了所有前缀（包括 `RE:` 和 `FWD:`）的项目主题。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="e64f0-p120">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="e64f0-p121">normalizedSubject 属性获取包含由电子邮件程序添加的任何标准前缀（如 `RE:` 和 `FW:`）的项目主题。若要获取包含完整前缀的项目主题，请使用 [`subject`](#subject-stringsubjectjavascriptapioutlookofficesubject) 属性。</span><span class="sxs-lookup"><span data-stu-id="e64f0-p121">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubjectjavascriptapioutlookofficesubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="e64f0-492">类型:</span><span class="sxs-lookup"><span data-stu-id="e64f0-492">Type:</span></span>

*   <span data-ttu-id="e64f0-493">String</span><span class="sxs-lookup"><span data-stu-id="e64f0-493">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="e64f0-494">要求</span><span class="sxs-lookup"><span data-stu-id="e64f0-494">Requirements</span></span>

|<span data-ttu-id="e64f0-495">要求</span><span class="sxs-lookup"><span data-stu-id="e64f0-495">Requirement</span></span>|<span data-ttu-id="e64f0-496">值</span><span class="sxs-lookup"><span data-stu-id="e64f0-496">Value</span></span>|
|---|---|
|[<span data-ttu-id="e64f0-497">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="e64f0-497">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e64f0-498">1.0</span><span class="sxs-lookup"><span data-stu-id="e64f0-498">1.0</span></span>|
|[<span data-ttu-id="e64f0-499">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="e64f0-499">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e64f0-500">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e64f0-500">ReadItem</span></span>|
|[<span data-ttu-id="e64f0-501">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="e64f0-501">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="e64f0-502">阅读</span><span class="sxs-lookup"><span data-stu-id="e64f0-502">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="e64f0-503">示例</span><span class="sxs-lookup"><span data-stu-id="e64f0-503">Example</span></span>

```javascript
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
```

####  <a name="notificationmessages-notificationmessagesjavascriptapioutlookofficenotificationmessages"></a><span data-ttu-id="e64f0-504">notificationMessages :[NotificationMessages](/javascript/api/outlook/office.notificationmessages)</span><span class="sxs-lookup"><span data-stu-id="e64f0-504">notificationMessages :[NotificationMessages](/javascript/api/outlook/office.notificationmessages)</span></span>

<span data-ttu-id="e64f0-505">获取项目的通知邮件。</span><span class="sxs-lookup"><span data-stu-id="e64f0-505">Gets the notification messages for an item.</span></span>

##### <a name="type"></a><span data-ttu-id="e64f0-506">类型：</span><span class="sxs-lookup"><span data-stu-id="e64f0-506">Type:</span></span>

*   [<span data-ttu-id="e64f0-507">NotificationMessages</span><span class="sxs-lookup"><span data-stu-id="e64f0-507">NotificationMessages</span></span>](/javascript/api/outlook/office.notificationmessages)

##### <a name="requirements"></a><span data-ttu-id="e64f0-508">要求</span><span class="sxs-lookup"><span data-stu-id="e64f0-508">Requirements</span></span>

|<span data-ttu-id="e64f0-509">要求</span><span class="sxs-lookup"><span data-stu-id="e64f0-509">Requirement</span></span>|<span data-ttu-id="e64f0-510">值</span><span class="sxs-lookup"><span data-stu-id="e64f0-510">Value</span></span>|
|---|---|
|[<span data-ttu-id="e64f0-511">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="e64f0-511">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e64f0-512">1.3</span><span class="sxs-lookup"><span data-stu-id="e64f0-512">1.3</span></span>|
|[<span data-ttu-id="e64f0-513">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="e64f0-513">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e64f0-514">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e64f0-514">ReadItem</span></span>|
|[<span data-ttu-id="e64f0-515">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="e64f0-515">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="e64f0-516">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="e64f0-516">Compose or read</span></span>|

####  <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="e64f0-517">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="e64f0-517">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="e64f0-518">提供对事件的可选与会者的访问权限。</span><span class="sxs-lookup"><span data-stu-id="e64f0-518">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="e64f0-519">对象的类型和访问级别取决于当前项目的模式。</span><span class="sxs-lookup"><span data-stu-id="e64f0-519">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="e64f0-520">阅读模式</span><span class="sxs-lookup"><span data-stu-id="e64f0-520">Read mode</span></span>

<span data-ttu-id="e64f0-521">`optionalAttendees` 属性返回一个数组，其中包含每个可选与会者的 `EmailAddressDetails` 对象。</span><span class="sxs-lookup"><span data-stu-id="e64f0-521">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="e64f0-522">撰写模式</span><span class="sxs-lookup"><span data-stu-id="e64f0-522">Compose mode</span></span>

<span data-ttu-id="e64f0-523">`optionalAttendees` 属性返回一个 `Recipients` 对象，该对象提供用于获取或更新可选与会者的方法。</span><span class="sxs-lookup"><span data-stu-id="e64f0-523">The `optionalAttendees` property returns a `Recipients` object that provides methods to get and set the optional attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="e64f0-524">类型：</span><span class="sxs-lookup"><span data-stu-id="e64f0-524">Type:</span></span>

*   <span data-ttu-id="e64f0-525">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="e64f0-525">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="e64f0-526">要求</span><span class="sxs-lookup"><span data-stu-id="e64f0-526">Requirements</span></span>

|<span data-ttu-id="e64f0-527">要求</span><span class="sxs-lookup"><span data-stu-id="e64f0-527">Requirement</span></span>|<span data-ttu-id="e64f0-528">值</span><span class="sxs-lookup"><span data-stu-id="e64f0-528">Value</span></span>|
|---|---|
|[<span data-ttu-id="e64f0-529">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="e64f0-529">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e64f0-530">1.0</span><span class="sxs-lookup"><span data-stu-id="e64f0-530">1.0</span></span>|
|[<span data-ttu-id="e64f0-531">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="e64f0-531">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e64f0-532">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e64f0-532">ReadItem</span></span>|
|[<span data-ttu-id="e64f0-533">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="e64f0-533">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="e64f0-534">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="e64f0-534">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="e64f0-535">示例</span><span class="sxs-lookup"><span data-stu-id="e64f0-535">Example</span></span>

```javascript
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

#### <a name="organizer-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsorganizerjavascriptapioutlookofficeorganizer"></a><span data-ttu-id="e64f0-536">organizer :[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[Organizer](/javascript/api/outlook/office.organizer)</span><span class="sxs-lookup"><span data-stu-id="e64f0-536">organizer :[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[Organizer](/javascript/api/outlook/office.organizer)</span></span>

<span data-ttu-id="e64f0-537">获取指定会议的组织者的电子邮件地址。</span><span class="sxs-lookup"><span data-stu-id="e64f0-537">Gets the email address of the meeting organizer for a specified meeting.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="e64f0-538">阅读模式</span><span class="sxs-lookup"><span data-stu-id="e64f0-538">Read mode</span></span>

<span data-ttu-id="e64f0-539">`organizer` 属性返回 [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) 对象，它表示会议组织者。</span><span class="sxs-lookup"><span data-stu-id="e64f0-539">The `organizer` property returns an [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) object that represents the meeting organizer.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="e64f0-540">撰写模式</span><span class="sxs-lookup"><span data-stu-id="e64f0-540">Compose mode</span></span>

<span data-ttu-id="e64f0-541">`organizer` 属性返回 [Organizer](/javascript/api/outlook/office.organizer) 对象，该对象提供获取组织者值的方法。</span><span class="sxs-lookup"><span data-stu-id="e64f0-541">The `organizer` property returns an [Organizer](/javascript/api/outlook/office.organizer) object that provides a method to get the organizer value.</span></span>

##### <a name="type"></a><span data-ttu-id="e64f0-542">类型：</span><span class="sxs-lookup"><span data-stu-id="e64f0-542">Type:</span></span>

*   <span data-ttu-id="e64f0-543">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [Organizer](/javascript/api/outlook/office.organizer)</span><span class="sxs-lookup"><span data-stu-id="e64f0-543">[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [Organizer](/javascript/api/outlook/office.organizer)</span></span>

##### <a name="requirements"></a><span data-ttu-id="e64f0-544">要求</span><span class="sxs-lookup"><span data-stu-id="e64f0-544">Requirements</span></span>

|<span data-ttu-id="e64f0-545">要求</span><span class="sxs-lookup"><span data-stu-id="e64f0-545">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="e64f0-546">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="e64f0-546">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e64f0-547">1.0</span><span class="sxs-lookup"><span data-stu-id="e64f0-547">1.0</span></span>|<span data-ttu-id="e64f0-548">1.7</span><span class="sxs-lookup"><span data-stu-id="e64f0-548">-17</span></span>|
|[<span data-ttu-id="e64f0-549">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="e64f0-549">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e64f0-550">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e64f0-550">ReadItem</span></span>|<span data-ttu-id="e64f0-551">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="e64f0-551">ReadWriteItem</span></span>|
|[<span data-ttu-id="e64f0-552">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="e64f0-552">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="e64f0-553">阅读</span><span class="sxs-lookup"><span data-stu-id="e64f0-553">Read</span></span>|<span data-ttu-id="e64f0-554">撰写</span><span class="sxs-lookup"><span data-stu-id="e64f0-554">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="e64f0-555">示例</span><span class="sxs-lookup"><span data-stu-id="e64f0-555">Example</span></span>

```javascript
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
```

#### <a name="nullable-recurrence-recurrencejavascriptapioutlookofficerecurrence"></a><span data-ttu-id="e64f0-556">(nullable) recurrence :[Recurrence](/javascript/api/outlook/office.recurrence)</span><span class="sxs-lookup"><span data-stu-id="e64f0-556">(nullable) recurrence :[Recurrence](/javascript/api/outlook/office.recurrence)</span></span>

<span data-ttu-id="e64f0-557">获取或设置约会的定期模式。</span><span class="sxs-lookup"><span data-stu-id="e64f0-557">Gets or sets the location of an appointment.</span></span> <span data-ttu-id="e64f0-558">获取或设置会议请求的定期模式。</span><span class="sxs-lookup"><span data-stu-id="e64f0-558">Gets the recurrence pattern of a meeting request.</span></span> <span data-ttu-id="e64f0-559">阅读撰写约会项目的模式。</span><span class="sxs-lookup"><span data-stu-id="e64f0-559">Read and compose modes for appointment items.</span></span> <span data-ttu-id="e64f0-560">会议请求项目的阅读模式。</span><span class="sxs-lookup"><span data-stu-id="e64f0-560">Read mode for meeting request items.</span></span>

<span data-ttu-id="e64f0-561">如果项目是一个系列或系列中的一个实例，则 `recurrence` 属性将返回定期约会的 [recurrence](/javascript/api/outlook/office.recurrence) 对象或会议请求。</span><span class="sxs-lookup"><span data-stu-id="e64f0-561">The `recurrence` property returns a [recurrence](/javascript/api/outlook/office.recurrence) object for recurring appointments or meetings requests if an item is a series or an instance in a series.</span></span> <span data-ttu-id="e64f0-562">针对单个约会和单个约会的会议请求返回 `null`。</span><span class="sxs-lookup"><span data-stu-id="e64f0-562">`null` is returned for single appointments and meeting requests of single appointments.</span></span> <span data-ttu-id="e64f0-563">针对非会议请求的邮件返回 `undefined`。</span><span class="sxs-lookup"><span data-stu-id="e64f0-563">`undefined` is returned for messages that are not meeting requests.</span></span>

> <span data-ttu-id="e64f0-564">注意：会议请求的 `itemClass` 值为 IPM.Schedule.Meeting.Request。</span><span class="sxs-lookup"><span data-stu-id="e64f0-564">Note: Meeting requests have an `itemClass` value of IPM.Schedule.Meeting.Request.</span></span>

> <span data-ttu-id="e64f0-565">注意：如果 recurrence 对象为 `null`，则这表示对象是单个约会或单个约会的会议请求，而不是系列的一部分。</span><span class="sxs-lookup"><span data-stu-id="e64f0-565">Note: If the recurrence object is `null`, this indicates that the object is a single appointment or a meeting request of a single appointment and NOT a part of a series.</span></span>

##### <a name="type"></a><span data-ttu-id="e64f0-566">类型：</span><span class="sxs-lookup"><span data-stu-id="e64f0-566">Type:</span></span>

* [<span data-ttu-id="e64f0-567">Recurrence</span><span class="sxs-lookup"><span data-stu-id="e64f0-567">recurrence</span></span>](/javascript/api/outlook/office.recurrence)

|<span data-ttu-id="e64f0-568">要求</span><span class="sxs-lookup"><span data-stu-id="e64f0-568">Requirement</span></span>|<span data-ttu-id="e64f0-569">值</span><span class="sxs-lookup"><span data-stu-id="e64f0-569">Value</span></span>|
|---|---|
|[<span data-ttu-id="e64f0-570">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="e64f0-570">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e64f0-571">1.7</span><span class="sxs-lookup"><span data-stu-id="e64f0-571">-17</span></span>|
|[<span data-ttu-id="e64f0-572">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="e64f0-572">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e64f0-573">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e64f0-573">ReadItem</span></span>|
|[<span data-ttu-id="e64f0-574">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="e64f0-574">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="e64f0-575">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="e64f0-575">Compose or read</span></span>|

####  <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="e64f0-576">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="e64f0-576">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="e64f0-577">提供对事件的必需与会者的访问权限。</span><span class="sxs-lookup"><span data-stu-id="e64f0-577">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="e64f0-578">对象的类型和访问级别取决于当前项目的模式。</span><span class="sxs-lookup"><span data-stu-id="e64f0-578">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="e64f0-579">阅读模式</span><span class="sxs-lookup"><span data-stu-id="e64f0-579">Read mode</span></span>

<span data-ttu-id="e64f0-580">`requiredAttendees` 属性返回一个数组，其中包含每个必需与会者的 `EmailAddressDetails` 对象。</span><span class="sxs-lookup"><span data-stu-id="e64f0-580">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="e64f0-581">撰写模式</span><span class="sxs-lookup"><span data-stu-id="e64f0-581">Compose mode</span></span>

<span data-ttu-id="e64f0-582">`requiredAttendees` 属性返回一个 `Recipients` 对象，该对象提供用于获取或更新必需与会者的方法。</span><span class="sxs-lookup"><span data-stu-id="e64f0-582">The `requiredAttendees` property returns a `Recipients` object that provides methods to get and set the required attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="e64f0-583">类型：</span><span class="sxs-lookup"><span data-stu-id="e64f0-583">Type:</span></span>

*   <span data-ttu-id="e64f0-584">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="e64f0-584">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="e64f0-585">要求</span><span class="sxs-lookup"><span data-stu-id="e64f0-585">Requirements</span></span>

|<span data-ttu-id="e64f0-586">要求</span><span class="sxs-lookup"><span data-stu-id="e64f0-586">Requirement</span></span>|<span data-ttu-id="e64f0-587">值</span><span class="sxs-lookup"><span data-stu-id="e64f0-587">Value</span></span>|
|---|---|
|[<span data-ttu-id="e64f0-588">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="e64f0-588">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e64f0-589">1.0</span><span class="sxs-lookup"><span data-stu-id="e64f0-589">1.0</span></span>|
|[<span data-ttu-id="e64f0-590">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="e64f0-590">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e64f0-591">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e64f0-591">ReadItem</span></span>|
|[<span data-ttu-id="e64f0-592">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="e64f0-592">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="e64f0-593">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="e64f0-593">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="e64f0-594">示例</span><span class="sxs-lookup"><span data-stu-id="e64f0-594">Example</span></span>

```javascript
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
}
```

#### <a name="sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetails"></a><span data-ttu-id="e64f0-595">sender :[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="e64f0-595">sender :[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)</span></span>

<span data-ttu-id="e64f0-p126">获取电子邮件发件人的电子邮件地址。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="e64f0-p126">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="e64f0-p127">[`from`](#from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsfromjavascriptapioutlookofficefrom) 和 `sender` 属性表示同一个人，邮件由代理人发送的除外。在此情况下，`from` 属性表示代理程序，而 sender 属性表示代理人。</span><span class="sxs-lookup"><span data-stu-id="e64f0-p127">The [`from`](#from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsfromjavascriptapioutlookofficefrom) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="e64f0-600">`sender` 属性中 `EmailAddressDetails` 对象的 `recipientType` 属性为 `undefined`。</span><span class="sxs-lookup"><span data-stu-id="e64f0-600">Note: The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="e64f0-601">类型：</span><span class="sxs-lookup"><span data-stu-id="e64f0-601">Type:</span></span>

*   [<span data-ttu-id="e64f0-602">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="e64f0-602">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="e64f0-603">要求</span><span class="sxs-lookup"><span data-stu-id="e64f0-603">Requirements</span></span>

|<span data-ttu-id="e64f0-604">要求</span><span class="sxs-lookup"><span data-stu-id="e64f0-604">Requirement</span></span>|<span data-ttu-id="e64f0-605">值</span><span class="sxs-lookup"><span data-stu-id="e64f0-605">Value</span></span>|
|---|---|
|[<span data-ttu-id="e64f0-606">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="e64f0-606">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e64f0-607">1.0</span><span class="sxs-lookup"><span data-stu-id="e64f0-607">1.0</span></span>|
|[<span data-ttu-id="e64f0-608">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="e64f0-608">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e64f0-609">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e64f0-609">ReadItem</span></span>|
|[<span data-ttu-id="e64f0-610">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="e64f0-610">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="e64f0-611">阅读</span><span class="sxs-lookup"><span data-stu-id="e64f0-611">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="e64f0-612">示例</span><span class="sxs-lookup"><span data-stu-id="e64f0-612">Example</span></span>

```javascript
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
```

#### <a name="nullable-seriesid-string"></a><span data-ttu-id="e64f0-613">(nullable) seriesId :String</span><span class="sxs-lookup"><span data-stu-id="e64f0-613">(nullable) seriesId :String</span></span>

<span data-ttu-id="e64f0-614">获取实例所属的系列的 ID。</span><span class="sxs-lookup"><span data-stu-id="e64f0-614">Gets the id of the series that an instance belongs to.</span></span>

<span data-ttu-id="e64f0-615">在 OWA 和 Outlook 中，`seriesId` 返回此项目所属的父（系列）项目的 Exchange Web 服务 (EWS) ID。</span><span class="sxs-lookup"><span data-stu-id="e64f0-615">In OWA and Outlook, the `seriesId` returns the Exchange Web Services (EWS) ID of the parent (series) item that this item belongs to.</span></span> <span data-ttu-id="e64f0-616">但是，在 iOS 和 Android 中，`seriesId` 返回父项目的其余部分 ID。</span><span class="sxs-lookup"><span data-stu-id="e64f0-616">However, in iOS and Android, the `seriesId` returns the REST ID of the parent item.</span></span>

> [!NOTE]
> <span data-ttu-id="e64f0-617">`seriesId` 属性返回的标识符与 Exchange Web 服务项目标识符相同。</span><span class="sxs-lookup"><span data-stu-id="e64f0-617">The identifier returned by the `seriesId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="e64f0-618">`seriesId` 属性与 Outlook REST API 使用的 Outlook ID 不同。</span><span class="sxs-lookup"><span data-stu-id="e64f0-618">The `seriesId` property is not identical to the Outlook IDs used by the Outlook REST API.</span></span> <span data-ttu-id="e64f0-619">使用此值进行 REST API 调用前，应使用 [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string) 对它进行转换。</span><span class="sxs-lookup"><span data-stu-id="e64f0-619">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="e64f0-620">有关详细信息，请参阅[从 Outlook 加载项使用 Outlook REST API](https://docs.microsoft.com/outlook/add-ins/use-rest-api)。</span><span class="sxs-lookup"><span data-stu-id="e64f0-620">For more details, see [Use the Outlook REST APIs from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/use-rest-api).</span></span>

<span data-ttu-id="e64f0-621">`seriesId` 属性对于没有父项目（如单个约会、系列项目或会议请求）的项目返回 `null`，对于非会议请求的任何其他项目，返回 `undefined`。</span><span class="sxs-lookup"><span data-stu-id="e64f0-621">The `seriesId` property returns `null` for items that do not have parent items such as single appointments, series items, or meeting requests and returns `undefined` for any other items that are not meeting requests.</span></span>

##### <a name="type"></a><span data-ttu-id="e64f0-622">类型：</span><span class="sxs-lookup"><span data-stu-id="e64f0-622">Type:</span></span>

* <span data-ttu-id="e64f0-623">String</span><span class="sxs-lookup"><span data-stu-id="e64f0-623">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="e64f0-624">要求</span><span class="sxs-lookup"><span data-stu-id="e64f0-624">Requirements</span></span>

|<span data-ttu-id="e64f0-625">要求</span><span class="sxs-lookup"><span data-stu-id="e64f0-625">Requirement</span></span>|<span data-ttu-id="e64f0-626">值</span><span class="sxs-lookup"><span data-stu-id="e64f0-626">Value</span></span>|
|---|---|
|[<span data-ttu-id="e64f0-627">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="e64f0-627">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e64f0-628">1.7</span><span class="sxs-lookup"><span data-stu-id="e64f0-628">-17</span></span>|
|[<span data-ttu-id="e64f0-629">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="e64f0-629">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e64f0-630">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e64f0-630">ReadItem</span></span>|
|[<span data-ttu-id="e64f0-631">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="e64f0-631">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="e64f0-632">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="e64f0-632">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="e64f0-633">示例</span><span class="sxs-lookup"><span data-stu-id="e64f0-633">Example</span></span>

```javascript
var seriesId = Office.context.mailbox.item.seriesId;
var isSeries = (seriesId == null);
```

####  <a name="start-datetimejavascriptapioutlookofficetime"></a><span data-ttu-id="e64f0-634">start :Date|[Time](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="e64f0-634">start :Date|[Time](/javascript/api/outlook/office.time)</span></span>

<span data-ttu-id="e64f0-635">获取或设置约会开始的日期和时间。</span><span class="sxs-lookup"><span data-stu-id="e64f0-635">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="e64f0-p130">将 `start` 属性表示为协调世界时 (UTC) 的日期和时间值。可使用 [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttime) 方法将该值转换为客户端的本地日期和时间。</span><span class="sxs-lookup"><span data-stu-id="e64f0-p130">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="e64f0-638">阅读模式</span><span class="sxs-lookup"><span data-stu-id="e64f0-638">Read mode</span></span>

<span data-ttu-id="e64f0-639">`start` 属性返回 `Date` 对象。</span><span class="sxs-lookup"><span data-stu-id="e64f0-639">The `start` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="e64f0-640">撰写模式</span><span class="sxs-lookup"><span data-stu-id="e64f0-640">Compose mode</span></span>

<span data-ttu-id="e64f0-641">`start` 属性返回 `Time` 对象。</span><span class="sxs-lookup"><span data-stu-id="e64f0-641">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="e64f0-642">使用 [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) 方法设置开始时间时，应使用 [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) 方法将客户端的本地时间转换为服务器的 UTC。</span><span class="sxs-lookup"><span data-stu-id="e64f0-642">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="e64f0-643">类型：</span><span class="sxs-lookup"><span data-stu-id="e64f0-643">Type:</span></span>

*   <span data-ttu-id="e64f0-644">Date | [Time](/javascript/api/outlook/office.time)</span><span class="sxs-lookup"><span data-stu-id="e64f0-644">Date | [Time](/javascript/api/outlook/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="e64f0-645">要求</span><span class="sxs-lookup"><span data-stu-id="e64f0-645">Requirements</span></span>

|<span data-ttu-id="e64f0-646">要求</span><span class="sxs-lookup"><span data-stu-id="e64f0-646">Requirement</span></span>|<span data-ttu-id="e64f0-647">值</span><span class="sxs-lookup"><span data-stu-id="e64f0-647">Value</span></span>|
|---|---|
|[<span data-ttu-id="e64f0-648">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="e64f0-648">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e64f0-649">1.0</span><span class="sxs-lookup"><span data-stu-id="e64f0-649">1.0</span></span>|
|[<span data-ttu-id="e64f0-650">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="e64f0-650">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e64f0-651">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e64f0-651">ReadItem</span></span>|
|[<span data-ttu-id="e64f0-652">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="e64f0-652">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="e64f0-653">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="e64f0-653">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="e64f0-654">示例</span><span class="sxs-lookup"><span data-stu-id="e64f0-654">Example</span></span>

<span data-ttu-id="e64f0-655">以下示例通过使用 `Time` 对象的 [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) 方法，设置撰写模式下约会的开始时间。</span><span class="sxs-lookup"><span data-stu-id="e64f0-655">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

####  <a name="subject-stringsubjectjavascriptapioutlookofficesubject"></a><span data-ttu-id="e64f0-656">subject :String|[Subject](/javascript/api/outlook/office.subject)</span><span class="sxs-lookup"><span data-stu-id="e64f0-656">subject :String|[Subject](/javascript/api/outlook/office.subject)</span></span>

<span data-ttu-id="e64f0-657">获取或设置显示在项目的主题字段中的说明。</span><span class="sxs-lookup"><span data-stu-id="e64f0-657">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="e64f0-658">`subject` 属性获取或设置由电子邮件服务器发送项目时的整个主题。</span><span class="sxs-lookup"><span data-stu-id="e64f0-658">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="e64f0-659">阅读模式</span><span class="sxs-lookup"><span data-stu-id="e64f0-659">Read mode</span></span>

<span data-ttu-id="e64f0-p131">`subject` 属性返回一个字符串。使用 [`normalizedSubject`](#normalizedsubject-string) 属性获取不带任何前导前缀（如 `RE:` 和 `FW:`）的主题。</span><span class="sxs-lookup"><span data-stu-id="e64f0-p131">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

```javascript
var subject = Office.context.mailbox.item.subject;
```

##### <a name="compose-mode"></a><span data-ttu-id="e64f0-662">撰写模式</span><span class="sxs-lookup"><span data-stu-id="e64f0-662">Compose mode</span></span>

<span data-ttu-id="e64f0-663">`subject` 属性返回一个 `Subject` 对象，该对象提供用于获取和设置主题的方法。</span><span class="sxs-lookup"><span data-stu-id="e64f0-663">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```javascript
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="e64f0-664">类型：</span><span class="sxs-lookup"><span data-stu-id="e64f0-664">Type:</span></span>

*   <span data-ttu-id="e64f0-665">String | [Subject](/javascript/api/outlook/office.subject)</span><span class="sxs-lookup"><span data-stu-id="e64f0-665">String | [Subject](/javascript/api/outlook/office.subject)</span></span>

##### <a name="requirements"></a><span data-ttu-id="e64f0-666">要求</span><span class="sxs-lookup"><span data-stu-id="e64f0-666">Requirements</span></span>

|<span data-ttu-id="e64f0-667">要求</span><span class="sxs-lookup"><span data-stu-id="e64f0-667">Requirement</span></span>|<span data-ttu-id="e64f0-668">值</span><span class="sxs-lookup"><span data-stu-id="e64f0-668">Value</span></span>|
|---|---|
|[<span data-ttu-id="e64f0-669">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="e64f0-669">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e64f0-670">1.0</span><span class="sxs-lookup"><span data-stu-id="e64f0-670">1.0</span></span>|
|[<span data-ttu-id="e64f0-671">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="e64f0-671">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e64f0-672">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e64f0-672">ReadItem</span></span>|
|[<span data-ttu-id="e64f0-673">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="e64f0-673">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="e64f0-674">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="e64f0-674">Compose or read</span></span>|

####  <a name="to-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a><span data-ttu-id="e64f0-675">to :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="e64f0-675">to :Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)</span></span>

<span data-ttu-id="e64f0-676">提供对邮件的“**收件人**”行上的收件人的访问权限。</span><span class="sxs-lookup"><span data-stu-id="e64f0-676">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="e64f0-677">对象的类型和访问级别取决于当前项目的模式。</span><span class="sxs-lookup"><span data-stu-id="e64f0-677">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="e64f0-678">阅读模式</span><span class="sxs-lookup"><span data-stu-id="e64f0-678">Read mode</span></span>

<span data-ttu-id="e64f0-p133">`to` 属性返回包含邮件的**收件人**行上所列的每个收件人的 `EmailAddressDetails` 对象的数组。集合上限为 100 个成员。</span><span class="sxs-lookup"><span data-stu-id="e64f0-p133">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="e64f0-681">撰写模式</span><span class="sxs-lookup"><span data-stu-id="e64f0-681">Compose mode</span></span>

<span data-ttu-id="e64f0-682">`to` 属性返回一个 `Recipients` 对象，该对象提供用于获取或更新邮件的“**收件人**”行上收件人的方法。</span><span class="sxs-lookup"><span data-stu-id="e64f0-682">The to`to` property returns a Recipients`Recipients` object that provides methods for manipulating the recipients on the **To** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="e64f0-683">类型：</span><span class="sxs-lookup"><span data-stu-id="e64f0-683">Type:</span></span>

*   <span data-ttu-id="e64f0-684">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="e64f0-684">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="e64f0-685">要求</span><span class="sxs-lookup"><span data-stu-id="e64f0-685">Requirements</span></span>

|<span data-ttu-id="e64f0-686">要求</span><span class="sxs-lookup"><span data-stu-id="e64f0-686">Requirement</span></span>|<span data-ttu-id="e64f0-687">值</span><span class="sxs-lookup"><span data-stu-id="e64f0-687">Value</span></span>|
|---|---|
|[<span data-ttu-id="e64f0-688">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="e64f0-688">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e64f0-689">1.0</span><span class="sxs-lookup"><span data-stu-id="e64f0-689">1.0</span></span>|
|[<span data-ttu-id="e64f0-690">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="e64f0-690">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e64f0-691">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e64f0-691">ReadItem</span></span>|
|[<span data-ttu-id="e64f0-692">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="e64f0-692">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="e64f0-693">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="e64f0-693">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="e64f0-694">示例</span><span class="sxs-lookup"><span data-stu-id="e64f0-694">Example</span></span>

```javascript
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

### <a name="methods"></a><span data-ttu-id="e64f0-695">方法</span><span class="sxs-lookup"><span data-stu-id="e64f0-695">Methods</span></span>

####  <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="e64f0-696">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="e64f0-696">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="e64f0-697">将文件作为附件添加到邮件或约会。</span><span class="sxs-lookup"><span data-stu-id="e64f0-697">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="e64f0-698">`addFileAttachmentAsync` 方法在指定的 URI 上载文件并将其附加到撰写窗体中的项目。</span><span class="sxs-lookup"><span data-stu-id="e64f0-698">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="e64f0-699">你随后可以将该标识符与 [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) 方法一同使用，以删除同一个会话中的附件。</span><span class="sxs-lookup"><span data-stu-id="e64f0-699">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="e64f0-700">参数：</span><span class="sxs-lookup"><span data-stu-id="e64f0-700">Parameters:</span></span>
|<span data-ttu-id="e64f0-701">名称</span><span class="sxs-lookup"><span data-stu-id="e64f0-701">Name</span></span>|<span data-ttu-id="e64f0-702">类型</span><span class="sxs-lookup"><span data-stu-id="e64f0-702">Type</span></span>|<span data-ttu-id="e64f0-703">属性</span><span class="sxs-lookup"><span data-stu-id="e64f0-703">Attributes</span></span>|<span data-ttu-id="e64f0-704">说明</span><span class="sxs-lookup"><span data-stu-id="e64f0-704">Description</span></span>|
|---|---|---|---|
|`uri`|<span data-ttu-id="e64f0-705">字符串</span><span class="sxs-lookup"><span data-stu-id="e64f0-705">String</span></span>||<span data-ttu-id="e64f0-p134">提供附加到邮件或约会的文件的位置的 URI。最大长度为 2048 个字符。</span><span class="sxs-lookup"><span data-stu-id="e64f0-p134">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`|<span data-ttu-id="e64f0-708">字符串</span><span class="sxs-lookup"><span data-stu-id="e64f0-708">String</span></span>||<span data-ttu-id="e64f0-p135">在附件上载过程中显示的附件名称。最大长度为 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="e64f0-p135">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="e64f0-711">Object</span><span class="sxs-lookup"><span data-stu-id="e64f0-711">Object</span></span>|<span data-ttu-id="e64f0-712">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="e64f0-712">&lt;optional&gt;</span></span>|<span data-ttu-id="e64f0-713">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="e64f0-713">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="e64f0-714">对象</span><span class="sxs-lookup"><span data-stu-id="e64f0-714">Object</span></span>|<span data-ttu-id="e64f0-715">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="e64f0-715">&lt;optional&gt;</span></span>|<span data-ttu-id="e64f0-716">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="e64f0-716">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.isInline`|<span data-ttu-id="e64f0-717">布尔值</span><span class="sxs-lookup"><span data-stu-id="e64f0-717">Boolean</span></span>|<span data-ttu-id="e64f0-718">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="e64f0-718">&lt;optional&gt;</span></span>|<span data-ttu-id="e64f0-719">如果为 `true`，则表示附件将在邮件正文中内联显示，并且不应显示在附件列表中。</span><span class="sxs-lookup"><span data-stu-id="e64f0-719">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`callback`|<span data-ttu-id="e64f0-720">函数</span><span class="sxs-lookup"><span data-stu-id="e64f0-720">function</span></span>|<span data-ttu-id="e64f0-721">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="e64f0-721">&lt;optional&gt;</span></span>|<span data-ttu-id="e64f0-722">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="e64f0-722">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="e64f0-723">如果成功，附件标识符将在 `asyncResult.value` 属性中提供。</span><span class="sxs-lookup"><span data-stu-id="e64f0-723">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="e64f0-724">如果上传附件失败，`asyncResult` 对象将包含一个提供错误说明的 `Error` 对象。</span><span class="sxs-lookup"><span data-stu-id="e64f0-724">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="e64f0-725">错误</span><span class="sxs-lookup"><span data-stu-id="e64f0-725">Errors</span></span>

|<span data-ttu-id="e64f0-726">错误代码</span><span class="sxs-lookup"><span data-stu-id="e64f0-726">Error code</span></span>|<span data-ttu-id="e64f0-727">说明</span><span class="sxs-lookup"><span data-stu-id="e64f0-727">Description</span></span>|
|------------|-------------|
|`AttachmentSizeExceeded`|<span data-ttu-id="e64f0-728">附件大小超过了允许的大小。</span><span class="sxs-lookup"><span data-stu-id="e64f0-728">The attachment is larger than allowed.</span></span>|
|`FileTypeNotSupported`|<span data-ttu-id="e64f0-729">该附件的扩展名不是允许的扩展名。</span><span class="sxs-lookup"><span data-stu-id="e64f0-729">The attachment has an extension that is not allowed.</span></span>|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="e64f0-730">邮件或约会具有的附件过多。</span><span class="sxs-lookup"><span data-stu-id="e64f0-730">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e64f0-731">要求</span><span class="sxs-lookup"><span data-stu-id="e64f0-731">Requirements</span></span>

|<span data-ttu-id="e64f0-732">要求</span><span class="sxs-lookup"><span data-stu-id="e64f0-732">Requirement</span></span>|<span data-ttu-id="e64f0-733">值</span><span class="sxs-lookup"><span data-stu-id="e64f0-733">Value</span></span>|
|---|---|
|[<span data-ttu-id="e64f0-734">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="e64f0-734">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e64f0-735">1.1</span><span class="sxs-lookup"><span data-stu-id="e64f0-735">1.1</span></span>|
|[<span data-ttu-id="e64f0-736">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="e64f0-736">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e64f0-737">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="e64f0-737">ReadWriteItem</span></span>|
|[<span data-ttu-id="e64f0-738">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="e64f0-738">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="e64f0-739">撰写</span><span class="sxs-lookup"><span data-stu-id="e64f0-739">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="e64f0-740">示例</span><span class="sxs-lookup"><span data-stu-id="e64f0-740">Examples</span></span>

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

<span data-ttu-id="e64f0-741">以下示例将图像文件添加为内联附件，并在邮件正文中引用该附件。</span><span class="sxs-lookup"><span data-stu-id="e64f0-741">The following example adds an image file as an inline attachment and references the attachment in the message body.</span></span>

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

#### <a name="addfileattachmentfrombase64asyncbase64file-attachmentname-options-callback"></a><span data-ttu-id="e64f0-742">addFileAttachmentFromBase64Async(base64File, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="e64f0-742">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="e64f0-743">将 base64 编码中的文件作为附件添加到消息或约会。</span><span class="sxs-lookup"><span data-stu-id="e64f0-743">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="e64f0-744">`addFileAttachmentFromBase64Async` 方法从 base64 编码上传文件，并将其附加到撰写表单中的项目。</span><span class="sxs-lookup"><span data-stu-id="e64f0-744">The `addFileAttachmentFromBase64Async` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span> <span data-ttu-id="e64f0-745">此方法返回 AsyncResult.value 对象中的附件标识符。</span><span class="sxs-lookup"><span data-stu-id="e64f0-745">This method returns the attachment identifier in the AsyncResult.value object.</span></span>

<span data-ttu-id="e64f0-746">随后可以将该标识符与 [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) 方法一同使用，以删除同一个会话中的附件。</span><span class="sxs-lookup"><span data-stu-id="e64f0-746">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="e64f0-747">参数：</span><span class="sxs-lookup"><span data-stu-id="e64f0-747">Parameters:</span></span>
|<span data-ttu-id="e64f0-748">名称</span><span class="sxs-lookup"><span data-stu-id="e64f0-748">Name</span></span>|<span data-ttu-id="e64f0-749">类型</span><span class="sxs-lookup"><span data-stu-id="e64f0-749">Type</span></span>|<span data-ttu-id="e64f0-750">属性</span><span class="sxs-lookup"><span data-stu-id="e64f0-750">Attributes</span></span>|<span data-ttu-id="e64f0-751">说明</span><span class="sxs-lookup"><span data-stu-id="e64f0-751">Description</span></span>|
|---|---|---|---|
|`base64File`|<span data-ttu-id="e64f0-752">字符串</span><span class="sxs-lookup"><span data-stu-id="e64f0-752">String</span></span>||<span data-ttu-id="e64f0-753">要添加到电子邮件或事件的图像或文件的 base64 编码内容。</span><span class="sxs-lookup"><span data-stu-id="e64f0-753">The base64 encoded content of an image or file to be added to an email or event.</span></span>|
|`attachmentName`|<span data-ttu-id="e64f0-754">字符串</span><span class="sxs-lookup"><span data-stu-id="e64f0-754">String</span></span>||<span data-ttu-id="e64f0-p137">在附件上载过程中显示的附件名称。最大长度为 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="e64f0-p137">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="e64f0-757">Object</span><span class="sxs-lookup"><span data-stu-id="e64f0-757">Object</span></span>|<span data-ttu-id="e64f0-758">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="e64f0-758">&lt;optional&gt;</span></span>|<span data-ttu-id="e64f0-759">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="e64f0-759">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="e64f0-760">对象</span><span class="sxs-lookup"><span data-stu-id="e64f0-760">Object</span></span>|<span data-ttu-id="e64f0-761">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="e64f0-761">&lt;optional&gt;</span></span>|<span data-ttu-id="e64f0-762">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="e64f0-762">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.isInline`|<span data-ttu-id="e64f0-763">布尔值</span><span class="sxs-lookup"><span data-stu-id="e64f0-763">Boolean</span></span>|<span data-ttu-id="e64f0-764">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="e64f0-764">&lt;optional&gt;</span></span>|<span data-ttu-id="e64f0-765">如果为 `true`，则表示附件将在邮件正文中内联显示，并且不应显示在附件列表中。</span><span class="sxs-lookup"><span data-stu-id="e64f0-765">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`callback`|<span data-ttu-id="e64f0-766">函数</span><span class="sxs-lookup"><span data-stu-id="e64f0-766">function</span></span>|<span data-ttu-id="e64f0-767">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="e64f0-767">&lt;optional&gt;</span></span>|<span data-ttu-id="e64f0-768">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="e64f0-768">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="e64f0-769">如果成功，附件标识符将在 `asyncResult.value` 属性中提供。</span><span class="sxs-lookup"><span data-stu-id="e64f0-769">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="e64f0-770">如果上传附件失败，`asyncResult` 对象将包含一个提供错误说明的 `Error` 对象。</span><span class="sxs-lookup"><span data-stu-id="e64f0-770">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="e64f0-771">错误</span><span class="sxs-lookup"><span data-stu-id="e64f0-771">Errors</span></span>

|<span data-ttu-id="e64f0-772">错误代码</span><span class="sxs-lookup"><span data-stu-id="e64f0-772">Error code</span></span>|<span data-ttu-id="e64f0-773">说明</span><span class="sxs-lookup"><span data-stu-id="e64f0-773">Description</span></span>|
|------------|-------------|
|`AttachmentSizeExceeded`|<span data-ttu-id="e64f0-774">附件大小超过了允许的大小。</span><span class="sxs-lookup"><span data-stu-id="e64f0-774">The attachment is larger than allowed.</span></span>|
|`FileTypeNotSupported`|<span data-ttu-id="e64f0-775">该附件的扩展名不是允许的扩展名。</span><span class="sxs-lookup"><span data-stu-id="e64f0-775">The attachment has an extension that is not allowed.</span></span>|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="e64f0-776">邮件或约会具有的附件过多。</span><span class="sxs-lookup"><span data-stu-id="e64f0-776">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e64f0-777">要求</span><span class="sxs-lookup"><span data-stu-id="e64f0-777">Requirements</span></span>

|<span data-ttu-id="e64f0-778">要求</span><span class="sxs-lookup"><span data-stu-id="e64f0-778">Requirement</span></span>|<span data-ttu-id="e64f0-779">值</span><span class="sxs-lookup"><span data-stu-id="e64f0-779">Value</span></span>|
|---|---|
|[<span data-ttu-id="e64f0-780">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="e64f0-780">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e64f0-781">预览</span><span class="sxs-lookup"><span data-stu-id="e64f0-781">Preview</span></span>|
|[<span data-ttu-id="e64f0-782">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="e64f0-782">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e64f0-783">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="e64f0-783">ReadWriteItem</span></span>|
|[<span data-ttu-id="e64f0-784">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="e64f0-784">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="e64f0-785">撰写</span><span class="sxs-lookup"><span data-stu-id="e64f0-785">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="e64f0-786">示例</span><span class="sxs-lookup"><span data-stu-id="e64f0-786">Examples</span></span>

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

####  <a name="addhandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="e64f0-787">addHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="e64f0-787">addHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="e64f0-788">添加支持事件的事件处理程序。</span><span class="sxs-lookup"><span data-stu-id="e64f0-788">Adds an event handler for a supported event.</span></span>

<span data-ttu-id="e64f0-789">当前，支持的事件类型是 `Office.EventType.AttachmentsChanged`、`Office.EventType.AppointmentTimeChanged`、`Office.EventType.RecipientsChanged` 和 `Office.EventType.RecurrenceChanged`</span><span class="sxs-lookup"><span data-stu-id="e64f0-789">Currently the supported event types are `Office.EventType.AttachmentsChanged`, `Office.EventType.AppointmentTimeChanged`, `Office.EventType.RecipientsChanged`, and `Office.EventType.RecurrenceChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="e64f0-790">参数：</span><span class="sxs-lookup"><span data-stu-id="e64f0-790">Parameters:</span></span>

| <span data-ttu-id="e64f0-791">名称</span><span class="sxs-lookup"><span data-stu-id="e64f0-791">Name</span></span> | <span data-ttu-id="e64f0-792">类型</span><span class="sxs-lookup"><span data-stu-id="e64f0-792">Type</span></span> | <span data-ttu-id="e64f0-793">属性</span><span class="sxs-lookup"><span data-stu-id="e64f0-793">Attributes</span></span> | <span data-ttu-id="e64f0-794">说明</span><span class="sxs-lookup"><span data-stu-id="e64f0-794">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="e64f0-795">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="e64f0-795">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="e64f0-796">应调用处理程序的事件。</span><span class="sxs-lookup"><span data-stu-id="e64f0-796">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="e64f0-797">函数</span><span class="sxs-lookup"><span data-stu-id="e64f0-797">Function</span></span> || <span data-ttu-id="e64f0-p138">用于处理事件的函数。此函数必须接受一个参数，即对象文本。参数上的 `type` 属性将匹配传递给 `addHandlerAsync` 的 `eventType` 参数。</span><span class="sxs-lookup"><span data-stu-id="e64f0-p138">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `addHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="e64f0-801">Object</span><span class="sxs-lookup"><span data-stu-id="e64f0-801">Object</span></span> | <span data-ttu-id="e64f0-802">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="e64f0-802">&lt;optional&gt;</span></span> | <span data-ttu-id="e64f0-803">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="e64f0-803">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="e64f0-804">对象</span><span class="sxs-lookup"><span data-stu-id="e64f0-804">Object</span></span> | <span data-ttu-id="e64f0-805">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="e64f0-805">&lt;optional&gt;</span></span> | <span data-ttu-id="e64f0-806">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="e64f0-806">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="e64f0-807">函数</span><span class="sxs-lookup"><span data-stu-id="e64f0-807">function</span></span>| <span data-ttu-id="e64f0-808">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="e64f0-808">&lt;optional&gt;</span></span>|<span data-ttu-id="e64f0-809">方法完成后，使用单个参数 `callback`（一个 [`asyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `AsyncResult` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="e64f0-809">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e64f0-810">Requirements</span><span class="sxs-lookup"><span data-stu-id="e64f0-810">Requirements</span></span>

|<span data-ttu-id="e64f0-811">要求</span><span class="sxs-lookup"><span data-stu-id="e64f0-811">Requirement</span></span>| <span data-ttu-id="e64f0-812">值</span><span class="sxs-lookup"><span data-stu-id="e64f0-812">Value</span></span>|
|---|---|
|[<span data-ttu-id="e64f0-813">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="e64f0-813">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e64f0-814">1.7</span><span class="sxs-lookup"><span data-stu-id="e64f0-814">-17</span></span> |
|[<span data-ttu-id="e64f0-815">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="e64f0-815">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e64f0-816">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e64f0-816">ReadItem</span></span> |
|[<span data-ttu-id="e64f0-817">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="e64f0-817">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="e64f0-818">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="e64f0-818">Compose or read</span></span> |

####  <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="e64f0-819">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="e64f0-819">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="e64f0-820">将 Exchange 项目（如邮件）作为附件添加到邮件或约会。</span><span class="sxs-lookup"><span data-stu-id="e64f0-820">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="e64f0-p139">`addItemAttachmentAsync` 方法将包含指定 Exchange 标识符的项目附加到撰写窗体中的项目。如果指定一个回调方法，此方法使用 `asyncResult` 参数调用，该参数包含一个附件标识符或代码，指示附加项目过程中出现的任何错误。可以使用 `options` 参数将状态信息传递给回调方法（如果需要）。</span><span class="sxs-lookup"><span data-stu-id="e64f0-p139">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="e64f0-824">随后可以将该标识符与 [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) 方法一同使用，以删除同一个会话中的附件。</span><span class="sxs-lookup"><span data-stu-id="e64f0-824">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="e64f0-825">如果 Office 加载项在 Outlook Web App 中运行，则 `addItemAttachmentAsync` 方法可以将项目附加到项目（正在编辑的项目除外）中；然而，不支持也不建议这样做。</span><span class="sxs-lookup"><span data-stu-id="e64f0-825">If your Office add-in is running in Outlook Web App, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="e64f0-826">参数：</span><span class="sxs-lookup"><span data-stu-id="e64f0-826">Parameters:</span></span>

|<span data-ttu-id="e64f0-827">名称</span><span class="sxs-lookup"><span data-stu-id="e64f0-827">Name</span></span>|<span data-ttu-id="e64f0-828">类型</span><span class="sxs-lookup"><span data-stu-id="e64f0-828">Type</span></span>|<span data-ttu-id="e64f0-829">属性</span><span class="sxs-lookup"><span data-stu-id="e64f0-829">Attributes</span></span>|<span data-ttu-id="e64f0-830">说明</span><span class="sxs-lookup"><span data-stu-id="e64f0-830">Description</span></span>|
|---|---|---|---|
|`itemId`|<span data-ttu-id="e64f0-831">字符串</span><span class="sxs-lookup"><span data-stu-id="e64f0-831">String</span></span>||<span data-ttu-id="e64f0-p140">要附加的项目的 Exchange 标识符。最大长度为 100 个字符。</span><span class="sxs-lookup"><span data-stu-id="e64f0-p140">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`|<span data-ttu-id="e64f0-834">字符串</span><span class="sxs-lookup"><span data-stu-id="e64f0-834">String</span></span>||<span data-ttu-id="e64f0-p141">要附加的项目的主题。最大长度为 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="e64f0-p141">The sujbect of the item to be attached. The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="e64f0-837">Object</span><span class="sxs-lookup"><span data-stu-id="e64f0-837">Object</span></span>|<span data-ttu-id="e64f0-838">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="e64f0-838">&lt;optional&gt;</span></span>|<span data-ttu-id="e64f0-839">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="e64f0-839">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="e64f0-840">对象</span><span class="sxs-lookup"><span data-stu-id="e64f0-840">Object</span></span>|<span data-ttu-id="e64f0-841">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="e64f0-841">&lt;optional&gt;</span></span>|<span data-ttu-id="e64f0-842">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="e64f0-842">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="e64f0-843">函数</span><span class="sxs-lookup"><span data-stu-id="e64f0-843">function</span></span>|<span data-ttu-id="e64f0-844">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="e64f0-844">&lt;optional&gt;</span></span>|<span data-ttu-id="e64f0-845">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="e64f0-845">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="e64f0-846">如果成功，附件标识符将在 `asyncResult.value` 属性中提供。</span><span class="sxs-lookup"><span data-stu-id="e64f0-846">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="e64f0-847">如果添加附件失败，`asyncResult` 对象将包含一个提供错误说明的 `Error` 对象。</span><span class="sxs-lookup"><span data-stu-id="e64f0-847">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="e64f0-848">错误</span><span class="sxs-lookup"><span data-stu-id="e64f0-848">Errors</span></span>

|<span data-ttu-id="e64f0-849">错误代码</span><span class="sxs-lookup"><span data-stu-id="e64f0-849">Error code</span></span>|<span data-ttu-id="e64f0-850">说明</span><span class="sxs-lookup"><span data-stu-id="e64f0-850">Description</span></span>|
|------------|-------------|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="e64f0-851">邮件或约会具有的附件过多。</span><span class="sxs-lookup"><span data-stu-id="e64f0-851">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e64f0-852">要求</span><span class="sxs-lookup"><span data-stu-id="e64f0-852">Requirements</span></span>

|<span data-ttu-id="e64f0-853">要求</span><span class="sxs-lookup"><span data-stu-id="e64f0-853">Requirement</span></span>|<span data-ttu-id="e64f0-854">值</span><span class="sxs-lookup"><span data-stu-id="e64f0-854">Value</span></span>|
|---|---|
|[<span data-ttu-id="e64f0-855">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="e64f0-855">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e64f0-856">1.1</span><span class="sxs-lookup"><span data-stu-id="e64f0-856">1.1</span></span>|
|[<span data-ttu-id="e64f0-857">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="e64f0-857">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e64f0-858">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="e64f0-858">ReadWriteItem</span></span>|
|[<span data-ttu-id="e64f0-859">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="e64f0-859">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="e64f0-860">撰写</span><span class="sxs-lookup"><span data-stu-id="e64f0-860">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="e64f0-861">示例</span><span class="sxs-lookup"><span data-stu-id="e64f0-861">Example</span></span>

<span data-ttu-id="e64f0-862">以下示例将现有的 Outlook 项目添加为名为 `My Attachment` 的附件。</span><span class="sxs-lookup"><span data-stu-id="e64f0-862">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

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

####  <a name="close"></a><span data-ttu-id="e64f0-863">close()</span><span class="sxs-lookup"><span data-stu-id="e64f0-863">close()</span></span>

<span data-ttu-id="e64f0-864">关闭当前正在撰写的项目。</span><span class="sxs-lookup"><span data-stu-id="e64f0-864">Closes the current item that is being composed.</span></span>

<span data-ttu-id="e64f0-p142">
            \`close\` 方法的行为取决于要撰写的项目的当前状态。如果项目具有未保存的更改，客户端将提示用户保存、放弃或取消关闭操作。</span><span class="sxs-lookup"><span data-stu-id="e64f0-p142">The behavior of the `close` method depends on the current state of the item being composed. If the item has unsaved changes, the client prompts the user to save, discard, or cancel the close action.</span></span>

> [!NOTE]
> <span data-ttu-id="e64f0-867">在 Outlook 网页版中，如果该项目是约会并且之前已使用 `saveAsync` 保存，则即使自上次保存项目后未发生任何更改，也会提示用户保存、放弃或取消。</span><span class="sxs-lookup"><span data-stu-id="e64f0-867">In Outlook on the web, if the item is an appointment and it has previously been saved using `saveAsync`, the user is prompted to save, discard, or cancel even if no changes have occurred since the item was last saved.</span></span>

<span data-ttu-id="e64f0-868">在 Outlook 桌面客户端中，如果邮件是内联答复，`close` 方法不起作用。</span><span class="sxs-lookup"><span data-stu-id="e64f0-868">In the Outlook desktop client, if the message is an inline reply, the `close` method has no effect.</span></span>

##### <a name="requirements"></a><span data-ttu-id="e64f0-869">要求</span><span class="sxs-lookup"><span data-stu-id="e64f0-869">Requirements</span></span>

|<span data-ttu-id="e64f0-870">要求</span><span class="sxs-lookup"><span data-stu-id="e64f0-870">Requirement</span></span>|<span data-ttu-id="e64f0-871">值</span><span class="sxs-lookup"><span data-stu-id="e64f0-871">Value</span></span>|
|---|---|
|[<span data-ttu-id="e64f0-872">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="e64f0-872">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e64f0-873">1.3</span><span class="sxs-lookup"><span data-stu-id="e64f0-873">1.3</span></span>|
|[<span data-ttu-id="e64f0-874">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="e64f0-874">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e64f0-875">受限</span><span class="sxs-lookup"><span data-stu-id="e64f0-875">Restricted</span></span>|
|[<span data-ttu-id="e64f0-876">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="e64f0-876">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="e64f0-877">撰写</span><span class="sxs-lookup"><span data-stu-id="e64f0-877">Compose</span></span>|

#### <a name="displayreplyallformformdata"></a><span data-ttu-id="e64f0-878">displayReplyAllForm(formData)</span><span class="sxs-lookup"><span data-stu-id="e64f0-878">displayReplyAllForm(formData)</span></span>

<span data-ttu-id="e64f0-879">显示答复窗体，其中包括所选邮件的发件人和所有收件人或所选约会的组织者和所有与会者。</span><span class="sxs-lookup"><span data-stu-id="e64f0-879">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="e64f0-880">在 Outlook for iOS 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="e64f0-880">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="e64f0-881">在 Outlook Web App 中，答复窗体显示为包含 3 列视图的弹出式窗体以及包含 2 列或 1 列视图的弹出式窗体。</span><span class="sxs-lookup"><span data-stu-id="e64f0-881">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="e64f0-882">如果任意字符串参数超出其限制，`displayReplyAllForm` 将引发异常。</span><span class="sxs-lookup"><span data-stu-id="e64f0-882">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="e64f0-p143">当在 `formData.attachments` 参数中指定附件时，Outlook 和 Outlook Web App 尝试下载所有附件并将其附加到答复窗体。如果无法添加任何附件，则在窗体 UI 中显示错误。如果这不可能，则不引发错误消息。</span><span class="sxs-lookup"><span data-stu-id="e64f0-p143">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="e64f0-886">参数：</span><span class="sxs-lookup"><span data-stu-id="e64f0-886">Parameters:</span></span>

|<span data-ttu-id="e64f0-887">名称</span><span class="sxs-lookup"><span data-stu-id="e64f0-887">Name</span></span>|<span data-ttu-id="e64f0-888">类型</span><span class="sxs-lookup"><span data-stu-id="e64f0-888">Type</span></span>|<span data-ttu-id="e64f0-889">属性</span><span class="sxs-lookup"><span data-stu-id="e64f0-889">Attributes</span></span>|<span data-ttu-id="e64f0-890">说明</span><span class="sxs-lookup"><span data-stu-id="e64f0-890">Description</span></span>|
|---|---|---|---|
|`formData`|<span data-ttu-id="e64f0-891">字符串 &#124; 对象</span><span class="sxs-lookup"><span data-stu-id="e64f0-891">String &#124; Object</span></span>||<span data-ttu-id="e64f0-p144">一个包含文本和 HTML 且表示答复窗体的正文的字符串。字符串限制为 32 KB。</span><span class="sxs-lookup"><span data-stu-id="e64f0-p144">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="e64f0-894">**或**</span><span class="sxs-lookup"><span data-stu-id="e64f0-894">**OR**</span></span><br/><span data-ttu-id="e64f0-p145">包含正文或附件数据和回调函数的对象。对象定义如下。</span><span class="sxs-lookup"><span data-stu-id="e64f0-p145">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span>|
|`formData.htmlBody`|<span data-ttu-id="e64f0-897">字符串</span><span class="sxs-lookup"><span data-stu-id="e64f0-897">String</span></span>|<span data-ttu-id="e64f0-898">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="e64f0-898">&lt;optional&gt;</span></span>|<span data-ttu-id="e64f0-p146">一个包含文本和 HTML 且表示答复窗体的正文的字符串。字符串限制为 32 KB。</span><span class="sxs-lookup"><span data-stu-id="e64f0-p146">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
|`formData.attachments`|<span data-ttu-id="e64f0-901">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="e64f0-901">Array.&lt;Object&gt;</span></span>|<span data-ttu-id="e64f0-902">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="e64f0-902">&lt;optional&gt;</span></span>|<span data-ttu-id="e64f0-903">JSON 对象（是文件或项目附件）的数组。</span><span class="sxs-lookup"><span data-stu-id="e64f0-903">An array of JSON objects that are either file or item attachments.</span></span>|
|`formData.attachments.type`|<span data-ttu-id="e64f0-904">字符串</span><span class="sxs-lookup"><span data-stu-id="e64f0-904">String</span></span>||<span data-ttu-id="e64f0-p147">指示附件的类型。必须是文件附件的 `file` 或项目附件的 `item`。</span><span class="sxs-lookup"><span data-stu-id="e64f0-p147">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span>|
|`formData.attachments.name`|<span data-ttu-id="e64f0-907">字符串</span><span class="sxs-lookup"><span data-stu-id="e64f0-907">String</span></span>||<span data-ttu-id="e64f0-908">一个包含附件的名称的字符串，最多包含 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="e64f0-908">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
|`formData.attachments.url`|<span data-ttu-id="e64f0-909">String</span><span class="sxs-lookup"><span data-stu-id="e64f0-909">String</span></span>||<span data-ttu-id="e64f0-p148">仅在将 `type` 设置为 `file` 时使用。文件的位置的 URI。</span><span class="sxs-lookup"><span data-stu-id="e64f0-p148">Only used if `type` is set to `file`. The URI of the location for the file.</span></span>|
|`formData.attachments.isInline`|<span data-ttu-id="e64f0-912">布尔</span><span class="sxs-lookup"><span data-stu-id="e64f0-912">Boolean</span></span>||<span data-ttu-id="e64f0-p149">仅在将 `type` 设置为 `file` 时使用。如果为 `true`，则表示附件将在邮件正文中内联显示，并且不应显示在附件列表中。</span><span class="sxs-lookup"><span data-stu-id="e64f0-p149">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`formData.attachments.itemId`|<span data-ttu-id="e64f0-915">String</span><span class="sxs-lookup"><span data-stu-id="e64f0-915">String</span></span>||<span data-ttu-id="e64f0-p150">仅在将 `type` 设置为 `item` 时使用。附件的 EWS 项目 ID。字符串最长为 100 个字符。</span><span class="sxs-lookup"><span data-stu-id="e64f0-p150">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span>|
|`callback`|<span data-ttu-id="e64f0-919">函数</span><span class="sxs-lookup"><span data-stu-id="e64f0-919">function</span></span>|<span data-ttu-id="e64f0-920">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="e64f0-920">&lt;optional&gt;</span></span>|<span data-ttu-id="e64f0-921">方法完成后，使用单个参数 `asyncResult`（一个 [AsyncResult](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="e64f0-921">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e64f0-922">要求</span><span class="sxs-lookup"><span data-stu-id="e64f0-922">Requirements</span></span>

|<span data-ttu-id="e64f0-923">要求</span><span class="sxs-lookup"><span data-stu-id="e64f0-923">Requirement</span></span>|<span data-ttu-id="e64f0-924">值</span><span class="sxs-lookup"><span data-stu-id="e64f0-924">Value</span></span>|
|---|---|
|[<span data-ttu-id="e64f0-925">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="e64f0-925">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e64f0-926">1.0</span><span class="sxs-lookup"><span data-stu-id="e64f0-926">1.0</span></span>|
|[<span data-ttu-id="e64f0-927">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="e64f0-927">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e64f0-928">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e64f0-928">ReadItem</span></span>|
|[<span data-ttu-id="e64f0-929">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="e64f0-929">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="e64f0-930">阅读</span><span class="sxs-lookup"><span data-stu-id="e64f0-930">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="e64f0-931">示例</span><span class="sxs-lookup"><span data-stu-id="e64f0-931">Examples</span></span>

<span data-ttu-id="e64f0-932">以下代码将一个字符串传递到 `displayReplyAllForm` 函数。</span><span class="sxs-lookup"><span data-stu-id="e64f0-932">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="e64f0-933">使用空白正文答复。</span><span class="sxs-lookup"><span data-stu-id="e64f0-933">Reply with an empty body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="e64f0-934">仅使用正文答复。</span><span class="sxs-lookup"><span data-stu-id="e64f0-934">Reply with just a body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="e64f0-935">使用正文和文件附件答复。</span><span class="sxs-lookup"><span data-stu-id="e64f0-935">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="e64f0-936">使用正文和项目附件答复。</span><span class="sxs-lookup"><span data-stu-id="e64f0-936">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="e64f0-937">使用正文、文件附件、项目附件和回调答复。</span><span class="sxs-lookup"><span data-stu-id="e64f0-937">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="displayreplyformformdata"></a><span data-ttu-id="e64f0-938">displayReplyForm(formData)</span><span class="sxs-lookup"><span data-stu-id="e64f0-938">displayReplyForm(formData)</span></span>

<span data-ttu-id="e64f0-939">显示答复窗体，其中仅包括所选邮件的发件人或所选约会的组织者。</span><span class="sxs-lookup"><span data-stu-id="e64f0-939">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="e64f0-940">在 Outlook for iOS 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="e64f0-940">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="e64f0-941">在 Outlook Web App 中，答复窗体显示为包含 3 列视图的弹出式窗体以及包含 2 列或 1 列视图的弹出式窗体。</span><span class="sxs-lookup"><span data-stu-id="e64f0-941">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="e64f0-942">如果任意字符串参数超出其限制，`displayReplyForm` 将引发异常。</span><span class="sxs-lookup"><span data-stu-id="e64f0-942">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="e64f0-p151">当在 `formData.attachments` 参数中指定附件时，Outlook 和 Outlook Web App 尝试下载所有附件并将其附加到答复窗体。如果无法添加任何附件，则在窗体 UI 中显示错误。如果这不可能，则不引发错误消息。</span><span class="sxs-lookup"><span data-stu-id="e64f0-p151">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="e64f0-946">参数：</span><span class="sxs-lookup"><span data-stu-id="e64f0-946">Parameters:</span></span>

|<span data-ttu-id="e64f0-947">名称</span><span class="sxs-lookup"><span data-stu-id="e64f0-947">Name</span></span>|<span data-ttu-id="e64f0-948">类型</span><span class="sxs-lookup"><span data-stu-id="e64f0-948">Type</span></span>|<span data-ttu-id="e64f0-949">属性</span><span class="sxs-lookup"><span data-stu-id="e64f0-949">Attributes</span></span>|<span data-ttu-id="e64f0-950">说明</span><span class="sxs-lookup"><span data-stu-id="e64f0-950">Description</span></span>|
|---|---|---|---|
|`formData`|<span data-ttu-id="e64f0-951">字符串 &#124; 对象</span><span class="sxs-lookup"><span data-stu-id="e64f0-951">String &#124; Object</span></span>||<span data-ttu-id="e64f0-p152">一个包含文本和 HTML 且表示答复窗体的正文的字符串。字符串限制为 32 KB。</span><span class="sxs-lookup"><span data-stu-id="e64f0-p152">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="e64f0-954">**或**</span><span class="sxs-lookup"><span data-stu-id="e64f0-954">**OR**</span></span><br/><span data-ttu-id="e64f0-p153">包含正文或附件数据和回调函数的对象。对象定义如下。</span><span class="sxs-lookup"><span data-stu-id="e64f0-p153">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span>|
|`formData.htmlBody`|<span data-ttu-id="e64f0-957">字符串</span><span class="sxs-lookup"><span data-stu-id="e64f0-957">String</span></span>|<span data-ttu-id="e64f0-958">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="e64f0-958">&lt;optional&gt;</span></span>|<span data-ttu-id="e64f0-p154">一个包含文本和 HTML 且表示答复窗体的正文的字符串。字符串限制为 32 KB。</span><span class="sxs-lookup"><span data-stu-id="e64f0-p154">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
|`formData.attachments`|<span data-ttu-id="e64f0-961">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="e64f0-961">Array.&lt;Object&gt;</span></span>|<span data-ttu-id="e64f0-962">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="e64f0-962">&lt;optional&gt;</span></span>|<span data-ttu-id="e64f0-963">JSON 对象（是文件或项目附件）的数组。</span><span class="sxs-lookup"><span data-stu-id="e64f0-963">An array of JSON objects that are either file or item attachments.</span></span>|
|`formData.attachments.type`|<span data-ttu-id="e64f0-964">字符串</span><span class="sxs-lookup"><span data-stu-id="e64f0-964">String</span></span>||<span data-ttu-id="e64f0-p155">指示附件的类型。必须是文件附件的 `file` 或项目附件的 `item`。</span><span class="sxs-lookup"><span data-stu-id="e64f0-p155">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span>|
|`formData.attachments.name`|<span data-ttu-id="e64f0-967">字符串</span><span class="sxs-lookup"><span data-stu-id="e64f0-967">String</span></span>||<span data-ttu-id="e64f0-968">一个包含附件的名称的字符串，最多包含 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="e64f0-968">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
|`formData.attachments.url`|<span data-ttu-id="e64f0-969">String</span><span class="sxs-lookup"><span data-stu-id="e64f0-969">String</span></span>||<span data-ttu-id="e64f0-p156">仅在将 `type` 设置为 `file` 时使用。文件的位置的 URI。</span><span class="sxs-lookup"><span data-stu-id="e64f0-p156">Only used if `type` is set to `file`. The URI of the location for the file.</span></span>|
|`formData.attachments.isInline`|<span data-ttu-id="e64f0-972">布尔</span><span class="sxs-lookup"><span data-stu-id="e64f0-972">Boolean</span></span>||<span data-ttu-id="e64f0-p157">仅在将 `type` 设置为 `file` 时使用。如果为 `true`，则表示附件将在邮件正文中内联显示，并且不应显示在附件列表中。</span><span class="sxs-lookup"><span data-stu-id="e64f0-p157">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`formData.attachments.itemId`|<span data-ttu-id="e64f0-975">String</span><span class="sxs-lookup"><span data-stu-id="e64f0-975">String</span></span>||<span data-ttu-id="e64f0-p158">仅在将 `type` 设置为 `item` 时使用。附件的 EWS 项目 ID。字符串最长为 100 个字符。</span><span class="sxs-lookup"><span data-stu-id="e64f0-p158">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span>|
|`callback`|<span data-ttu-id="e64f0-979">函数</span><span class="sxs-lookup"><span data-stu-id="e64f0-979">function</span></span>|<span data-ttu-id="e64f0-980">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="e64f0-980">&lt;optional&gt;</span></span>|<span data-ttu-id="e64f0-981">方法完成后，使用单个参数 `asyncResult`（一个 [AsyncResult](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="e64f0-981">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e64f0-982">要求</span><span class="sxs-lookup"><span data-stu-id="e64f0-982">Requirements</span></span>

|<span data-ttu-id="e64f0-983">要求</span><span class="sxs-lookup"><span data-stu-id="e64f0-983">Requirement</span></span>|<span data-ttu-id="e64f0-984">值</span><span class="sxs-lookup"><span data-stu-id="e64f0-984">Value</span></span>|
|---|---|
|[<span data-ttu-id="e64f0-985">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="e64f0-985">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e64f0-986">1.0</span><span class="sxs-lookup"><span data-stu-id="e64f0-986">1.0</span></span>|
|[<span data-ttu-id="e64f0-987">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="e64f0-987">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e64f0-988">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e64f0-988">ReadItem</span></span>|
|[<span data-ttu-id="e64f0-989">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="e64f0-989">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="e64f0-990">阅读</span><span class="sxs-lookup"><span data-stu-id="e64f0-990">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="e64f0-991">示例</span><span class="sxs-lookup"><span data-stu-id="e64f0-991">Examples</span></span>

<span data-ttu-id="e64f0-992">以下代码将一个字符串传递到 `displayReplyForm` 函数。</span><span class="sxs-lookup"><span data-stu-id="e64f0-992">The following code passes a string to the `displayReplyForm` function.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="e64f0-993">使用空白正文答复。</span><span class="sxs-lookup"><span data-stu-id="e64f0-993">Reply with an empty body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="e64f0-994">仅使用正文答复。</span><span class="sxs-lookup"><span data-stu-id="e64f0-994">Reply with just a body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="e64f0-995">使用正文和文件附件答复。</span><span class="sxs-lookup"><span data-stu-id="e64f0-995">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="e64f0-996">使用正文和项目附件答复。</span><span class="sxs-lookup"><span data-stu-id="e64f0-996">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="e64f0-997">使用正文、文件附件、项目附件和回调答复。</span><span class="sxs-lookup"><span data-stu-id="e64f0-997">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="getattachmentcontentasyncattachmentid-options-callback--attachmentcontentjavascriptapioutlookofficeattachmentcontent"></a><span data-ttu-id="e64f0-998">getAttachmentContentAsync(attachmentId, [options], callback) → [AttachmentContent](/javascript/api/outlook/office.attachmentcontent)</span><span class="sxs-lookup"><span data-stu-id="e64f0-998">getAttachmentContentAsync(attachmentId, [options], callback) → [AttachmentContent](/javascript/api/outlook/office.attachmentcontent)</span></span>

<span data-ttu-id="e64f0-999">从消息或约会中获取指定的附件，并将其作为 `AttachmentContent` 对象返回。</span><span class="sxs-lookup"><span data-stu-id="e64f0-999">Gets the specified attachment from a message or appointment and returns it as an `AttachmentContent` object.</span></span>

<span data-ttu-id="e64f0-1000">`getAttachmentContentAsync` 方法获取项目中带指定标识符的附件。</span><span class="sxs-lookup"><span data-stu-id="e64f0-1000">The `getAttachmentContentAsync` method removes the attachment with the specified identifier from the item.</span></span> <span data-ttu-id="e64f0-1001">作为最佳做法，应使用标识符检索同一会话中的附件，在该会话中，使用 `getAttachmentsAsync` 或 `item.attachments` 调用检索附件 ID。</span><span class="sxs-lookup"><span data-stu-id="e64f0-1001">As a best practice, you should use the identifier to retrieve an attachment in the same session that the attachmentIds were retrieved with the `getAttachmentsAsync` or `item.attachments` call.</span></span> <span data-ttu-id="e64f0-1002">在 Outlook Web App 和适用于设备的 OWA 中，附件标识符只在同一个会话中才有效。</span><span class="sxs-lookup"><span data-stu-id="e64f0-1002">In Outlook Web App and OWA for Devices, the attachment identifier is valid only within the same session.</span></span> <span data-ttu-id="e64f0-1003">当用户关闭应用，或者如果用户开始在内嵌窗体中撰写，则随后弹出的窗体中继续在单独的窗口撰写时，会话即结束。</span><span class="sxs-lookup"><span data-stu-id="e64f0-1003">A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="e64f0-1004">参数：</span><span class="sxs-lookup"><span data-stu-id="e64f0-1004">Parameters:</span></span>

|<span data-ttu-id="e64f0-1005">名称</span><span class="sxs-lookup"><span data-stu-id="e64f0-1005">Name</span></span>|<span data-ttu-id="e64f0-1006">类型</span><span class="sxs-lookup"><span data-stu-id="e64f0-1006">Type</span></span>|<span data-ttu-id="e64f0-1007">属性</span><span class="sxs-lookup"><span data-stu-id="e64f0-1007">Attributes</span></span>|<span data-ttu-id="e64f0-1008">说明</span><span class="sxs-lookup"><span data-stu-id="e64f0-1008">Description</span></span>|
|---|---|---|---|
|`attachmentId`|<span data-ttu-id="e64f0-1009">字符串</span><span class="sxs-lookup"><span data-stu-id="e64f0-1009">String</span></span>||<span data-ttu-id="e64f0-1010">要获取的附件的标识符。</span><span class="sxs-lookup"><span data-stu-id="e64f0-1010">The identifier of the attachment you want to get.</span></span> <span data-ttu-id="e64f0-1011">字符串的最大长度为 100 个字符。</span><span class="sxs-lookup"><span data-stu-id="e64f0-1011">The maximum length of the string is 100 characters.</span></span>|
|`options`|<span data-ttu-id="e64f0-1012">对象</span><span class="sxs-lookup"><span data-stu-id="e64f0-1012">Object</span></span>|<span data-ttu-id="e64f0-1013">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="e64f0-1013">&lt;optional&gt;</span></span>|<span data-ttu-id="e64f0-1014">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="e64f0-1014">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="e64f0-1015">对象</span><span class="sxs-lookup"><span data-stu-id="e64f0-1015">Object</span></span>|<span data-ttu-id="e64f0-1016">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="e64f0-1016">&lt;optional&gt;</span></span>|<span data-ttu-id="e64f0-1017">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="e64f0-1017">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="e64f0-1018">函数</span><span class="sxs-lookup"><span data-stu-id="e64f0-1018">function</span></span>|<span data-ttu-id="e64f0-1019">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="e64f0-1019">&lt;optional&gt;</span></span>|<span data-ttu-id="e64f0-1020">方法完成后，使用单个参数 `asyncResult`（一个 [AsyncResult](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="e64f0-1020">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e64f0-1021">要求</span><span class="sxs-lookup"><span data-stu-id="e64f0-1021">Requirements</span></span>

|<span data-ttu-id="e64f0-1022">要求</span><span class="sxs-lookup"><span data-stu-id="e64f0-1022">Requirement</span></span>|<span data-ttu-id="e64f0-1023">值</span><span class="sxs-lookup"><span data-stu-id="e64f0-1023">Value</span></span>|
|---|---|
|[<span data-ttu-id="e64f0-1024">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="e64f0-1024">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e64f0-1025">预览</span><span class="sxs-lookup"><span data-stu-id="e64f0-1025">Preview</span></span>|
|[<span data-ttu-id="e64f0-1026">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="e64f0-1026">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e64f0-1027">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e64f0-1027">ReadItem</span></span>|
|[<span data-ttu-id="e64f0-1028">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="e64f0-1028">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="e64f0-1029">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="e64f0-1029">Compose or read</span></span>|

##### <a name="returns"></a><span data-ttu-id="e64f0-1030">返回：</span><span class="sxs-lookup"><span data-stu-id="e64f0-1030">Returns:</span></span>

<span data-ttu-id="e64f0-1031">类型：[AttachmentContent](/javascript/api/outlook/office.attachmentcontent)</span><span class="sxs-lookup"><span data-stu-id="e64f0-1031">Type: [AttachmentContent](/javascript/api/outlook/office.attachmentcontent)</span></span>

##### <a name="example"></a><span data-ttu-id="e64f0-1032">示例</span><span class="sxs-lookup"><span data-stu-id="e64f0-1032">Example</span></span>

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

#### <a name="getattachmentsasyncoptions-callback--arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetails"></a><span data-ttu-id="e64f0-1033">getAttachmentsAsync([options], callback) → Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="e64f0-1033">getAttachmentsAsync([options], callback) → Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span></span>

<span data-ttu-id="e64f0-1034">获取项目的附件作为数组。</span><span class="sxs-lookup"><span data-stu-id="e64f0-1034">Gets the item's attachments as an array.</span></span> <span data-ttu-id="e64f0-1035">仅限撰写模式。</span><span class="sxs-lookup"><span data-stu-id="e64f0-1035">Compose mode only.</span></span>

##### <a name="parameters"></a><span data-ttu-id="e64f0-1036">参数：</span><span class="sxs-lookup"><span data-stu-id="e64f0-1036">Parameters:</span></span>

|<span data-ttu-id="e64f0-1037">名称</span><span class="sxs-lookup"><span data-stu-id="e64f0-1037">Name</span></span>|<span data-ttu-id="e64f0-1038">类型</span><span class="sxs-lookup"><span data-stu-id="e64f0-1038">Type</span></span>|<span data-ttu-id="e64f0-1039">属性</span><span class="sxs-lookup"><span data-stu-id="e64f0-1039">Attributes</span></span>|<span data-ttu-id="e64f0-1040">说明</span><span class="sxs-lookup"><span data-stu-id="e64f0-1040">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="e64f0-1041">Object</span><span class="sxs-lookup"><span data-stu-id="e64f0-1041">Object</span></span>|<span data-ttu-id="e64f0-1042">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="e64f0-1042">&lt;optional&gt;</span></span>|<span data-ttu-id="e64f0-1043">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="e64f0-1043">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="e64f0-1044">对象</span><span class="sxs-lookup"><span data-stu-id="e64f0-1044">Object</span></span>|<span data-ttu-id="e64f0-1045">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="e64f0-1045">&lt;optional&gt;</span></span>|<span data-ttu-id="e64f0-1046">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="e64f0-1046">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="e64f0-1047">函数</span><span class="sxs-lookup"><span data-stu-id="e64f0-1047">function</span></span>|<span data-ttu-id="e64f0-1048">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="e64f0-1048">&lt;optional&gt;</span></span>|<span data-ttu-id="e64f0-1049">方法完成后，使用单个参数 `asyncResult`（一个 [AsyncResult](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="e64f0-1049">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e64f0-1050">要求</span><span class="sxs-lookup"><span data-stu-id="e64f0-1050">Requirements</span></span>

|<span data-ttu-id="e64f0-1051">要求</span><span class="sxs-lookup"><span data-stu-id="e64f0-1051">Requirement</span></span>|<span data-ttu-id="e64f0-1052">值</span><span class="sxs-lookup"><span data-stu-id="e64f0-1052">Value</span></span>|
|---|---|
|[<span data-ttu-id="e64f0-1053">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="e64f0-1053">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e64f0-1054">预览</span><span class="sxs-lookup"><span data-stu-id="e64f0-1054">Preview</span></span>|
|[<span data-ttu-id="e64f0-1055">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="e64f0-1055">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e64f0-1056">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e64f0-1056">ReadItem</span></span>|
|[<span data-ttu-id="e64f0-1057">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="e64f0-1057">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="e64f0-1058">撰写</span><span class="sxs-lookup"><span data-stu-id="e64f0-1058">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="e64f0-1059">返回：</span><span class="sxs-lookup"><span data-stu-id="e64f0-1059">Returns:</span></span>

<span data-ttu-id="e64f0-1060">类型：Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="e64f0-1060">Type: Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)></span></span>

##### <a name="example"></a><span data-ttu-id="e64f0-1061">示例</span><span class="sxs-lookup"><span data-stu-id="e64f0-1061">Example</span></span>

<span data-ttu-id="e64f0-1062">以下示例使用当前项目上所有附件的详细信息构成 HTML 字符串。</span><span class="sxs-lookup"><span data-stu-id="e64f0-1062">The following code builds an HTML string with details of all attachments on the current item.</span></span>

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

#### <a name="getentities--entitiesjavascriptapioutlookofficeentities"></a><span data-ttu-id="e64f0-1063">getEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="e64f0-1063">getEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span></span>

<span data-ttu-id="e64f0-1064">获取在所选项目的正文中找到的实体。</span><span class="sxs-lookup"><span data-stu-id="e64f0-1064">Gets the entities found in the selected item.</span></span>

> [!NOTE]
> <span data-ttu-id="e64f0-1065">在 Outlook for iOS 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="e64f0-1065">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="e64f0-1066">要求</span><span class="sxs-lookup"><span data-stu-id="e64f0-1066">Requirements</span></span>

|<span data-ttu-id="e64f0-1067">要求</span><span class="sxs-lookup"><span data-stu-id="e64f0-1067">Requirement</span></span>|<span data-ttu-id="e64f0-1068">值</span><span class="sxs-lookup"><span data-stu-id="e64f0-1068">Value</span></span>|
|---|---|
|[<span data-ttu-id="e64f0-1069">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="e64f0-1069">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e64f0-1070">1.0</span><span class="sxs-lookup"><span data-stu-id="e64f0-1070">1.0</span></span>|
|[<span data-ttu-id="e64f0-1071">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="e64f0-1071">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e64f0-1072">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e64f0-1072">ReadItem</span></span>|
|[<span data-ttu-id="e64f0-1073">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="e64f0-1073">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="e64f0-1074">阅读</span><span class="sxs-lookup"><span data-stu-id="e64f0-1074">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="e64f0-1075">返回：</span><span class="sxs-lookup"><span data-stu-id="e64f0-1075">Returns:</span></span>

<span data-ttu-id="e64f0-1076">类型：[Entities](/javascript/api/outlook/office.entities)</span><span class="sxs-lookup"><span data-stu-id="e64f0-1076">Type: [Entities](/javascript/api/outlook/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="e64f0-1077">示例</span><span class="sxs-lookup"><span data-stu-id="e64f0-1077">Example</span></span>

<span data-ttu-id="e64f0-1078">以下示例访问当前项目的正文中的联系人实体。</span><span class="sxs-lookup"><span data-stu-id="e64f0-1078">The following example accesses the contacts entities on the current item.</span></span>

```javascript
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlookofficecontactmeetingsuggestionjavascriptapioutlookofficemeetingsuggestionphonenumberjavascriptapioutlookofficephonenumbertasksuggestionjavascriptapioutlookofficetasksuggestion"></a><span data-ttu-id="e64f0-1079">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="e64f0-1079">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span></span>

<span data-ttu-id="e64f0-1080">获取所选项目的正文中找到的指定实体类型的所有实体的数组。</span><span class="sxs-lookup"><span data-stu-id="e64f0-1080">Gets an array of all the entities of the specified entity type found in the selected item.</span></span>

> [!NOTE]
> <span data-ttu-id="e64f0-1081">在 Outlook for iOS 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="e64f0-1081">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="e64f0-1082">参数：</span><span class="sxs-lookup"><span data-stu-id="e64f0-1082">Parameters:</span></span>

|<span data-ttu-id="e64f0-1083">名称</span><span class="sxs-lookup"><span data-stu-id="e64f0-1083">Name</span></span>|<span data-ttu-id="e64f0-1084">类型</span><span class="sxs-lookup"><span data-stu-id="e64f0-1084">Type</span></span>|<span data-ttu-id="e64f0-1085">说明</span><span class="sxs-lookup"><span data-stu-id="e64f0-1085">Description</span></span>|
|---|---|---|
|`entityType`|[<span data-ttu-id="e64f0-1086">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="e64f0-1086">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook/office.mailboxenums.entitytype)|<span data-ttu-id="e64f0-1087">EntityType 枚举值之一。</span><span class="sxs-lookup"><span data-stu-id="e64f0-1087">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e64f0-1088">要求</span><span class="sxs-lookup"><span data-stu-id="e64f0-1088">Requirements</span></span>

|<span data-ttu-id="e64f0-1089">要求</span><span class="sxs-lookup"><span data-stu-id="e64f0-1089">Requirement</span></span>|<span data-ttu-id="e64f0-1090">值</span><span class="sxs-lookup"><span data-stu-id="e64f0-1090">Value</span></span>|
|---|---|
|[<span data-ttu-id="e64f0-1091">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="e64f0-1091">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e64f0-1092">1.0</span><span class="sxs-lookup"><span data-stu-id="e64f0-1092">1.0</span></span>|
|[<span data-ttu-id="e64f0-1093">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="e64f0-1093">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e64f0-1094">受限</span><span class="sxs-lookup"><span data-stu-id="e64f0-1094">Restricted</span></span>|
|[<span data-ttu-id="e64f0-1095">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="e64f0-1095">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="e64f0-1096">阅读</span><span class="sxs-lookup"><span data-stu-id="e64f0-1096">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="e64f0-1097">返回：</span><span class="sxs-lookup"><span data-stu-id="e64f0-1097">Returns:</span></span>

<span data-ttu-id="e64f0-1098">如果在 `entityType` 中传递的值不是 `EntityType` 枚举的有效成员，该方法返回 null。</span><span class="sxs-lookup"><span data-stu-id="e64f0-1098">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="e64f0-1099">如果指定类型的任何实体都不存在于该项目的正文中，该方法将返回空数组。</span><span class="sxs-lookup"><span data-stu-id="e64f0-1099">If no entities of the specified type are present on the item, the method returns an empty array.</span></span> <span data-ttu-id="e64f0-1100">否则，返回的数组中对象的类型取决于 `entityType` 参数中请求实体的类型。</span><span class="sxs-lookup"><span data-stu-id="e64f0-1100">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="e64f0-1101">当使用此方法的最低权限级别**受限**时，某些实体类型需要 **ReadItem** 才能进行访问，如下表中所指定。</span><span class="sxs-lookup"><span data-stu-id="e64f0-1101">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

|<span data-ttu-id="e64f0-1102">`entityType` 的值</span><span class="sxs-lookup"><span data-stu-id="e64f0-1102">Value of `entityType`</span></span>|<span data-ttu-id="e64f0-1103">返回的数组中对象的类型</span><span class="sxs-lookup"><span data-stu-id="e64f0-1103">Type of objects in returned array</span></span>|<span data-ttu-id="e64f0-1104">所需权限级别</span><span class="sxs-lookup"><span data-stu-id="e64f0-1104">Required Permission Level</span></span>|
|---|---|---|
|`Address`|<span data-ttu-id="e64f0-1105">String</span><span class="sxs-lookup"><span data-stu-id="e64f0-1105">String</span></span>|<span data-ttu-id="e64f0-1106">**受限**</span><span class="sxs-lookup"><span data-stu-id="e64f0-1106">**Restricted**</span></span>|
|`Contact`|<span data-ttu-id="e64f0-1107">Contact</span><span class="sxs-lookup"><span data-stu-id="e64f0-1107">Contact</span></span>|<span data-ttu-id="e64f0-1108">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="e64f0-1108">**ReadItem**</span></span>|
|`EmailAddress`|<span data-ttu-id="e64f0-1109">String</span><span class="sxs-lookup"><span data-stu-id="e64f0-1109">String</span></span>|<span data-ttu-id="e64f0-1110">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="e64f0-1110">**ReadItem**</span></span>|
|`MeetingSuggestion`|<span data-ttu-id="e64f0-1111">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="e64f0-1111">MeetingSuggestion</span></span>|<span data-ttu-id="e64f0-1112">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="e64f0-1112">**ReadItem**</span></span>|
|`PhoneNumber`|<span data-ttu-id="e64f0-1113">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="e64f0-1113">PhoneNumber</span></span>|<span data-ttu-id="e64f0-1114">**受限**</span><span class="sxs-lookup"><span data-stu-id="e64f0-1114">**Restricted**</span></span>|
|`TaskSuggestion`|<span data-ttu-id="e64f0-1115">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="e64f0-1115">TaskSuggestion</span></span>|<span data-ttu-id="e64f0-1116">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="e64f0-1116">**ReadItem**</span></span>|
|`URL`|<span data-ttu-id="e64f0-1117">String</span><span class="sxs-lookup"><span data-stu-id="e64f0-1117">String</span></span>|<span data-ttu-id="e64f0-1118">**受限**</span><span class="sxs-lookup"><span data-stu-id="e64f0-1118">**Restricted**</span></span>|

<span data-ttu-id="e64f0-1119">类型：Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="e64f0-1119">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span></span>

##### <a name="example"></a><span data-ttu-id="e64f0-1120">示例</span><span class="sxs-lookup"><span data-stu-id="e64f0-1120">Example</span></span>

<span data-ttu-id="e64f0-1121">以下示例显示了如何访问表示当前项目的正文中的邮政地址的字符串数组。</span><span class="sxs-lookup"><span data-stu-id="e64f0-1121">The following example shows how to access an array of strings that represent postal addresses in the subject or body of the current item.</span></span>

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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlookofficecontactmeetingsuggestionjavascriptapioutlookofficemeetingsuggestionphonenumberjavascriptapioutlookofficephonenumbertasksuggestionjavascriptapioutlookofficetasksuggestion"></a><span data-ttu-id="e64f0-1122">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="e64f0-1122">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}</span></span>

<span data-ttu-id="e64f0-1123">返回传递清单 XML 文件中定义的命名筛选器的所选项目中的已知实体。</span><span class="sxs-lookup"><span data-stu-id="e64f0-1123">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="e64f0-1124">在 Outlook for iOS 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="e64f0-1124">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="e64f0-1125">`getFilteredEntitiesByName` 方法返回匹配在具有指定 `FilterName` 元素值的清单 XML 文件中的 [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) 规则元素中定义的正则表达式的实体。</span><span class="sxs-lookup"><span data-stu-id="e64f0-1125">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="e64f0-1126">参数：</span><span class="sxs-lookup"><span data-stu-id="e64f0-1126">Parameters:</span></span>

|<span data-ttu-id="e64f0-1127">名称</span><span class="sxs-lookup"><span data-stu-id="e64f0-1127">Name</span></span>|<span data-ttu-id="e64f0-1128">类型</span><span class="sxs-lookup"><span data-stu-id="e64f0-1128">Type</span></span>|<span data-ttu-id="e64f0-1129">描述</span><span class="sxs-lookup"><span data-stu-id="e64f0-1129">Description</span></span>|
|---|---|---|
|`name`|<span data-ttu-id="e64f0-1130">字符串</span><span class="sxs-lookup"><span data-stu-id="e64f0-1130">String</span></span>|<span data-ttu-id="e64f0-1131">定义筛选器匹配的 `ItemHasKnownEntity` 规则元素的名称。</span><span class="sxs-lookup"><span data-stu-id="e64f0-1131">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e64f0-1132">要求</span><span class="sxs-lookup"><span data-stu-id="e64f0-1132">Requirements</span></span>

|<span data-ttu-id="e64f0-1133">要求</span><span class="sxs-lookup"><span data-stu-id="e64f0-1133">Requirement</span></span>|<span data-ttu-id="e64f0-1134">值</span><span class="sxs-lookup"><span data-stu-id="e64f0-1134">Value</span></span>|
|---|---|
|[<span data-ttu-id="e64f0-1135">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="e64f0-1135">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e64f0-1136">1.0</span><span class="sxs-lookup"><span data-stu-id="e64f0-1136">1.0</span></span>|
|[<span data-ttu-id="e64f0-1137">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="e64f0-1137">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e64f0-1138">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e64f0-1138">ReadItem</span></span>|
|[<span data-ttu-id="e64f0-1139">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="e64f0-1139">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="e64f0-1140">阅读</span><span class="sxs-lookup"><span data-stu-id="e64f0-1140">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="e64f0-1141">返回：</span><span class="sxs-lookup"><span data-stu-id="e64f0-1141">Returns:</span></span>

<span data-ttu-id="e64f0-p163">如果具有匹配 `name` 参数的 `FilterName` 元素值的清单中没有任何 `ItemHasKnownEntity` 元素，则该方法返回 `null`。如果 `name` 参数匹配清单中的 `ItemHasKnownEntity` 元素，但在匹配的当前项目中没有实体，则该方法返回一个空数组。</span><span class="sxs-lookup"><span data-stu-id="e64f0-p163">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="e64f0-1144">类型：Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="e64f0-1144">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))></span></span>

#### <a name="getinitializationcontextasyncoptions-callback"></a><span data-ttu-id="e64f0-1145">getInitializationContextAsync([options], [callback])</span><span class="sxs-lookup"><span data-stu-id="e64f0-1145">getInitializationContextAsync([options], [callback])</span></span>

<span data-ttu-id="e64f0-1146">当[可操作邮件激活](https://docs.microsoft.com/outlook/actionable-messages/invoke-add-in-from-actionable-message)加载项时，获取传递的初始化数据。</span><span class="sxs-lookup"><span data-stu-id="e64f0-1146">Gets initialization data passed when the add-in is [activated by an actionable message](https://docs.microsoft.com/outlook/actionable-messages/invoke-add-in-from-actionable-message).</span></span>

> [!NOTE]
> <span data-ttu-id="e64f0-1147">仅 Outlook 2016 for Windows 或更高版本（高于 16.0.8413.1000 的即点即用版本）和适用于 Office 365 的 Outlook 网页版支持此方法。</span><span class="sxs-lookup"><span data-stu-id="e64f0-1147">Note: This method is only supported by Outlook 2016 for Windows (Click-to-Run versions greater than 16.0.8413.1000) and Outlook on the web for Office 365.</span></span>

##### <a name="parameters"></a><span data-ttu-id="e64f0-1148">参数：</span><span class="sxs-lookup"><span data-stu-id="e64f0-1148">Parameters:</span></span>
|<span data-ttu-id="e64f0-1149">名称</span><span class="sxs-lookup"><span data-stu-id="e64f0-1149">Name</span></span>|<span data-ttu-id="e64f0-1150">类型</span><span class="sxs-lookup"><span data-stu-id="e64f0-1150">Type</span></span>|<span data-ttu-id="e64f0-1151">属性</span><span class="sxs-lookup"><span data-stu-id="e64f0-1151">Attributes</span></span>|<span data-ttu-id="e64f0-1152">说明</span><span class="sxs-lookup"><span data-stu-id="e64f0-1152">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="e64f0-1153">Object</span><span class="sxs-lookup"><span data-stu-id="e64f0-1153">Object</span></span>|<span data-ttu-id="e64f0-1154">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="e64f0-1154">&lt;optional&gt;</span></span>|<span data-ttu-id="e64f0-1155">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="e64f0-1155">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="e64f0-1156">对象</span><span class="sxs-lookup"><span data-stu-id="e64f0-1156">Object</span></span>|<span data-ttu-id="e64f0-1157">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="e64f0-1157">&lt;optional&gt;</span></span>|<span data-ttu-id="e64f0-1158">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="e64f0-1158">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="e64f0-1159">函数</span><span class="sxs-lookup"><span data-stu-id="e64f0-1159">function</span></span>|<span data-ttu-id="e64f0-1160">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="e64f0-1160">&lt;optional&gt;</span></span>|<span data-ttu-id="e64f0-1161">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="e64f0-1161">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="e64f0-1162">成功后，`asyncResult.value` 属性便以字符串形式提供初始化数据。</span><span class="sxs-lookup"><span data-stu-id="e64f0-1162">On success, the intialization data is provided in the `asyncResult.value` property as a string.</span></span><br/><span data-ttu-id="e64f0-1163">如果没有初始化上下文，`asyncResult` 对象包含 `Error` 对象，并将它的 `code` 和 `name` 属性分别设置为 `9020` 和 `GenericResponseError`。</span><span class="sxs-lookup"><span data-stu-id="e64f0-1163">If there is no initialization context, the `asyncResult` object will contain an `Error` object with its `code` property set to `9020` and its `name` property set to `GenericResponseError`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e64f0-1164">要求</span><span class="sxs-lookup"><span data-stu-id="e64f0-1164">Requirements</span></span>

|<span data-ttu-id="e64f0-1165">要求</span><span class="sxs-lookup"><span data-stu-id="e64f0-1165">Requirement</span></span>|<span data-ttu-id="e64f0-1166">值</span><span class="sxs-lookup"><span data-stu-id="e64f0-1166">Value</span></span>|
|---|---|
|[<span data-ttu-id="e64f0-1167">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="e64f0-1167">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e64f0-1168">预览</span><span class="sxs-lookup"><span data-stu-id="e64f0-1168">Preview</span></span>|
|[<span data-ttu-id="e64f0-1169">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="e64f0-1169">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e64f0-1170">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e64f0-1170">ReadItem</span></span>|
|[<span data-ttu-id="e64f0-1171">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="e64f0-1171">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="e64f0-1172">阅读</span><span class="sxs-lookup"><span data-stu-id="e64f0-1172">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="e64f0-1173">示例</span><span class="sxs-lookup"><span data-stu-id="e64f0-1173">Example</span></span>

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

#### <a name="getregexmatches--object"></a><span data-ttu-id="e64f0-1174">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="e64f0-1174">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="e64f0-1175">返回所选项目中匹配在清单 XML 文件中定义的正则表达式的字符串值。</span><span class="sxs-lookup"><span data-stu-id="e64f0-1175">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="e64f0-1176">在 Outlook for iOS 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="e64f0-1176">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="e64f0-p164">`getRegExMatches` 方法返回匹配在清单 XML 文件中的每个 `ItemHasRegularExpressionMatch` 或 `ItemHasKnownEntity` 规则元素中定义的正则表达式的字符串。对于 `ItemHasRegularExpressionMatch` 规则，匹配字符串必须发生在该规则指定的项目的属性中。`PropertyName` 简单类型定义支持的属性。</span><span class="sxs-lookup"><span data-stu-id="e64f0-p164">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="e64f0-1180">例如，考虑一个外接程序清单具有以下 `Rule` 元素：</span><span class="sxs-lookup"><span data-stu-id="e64f0-1180">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="e64f0-1181">从 `getRegExMatches` 返回的对象应有两个属性：`fruits` 和 `veggies`。</span><span class="sxs-lookup"><span data-stu-id="e64f0-1181">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="e64f0-p165">如果在项目的正文属性上指定 `ItemHasRegularExpressionMatch` 规则，则正则表达式应进一步筛选正文，不应尝试返回该项目的整个正文。使用正则表达式（如 `.*`）获取项目的整个正文并不总是返回预期的结果。而是使用 [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) 方法检索整个正文。</span><span class="sxs-lookup"><span data-stu-id="e64f0-p165">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="e64f0-1185">Requirements</span><span class="sxs-lookup"><span data-stu-id="e64f0-1185">Requirements</span></span>

|<span data-ttu-id="e64f0-1186">要求</span><span class="sxs-lookup"><span data-stu-id="e64f0-1186">Requirement</span></span>|<span data-ttu-id="e64f0-1187">值</span><span class="sxs-lookup"><span data-stu-id="e64f0-1187">Value</span></span>|
|---|---|
|[<span data-ttu-id="e64f0-1188">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="e64f0-1188">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e64f0-1189">1.0</span><span class="sxs-lookup"><span data-stu-id="e64f0-1189">1.0</span></span>|
|[<span data-ttu-id="e64f0-1190">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="e64f0-1190">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e64f0-1191">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e64f0-1191">ReadItem</span></span>|
|[<span data-ttu-id="e64f0-1192">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="e64f0-1192">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="e64f0-1193">阅读</span><span class="sxs-lookup"><span data-stu-id="e64f0-1193">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="e64f0-1194">返回：</span><span class="sxs-lookup"><span data-stu-id="e64f0-1194">Returns:</span></span>

<span data-ttu-id="e64f0-p166">一个包含匹配在清单 XML 文件中定义的正则表达式的字符串数组的对象。每个数组的名称等于匹配 `ItemHasRegularExpressionMatch` 规则的 `RegExName` 属性或匹配 `ItemHasKnownEntity` 规则的 `FilterName` 属性的相应值。</span><span class="sxs-lookup"><span data-stu-id="e64f0-p166">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<dl class="param-type"><span data-ttu-id="e64f0-1197">

<dt>
类型</dt>


</span><span class="sxs-lookup"><span data-stu-id="e64f0-1197">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="e64f0-1198">对象</span><span class="sxs-lookup"><span data-stu-id="e64f0-1198">Object</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="e64f0-1199">示例</span><span class="sxs-lookup"><span data-stu-id="e64f0-1199">Example</span></span>

<span data-ttu-id="e64f0-1200">以下示例显示了如何访问正则表达式规则元素 `fruits` 和 `veggies` 的匹配项的数组，这些元素在清单中指定。</span><span class="sxs-lookup"><span data-stu-id="e64f0-1200">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```javascript
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veges = allMatches.veggies;
```

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="e64f0-1201">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span><span class="sxs-lookup"><span data-stu-id="e64f0-1201">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="e64f0-1202">返回匹配在清单 XML 文件中定义的命名正则表达式的所选项目中的字符串值。</span><span class="sxs-lookup"><span data-stu-id="e64f0-1202">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="e64f0-1203">在 Outlook for iOS 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="e64f0-1203">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="e64f0-1204">`getRegExMatchesByName` 方法返回匹配在具有指定 `RegExName` 元素值的清单 XML 文件中的 `ItemHasRegularExpressionMatch` 规则元素中定义的正则表达式的字符串。</span><span class="sxs-lookup"><span data-stu-id="e64f0-1204">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="e64f0-p167">如果在项目的正文属性上指定 `ItemHasRegularExpressionMatch` 规则，则正则表达式应进一步筛选正文，不应尝试返回该项目的整个正文。使用正则表达式（如 `.*`）获取项目的整个正文并不总是返回预期的结果。</span><span class="sxs-lookup"><span data-stu-id="e64f0-p167">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="e64f0-1207">参数：</span><span class="sxs-lookup"><span data-stu-id="e64f0-1207">Parameters:</span></span>

|<span data-ttu-id="e64f0-1208">名称</span><span class="sxs-lookup"><span data-stu-id="e64f0-1208">Name</span></span>|<span data-ttu-id="e64f0-1209">类型</span><span class="sxs-lookup"><span data-stu-id="e64f0-1209">Type</span></span>|<span data-ttu-id="e64f0-1210">描述</span><span class="sxs-lookup"><span data-stu-id="e64f0-1210">Description</span></span>|
|---|---|---|
|`name`|<span data-ttu-id="e64f0-1211">字符串</span><span class="sxs-lookup"><span data-stu-id="e64f0-1211">String</span></span>|<span data-ttu-id="e64f0-1212">定义筛选器匹配的 `ItemHasRegularExpressionMatch` 规则元素的名称。</span><span class="sxs-lookup"><span data-stu-id="e64f0-1212">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e64f0-1213">要求</span><span class="sxs-lookup"><span data-stu-id="e64f0-1213">Requirements</span></span>

|<span data-ttu-id="e64f0-1214">要求</span><span class="sxs-lookup"><span data-stu-id="e64f0-1214">Requirement</span></span>|<span data-ttu-id="e64f0-1215">值</span><span class="sxs-lookup"><span data-stu-id="e64f0-1215">Value</span></span>|
|---|---|
|[<span data-ttu-id="e64f0-1216">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="e64f0-1216">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e64f0-1217">1.0</span><span class="sxs-lookup"><span data-stu-id="e64f0-1217">1.0</span></span>|
|[<span data-ttu-id="e64f0-1218">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="e64f0-1218">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e64f0-1219">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e64f0-1219">ReadItem</span></span>|
|[<span data-ttu-id="e64f0-1220">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="e64f0-1220">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="e64f0-1221">阅读</span><span class="sxs-lookup"><span data-stu-id="e64f0-1221">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="e64f0-1222">返回：</span><span class="sxs-lookup"><span data-stu-id="e64f0-1222">Returns:</span></span>

<span data-ttu-id="e64f0-1223">一个包含匹配在清单 XML 文件中定义的正则表达式的字符串的数组。</span><span class="sxs-lookup"><span data-stu-id="e64f0-1223">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<dl class="param-type"><span data-ttu-id="e64f0-1224">

<dt>类型</dt>

</span><span class="sxs-lookup"><span data-stu-id="e64f0-1224">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="e64f0-1225">Array.< String ></span><span class="sxs-lookup"><span data-stu-id="e64f0-1225">Array.< String ></span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="e64f0-1226">示例</span><span class="sxs-lookup"><span data-stu-id="e64f0-1226">Example</span></span>

```javascript
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

####  <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="e64f0-1227">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="e64f0-1227">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="e64f0-1228">以异步方式返回邮件的主题或正文中选定的数据。</span><span class="sxs-lookup"><span data-stu-id="e64f0-1228">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="e64f0-p168">如果没有选定内容，但光标位于正文或主题中，此方法将会为所选数据返回 null。如果选定的是字段，而不是正文或主题，则此方法返回 `InvalidSelection` 错误。</span><span class="sxs-lookup"><span data-stu-id="e64f0-p168">If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="e64f0-1231">参数：</span><span class="sxs-lookup"><span data-stu-id="e64f0-1231">Parameters:</span></span>

|<span data-ttu-id="e64f0-1232">名称</span><span class="sxs-lookup"><span data-stu-id="e64f0-1232">Name</span></span>|<span data-ttu-id="e64f0-1233">类型</span><span class="sxs-lookup"><span data-stu-id="e64f0-1233">Type</span></span>|<span data-ttu-id="e64f0-1234">属性</span><span class="sxs-lookup"><span data-stu-id="e64f0-1234">Attributes</span></span>|<span data-ttu-id="e64f0-1235">说明</span><span class="sxs-lookup"><span data-stu-id="e64f0-1235">Description</span></span>|
|---|---|---|---|
|`coercionType`|[<span data-ttu-id="e64f0-1236">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="e64f0-1236">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="e64f0-p169">请求数据的格式。如果为文本，则此方法返回纯文本作为字符串，删除任何显示的 HTML 标记。如果为 HTML，则此方法返回所选文本，不论是纯文本还是 HTML。</span><span class="sxs-lookup"><span data-stu-id="e64f0-p169">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`|<span data-ttu-id="e64f0-1240">对象</span><span class="sxs-lookup"><span data-stu-id="e64f0-1240">Object</span></span>|<span data-ttu-id="e64f0-1241">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="e64f0-1241">&lt;optional&gt;</span></span>|<span data-ttu-id="e64f0-1242">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="e64f0-1242">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="e64f0-1243">对象</span><span class="sxs-lookup"><span data-stu-id="e64f0-1243">Object</span></span>|<span data-ttu-id="e64f0-1244">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="e64f0-1244">&lt;optional&gt;</span></span>|<span data-ttu-id="e64f0-1245">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="e64f0-1245">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="e64f0-1246">函数</span><span class="sxs-lookup"><span data-stu-id="e64f0-1246">function</span></span>||<span data-ttu-id="e64f0-1247">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="e64f0-1247">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="e64f0-1248">若要从回调方法访问所选数据，请调用 `asyncResult.value.data`。</span><span class="sxs-lookup"><span data-stu-id="e64f0-1248">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="e64f0-1249">若要访问选定内容的源属性，请调用 `asyncResult.value.sourceProperty`，这将为 `body` 或 `subject`。</span><span class="sxs-lookup"><span data-stu-id="e64f0-1249">To access the source property that the selection comes from, call , which will be either  or .|</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e64f0-1250">要求</span><span class="sxs-lookup"><span data-stu-id="e64f0-1250">Requirements</span></span>

|<span data-ttu-id="e64f0-1251">要求</span><span class="sxs-lookup"><span data-stu-id="e64f0-1251">Requirement</span></span>|<span data-ttu-id="e64f0-1252">值</span><span class="sxs-lookup"><span data-stu-id="e64f0-1252">Value</span></span>|
|---|---|
|[<span data-ttu-id="e64f0-1253">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="e64f0-1253">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e64f0-1254">1.2</span><span class="sxs-lookup"><span data-stu-id="e64f0-1254">1.2</span></span>|
|[<span data-ttu-id="e64f0-1255">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="e64f0-1255">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e64f0-1256">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="e64f0-1256">ReadWriteItem</span></span>|
|[<span data-ttu-id="e64f0-1257">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="e64f0-1257">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="e64f0-1258">撰写</span><span class="sxs-lookup"><span data-stu-id="e64f0-1258">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="e64f0-1259">返回：</span><span class="sxs-lookup"><span data-stu-id="e64f0-1259">Returns:</span></span>

<span data-ttu-id="e64f0-1260">作为字符串的所选数据的格式由 `coercionType` 确定。</span><span class="sxs-lookup"><span data-stu-id="e64f0-1260">The selected data as a string with format determined by `coercionType`.</span></span>

<dl class="param-type"><span data-ttu-id="e64f0-1261">

<dt>
类型</dt>


</span><span class="sxs-lookup"><span data-stu-id="e64f0-1261">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="e64f0-1262">字符串</span><span class="sxs-lookup"><span data-stu-id="e64f0-1262">String</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="e64f0-1263">示例</span><span class="sxs-lookup"><span data-stu-id="e64f0-1263">Example</span></span>

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

#### <a name="getselectedentities--entitiesjavascriptapioutlookofficeentities"></a><span data-ttu-id="e64f0-1264">getSelectedEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="e64f0-1264">getSelectedEntities() → {[Entities](/javascript/api/outlook/office.entities)}</span></span>

<span data-ttu-id="e64f0-p171">获取在用户已选择的突出显示匹配项中找到的实体。突出显示匹配项适用于[上下文加载项](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins)。</span><span class="sxs-lookup"><span data-stu-id="e64f0-p171">Gets the entities found in a highlighted match a user has selected. Highlighted matches apply to [contextual add-ins](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="e64f0-1267">在 Outlook for iOS 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="e64f0-1267">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="e64f0-1268">要求</span><span class="sxs-lookup"><span data-stu-id="e64f0-1268">Requirements</span></span>

|<span data-ttu-id="e64f0-1269">要求</span><span class="sxs-lookup"><span data-stu-id="e64f0-1269">Requirement</span></span>|<span data-ttu-id="e64f0-1270">值</span><span class="sxs-lookup"><span data-stu-id="e64f0-1270">Value</span></span>|
|---|---|
|[<span data-ttu-id="e64f0-1271">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="e64f0-1271">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e64f0-1272">1.6</span><span class="sxs-lookup"><span data-stu-id="e64f0-1272">-16</span></span>|
|[<span data-ttu-id="e64f0-1273">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="e64f0-1273">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e64f0-1274">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e64f0-1274">ReadItem</span></span>|
|[<span data-ttu-id="e64f0-1275">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="e64f0-1275">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="e64f0-1276">阅读</span><span class="sxs-lookup"><span data-stu-id="e64f0-1276">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="e64f0-1277">返回：</span><span class="sxs-lookup"><span data-stu-id="e64f0-1277">Returns:</span></span>

<span data-ttu-id="e64f0-1278">类型：[Entities](/javascript/api/outlook/office.entities)</span><span class="sxs-lookup"><span data-stu-id="e64f0-1278">Type: [Entities](/javascript/api/outlook/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="e64f0-1279">示例</span><span class="sxs-lookup"><span data-stu-id="e64f0-1279">Example</span></span>

<span data-ttu-id="e64f0-1280">以下示例访问用户选择的突出显示匹配项中的地址实体。</span><span class="sxs-lookup"><span data-stu-id="e64f0-1280">The following example accesses the addresses entities in the highlighted match selected by the user.</span></span>

```javascript
var contacts = Office.context.mailbox.item.getSelectedEntities().addresses;
```

#### <a name="getselectedregexmatches--object"></a><span data-ttu-id="e64f0-1281">getSelectedRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="e64f0-1281">getSelectedRegExMatches() → {Object}</span></span>

<span data-ttu-id="e64f0-p172">返回突出显示匹配项中匹配在清单 XML 文件中定义的正则表达式的字符串值。突出显示匹配项适用于[上下文外接程序](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins)。</span><span class="sxs-lookup"><span data-stu-id="e64f0-p172">Returns string values in a highlighted match that match the regular expressions defined in the manifest XML file. Highlighted matches apply to [contextual add-ins](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="e64f0-1284">在 Outlook for iOS 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="e64f0-1284">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="e64f0-p173">`getSelectedRegExMatches` 方法返回匹配在清单 XML 文件中的每个 `ItemHasRegularExpressionMatch` 或 `ItemHasKnownEntity` 规则元素中定义的正则表达式的字符串。对于 `ItemHasRegularExpressionMatch` 规则，匹配字符串必须发生在该规则指定的项目的属性中。`PropertyName` 简单类型定义支持的属性。</span><span class="sxs-lookup"><span data-stu-id="e64f0-p173">The `getSelectedRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="e64f0-1288">例如，考虑一个外接程序清单具有以下 `Rule` 元素：</span><span class="sxs-lookup"><span data-stu-id="e64f0-1288">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="e64f0-1289">从 `getRegExMatches` 返回的对象应有两个属性：`fruits` 和 `veggies`。</span><span class="sxs-lookup"><span data-stu-id="e64f0-1289">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="e64f0-p174">如果在项目的正文属性上指定 `ItemHasRegularExpressionMatch` 规则，则正则表达式应进一步筛选正文，不应尝试返回该项目的整个正文。使用正则表达式（如 `.*`）获取项目的整个正文并不总是返回预期的结果。而是使用 [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) 方法检索整个正文。</span><span class="sxs-lookup"><span data-stu-id="e64f0-p174">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="e64f0-1293">Requirements</span><span class="sxs-lookup"><span data-stu-id="e64f0-1293">Requirements</span></span>

|<span data-ttu-id="e64f0-1294">要求</span><span class="sxs-lookup"><span data-stu-id="e64f0-1294">Requirement</span></span>|<span data-ttu-id="e64f0-1295">值</span><span class="sxs-lookup"><span data-stu-id="e64f0-1295">Value</span></span>|
|---|---|
|[<span data-ttu-id="e64f0-1296">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="e64f0-1296">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e64f0-1297">1.6</span><span class="sxs-lookup"><span data-stu-id="e64f0-1297">-16</span></span>|
|[<span data-ttu-id="e64f0-1298">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="e64f0-1298">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e64f0-1299">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e64f0-1299">ReadItem</span></span>|
|[<span data-ttu-id="e64f0-1300">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="e64f0-1300">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="e64f0-1301">阅读</span><span class="sxs-lookup"><span data-stu-id="e64f0-1301">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="e64f0-1302">返回：</span><span class="sxs-lookup"><span data-stu-id="e64f0-1302">Returns:</span></span>

<span data-ttu-id="e64f0-p175">一个包含匹配在清单 XML 文件中定义的正则表达式的字符串数组的对象。每个数组的名称等于匹配 `ItemHasRegularExpressionMatch` 规则的 `RegExName` 属性或匹配 `ItemHasKnownEntity` 规则的 `FilterName` 属性的相应值。</span><span class="sxs-lookup"><span data-stu-id="e64f0-p175">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

##### <a name="example"></a><span data-ttu-id="e64f0-1305">示例</span><span class="sxs-lookup"><span data-stu-id="e64f0-1305">Example</span></span>

<span data-ttu-id="e64f0-1306">以下示例显示了如何访问正则表达式规则元素 `fruits` 和 `veggies` 的匹配项的数组，这些元素在清单中指定。</span><span class="sxs-lookup"><span data-stu-id="e64f0-1306">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```javascript
var selectedMatches = Office.context.mailbox.item.getSelectedRegExMatches();
var fruits = selectedMatches.fruits;
var veggies = selectedMatches.veggies;
```

#### <a name="getsharedpropertiesasyncoptions-callback"></a><span data-ttu-id="e64f0-1307">getSharedPropertiesAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="e64f0-1307">getSharedPropertiesAsync([options], callback)</span></span>

<span data-ttu-id="e64f0-1308">获取共享文件夹、日历或邮箱中所选约会或邮件的属性。</span><span class="sxs-lookup"><span data-stu-id="e64f0-1308">Gets the properties of the selected appointment or message in a shared folder, calendar, or mailbox.</span></span>

##### <a name="parameters"></a><span data-ttu-id="e64f0-1309">参数：</span><span class="sxs-lookup"><span data-stu-id="e64f0-1309">Parameters:</span></span>

|<span data-ttu-id="e64f0-1310">名称</span><span class="sxs-lookup"><span data-stu-id="e64f0-1310">Name</span></span>|<span data-ttu-id="e64f0-1311">类型</span><span class="sxs-lookup"><span data-stu-id="e64f0-1311">Type</span></span>|<span data-ttu-id="e64f0-1312">属性</span><span class="sxs-lookup"><span data-stu-id="e64f0-1312">Attributes</span></span>|<span data-ttu-id="e64f0-1313">说明</span><span class="sxs-lookup"><span data-stu-id="e64f0-1313">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="e64f0-1314">Object</span><span class="sxs-lookup"><span data-stu-id="e64f0-1314">Object</span></span>|<span data-ttu-id="e64f0-1315">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="e64f0-1315">&lt;optional&gt;</span></span>|<span data-ttu-id="e64f0-1316">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="e64f0-1316">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="e64f0-1317">对象</span><span class="sxs-lookup"><span data-stu-id="e64f0-1317">Object</span></span>|<span data-ttu-id="e64f0-1318">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="e64f0-1318">&lt;optional&gt;</span></span>|<span data-ttu-id="e64f0-1319">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="e64f0-1319">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="e64f0-1320">函数</span><span class="sxs-lookup"><span data-stu-id="e64f0-1320">function</span></span>||<span data-ttu-id="e64f0-1321">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="e64f0-1321">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="e64f0-1322">共享属性作为 `asyncResult.value` 属性中的 [`SharedProperties`](/javascript/api/outlook/office.sharedproperties) 对象提供。</span><span class="sxs-lookup"><span data-stu-id="e64f0-1322">The custom properties are provided as a [`SharedProperties`](/javascript/api/outlook/office.sharedproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="e64f0-1323">此对象可用于获取项目的共享属性。</span><span class="sxs-lookup"><span data-stu-id="e64f0-1323">This object can be used to get the item's shared properties.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e64f0-1324">要求</span><span class="sxs-lookup"><span data-stu-id="e64f0-1324">Requirements</span></span>

|<span data-ttu-id="e64f0-1325">要求</span><span class="sxs-lookup"><span data-stu-id="e64f0-1325">Requirement</span></span>|<span data-ttu-id="e64f0-1326">值</span><span class="sxs-lookup"><span data-stu-id="e64f0-1326">Value</span></span>|
|---|---|
|[<span data-ttu-id="e64f0-1327">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="e64f0-1327">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e64f0-1328">预览</span><span class="sxs-lookup"><span data-stu-id="e64f0-1328">Preview</span></span>|
|[<span data-ttu-id="e64f0-1329">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="e64f0-1329">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e64f0-1330">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e64f0-1330">ReadItem</span></span>|
|[<span data-ttu-id="e64f0-1331">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="e64f0-1331">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="e64f0-1332">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="e64f0-1332">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="e64f0-1333">示例</span><span class="sxs-lookup"><span data-stu-id="e64f0-1333">Example</span></span>

```js
Office.context.mailbox.item.getSharedPropertiesAsync(callback);
function callback (asyncResult) {
  var context=asyncResult.context;
  var sharedProperties = asyncResult.value;
}
```

####  <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="e64f0-1334">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="e64f0-1334">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="e64f0-1335">异步加载所选项目上此外接程序的自定义属性。</span><span class="sxs-lookup"><span data-stu-id="e64f0-1335">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="e64f0-p177">自定义属性基于每个应用、每个项目存储为键/值对。此方法在回调中返回 `CustomProperties` 对象，该回调提供访问特定于当前项目和当前外接程序的自定义属性的方法。自定义属性未在项目上加密，因此这不应用作安全存储。</span><span class="sxs-lookup"><span data-stu-id="e64f0-p177">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="e64f0-1339">参数：</span><span class="sxs-lookup"><span data-stu-id="e64f0-1339">Parameters:</span></span>

|<span data-ttu-id="e64f0-1340">名称</span><span class="sxs-lookup"><span data-stu-id="e64f0-1340">Name</span></span>|<span data-ttu-id="e64f0-1341">类型</span><span class="sxs-lookup"><span data-stu-id="e64f0-1341">Type</span></span>|<span data-ttu-id="e64f0-1342">属性</span><span class="sxs-lookup"><span data-stu-id="e64f0-1342">Attributes</span></span>|<span data-ttu-id="e64f0-1343">说明</span><span class="sxs-lookup"><span data-stu-id="e64f0-1343">Description</span></span>|
|---|---|---|---|
|`callback`|<span data-ttu-id="e64f0-1344">函数</span><span class="sxs-lookup"><span data-stu-id="e64f0-1344">function</span></span>||<span data-ttu-id="e64f0-1345">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="e64f0-1345">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="e64f0-1346">自定义属性作为 `asyncResult.value` 属性中的 [`CustomProperties`](/javascript/api/outlook/office.customproperties) 对象提供。</span><span class="sxs-lookup"><span data-stu-id="e64f0-1346">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook/office.customproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="e64f0-1347">此对象可用于获取、设置以及从项目中删除自定义属性，并将自定义属性集的更改重新保存到服务器。</span><span class="sxs-lookup"><span data-stu-id="e64f0-1347">The custom properties are provided as a CustomProperties object in the asyncResult.value property. This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`|<span data-ttu-id="e64f0-1348">对象</span><span class="sxs-lookup"><span data-stu-id="e64f0-1348">Object</span></span>|<span data-ttu-id="e64f0-1349">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="e64f0-1349">&lt;optional&gt;</span></span>|<span data-ttu-id="e64f0-1350">开发人员可以提供他们想要在回调函数中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="e64f0-1350">Developers can provide any object they wish to access in the callback method.</span></span> <span data-ttu-id="e64f0-1351">此对象可以通过回调函数中的 `asyncResult.asyncContext` 属性进行访问。</span><span class="sxs-lookup"><span data-stu-id="e64f0-1351">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.|</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e64f0-1352">要求</span><span class="sxs-lookup"><span data-stu-id="e64f0-1352">Requirements</span></span>

|<span data-ttu-id="e64f0-1353">要求</span><span class="sxs-lookup"><span data-stu-id="e64f0-1353">Requirement</span></span>|<span data-ttu-id="e64f0-1354">值</span><span class="sxs-lookup"><span data-stu-id="e64f0-1354">Value</span></span>|
|---|---|
|[<span data-ttu-id="e64f0-1355">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="e64f0-1355">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e64f0-1356">1.0</span><span class="sxs-lookup"><span data-stu-id="e64f0-1356">1.0</span></span>|
|[<span data-ttu-id="e64f0-1357">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="e64f0-1357">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e64f0-1358">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e64f0-1358">ReadItem</span></span>|
|[<span data-ttu-id="e64f0-1359">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="e64f0-1359">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="e64f0-1360">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="e64f0-1360">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="e64f0-1361">示例</span><span class="sxs-lookup"><span data-stu-id="e64f0-1361">Example</span></span>

<span data-ttu-id="e64f0-p180">以下代码示例显示了如何使用 `loadCustomPropertiesAsync` 方法异步加载特定于当前项目的自定义属性。该示例还显示了如何使用 `CustomProperties.saveAsync` 方法将这些属性重新保存到服务器。加载自定义属性后，该代码示例将使用 `CustomProperties.get` 方法读取自定义属性 `myProp`，使用 `CustomProperties.set` 方法写入自定义属性 `otherProp`，最后调用 `saveAsync` 方法保存这些自定义属性。</span><span class="sxs-lookup"><span data-stu-id="e64f0-p180">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

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

####  <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="e64f0-1365">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="e64f0-1365">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="e64f0-1366">将附件从邮件或约会中删除。</span><span class="sxs-lookup"><span data-stu-id="e64f0-1366">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="e64f0-1367">`removeAttachmentAsync` 方法删除项目中带指定标识符的附件。</span><span class="sxs-lookup"><span data-stu-id="e64f0-1367">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item.</span></span> <span data-ttu-id="e64f0-1368">最佳做法是，仅当同一个邮件应用程序在同一会话中添加了一个附件时，你才应使用该附件标识符来删除该附件。</span><span class="sxs-lookup"><span data-stu-id="e64f0-1368">As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session.</span></span> <span data-ttu-id="e64f0-1369">在 Outlook Web App 和适用于设备的 OWA 中，附件标识符只在同一个会话中才有效。</span><span class="sxs-lookup"><span data-stu-id="e64f0-1369">In Outlook Web App and OWA for Devices, the attachment identifier is valid only within the same session.</span></span> <span data-ttu-id="e64f0-1370">当用户关闭应用，或者如果用户开始在内嵌窗体中撰写，则随后弹出的窗体中继续在单独的窗口撰写时，会话即结束。</span><span class="sxs-lookup"><span data-stu-id="e64f0-1370">A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="e64f0-1371">参数：</span><span class="sxs-lookup"><span data-stu-id="e64f0-1371">Parameters:</span></span>

|<span data-ttu-id="e64f0-1372">名称</span><span class="sxs-lookup"><span data-stu-id="e64f0-1372">Name</span></span>|<span data-ttu-id="e64f0-1373">类型</span><span class="sxs-lookup"><span data-stu-id="e64f0-1373">Type</span></span>|<span data-ttu-id="e64f0-1374">属性</span><span class="sxs-lookup"><span data-stu-id="e64f0-1374">Attributes</span></span>|<span data-ttu-id="e64f0-1375">说明</span><span class="sxs-lookup"><span data-stu-id="e64f0-1375">Description</span></span>|
|---|---|---|---|
|`attachmentId`|<span data-ttu-id="e64f0-1376">字符串</span><span class="sxs-lookup"><span data-stu-id="e64f0-1376">String</span></span>||<span data-ttu-id="e64f0-p182">要删除的附件的标识符。字符串的最大长度为 100 个字符。</span><span class="sxs-lookup"><span data-stu-id="e64f0-p182">The identifier of the attachment to remove. The maximum length of the string is 100 characters.</span></span>|
|`options`|<span data-ttu-id="e64f0-1379">Object</span><span class="sxs-lookup"><span data-stu-id="e64f0-1379">Object</span></span>|<span data-ttu-id="e64f0-1380">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="e64f0-1380">&lt;optional&gt;</span></span>|<span data-ttu-id="e64f0-1381">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="e64f0-1381">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="e64f0-1382">对象</span><span class="sxs-lookup"><span data-stu-id="e64f0-1382">Object</span></span>|<span data-ttu-id="e64f0-1383">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="e64f0-1383">&lt;optional&gt;</span></span>|<span data-ttu-id="e64f0-1384">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="e64f0-1384">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="e64f0-1385">函数</span><span class="sxs-lookup"><span data-stu-id="e64f0-1385">function</span></span>|<span data-ttu-id="e64f0-1386">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="e64f0-1386">&lt;optional&gt;</span></span>|<span data-ttu-id="e64f0-1387">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="e64f0-1387">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="e64f0-1388">如果删除附件失败，`asyncResult.error` 属性将包含一个说明失败原因的错误代码。</span><span class="sxs-lookup"><span data-stu-id="e64f0-1388">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="e64f0-1389">错误</span><span class="sxs-lookup"><span data-stu-id="e64f0-1389">Errors</span></span>

|<span data-ttu-id="e64f0-1390">错误代码</span><span class="sxs-lookup"><span data-stu-id="e64f0-1390">Error code</span></span>|<span data-ttu-id="e64f0-1391">说明</span><span class="sxs-lookup"><span data-stu-id="e64f0-1391">Description</span></span>|
|------------|-------------|
|`InvalidAttachmentId`|<span data-ttu-id="e64f0-1392">附件标识符不存在。</span><span class="sxs-lookup"><span data-stu-id="e64f0-1392">The attachment identifier does not exist.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e64f0-1393">要求</span><span class="sxs-lookup"><span data-stu-id="e64f0-1393">Requirements</span></span>

|<span data-ttu-id="e64f0-1394">要求</span><span class="sxs-lookup"><span data-stu-id="e64f0-1394">Requirement</span></span>|<span data-ttu-id="e64f0-1395">值</span><span class="sxs-lookup"><span data-stu-id="e64f0-1395">Value</span></span>|
|---|---|
|[<span data-ttu-id="e64f0-1396">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="e64f0-1396">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e64f0-1397">1.1</span><span class="sxs-lookup"><span data-stu-id="e64f0-1397">1.1</span></span>|
|[<span data-ttu-id="e64f0-1398">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="e64f0-1398">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e64f0-1399">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="e64f0-1399">ReadWriteItem</span></span>|
|[<span data-ttu-id="e64f0-1400">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="e64f0-1400">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="e64f0-1401">撰写</span><span class="sxs-lookup"><span data-stu-id="e64f0-1401">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="e64f0-1402">示例</span><span class="sxs-lookup"><span data-stu-id="e64f0-1402">Example</span></span>

<span data-ttu-id="e64f0-1403">以下代码删除包含标识符 '0' 的附件。</span><span class="sxs-lookup"><span data-stu-id="e64f0-1403">The following code removes an attachment with an identifier of '0'.</span></span>

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

####  <a name="removehandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="e64f0-1404">removeHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="e64f0-1404">addHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="e64f0-1405">删除支持事件的事件处理程序。</span><span class="sxs-lookup"><span data-stu-id="e64f0-1405">Removes an event handler for a</span></span>

<span data-ttu-id="e64f0-1406">当前，支持的事件类型是 `Office.EventType.AttachmentsChanged`、`Office.EventType.AppointmentTimeChanged`、`Office.EventType.RecipientsChanged` 和 `Office.EventType.RecurrenceChanged`</span><span class="sxs-lookup"><span data-stu-id="e64f0-1406">Currently the supported event types are `Office.EventType.AttachmentsChanged`, `Office.EventType.AppointmentTimeChanged`, `Office.EventType.RecipientsChanged`, and `Office.EventType.RecurrenceChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="e64f0-1407">参数：</span><span class="sxs-lookup"><span data-stu-id="e64f0-1407">Parameters:</span></span>

| <span data-ttu-id="e64f0-1408">名称</span><span class="sxs-lookup"><span data-stu-id="e64f0-1408">Name</span></span> | <span data-ttu-id="e64f0-1409">类型</span><span class="sxs-lookup"><span data-stu-id="e64f0-1409">Type</span></span> | <span data-ttu-id="e64f0-1410">属性</span><span class="sxs-lookup"><span data-stu-id="e64f0-1410">Attributes</span></span> | <span data-ttu-id="e64f0-1411">说明</span><span class="sxs-lookup"><span data-stu-id="e64f0-1411">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="e64f0-1412">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="e64f0-1412">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="e64f0-1413">应撤销处理程序的事件。</span><span class="sxs-lookup"><span data-stu-id="e64f0-1413">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="e64f0-1414">函数</span><span class="sxs-lookup"><span data-stu-id="e64f0-1414">Function</span></span> || <span data-ttu-id="e64f0-p183">用于处理事件的函数。此函数必须接受一个参数，即对象文本。参数上的 `type` 属性将匹配传递给 `removeHandlerAsync` 的 `eventType` 参数。</span><span class="sxs-lookup"><span data-stu-id="e64f0-p183">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `removeHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="e64f0-1418">Object</span><span class="sxs-lookup"><span data-stu-id="e64f0-1418">Object</span></span> | <span data-ttu-id="e64f0-1419">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="e64f0-1419">&lt;optional&gt;</span></span> | <span data-ttu-id="e64f0-1420">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="e64f0-1420">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="e64f0-1421">对象</span><span class="sxs-lookup"><span data-stu-id="e64f0-1421">Object</span></span> | <span data-ttu-id="e64f0-1422">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="e64f0-1422">&lt;optional&gt;</span></span> | <span data-ttu-id="e64f0-1423">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="e64f0-1423">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="e64f0-1424">函数</span><span class="sxs-lookup"><span data-stu-id="e64f0-1424">function</span></span>| <span data-ttu-id="e64f0-1425">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="e64f0-1425">&lt;optional&gt;</span></span>|<span data-ttu-id="e64f0-1426">方法完成后，使用单个参数 `callback`（一个 [`asyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `AsyncResult` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="e64f0-1426">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e64f0-1427">Requirements</span><span class="sxs-lookup"><span data-stu-id="e64f0-1427">Requirements</span></span>

|<span data-ttu-id="e64f0-1428">要求</span><span class="sxs-lookup"><span data-stu-id="e64f0-1428">Requirement</span></span>| <span data-ttu-id="e64f0-1429">值</span><span class="sxs-lookup"><span data-stu-id="e64f0-1429">Value</span></span>|
|---|---|
|[<span data-ttu-id="e64f0-1430">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="e64f0-1430">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e64f0-1431">1.7</span><span class="sxs-lookup"><span data-stu-id="e64f0-1431">-17</span></span> |
|[<span data-ttu-id="e64f0-1432">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="e64f0-1432">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e64f0-1433">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e64f0-1433">ReadItem</span></span> |
|[<span data-ttu-id="e64f0-1434">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="e64f0-1434">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="e64f0-1435">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="e64f0-1435">Compose or read</span></span> |

####  <a name="saveasyncoptions-callback"></a><span data-ttu-id="e64f0-1436">saveAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="e64f0-1436">saveAsync([options], callback)</span></span>

<span data-ttu-id="e64f0-1437">异步保存项目。</span><span class="sxs-lookup"><span data-stu-id="e64f0-1437">Asynchronously saves an item.</span></span>

<span data-ttu-id="e64f0-p184">调用时，此方法将当前邮件保存为草稿，并通过回调方法返回项目 ID。在 Outlook Web App 或 Outlook 联机模式下，该项目被保存到服务器中。在 Outlook 缓存模式下，该项目被保存到本地缓存中。</span><span class="sxs-lookup"><span data-stu-id="e64f0-p184">When invoked, this method saves the current message as a draft and returns the item id via the callback method. In Outlook Web App or Outlook in online mode, the item is saved to the server. In Outlook in cached mode, the item is saved to the local cache.</span></span>

> [!NOTE]
> <span data-ttu-id="e64f0-1441">如果加载项在撰写模式下对某个项目调用 `saveAsync` 来获得 `itemId`，以便与 EWS 或 REST API 一同使用，请注意，当 Outlook 处于高速缓存模式时，可能需要一段时间项目才能真正同步到服务器。</span><span class="sxs-lookup"><span data-stu-id="e64f0-1441">Note: If your add-in calls `saveAsync` on an item in compose mode in order to get an `itemId` to use with EWS or the REST API, be aware that when Outlook is in cached mode, it may take some time before the item is actually synced to the server. Until the item is synced, using the  will return an error.</span></span> <span data-ttu-id="e64f0-1442">在项目同步前，使用 `itemId` 将返回一个错误。</span><span class="sxs-lookup"><span data-stu-id="e64f0-1442">Until the item is synced, using the `itemId` will return an error.</span></span>

<span data-ttu-id="e64f0-p186">由于约会没有草稿状态，如果以撰写模式在约会中调用 `saveAsync`，则该项将被保存为用户日历中的正常约会。对于之前未保存过的新约会，则不会发送邀请。保存现有约会将向添加或删除的与会者发送更新。</span><span class="sxs-lookup"><span data-stu-id="e64f0-p186">Since appointments have no draft state, if `saveAsync` is called on an appointment in compose mode, the item will be saved as a normal appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent. Saving an existing appointment will send an update to added or removed attendees.</span></span>

> [!NOTE]
> <span data-ttu-id="e64f0-1446">以下客户端在撰写模式下对约会上的 `saveAsync` 具有不同的行为：</span><span class="sxs-lookup"><span data-stu-id="e64f0-1446">The following clients have different behavior for `saveAsync` on appointments in compose mode:</span></span>
>
> - <span data-ttu-id="e64f0-1447">Mac Outlook 不支持在撰写模式下对会议执行 `saveAsync` 操作。</span><span class="sxs-lookup"><span data-stu-id="e64f0-1447">Note: Mac Outlook does not support `saveAsync` on a meeting in compose mode. Calling  on a meeting in Mac Outlook will return an error.</span></span> <span data-ttu-id="e64f0-1448">对 Mac Outlook 中的会议调用 `saveAsync` 将会返回一个错误。</span><span class="sxs-lookup"><span data-stu-id="e64f0-1448">Note: Mac Outlook does not support  on a meeting in compose mode. Calling `saveAsync` on a meeting in Mac Outlook will return an error.</span></span>
> - <span data-ttu-id="e64f0-1449">在撰写模式下的约会上调用 `saveAsync` 时，Outlook 网页版始终发送邀请或更新。</span><span class="sxs-lookup"><span data-stu-id="e64f0-1449">Outlook on the web always sends an invitation or update when `saveAsync` is called on an appointment in compose mode.</span></span>

##### <a name="parameters"></a><span data-ttu-id="e64f0-1450">参数：</span><span class="sxs-lookup"><span data-stu-id="e64f0-1450">Parameters:</span></span>

|<span data-ttu-id="e64f0-1451">名称</span><span class="sxs-lookup"><span data-stu-id="e64f0-1451">Name</span></span>|<span data-ttu-id="e64f0-1452">类型</span><span class="sxs-lookup"><span data-stu-id="e64f0-1452">Type</span></span>|<span data-ttu-id="e64f0-1453">属性</span><span class="sxs-lookup"><span data-stu-id="e64f0-1453">Attributes</span></span>|<span data-ttu-id="e64f0-1454">说明</span><span class="sxs-lookup"><span data-stu-id="e64f0-1454">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="e64f0-1455">Object</span><span class="sxs-lookup"><span data-stu-id="e64f0-1455">Object</span></span>|<span data-ttu-id="e64f0-1456">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="e64f0-1456">&lt;optional&gt;</span></span>|<span data-ttu-id="e64f0-1457">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="e64f0-1457">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="e64f0-1458">对象</span><span class="sxs-lookup"><span data-stu-id="e64f0-1458">Object</span></span>|<span data-ttu-id="e64f0-1459">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="e64f0-1459">&lt;optional&gt;</span></span>|<span data-ttu-id="e64f0-1460">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="e64f0-1460">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="e64f0-1461">函数</span><span class="sxs-lookup"><span data-stu-id="e64f0-1461">function</span></span>||<span data-ttu-id="e64f0-1462">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="e64f0-1462">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="e64f0-1463">如果成功，该项目标识符将在 `asyncResult.value` 属性中提供。</span><span class="sxs-lookup"><span data-stu-id="e64f0-1463">On success, the item identifier is provided in the `asyncResult.value` property.|</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e64f0-1464">要求</span><span class="sxs-lookup"><span data-stu-id="e64f0-1464">Requirements</span></span>

|<span data-ttu-id="e64f0-1465">要求</span><span class="sxs-lookup"><span data-stu-id="e64f0-1465">Requirement</span></span>|<span data-ttu-id="e64f0-1466">值</span><span class="sxs-lookup"><span data-stu-id="e64f0-1466">Value</span></span>|
|---|---|
|[<span data-ttu-id="e64f0-1467">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="e64f0-1467">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e64f0-1468">1.3</span><span class="sxs-lookup"><span data-stu-id="e64f0-1468">1.3</span></span>|
|[<span data-ttu-id="e64f0-1469">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="e64f0-1469">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e64f0-1470">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="e64f0-1470">ReadWriteItem</span></span>|
|[<span data-ttu-id="e64f0-1471">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="e64f0-1471">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="e64f0-1472">撰写</span><span class="sxs-lookup"><span data-stu-id="e64f0-1472">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="e64f0-1473">示例</span><span class="sxs-lookup"><span data-stu-id="e64f0-1473">Examples</span></span>

```javascript
Office.context.mailbox.item.saveAsync(
  function callback(result) {
    // Process the result
  });
```

<span data-ttu-id="e64f0-p188">下面是传递给回调函数的 `result` 参数的示例。`value` 属性包含的项目的项目 ID。</span><span class="sxs-lookup"><span data-stu-id="e64f0-p188">The following is an example of the `result` parameter passed to the callback function. The `value` property contains the item ID of the item.</span></span>

```json
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

####  <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="e64f0-1476">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="e64f0-1476">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="e64f0-1477">以异步方式将数据插入到邮件的正文或主题中。</span><span class="sxs-lookup"><span data-stu-id="e64f0-1477">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="e64f0-p189">`setSelectedDataAsync` 方法将指定字符串插入到项目主题或正文的光标位置，或者，如果在编辑器中已选择文本，则该方法将替换选择的文本。如果光标不在正文或主题字段中，则返回错误。插入之后，光标会位于插入内容的末尾。</span><span class="sxs-lookup"><span data-stu-id="e64f0-p189">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="e64f0-1481">参数：</span><span class="sxs-lookup"><span data-stu-id="e64f0-1481">Parameters:</span></span>

|<span data-ttu-id="e64f0-1482">名称</span><span class="sxs-lookup"><span data-stu-id="e64f0-1482">Name</span></span>|<span data-ttu-id="e64f0-1483">类型</span><span class="sxs-lookup"><span data-stu-id="e64f0-1483">Type</span></span>|<span data-ttu-id="e64f0-1484">属性</span><span class="sxs-lookup"><span data-stu-id="e64f0-1484">Attributes</span></span>|<span data-ttu-id="e64f0-1485">说明</span><span class="sxs-lookup"><span data-stu-id="e64f0-1485">Description</span></span>|
|---|---|---|---|
|`data`|<span data-ttu-id="e64f0-1486">字符串</span><span class="sxs-lookup"><span data-stu-id="e64f0-1486">String</span></span>||<span data-ttu-id="e64f0-p190">要插入的数据。数据不得超过 1,000,000 个字符。如果传入的数据超过 1,000,000 个字符，则会引发 `ArgumentOutOfRange` 异常。</span><span class="sxs-lookup"><span data-stu-id="e64f0-p190">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`|<span data-ttu-id="e64f0-1490">Object</span><span class="sxs-lookup"><span data-stu-id="e64f0-1490">Object</span></span>|<span data-ttu-id="e64f0-1491">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="e64f0-1491">&lt;optional&gt;</span></span>|<span data-ttu-id="e64f0-1492">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="e64f0-1492">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="e64f0-1493">对象</span><span class="sxs-lookup"><span data-stu-id="e64f0-1493">Object</span></span>|<span data-ttu-id="e64f0-1494">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="e64f0-1494">&lt;optional&gt;</span></span>|<span data-ttu-id="e64f0-1495">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="e64f0-1495">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`|[<span data-ttu-id="e64f0-1496">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="e64f0-1496">Office.CoercionType</span></span>](office.md#coerciontype-string)|<span data-ttu-id="e64f0-1497">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="e64f0-1497">&lt;optional&gt;</span></span>|<span data-ttu-id="e64f0-p191">如果为 `text`，则在 Outlook Web App 和 Outlook 中应用当前样式。如果该字段是 HTML 编辑器，则仅插入文本数据，即使数据为 HTML。</span><span class="sxs-lookup"><span data-stu-id="e64f0-p191">If `text`, the current style is applied in Outlook Web App and Outlook. If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="e64f0-p192">如果 `html` 和该字段支持 HTML（主题不支持），则在 Outlook Web App 中应用当前样式，而在 Outlook 中应用默认样式。如果该字段是文本字段，则返回 `InvalidDataFormat` 错误。</span><span class="sxs-lookup"><span data-stu-id="e64f0-p192">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook Web App and the default style is applied in Outlook. If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="e64f0-1502">如果未设置 `coercionType`，则结果取决于该字段：如果该字段是 HTML，则使用 HTML；如果该字段是文本，则使用纯文本。</span><span class="sxs-lookup"><span data-stu-id="e64f0-1502">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`|<span data-ttu-id="e64f0-1503">function</span><span class="sxs-lookup"><span data-stu-id="e64f0-1503">function</span></span>||<span data-ttu-id="e64f0-1504">方法完成后，使用单个参数 `callback`（一个 [`asyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `AsyncResult` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="e64f0-1504">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e64f0-1505">Requirements</span><span class="sxs-lookup"><span data-stu-id="e64f0-1505">Requirements</span></span>

|<span data-ttu-id="e64f0-1506">要求</span><span class="sxs-lookup"><span data-stu-id="e64f0-1506">Requirement</span></span>|<span data-ttu-id="e64f0-1507">值</span><span class="sxs-lookup"><span data-stu-id="e64f0-1507">Value</span></span>|
|---|---|
|[<span data-ttu-id="e64f0-1508">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="e64f0-1508">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="e64f0-1509">1.2</span><span class="sxs-lookup"><span data-stu-id="e64f0-1509">1.2</span></span>|
|[<span data-ttu-id="e64f0-1510">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="e64f0-1510">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="e64f0-1511">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="e64f0-1511">ReadWriteItem</span></span>|
|[<span data-ttu-id="e64f0-1512">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="e64f0-1512">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="e64f0-1513">撰写</span><span class="sxs-lookup"><span data-stu-id="e64f0-1513">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="e64f0-1514">示例</span><span class="sxs-lookup"><span data-stu-id="e64f0-1514">Example</span></span>

```javascript
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```