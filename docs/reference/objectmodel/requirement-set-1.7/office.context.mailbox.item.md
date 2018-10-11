
# <a name="item"></a><span data-ttu-id="c7a76-101">item</span><span class="sxs-lookup"><span data-stu-id="c7a76-101">item</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmditem"></a><span data-ttu-id="c7a76-102">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span><span class="sxs-lookup"><span data-stu-id="c7a76-102">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span></span>

<span data-ttu-id="c7a76-p101">`item`命名空间用于访问当前选定的邮件、会议请求或安排。可以通过使用[itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlook17officemailboxenumsitemtype) 属性确定`item`的类型。</span><span class="sxs-lookup"><span data-stu-id="c7a76-p101">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlook17officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="c7a76-105">要求</span><span class="sxs-lookup"><span data-stu-id="c7a76-105">Requirements</span></span>

|<span data-ttu-id="c7a76-106">要求</span><span class="sxs-lookup"><span data-stu-id="c7a76-106">Requirement</span></span>|<span data-ttu-id="c7a76-107">值</span><span class="sxs-lookup"><span data-stu-id="c7a76-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="c7a76-108">最低的邮箱要求集版本</span><span class="sxs-lookup"><span data-stu-id="c7a76-108">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c7a76-109">1.0</span><span class="sxs-lookup"><span data-stu-id="c7a76-109">1.0</span></span>|
|[<span data-ttu-id="c7a76-110">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="c7a76-110">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c7a76-111">受限</span><span class="sxs-lookup"><span data-stu-id="c7a76-111">Restricted</span></span>|
|[<span data-ttu-id="c7a76-112">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c7a76-112">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="c7a76-113">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="c7a76-113">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="c7a76-114">成员和方法</span><span class="sxs-lookup"><span data-stu-id="c7a76-114">Members and methods</span></span>

| <span data-ttu-id="c7a76-115">成员</span><span class="sxs-lookup"><span data-stu-id="c7a76-115">Member</span></span> | <span data-ttu-id="c7a76-116">类型</span><span class="sxs-lookup"><span data-stu-id="c7a76-116">Type</span></span> |
|--------|------|
| [<span data-ttu-id="c7a76-117">attachments</span><span class="sxs-lookup"><span data-stu-id="c7a76-117">attachments</span></span>](#attachments-arrayattachmentdetailsjavascriptapioutlook17officeattachmentdetails) | <span data-ttu-id="c7a76-118">成员</span><span class="sxs-lookup"><span data-stu-id="c7a76-118">Member</span></span> |
| [<span data-ttu-id="c7a76-119">bcc</span><span class="sxs-lookup"><span data-stu-id="c7a76-119">bcc</span></span>](#bcc-recipientsjavascriptapioutlook17officerecipients) | <span data-ttu-id="c7a76-120">成员</span><span class="sxs-lookup"><span data-stu-id="c7a76-120">Member</span></span> |
| [<span data-ttu-id="c7a76-121">body</span><span class="sxs-lookup"><span data-stu-id="c7a76-121">body</span></span>](#body-bodyjavascriptapioutlook17officebody) | <span data-ttu-id="c7a76-122">成员</span><span class="sxs-lookup"><span data-stu-id="c7a76-122">Member</span></span> |
| [<span data-ttu-id="c7a76-123">cc</span><span class="sxs-lookup"><span data-stu-id="c7a76-123">cc</span></span>](#cc-arrayemailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsrecipientsjavascriptapioutlook17officerecipients) | <span data-ttu-id="c7a76-124">成员</span><span class="sxs-lookup"><span data-stu-id="c7a76-124">Member</span></span> |
| [<span data-ttu-id="c7a76-125">conversationId</span><span class="sxs-lookup"><span data-stu-id="c7a76-125">conversationId</span></span>](#nullable-conversationid-string) | <span data-ttu-id="c7a76-126">成员</span><span class="sxs-lookup"><span data-stu-id="c7a76-126">Member</span></span> |
| [<span data-ttu-id="c7a76-127">dateTimeCreated</span><span class="sxs-lookup"><span data-stu-id="c7a76-127">dateTimeCreated</span></span>](#datetimecreated-date) | <span data-ttu-id="c7a76-128">成员</span><span class="sxs-lookup"><span data-stu-id="c7a76-128">Member</span></span> |
| [<span data-ttu-id="c7a76-129">dateTimeModified</span><span class="sxs-lookup"><span data-stu-id="c7a76-129">dateTimeModified</span></span>](#datetimemodified-date) | <span data-ttu-id="c7a76-130">成员</span><span class="sxs-lookup"><span data-stu-id="c7a76-130">Member</span></span> |
| [<span data-ttu-id="c7a76-131">end</span><span class="sxs-lookup"><span data-stu-id="c7a76-131">end</span></span>](#end-datetimejavascriptapioutlook17officetime) | <span data-ttu-id="c7a76-132">成员</span><span class="sxs-lookup"><span data-stu-id="c7a76-132">Member</span></span> |
| [<span data-ttu-id="c7a76-133">from</span><span class="sxs-lookup"><span data-stu-id="c7a76-133">from</span></span>](#from-emailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsfromjavascriptapioutlook17officefrom) | <span data-ttu-id="c7a76-134">成员</span><span class="sxs-lookup"><span data-stu-id="c7a76-134">Member</span></span> |
| [<span data-ttu-id="c7a76-135">internetMessageId</span><span class="sxs-lookup"><span data-stu-id="c7a76-135">internetMessageId</span></span>](#internetmessageid-string) | <span data-ttu-id="c7a76-136">成员</span><span class="sxs-lookup"><span data-stu-id="c7a76-136">Member</span></span> |
| [<span data-ttu-id="c7a76-137">itemClass</span><span class="sxs-lookup"><span data-stu-id="c7a76-137">itemClass</span></span>](#itemclass-string) | <span data-ttu-id="c7a76-138">成员</span><span class="sxs-lookup"><span data-stu-id="c7a76-138">Member</span></span> |
| [<span data-ttu-id="c7a76-139">itemId</span><span class="sxs-lookup"><span data-stu-id="c7a76-139">itemId</span></span>](#nullable-itemid-string) | <span data-ttu-id="c7a76-140">成员</span><span class="sxs-lookup"><span data-stu-id="c7a76-140">Member</span></span> |
| [<span data-ttu-id="c7a76-141">itemType</span><span class="sxs-lookup"><span data-stu-id="c7a76-141">itemType</span></span>](#itemtype-officemailboxenumsitemtypejavascriptapioutlook17officemailboxenumsitemtype) | <span data-ttu-id="c7a76-142">成员</span><span class="sxs-lookup"><span data-stu-id="c7a76-142">Member</span></span> |
| [<span data-ttu-id="c7a76-143">location</span><span class="sxs-lookup"><span data-stu-id="c7a76-143">location</span></span>](#location-stringlocationjavascriptapioutlook17officelocation) | <span data-ttu-id="c7a76-144">成员</span><span class="sxs-lookup"><span data-stu-id="c7a76-144">Member</span></span> |
| [<span data-ttu-id="c7a76-145">normalizedSubject</span><span class="sxs-lookup"><span data-stu-id="c7a76-145">normalizedSubject</span></span>](#normalizedsubject-string) | <span data-ttu-id="c7a76-146">成员</span><span class="sxs-lookup"><span data-stu-id="c7a76-146">Member</span></span> |
| [<span data-ttu-id="c7a76-147">NotificationMessages</span><span class="sxs-lookup"><span data-stu-id="c7a76-147">notificationMessages</span></span>](#notificationmessages-notificationmessagesjavascriptapioutlook17officenotificationmessages) | <span data-ttu-id="c7a76-148">成员</span><span class="sxs-lookup"><span data-stu-id="c7a76-148">Member</span></span> |
| [<span data-ttu-id="c7a76-149">optionalAttendees</span><span class="sxs-lookup"><span data-stu-id="c7a76-149">optionalAttendees</span></span>](#optionalattendees-arrayemailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsrecipientsjavascriptapioutlook17officerecipients) | <span data-ttu-id="c7a76-150">成员</span><span class="sxs-lookup"><span data-stu-id="c7a76-150">Member</span></span> |
| [<span data-ttu-id="c7a76-151">organizer</span><span class="sxs-lookup"><span data-stu-id="c7a76-151">organizer</span></span>](#organizer-emailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsorganizerjavascriptapioutlook17officeorganizer) | <span data-ttu-id="c7a76-152">成员</span><span class="sxs-lookup"><span data-stu-id="c7a76-152">Member</span></span> |
| [<span data-ttu-id="c7a76-153">recurrence</span><span class="sxs-lookup"><span data-stu-id="c7a76-153">recurrence</span></span>](#nullable-recurrence-recurrencejavascriptapioutlook17officerecurrence) | <span data-ttu-id="c7a76-154">成员</span><span class="sxs-lookup"><span data-stu-id="c7a76-154">Member</span></span> |
| [<span data-ttu-id="c7a76-155">requiredAttendees</span><span class="sxs-lookup"><span data-stu-id="c7a76-155">requiredAttendees</span></span>](#requiredattendees-arrayemailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsrecipientsjavascriptapioutlook17officerecipients) | <span data-ttu-id="c7a76-156">成员</span><span class="sxs-lookup"><span data-stu-id="c7a76-156">Member</span></span> |
| [<span data-ttu-id="c7a76-157">sender</span><span class="sxs-lookup"><span data-stu-id="c7a76-157">sender</span></span>](#sender-emailaddressdetailsjavascriptapioutlook17officeemailaddressdetails) | <span data-ttu-id="c7a76-158">成员</span><span class="sxs-lookup"><span data-stu-id="c7a76-158">Member</span></span> |
| [<span data-ttu-id="c7a76-159">seriesId</span><span class="sxs-lookup"><span data-stu-id="c7a76-159">seriesId</span></span>](#nullable-seriesid-string) | <span data-ttu-id="c7a76-160">成员</span><span class="sxs-lookup"><span data-stu-id="c7a76-160">Member</span></span> |
| [<span data-ttu-id="c7a76-161">start</span><span class="sxs-lookup"><span data-stu-id="c7a76-161">start</span></span>](#start-datetimejavascriptapioutlook17officetime) | <span data-ttu-id="c7a76-162">成员</span><span class="sxs-lookup"><span data-stu-id="c7a76-162">Member</span></span> |
| [<span data-ttu-id="c7a76-163">subject</span><span class="sxs-lookup"><span data-stu-id="c7a76-163">subject</span></span>](#subject-stringsubjectjavascriptapioutlook17officesubject) | <span data-ttu-id="c7a76-164">成员</span><span class="sxs-lookup"><span data-stu-id="c7a76-164">Member</span></span> |
| [<span data-ttu-id="c7a76-165">to</span><span class="sxs-lookup"><span data-stu-id="c7a76-165">to</span></span>](#to-arrayemailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsrecipientsjavascriptapioutlook17officerecipients) | <span data-ttu-id="c7a76-166">成员</span><span class="sxs-lookup"><span data-stu-id="c7a76-166">Member</span></span> |
| [<span data-ttu-id="c7a76-167">addFileAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="c7a76-167">addFileAttachmentAsync</span></span>](#addfileattachmentasyncuri-attachmentname-options-callback) | <span data-ttu-id="c7a76-168">方法</span><span class="sxs-lookup"><span data-stu-id="c7a76-168">Method</span></span> |
| [<span data-ttu-id="c7a76-169">addHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="c7a76-169">addHandlerAsync</span></span>](#addhandlerasynceventtype-handler-options-callback) | <span data-ttu-id="c7a76-170">方法</span><span class="sxs-lookup"><span data-stu-id="c7a76-170">Method</span></span> |
| [<span data-ttu-id="c7a76-171">addItemAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="c7a76-171">addItemAttachmentAsync</span></span>](#additemattachmentasyncitemid-attachmentname-options-callback) | <span data-ttu-id="c7a76-172">方法</span><span class="sxs-lookup"><span data-stu-id="c7a76-172">Method</span></span> |
| [<span data-ttu-id="c7a76-173">close</span><span class="sxs-lookup"><span data-stu-id="c7a76-173">close</span></span>](#close) | <span data-ttu-id="c7a76-174">方法</span><span class="sxs-lookup"><span data-stu-id="c7a76-174">Method</span></span> |
| [<span data-ttu-id="c7a76-175">displayReplyAllForm</span><span class="sxs-lookup"><span data-stu-id="c7a76-175">displayReplyAllForm</span></span>](#displayreplyallformformdata) | <span data-ttu-id="c7a76-176">方法</span><span class="sxs-lookup"><span data-stu-id="c7a76-176">Method</span></span> |
| [<span data-ttu-id="c7a76-177">displayReplyForm</span><span class="sxs-lookup"><span data-stu-id="c7a76-177">displayReplyForm</span></span>](#displayreplyformformdata) | <span data-ttu-id="c7a76-178">方法</span><span class="sxs-lookup"><span data-stu-id="c7a76-178">Method</span></span> |
| [<span data-ttu-id="c7a76-179">getEntities</span><span class="sxs-lookup"><span data-stu-id="c7a76-179">getEntities</span></span>](#getentities--entitiesjavascriptapioutlook17officeentities) | <span data-ttu-id="c7a76-180">方法</span><span class="sxs-lookup"><span data-stu-id="c7a76-180">Method</span></span> |
| [<span data-ttu-id="c7a76-181">getEntitiesByType</span><span class="sxs-lookup"><span data-stu-id="c7a76-181">getEntitiesByType</span></span>](#getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlook17officecontactmeetingsuggestionjavascriptapioutlook17officemeetingsuggestionphonenumberjavascriptapioutlook17officephonenumbertasksuggestionjavascriptapioutlook17officetasksuggestion) | <span data-ttu-id="c7a76-182">方法</span><span class="sxs-lookup"><span data-stu-id="c7a76-182">Method</span></span> |
| [<span data-ttu-id="c7a76-183">getFilteredEntitiesByName</span><span class="sxs-lookup"><span data-stu-id="c7a76-183">getFilteredEntitiesByName</span></span>](#getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlook17officecontactmeetingsuggestionjavascriptapioutlook17officemeetingsuggestionphonenumberjavascriptapioutlook17officephonenumbertasksuggestionjavascriptapioutlook17officetasksuggestion) | <span data-ttu-id="c7a76-184">方法</span><span class="sxs-lookup"><span data-stu-id="c7a76-184">Method</span></span> |
| [<span data-ttu-id="c7a76-185">getRegExMatches</span><span class="sxs-lookup"><span data-stu-id="c7a76-185">getRegExMatches</span></span>](#getregexmatches--object) | <span data-ttu-id="c7a76-186">方法</span><span class="sxs-lookup"><span data-stu-id="c7a76-186">Method</span></span> |
| [<span data-ttu-id="c7a76-187">getRegExMatchesByName</span><span class="sxs-lookup"><span data-stu-id="c7a76-187">getRegExMatchesByName</span></span>](#getregexmatchesbynamename--nullable-array-string-) | <span data-ttu-id="c7a76-188">方法</span><span class="sxs-lookup"><span data-stu-id="c7a76-188">Method</span></span> |
| [<span data-ttu-id="c7a76-189">getSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="c7a76-189">getSelectedDataAsync</span></span>](#getselecteddataasynccoerciontype-options-callback--string) | <span data-ttu-id="c7a76-190">方法</span><span class="sxs-lookup"><span data-stu-id="c7a76-190">Method</span></span> |
| [<span data-ttu-id="c7a76-191">getSelectedEntities</span><span class="sxs-lookup"><span data-stu-id="c7a76-191">getSelectedEntities</span></span>](#getselectedentities--entitiesjavascriptapioutlook17officeentities) | <span data-ttu-id="c7a76-192">方法</span><span class="sxs-lookup"><span data-stu-id="c7a76-192">Method</span></span> |
| [<span data-ttu-id="c7a76-193">getSelectedRegExMatches</span><span class="sxs-lookup"><span data-stu-id="c7a76-193">getSelectedRegExMatches</span></span>](#getselectedregexmatches--object) | <span data-ttu-id="c7a76-194">方法</span><span class="sxs-lookup"><span data-stu-id="c7a76-194">Method</span></span> |
| [<span data-ttu-id="c7a76-195">loadCustomPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="c7a76-195">loadCustomPropertiesAsync</span></span>](#loadcustompropertiesasynccallback-usercontext) | <span data-ttu-id="c7a76-196">方法</span><span class="sxs-lookup"><span data-stu-id="c7a76-196">Method</span></span> |
| [<span data-ttu-id="c7a76-197">removeAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="c7a76-197">removeAttachmentAsync</span></span>](#removeattachmentasyncattachmentid-options-callback) | <span data-ttu-id="c7a76-198">方法</span><span class="sxs-lookup"><span data-stu-id="c7a76-198">Method</span></span> |
| [<span data-ttu-id="c7a76-199">removeHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="c7a76-199">removeHandlerAsync</span></span>](#removehandlerasynceventtype-handler-options-callback) | <span data-ttu-id="c7a76-200">方法</span><span class="sxs-lookup"><span data-stu-id="c7a76-200">Method</span></span> |
| [<span data-ttu-id="c7a76-201">saveAsync</span><span class="sxs-lookup"><span data-stu-id="c7a76-201">saveAsync</span></span>](#saveasyncoptions-callback) | <span data-ttu-id="c7a76-202">方法</span><span class="sxs-lookup"><span data-stu-id="c7a76-202">Method</span></span> |
| [<span data-ttu-id="c7a76-203">setSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="c7a76-203">setSelectedDataAsync</span></span>](#setselecteddataasyncdata-options-callback) | <span data-ttu-id="c7a76-204">方法</span><span class="sxs-lookup"><span data-stu-id="c7a76-204">Method</span></span> |

### <a name="example"></a><span data-ttu-id="c7a76-205">示例</span><span class="sxs-lookup"><span data-stu-id="c7a76-205">Example</span></span>

<span data-ttu-id="c7a76-206">以下 JavaScript 代码示例显示了如何访问 Outlook 中当前项目的 `subject` 属性。</span><span class="sxs-lookup"><span data-stu-id="c7a76-206">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

```
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

### <a name="members"></a><span data-ttu-id="c7a76-207">成员</span><span class="sxs-lookup"><span data-stu-id="c7a76-207">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlook17officeattachmentdetails"></a><span data-ttu-id="c7a76-208">attachments :数组.<[AttachmentDetails](/javascript/api/outlook_1_7/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="c7a76-208">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_7/office.attachmentdetails)></span></span>

<span data-ttu-id="c7a76-p102">获取项目的附件数组。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="c7a76-p102">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="c7a76-211">某些类型的文件因潜在的安全问题被 Outlook 阻止，因此没有返回。</span><span class="sxs-lookup"><span data-stu-id="c7a76-211">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="c7a76-212">有关详细信息，请参阅 [在 Outlook 中被阻止的附件](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519)。</span><span class="sxs-lookup"><span data-stu-id="c7a76-212">For more information see [Payments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="c7a76-213">类型：</span><span class="sxs-lookup"><span data-stu-id="c7a76-213">Type:</span></span>

*   <span data-ttu-id="c7a76-214">数组.<[AttachmentDetails](/javascript/api/outlook_1_7/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="c7a76-214">Array.<[AttachmentDetails](/javascript/api/outlook_1_7/office.attachmentdetails)></span></span>

##### <a name="requirements"></a><span data-ttu-id="c7a76-215">要求</span><span class="sxs-lookup"><span data-stu-id="c7a76-215">Requirements</span></span>

|<span data-ttu-id="c7a76-216">要求</span><span class="sxs-lookup"><span data-stu-id="c7a76-216">Requirement</span></span>|<span data-ttu-id="c7a76-217">值</span><span class="sxs-lookup"><span data-stu-id="c7a76-217">Value</span></span>|
|---|---|
|[<span data-ttu-id="c7a76-218">最低的邮箱要求集版本</span><span class="sxs-lookup"><span data-stu-id="c7a76-218">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c7a76-219">1.0</span><span class="sxs-lookup"><span data-stu-id="c7a76-219">1.0</span></span>|
|[<span data-ttu-id="c7a76-220">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="c7a76-220">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c7a76-221">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c7a76-221">ReadItem</span></span>|
|[<span data-ttu-id="c7a76-222">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c7a76-222">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="c7a76-223">阅读</span><span class="sxs-lookup"><span data-stu-id="c7a76-223">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c7a76-224">示例</span><span class="sxs-lookup"><span data-stu-id="c7a76-224">Example</span></span>

<span data-ttu-id="c7a76-225">以下代码使用当前项目上所有附件的详细信息构成 HTML 字符串。</span><span class="sxs-lookup"><span data-stu-id="c7a76-225">The following code builds an HTML string with details of all attachments on the current item.</span></span>

```
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

####  <a name="bcc-recipientsjavascriptapioutlook17officerecipients"></a><span data-ttu-id="c7a76-226">密件抄送：[收件人](/javascript/api/outlook_1_7/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="c7a76-226">bcc :[Recipients](/javascript/api/outlook_1_7/office.recipients)</span></span>

<span data-ttu-id="c7a76-227">获取一个对象，提供用于获取或更新邮件的密件抄送 （密件抄送副本） 行上收件人的方法。</span><span class="sxs-lookup"><span data-stu-id="c7a76-227">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="c7a76-228">仅限撰写模式。</span><span class="sxs-lookup"><span data-stu-id="c7a76-228">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="c7a76-229">类型：</span><span class="sxs-lookup"><span data-stu-id="c7a76-229">Type:</span></span>

*   [<span data-ttu-id="c7a76-230">收件人</span><span class="sxs-lookup"><span data-stu-id="c7a76-230">Recipients</span></span>](/javascript/api/outlook_1_7/office.recipients)

##### <a name="requirements"></a><span data-ttu-id="c7a76-231">要求</span><span class="sxs-lookup"><span data-stu-id="c7a76-231">Requirements</span></span>

|<span data-ttu-id="c7a76-232">要求</span><span class="sxs-lookup"><span data-stu-id="c7a76-232">Requirement</span></span>|<span data-ttu-id="c7a76-233">值</span><span class="sxs-lookup"><span data-stu-id="c7a76-233">Value</span></span>|
|---|---|
|[<span data-ttu-id="c7a76-234">最低的邮箱要求集版本</span><span class="sxs-lookup"><span data-stu-id="c7a76-234">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c7a76-235">1.1</span><span class="sxs-lookup"><span data-stu-id="c7a76-235">1.1</span></span>|
|[<span data-ttu-id="c7a76-236">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="c7a76-236">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c7a76-237">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c7a76-237">ReadItem</span></span>|
|[<span data-ttu-id="c7a76-238">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c7a76-238">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="c7a76-239">撰写</span><span class="sxs-lookup"><span data-stu-id="c7a76-239">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="c7a76-240">示例</span><span class="sxs-lookup"><span data-stu-id="c7a76-240">Example</span></span>

```
Office.context.mailbox.item.bcc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.bcc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.bcc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfBccRecipients = asyncResult.value;
}
```

####  <a name="body-bodyjavascriptapioutlook17officebody"></a><span data-ttu-id="c7a76-241">正文：[正文](/javascript/api/outlook_1_7/office.body)</span><span class="sxs-lookup"><span data-stu-id="c7a76-241">body :[Body](/javascript/api/outlook_1_7/office.body)</span></span>

<span data-ttu-id="c7a76-242">获取一个提供用于处理项目正文的方法的对象。</span><span class="sxs-lookup"><span data-stu-id="c7a76-242">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="c7a76-243">类型：</span><span class="sxs-lookup"><span data-stu-id="c7a76-243">Type:</span></span>

*   [<span data-ttu-id="c7a76-244">Body</span><span class="sxs-lookup"><span data-stu-id="c7a76-244">Body</span></span>](/javascript/api/outlook_1_7/office.body)

##### <a name="requirements"></a><span data-ttu-id="c7a76-245">要求</span><span class="sxs-lookup"><span data-stu-id="c7a76-245">Requirements</span></span>

|<span data-ttu-id="c7a76-246">要求</span><span class="sxs-lookup"><span data-stu-id="c7a76-246">Requirement</span></span>|<span data-ttu-id="c7a76-247">值</span><span class="sxs-lookup"><span data-stu-id="c7a76-247">Value</span></span>|
|---|---|
|[<span data-ttu-id="c7a76-248">最低的邮箱要求集版本</span><span class="sxs-lookup"><span data-stu-id="c7a76-248">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c7a76-249">1.1</span><span class="sxs-lookup"><span data-stu-id="c7a76-249">1.1</span></span>|
|[<span data-ttu-id="c7a76-250">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="c7a76-250">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c7a76-251">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c7a76-251">ReadItem</span></span>|
|[<span data-ttu-id="c7a76-252">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c7a76-252">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="c7a76-253">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="c7a76-253">Compose or read</span></span>|

####  <a name="cc-arrayemailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsrecipientsjavascriptapioutlook17officerecipients"></a><span data-ttu-id="c7a76-254">cc :数组. <[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)>|[收件人](/javascript/api/outlook_1_7/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="c7a76-254">cc :Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_7/office.recipients)</span></span>

<span data-ttu-id="c7a76-255">提供对邮件抄送 (cc) 收件人的访问。</span><span class="sxs-lookup"><span data-stu-id="c7a76-255">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="c7a76-256">对象的类型和访问级别取决于当前项的模式。</span><span class="sxs-lookup"><span data-stu-id="c7a76-256">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="c7a76-257">阅读模式</span><span class="sxs-lookup"><span data-stu-id="c7a76-257">Read mode</span></span>

<span data-ttu-id="c7a76-p106">`cc`属性返回包含邮件的**抄送**行上所列每个收件人的 `EmailAddressDetails` 对象。集合上限为 100 个成员。</span><span class="sxs-lookup"><span data-stu-id="c7a76-p106">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="c7a76-260">撰写模式</span><span class="sxs-lookup"><span data-stu-id="c7a76-260">Compose mode</span></span>

<span data-ttu-id="c7a76-261">`cc` 属性返回 `Recipients` 对象，该对象提供用于处理邮件**抄送**行上收件人的方法。</span><span class="sxs-lookup"><span data-stu-id="c7a76-261">The `cc` property returns a `Recipients` object that provides methods for manipulating the recipients on the **Cc** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="c7a76-262">类型：</span><span class="sxs-lookup"><span data-stu-id="c7a76-262">Type:</span></span>

*   <span data-ttu-id="c7a76-263">数组.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)> |[收件人](/javascript/api/outlook_1_7/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="c7a76-263">Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_7/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="c7a76-264">要求</span><span class="sxs-lookup"><span data-stu-id="c7a76-264">Requirements</span></span>

|<span data-ttu-id="c7a76-265">要求</span><span class="sxs-lookup"><span data-stu-id="c7a76-265">Requirement</span></span>|<span data-ttu-id="c7a76-266">值</span><span class="sxs-lookup"><span data-stu-id="c7a76-266">Value</span></span>|
|---|---|
|[<span data-ttu-id="c7a76-267">最低的邮箱要求集版本</span><span class="sxs-lookup"><span data-stu-id="c7a76-267">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c7a76-268">1.0</span><span class="sxs-lookup"><span data-stu-id="c7a76-268">1.0</span></span>|
|[<span data-ttu-id="c7a76-269">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="c7a76-269">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c7a76-270">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c7a76-270">ReadItem</span></span>|
|[<span data-ttu-id="c7a76-271">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c7a76-271">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="c7a76-272">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="c7a76-272">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="c7a76-273">示例</span><span class="sxs-lookup"><span data-stu-id="c7a76-273">Example</span></span>

```
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

####  <a name="nullable-conversationid-string"></a><span data-ttu-id="c7a76-274">（可为空）conversationId :字符串</span><span class="sxs-lookup"><span data-stu-id="c7a76-274">(nullable) conversationId :String</span></span>

<span data-ttu-id="c7a76-275">获取包含特定消息的电子邮件会话的标识符。</span><span class="sxs-lookup"><span data-stu-id="c7a76-275">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="c7a76-p107">如果在阅读窗体或撰写窗体的回复中激活邮件应用程序，则此属性可以获得一个整数值。如果用户随后更改了回复邮件的主题（若发送回复），则该邮件的对话 ID 将改变且之前获取的值将不适用。</span><span class="sxs-lookup"><span data-stu-id="c7a76-p107">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="c7a76-p108">对于撰写窗体的新项目，此属性获得一个 null 值。如果用户设置一个主题并保存该项目，`conversationId` 属性将返回一个值。</span><span class="sxs-lookup"><span data-stu-id="c7a76-p108">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="c7a76-280">类型：</span><span class="sxs-lookup"><span data-stu-id="c7a76-280">Type:</span></span>

*   <span data-ttu-id="c7a76-281">字符串</span><span class="sxs-lookup"><span data-stu-id="c7a76-281">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="c7a76-282">要求</span><span class="sxs-lookup"><span data-stu-id="c7a76-282">Requirements</span></span>

|<span data-ttu-id="c7a76-283">要求</span><span class="sxs-lookup"><span data-stu-id="c7a76-283">Requirement</span></span>|<span data-ttu-id="c7a76-284">值</span><span class="sxs-lookup"><span data-stu-id="c7a76-284">Value</span></span>|
|---|---|
|[<span data-ttu-id="c7a76-285">最低的邮箱要求集版本</span><span class="sxs-lookup"><span data-stu-id="c7a76-285">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c7a76-286">1.0</span><span class="sxs-lookup"><span data-stu-id="c7a76-286">1.0</span></span>|
|[<span data-ttu-id="c7a76-287">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="c7a76-287">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c7a76-288">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c7a76-288">ReadItem</span></span>|
|[<span data-ttu-id="c7a76-289">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c7a76-289">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="c7a76-290">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="c7a76-290">Compose or read</span></span>|

#### <a name="datetimecreated-date"></a><span data-ttu-id="c7a76-291">dateTimeCreated：日期</span><span class="sxs-lookup"><span data-stu-id="c7a76-291">dateTimeCreated :Date</span></span>

<span data-ttu-id="c7a76-p109">获取项目创建的日期和时间。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="c7a76-p109">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="c7a76-294">类型：</span><span class="sxs-lookup"><span data-stu-id="c7a76-294">Type:</span></span>

*   <span data-ttu-id="c7a76-295">日期</span><span class="sxs-lookup"><span data-stu-id="c7a76-295">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="c7a76-296">要求</span><span class="sxs-lookup"><span data-stu-id="c7a76-296">Requirements</span></span>

|<span data-ttu-id="c7a76-297">要求</span><span class="sxs-lookup"><span data-stu-id="c7a76-297">Requirement</span></span>|<span data-ttu-id="c7a76-298">值</span><span class="sxs-lookup"><span data-stu-id="c7a76-298">Value</span></span>|
|---|---|
|[<span data-ttu-id="c7a76-299">最低的邮箱要求集版本</span><span class="sxs-lookup"><span data-stu-id="c7a76-299">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c7a76-300">1.0</span><span class="sxs-lookup"><span data-stu-id="c7a76-300">1.0</span></span>|
|[<span data-ttu-id="c7a76-301">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="c7a76-301">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c7a76-302">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c7a76-302">ReadItem</span></span>|
|[<span data-ttu-id="c7a76-303">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c7a76-303">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="c7a76-304">阅读</span><span class="sxs-lookup"><span data-stu-id="c7a76-304">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c7a76-305">示例</span><span class="sxs-lookup"><span data-stu-id="c7a76-305">Example</span></span>

```
var created = Office.context.mailbox.item.dateTimeCreated;
```

#### <a name="datetimemodified-date"></a><span data-ttu-id="c7a76-306">dateTimeModified： 日期</span><span class="sxs-lookup"><span data-stu-id="c7a76-306">dateTimeModified :Date</span></span>

<span data-ttu-id="c7a76-p110">获取项目最近一次修改的日期和时间。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="c7a76-p110">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="c7a76-309">在 iOS 版 Outlook 或  Android 版 Outlook 中不支持此成员。</span><span class="sxs-lookup"><span data-stu-id="c7a76-309">Note: This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="type"></a><span data-ttu-id="c7a76-310">类型：</span><span class="sxs-lookup"><span data-stu-id="c7a76-310">Type:</span></span>

*   <span data-ttu-id="c7a76-311">日期</span><span class="sxs-lookup"><span data-stu-id="c7a76-311">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="c7a76-312">要求</span><span class="sxs-lookup"><span data-stu-id="c7a76-312">Requirements</span></span>

|<span data-ttu-id="c7a76-313">要求</span><span class="sxs-lookup"><span data-stu-id="c7a76-313">Requirement</span></span>|<span data-ttu-id="c7a76-314">值</span><span class="sxs-lookup"><span data-stu-id="c7a76-314">Value</span></span>|
|---|---|
|[<span data-ttu-id="c7a76-315">最低的邮箱要求集版本</span><span class="sxs-lookup"><span data-stu-id="c7a76-315">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c7a76-316">1.0</span><span class="sxs-lookup"><span data-stu-id="c7a76-316">1.0</span></span>|
|[<span data-ttu-id="c7a76-317">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="c7a76-317">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c7a76-318">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c7a76-318">ReadItem</span></span>|
|[<span data-ttu-id="c7a76-319">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c7a76-319">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="c7a76-320">阅读</span><span class="sxs-lookup"><span data-stu-id="c7a76-320">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c7a76-321">示例</span><span class="sxs-lookup"><span data-stu-id="c7a76-321">Example</span></span>

```
var modified = Office.context.mailbox.item.dateTimeModified;
```

####  <a name="end-datetimejavascriptapioutlook17officetime"></a><span data-ttu-id="c7a76-322">end :日期 |[时间](/javascript/api/outlook_1_7/office.time)</span><span class="sxs-lookup"><span data-stu-id="c7a76-322">end :Date|[Time](/javascript/api/outlook_1_7/office.time)</span></span>

<span data-ttu-id="c7a76-323">获取或设置约会结束的日期和时间。</span><span class="sxs-lookup"><span data-stu-id="c7a76-323">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="c7a76-p111">将 `end` 属性表示为协调世界时 (UTC) 的日期和时间值。可使用 [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook17officelocalclienttime) 方法将 end 属性值转换为客户端的本地日期和时间。</span><span class="sxs-lookup"><span data-stu-id="c7a76-p111">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook17officelocalclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="c7a76-326">阅读模式</span><span class="sxs-lookup"><span data-stu-id="c7a76-326">Read mode</span></span>

<span data-ttu-id="c7a76-327">`end` 属性返回 `Date` 对象。</span><span class="sxs-lookup"><span data-stu-id="c7a76-327">The `end` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="c7a76-328">撰写模式</span><span class="sxs-lookup"><span data-stu-id="c7a76-328">Compose mode</span></span>

<span data-ttu-id="c7a76-329">`end` 属性返回 `Time` 对象。</span><span class="sxs-lookup"><span data-stu-id="c7a76-329">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="c7a76-330">使用 [`Time.setAsync`](/javascript/api/outlook_1_7/office.time#setasync-datetime--options--callback-)   方法设置结束时间时，应使用  [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date)  方法将客户端的本地时间转换为服务器的 UTC。</span><span class="sxs-lookup"><span data-stu-id="c7a76-330">When you use the [`Time.setAsync`](/javascript/api/outlook_1_7/office.time#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="c7a76-331">类型：</span><span class="sxs-lookup"><span data-stu-id="c7a76-331">Type:</span></span>

*   <span data-ttu-id="c7a76-332">日期 | [时间](/javascript/api/outlook_1_7/office.time)</span><span class="sxs-lookup"><span data-stu-id="c7a76-332">Date | [Time](/javascript/api/outlook_1_7/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="c7a76-333">要求</span><span class="sxs-lookup"><span data-stu-id="c7a76-333">Requirements</span></span>

|<span data-ttu-id="c7a76-334">要求</span><span class="sxs-lookup"><span data-stu-id="c7a76-334">Requirement</span></span>|<span data-ttu-id="c7a76-335">值</span><span class="sxs-lookup"><span data-stu-id="c7a76-335">Value</span></span>|
|---|---|
|[<span data-ttu-id="c7a76-336">最低的邮箱要求集版本</span><span class="sxs-lookup"><span data-stu-id="c7a76-336">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c7a76-337">1.0</span><span class="sxs-lookup"><span data-stu-id="c7a76-337">1.0</span></span>|
|[<span data-ttu-id="c7a76-338">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="c7a76-338">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c7a76-339">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c7a76-339">ReadItem</span></span>|
|[<span data-ttu-id="c7a76-340">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c7a76-340">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="c7a76-341">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="c7a76-341">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="c7a76-342">示例</span><span class="sxs-lookup"><span data-stu-id="c7a76-342">Example</span></span>

<span data-ttu-id="c7a76-343">以下示例使用 `Time`  对象的 [`setAsync`](/javascript/api/outlook_1_7/office.time#setasync-datetime--options--callback-)  方法在撰写模式下设置约会的结束时间。</span><span class="sxs-lookup"><span data-stu-id="c7a76-343">The following example sets the end time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_7/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

```
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

#### <a name="from-emailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsfromjavascriptapioutlook17officefrom"></a><span data-ttu-id="c7a76-344"> from :[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)|[发件人](/javascript/api/outlook_1_7/office.from)</span><span class="sxs-lookup"><span data-stu-id="c7a76-344">from :[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)|[From](/javascript/api/outlook_1_7/office.from)</span></span>

<span data-ttu-id="c7a76-345">获取发件人电子邮件地址。</span><span class="sxs-lookup"><span data-stu-id="c7a76-345">Gets the email address of the sender of a message.</span></span>

<span data-ttu-id="c7a76-p112">除非邮件是由代理人发送，否则 `from` 和 [`sender`](#sender-emailaddressdetailsjavascriptapioutlook17officeemailaddressdetails) 属性表示同一个人。在这种情况下， `from` 属性表示代理，发件人属性表示代理人。</span><span class="sxs-lookup"><span data-stu-id="c7a76-p112">The `from` and [`sender`](#sender-emailaddressdetailsjavascriptapioutlook17officeemailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="c7a76-348">`from` 属性内 `EmailAddressDetails` 对象的 `recipientType` 属性是 `undefined` 。</span><span class="sxs-lookup"><span data-stu-id="c7a76-348">Note: The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="c7a76-349">阅读模式</span><span class="sxs-lookup"><span data-stu-id="c7a76-349">Read mode</span></span>

<span data-ttu-id="c7a76-350">`from` 属性返回 `EmailAddressDetails` 对象。</span><span class="sxs-lookup"><span data-stu-id="c7a76-350">The AssignToCategory`from` property always returns an AssignToCategoryRuleAction`EmailAddressDetails` object.</span></span>

```
var subject = Office.context.mailbox.item.from;
```

##### <a name="compose-mode"></a><span data-ttu-id="c7a76-351">撰写模式</span><span class="sxs-lookup"><span data-stu-id="c7a76-351">Compose mode</span></span>

<span data-ttu-id="c7a76-352">`from` 属性返回 `From` 对象，该对象提供获取 from 值的方法。</span><span class="sxs-lookup"><span data-stu-id="c7a76-352">The `from` property returns a `From` object that provides a method to get the from value.</span></span>

```
Office.context.mailbox.item.from.getAsync(callback);

function callback(asyncResult) {
  var from = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="c7a76-353">类型：</span><span class="sxs-lookup"><span data-stu-id="c7a76-353">Type:</span></span>

*   <span data-ttu-id="c7a76-354">[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails) | [发件人](/javascript/api/outlook_1_7/office.from)</span><span class="sxs-lookup"><span data-stu-id="c7a76-354">[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails) | [From](/javascript/api/outlook_1_7/office.from)</span></span>

##### <a name="requirements"></a><span data-ttu-id="c7a76-355">要求</span><span class="sxs-lookup"><span data-stu-id="c7a76-355">Requirements</span></span>

|<span data-ttu-id="c7a76-356">要求</span><span class="sxs-lookup"><span data-stu-id="c7a76-356">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="c7a76-357">最低的邮箱要求集版本</span><span class="sxs-lookup"><span data-stu-id="c7a76-357">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c7a76-358">1.0</span><span class="sxs-lookup"><span data-stu-id="c7a76-358">1.0</span></span>|<span data-ttu-id="c7a76-359">1.7</span><span class="sxs-lookup"><span data-stu-id="c7a76-359">-17</span></span>|
|[<span data-ttu-id="c7a76-360">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="c7a76-360">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c7a76-361">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c7a76-361">ReadItem</span></span>|<span data-ttu-id="c7a76-362">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="c7a76-362">ReadWriteItem</span></span>|
|[<span data-ttu-id="c7a76-363">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c7a76-363">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="c7a76-364">阅读</span><span class="sxs-lookup"><span data-stu-id="c7a76-364">Read</span></span>|<span data-ttu-id="c7a76-365">撰写</span><span class="sxs-lookup"><span data-stu-id="c7a76-365">Compose</span></span>|

#### <a name="internetmessageid-string"></a><span data-ttu-id="c7a76-366">internetMessageId :String</span><span class="sxs-lookup"><span data-stu-id="c7a76-366">internetMessageId :String</span></span>

<span data-ttu-id="c7a76-p113">获取电子邮件的 Internet 消息标识符。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="c7a76-p113">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="c7a76-369">类型：</span><span class="sxs-lookup"><span data-stu-id="c7a76-369">Type:</span></span>

*   <span data-ttu-id="c7a76-370">字符串</span><span class="sxs-lookup"><span data-stu-id="c7a76-370">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="c7a76-371">要求</span><span class="sxs-lookup"><span data-stu-id="c7a76-371">Requirements</span></span>

|<span data-ttu-id="c7a76-372">要求</span><span class="sxs-lookup"><span data-stu-id="c7a76-372">Requirement</span></span>|<span data-ttu-id="c7a76-373">值</span><span class="sxs-lookup"><span data-stu-id="c7a76-373">Value</span></span>|
|---|---|
|[<span data-ttu-id="c7a76-374">最低的邮箱要求集版本</span><span class="sxs-lookup"><span data-stu-id="c7a76-374">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c7a76-375">1.0</span><span class="sxs-lookup"><span data-stu-id="c7a76-375">1.0</span></span>|
|[<span data-ttu-id="c7a76-376">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="c7a76-376">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c7a76-377">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c7a76-377">ReadItem</span></span>|
|[<span data-ttu-id="c7a76-378">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c7a76-378">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="c7a76-379">阅读</span><span class="sxs-lookup"><span data-stu-id="c7a76-379">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c7a76-380">示例</span><span class="sxs-lookup"><span data-stu-id="c7a76-380">Example</span></span>

```
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

#### <a name="itemclass-string"></a><span data-ttu-id="c7a76-381">itemClass：字符串</span><span class="sxs-lookup"><span data-stu-id="c7a76-381">itemClass :String</span></span>

<span data-ttu-id="c7a76-p114">获取选定项目的 Exchange Web 服务项目类。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="c7a76-p114">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="c7a76-p115">`itemClass` 属性指定所选项目的邮件类别。以下是邮件或约会项目的默认邮件类别。</span><span class="sxs-lookup"><span data-stu-id="c7a76-p115">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

|<span data-ttu-id="c7a76-386">类型</span><span class="sxs-lookup"><span data-stu-id="c7a76-386">Type</span></span>|<span data-ttu-id="c7a76-387">说明</span><span class="sxs-lookup"><span data-stu-id="c7a76-387">Description</span></span>|<span data-ttu-id="c7a76-388">项目类</span><span class="sxs-lookup"><span data-stu-id="c7a76-388">item class</span></span>|
|---|---|---|
|<span data-ttu-id="c7a76-389">约会项目</span><span class="sxs-lookup"><span data-stu-id="c7a76-389">Appointment items</span></span>|<span data-ttu-id="c7a76-390">这些是项目类 `IPM.Appointment` 或 `IPM.Appointment.Occurence` 的日历项目。</span><span class="sxs-lookup"><span data-stu-id="c7a76-390">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurence`.</span></span>|`IPM.Appointment`<br />`IPM.Appointment.Occurence`|
|<span data-ttu-id="c7a76-391">邮件项目</span><span class="sxs-lookup"><span data-stu-id="c7a76-391">Message items</span></span>|<span data-ttu-id="c7a76-392">这些项目包括具有默认邮件类别 `IPM.Note` 的电子邮件，以及使用 `IPM.Schedule.Meeting` 作为基础邮件类别的会议请求、响应和取消。</span><span class="sxs-lookup"><span data-stu-id="c7a76-392">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span>|`IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled`|

<span data-ttu-id="c7a76-393">你可以创建用于扩展默认邮件类别的自定义邮件类别，例如，自定义约会邮件类别 `IPM.Appointment.Contoso`。</span><span class="sxs-lookup"><span data-stu-id="c7a76-393">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="c7a76-394">类型：</span><span class="sxs-lookup"><span data-stu-id="c7a76-394">Type:</span></span>

*   <span data-ttu-id="c7a76-395">字符串</span><span class="sxs-lookup"><span data-stu-id="c7a76-395">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="c7a76-396">要求</span><span class="sxs-lookup"><span data-stu-id="c7a76-396">Requirements</span></span>

|<span data-ttu-id="c7a76-397">要求</span><span class="sxs-lookup"><span data-stu-id="c7a76-397">Requirement</span></span>|<span data-ttu-id="c7a76-398">值</span><span class="sxs-lookup"><span data-stu-id="c7a76-398">Value</span></span>|
|---|---|
|[<span data-ttu-id="c7a76-399">最低的邮箱要求集版本</span><span class="sxs-lookup"><span data-stu-id="c7a76-399">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c7a76-400">1.0</span><span class="sxs-lookup"><span data-stu-id="c7a76-400">1.0</span></span>|
|[<span data-ttu-id="c7a76-401">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="c7a76-401">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c7a76-402">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c7a76-402">ReadItem</span></span>|
|[<span data-ttu-id="c7a76-403">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c7a76-403">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="c7a76-404">阅读</span><span class="sxs-lookup"><span data-stu-id="c7a76-404">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c7a76-405">示例</span><span class="sxs-lookup"><span data-stu-id="c7a76-405">Example</span></span>

```
var itemClass = Office.context.mailbox.item.itemClass;
```

#### <a name="nullable-itemid-string"></a><span data-ttu-id="c7a76-406">（可为空）itemId :字符串</span><span class="sxs-lookup"><span data-stu-id="c7a76-406">(nullable) itemId :String</span></span>

<span data-ttu-id="c7a76-p116">获取当前项目的 Exchange Web 服务项目标识符。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="c7a76-p116">Gets the Exchange Web Services item identifier for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="c7a76-409">`itemId` 属性返回的标识符与 Exchange Web 服务项目标识符相同。</span><span class="sxs-lookup"><span data-stu-id="c7a76-409">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="c7a76-410">`itemId` 属性与 Outlook 条目 ID 或使用 Outlook REST API 的 ID不同。</span><span class="sxs-lookup"><span data-stu-id="c7a76-410">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="c7a76-411">使用此值的 REST API 调用之前，应使用 [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string)将其转换。</span><span class="sxs-lookup"><span data-stu-id="c7a76-411">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="c7a76-412">有关详细信息，请参阅 [从 Outlook 外接程序使用 Outlook REST API](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id)。</span><span class="sxs-lookup"><span data-stu-id="c7a76-412">For more details, see [Use the Outlook REST APIs from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

<span data-ttu-id="c7a76-p118">`itemId` 属性在撰写模式下不可用。如果需要项目标识符，[`saveAsync`](#saveasyncoptions-callback) 方法可用于将项目保存到存储，这将在回调函数的 [`AsyncResult.value`](/javascript/api/office/office.asyncresult) 参数中返回项目标识符。</span><span class="sxs-lookup"><span data-stu-id="c7a76-p118">The `itemId` property is not available in compose mode. If an item identifier is required, the [`saveAsync`](#saveasyncoptions-callback) method can be used to save the item to the store, which will return the item identifier in the [`AsyncResult.value`](/javascript/api/office/office.asyncresult) parameter in the callback function.</span></span>

##### <a name="type"></a><span data-ttu-id="c7a76-415">类型：</span><span class="sxs-lookup"><span data-stu-id="c7a76-415">Type:</span></span>

*   <span data-ttu-id="c7a76-416">字符串</span><span class="sxs-lookup"><span data-stu-id="c7a76-416">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="c7a76-417">要求</span><span class="sxs-lookup"><span data-stu-id="c7a76-417">Requirements</span></span>

|<span data-ttu-id="c7a76-418">要求</span><span class="sxs-lookup"><span data-stu-id="c7a76-418">Requirement</span></span>|<span data-ttu-id="c7a76-419">值</span><span class="sxs-lookup"><span data-stu-id="c7a76-419">Value</span></span>|
|---|---|
|[<span data-ttu-id="c7a76-420">最低的邮箱要求集版本</span><span class="sxs-lookup"><span data-stu-id="c7a76-420">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c7a76-421">1.0</span><span class="sxs-lookup"><span data-stu-id="c7a76-421">1.0</span></span>|
|[<span data-ttu-id="c7a76-422">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="c7a76-422">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c7a76-423">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c7a76-423">ReadItem</span></span>|
|[<span data-ttu-id="c7a76-424">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c7a76-424">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="c7a76-425">阅读</span><span class="sxs-lookup"><span data-stu-id="c7a76-425">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c7a76-426">示例</span><span class="sxs-lookup"><span data-stu-id="c7a76-426">Example</span></span>

<span data-ttu-id="c7a76-p119">以下代码检查项目标识符是否存在。如果 `itemId` 属性返回 `null` 或 `undefined`，则将项目保存到存储，并从异步结果中获取项目标识符。</span><span class="sxs-lookup"><span data-stu-id="c7a76-p119">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

```
var itemId = Office.context.mailbox.item.itemId;
if (itemId === null || itemId == undefined) {
  Office.context.mailbox.item.saveAsync(function(result){
    itemId = result.value;
  });
}
```

####  <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlook17officemailboxenumsitemtype"></a><span data-ttu-id="c7a76-429">itemType:[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_7/office.mailboxenums.itemtype)</span><span class="sxs-lookup"><span data-stu-id="c7a76-429">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_7/office.mailboxenums.itemtype)</span></span>

<span data-ttu-id="c7a76-430">获取实例代表项的类型。</span><span class="sxs-lookup"><span data-stu-id="c7a76-430">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="c7a76-431">`itemType` 属性返回其中一个 `ItemType` 枚举值，指示 `item` 对象实例是邮件还是约会。</span><span class="sxs-lookup"><span data-stu-id="c7a76-431">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="c7a76-432">类型：</span><span class="sxs-lookup"><span data-stu-id="c7a76-432">Type:</span></span>

*   [<span data-ttu-id="c7a76-433">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="c7a76-433">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook_1_7/office.mailboxenums.itemtype)

##### <a name="requirements"></a><span data-ttu-id="c7a76-434">要求</span><span class="sxs-lookup"><span data-stu-id="c7a76-434">Requirements</span></span>

|<span data-ttu-id="c7a76-435">要求</span><span class="sxs-lookup"><span data-stu-id="c7a76-435">Requirement</span></span>|<span data-ttu-id="c7a76-436">值</span><span class="sxs-lookup"><span data-stu-id="c7a76-436">Value</span></span>|
|---|---|
|[<span data-ttu-id="c7a76-437">最低的邮箱要求集版本</span><span class="sxs-lookup"><span data-stu-id="c7a76-437">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c7a76-438">1.0</span><span class="sxs-lookup"><span data-stu-id="c7a76-438">1.0</span></span>|
|[<span data-ttu-id="c7a76-439">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="c7a76-439">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c7a76-440">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c7a76-440">ReadItem</span></span>|
|[<span data-ttu-id="c7a76-441">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c7a76-441">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="c7a76-442">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="c7a76-442">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="c7a76-443">示例</span><span class="sxs-lookup"><span data-stu-id="c7a76-443">Example</span></span>

```
if (Office.context.mailbox.item.itemType == Office.MailboxEnums.ItemType.Message)
  // do something
else
  // do something else
```

####  <a name="location-stringlocationjavascriptapioutlook17officelocation"></a><span data-ttu-id="c7a76-444">location :字符串 |[位置](/javascript/api/outlook_1_7/office.location)</span><span class="sxs-lookup"><span data-stu-id="c7a76-444">location :String|[Location](/javascript/api/outlook_1_7/office.location)</span></span>

<span data-ttu-id="c7a76-445">获取或设置约会的位置。</span><span class="sxs-lookup"><span data-stu-id="c7a76-445">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="c7a76-446">阅读模式</span><span class="sxs-lookup"><span data-stu-id="c7a76-446">Read mode</span></span>

<span data-ttu-id="c7a76-447">`location` 属性返回一个包含约会位置的字符串。</span><span class="sxs-lookup"><span data-stu-id="c7a76-447">The `location` property returns a string that contains the location of the appointment.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="c7a76-448">撰写模式</span><span class="sxs-lookup"><span data-stu-id="c7a76-448">Compose mode</span></span>

<span data-ttu-id="c7a76-449">`location` 属性返回一个 `Location` 对象，该对象提供用于获取和设置约会位置的方法。</span><span class="sxs-lookup"><span data-stu-id="c7a76-449">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="c7a76-450">类型：</span><span class="sxs-lookup"><span data-stu-id="c7a76-450">Type:</span></span>

*   <span data-ttu-id="c7a76-451">字符串 | [位置](/javascript/api/outlook_1_7/office.location)</span><span class="sxs-lookup"><span data-stu-id="c7a76-451">String | [Location](/javascript/api/outlook_1_7/office.location)</span></span>

##### <a name="requirements"></a><span data-ttu-id="c7a76-452">要求</span><span class="sxs-lookup"><span data-stu-id="c7a76-452">Requirements</span></span>

|<span data-ttu-id="c7a76-453">要求</span><span class="sxs-lookup"><span data-stu-id="c7a76-453">Requirement</span></span>|<span data-ttu-id="c7a76-454">值</span><span class="sxs-lookup"><span data-stu-id="c7a76-454">Value</span></span>|
|---|---|
|[<span data-ttu-id="c7a76-455">最低的邮箱要求集版本</span><span class="sxs-lookup"><span data-stu-id="c7a76-455">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c7a76-456">1.0</span><span class="sxs-lookup"><span data-stu-id="c7a76-456">1.0</span></span>|
|[<span data-ttu-id="c7a76-457">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="c7a76-457">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c7a76-458">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c7a76-458">ReadItem</span></span>|
|[<span data-ttu-id="c7a76-459">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c7a76-459">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="c7a76-460">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="c7a76-460">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="c7a76-461">示例</span><span class="sxs-lookup"><span data-stu-id="c7a76-461">Example</span></span>

```
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

#### <a name="normalizedsubject-string"></a><span data-ttu-id="c7a76-462">normalizedSubject :字符串</span><span class="sxs-lookup"><span data-stu-id="c7a76-462">normalizedSubject :String</span></span>

<span data-ttu-id="c7a76-p120">获取删除了所有前缀（包括 `RE:` 和 `FWD:`）的项目主题。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="c7a76-p120">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="c7a76-p121">normalizedSubject 属性获取包含由电子邮件程序添加的任何标准前缀（如 `RE:` 和 `FW:`）的项目主题。若要获取包含完整前缀的项目主题，请使用 [`subject`](#subject-stringsubjectjavascriptapioutlook17officesubject) 属性。</span><span class="sxs-lookup"><span data-stu-id="c7a76-p121">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubjectjavascriptapioutlook17officesubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="c7a76-467">类型：</span><span class="sxs-lookup"><span data-stu-id="c7a76-467">Type:</span></span>

*   <span data-ttu-id="c7a76-468">字符串</span><span class="sxs-lookup"><span data-stu-id="c7a76-468">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="c7a76-469">要求</span><span class="sxs-lookup"><span data-stu-id="c7a76-469">Requirements</span></span>

|<span data-ttu-id="c7a76-470">要求</span><span class="sxs-lookup"><span data-stu-id="c7a76-470">Requirement</span></span>|<span data-ttu-id="c7a76-471">值</span><span class="sxs-lookup"><span data-stu-id="c7a76-471">Value</span></span>|
|---|---|
|[<span data-ttu-id="c7a76-472">最低的邮箱要求集版本</span><span class="sxs-lookup"><span data-stu-id="c7a76-472">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c7a76-473">1.0</span><span class="sxs-lookup"><span data-stu-id="c7a76-473">1.0</span></span>|
|[<span data-ttu-id="c7a76-474">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="c7a76-474">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c7a76-475">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c7a76-475">ReadItem</span></span>|
|[<span data-ttu-id="c7a76-476">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c7a76-476">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="c7a76-477">阅读</span><span class="sxs-lookup"><span data-stu-id="c7a76-477">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c7a76-478">示例</span><span class="sxs-lookup"><span data-stu-id="c7a76-478">Example</span></span>

```
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
```

####  <a name="notificationmessages-notificationmessagesjavascriptapioutlook17officenotificationmessages"></a><span data-ttu-id="c7a76-479">notificationMessages :[NotificationMessages](/javascript/api/outlook_1_7/office.notificationmessages)</span><span class="sxs-lookup"><span data-stu-id="c7a76-479">notificationMessages :[NotificationMessages](/javascript/api/outlook_1_7/office.notificationmessages)</span></span>

<span data-ttu-id="c7a76-480">获取一个项目的通知邮件。</span><span class="sxs-lookup"><span data-stu-id="c7a76-480">Gets the notification messages for an item.</span></span>

##### <a name="type"></a><span data-ttu-id="c7a76-481">类型：</span><span class="sxs-lookup"><span data-stu-id="c7a76-481">Type:</span></span>

*   [<span data-ttu-id="c7a76-482">NotificationMessages</span><span class="sxs-lookup"><span data-stu-id="c7a76-482">NotificationMessages</span></span>](/javascript/api/outlook_1_7/office.notificationmessages)

##### <a name="requirements"></a><span data-ttu-id="c7a76-483">要求</span><span class="sxs-lookup"><span data-stu-id="c7a76-483">Requirements</span></span>

|<span data-ttu-id="c7a76-484">要求</span><span class="sxs-lookup"><span data-stu-id="c7a76-484">Requirement</span></span>|<span data-ttu-id="c7a76-485">值</span><span class="sxs-lookup"><span data-stu-id="c7a76-485">Value</span></span>|
|---|---|
|[<span data-ttu-id="c7a76-486">最低的邮箱要求集版本</span><span class="sxs-lookup"><span data-stu-id="c7a76-486">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c7a76-487">1.3</span><span class="sxs-lookup"><span data-stu-id="c7a76-487">1.3</span></span>|
|[<span data-ttu-id="c7a76-488">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="c7a76-488">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c7a76-489">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c7a76-489">ReadItem</span></span>|
|[<span data-ttu-id="c7a76-490">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c7a76-490">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="c7a76-491">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="c7a76-491">Compose or read</span></span>|

####  <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsrecipientsjavascriptapioutlook17officerecipients"></a><span data-ttu-id="c7a76-492">optionalAttendees :数组.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)>|[收件人](/javascript/api/outlook_1_7/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="c7a76-492">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_7/office.recipients)</span></span>

<span data-ttu-id="c7a76-493">提供对事件可选与会者的访问。</span><span class="sxs-lookup"><span data-stu-id="c7a76-493">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="c7a76-494">对象的类型和访问级别取决于当前项的模式。</span><span class="sxs-lookup"><span data-stu-id="c7a76-494">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="c7a76-495">阅读模式</span><span class="sxs-lookup"><span data-stu-id="c7a76-495">Read mode</span></span>

<span data-ttu-id="c7a76-496">`optionalAttendees` 属性返回一个数组，其中包含每个可选与会者的 `EmailAddressDetails` 对象。</span><span class="sxs-lookup"><span data-stu-id="c7a76-496">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="c7a76-497">撰写模式</span><span class="sxs-lookup"><span data-stu-id="c7a76-497">Compose mode</span></span>

<span data-ttu-id="c7a76-498">`optionalAttendees` 属性返回一个 `Recipients` 对象，该对象提供用于获取和设置可选与会者的方法。</span><span class="sxs-lookup"><span data-stu-id="c7a76-498">The `optionalAttendees` property returns a `Recipients` object that provides methods to get and set the optional attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="c7a76-499">类型：</span><span class="sxs-lookup"><span data-stu-id="c7a76-499">Type:</span></span>

*   <span data-ttu-id="c7a76-500">数组.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)> |[收件人](/javascript/api/outlook_1_7/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="c7a76-500">Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_7/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="c7a76-501">要求</span><span class="sxs-lookup"><span data-stu-id="c7a76-501">Requirements</span></span>

|<span data-ttu-id="c7a76-502">要求</span><span class="sxs-lookup"><span data-stu-id="c7a76-502">Requirement</span></span>|<span data-ttu-id="c7a76-503">值</span><span class="sxs-lookup"><span data-stu-id="c7a76-503">Value</span></span>|
|---|---|
|[<span data-ttu-id="c7a76-504">最低的邮箱要求集版本</span><span class="sxs-lookup"><span data-stu-id="c7a76-504">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c7a76-505">1.0</span><span class="sxs-lookup"><span data-stu-id="c7a76-505">1.0</span></span>|
|[<span data-ttu-id="c7a76-506">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="c7a76-506">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c7a76-507">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c7a76-507">ReadItem</span></span>|
|[<span data-ttu-id="c7a76-508">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c7a76-508">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="c7a76-509">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="c7a76-509">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="c7a76-510">示例</span><span class="sxs-lookup"><span data-stu-id="c7a76-510">Example</span></span>

```
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

#### <a name="organizer-emailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsorganizerjavascriptapioutlook17officeorganizer"></a><span data-ttu-id="c7a76-511">organizer :[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)|[组织者](/javascript/api/outlook_1_7/office.organizer)</span><span class="sxs-lookup"><span data-stu-id="c7a76-511">organizer :[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)|[Organizer](/javascript/api/outlook_1_7/office.organizer)</span></span>

<span data-ttu-id="c7a76-512">获取特定会议组织者的电子邮件地址。</span><span class="sxs-lookup"><span data-stu-id="c7a76-512">Gets the email address of the meeting organizer for a specified meeting.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="c7a76-513">阅读模式</span><span class="sxs-lookup"><span data-stu-id="c7a76-513">Read mode</span></span>

<span data-ttu-id="c7a76-514">`organizer` 属性返回 [EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails) 对象，该对象表示会议组织者。</span><span class="sxs-lookup"><span data-stu-id="c7a76-514">The `organizer` property returns an [EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails) object that represents the meeting organizer.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="c7a76-515">撰写模式</span><span class="sxs-lookup"><span data-stu-id="c7a76-515">Compose mode</span></span>

<span data-ttu-id="c7a76-516">`organizer` 属性返回 [组织者](/javascript/api/outlook_1_7/office.organizer) 对象，它返回获取 organizer 值的方法。</span><span class="sxs-lookup"><span data-stu-id="c7a76-516">The `organizer` property returns an [Organizer](/javascript/api/outlook_1_7/office.organizer) object that provides a method to get the organizer value.</span></span>

##### <a name="type"></a><span data-ttu-id="c7a76-517">类型：</span><span class="sxs-lookup"><span data-stu-id="c7a76-517">Type:</span></span>

*   <span data-ttu-id="c7a76-518">[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails) | [组织者](/javascript/api/outlook_1_7/office.organizer)</span><span class="sxs-lookup"><span data-stu-id="c7a76-518">[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails) | [Organizer](/javascript/api/outlook_1_7/office.organizer)</span></span>

##### <a name="requirements"></a><span data-ttu-id="c7a76-519">要求</span><span class="sxs-lookup"><span data-stu-id="c7a76-519">Requirements</span></span>

|<span data-ttu-id="c7a76-520">要求</span><span class="sxs-lookup"><span data-stu-id="c7a76-520">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="c7a76-521">最低的邮箱要求集版本</span><span class="sxs-lookup"><span data-stu-id="c7a76-521">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c7a76-522">1.0</span><span class="sxs-lookup"><span data-stu-id="c7a76-522">1.0</span></span>|<span data-ttu-id="c7a76-523">1.7</span><span class="sxs-lookup"><span data-stu-id="c7a76-523">-17</span></span>|
|[<span data-ttu-id="c7a76-524">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="c7a76-524">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c7a76-525">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c7a76-525">ReadItem</span></span>|<span data-ttu-id="c7a76-526">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="c7a76-526">ReadWriteItem</span></span>|
|[<span data-ttu-id="c7a76-527">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c7a76-527">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="c7a76-528">阅读</span><span class="sxs-lookup"><span data-stu-id="c7a76-528">Read</span></span>|<span data-ttu-id="c7a76-529">撰写</span><span class="sxs-lookup"><span data-stu-id="c7a76-529">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="c7a76-530">示例</span><span class="sxs-lookup"><span data-stu-id="c7a76-530">Example</span></span>

```
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
```

#### <a name="nullable-recurrence-recurrencejavascriptapioutlook17officerecurrence"></a><span data-ttu-id="c7a76-531">（可为空）recurrence :[重复周期](/javascript/api/outlook_1_7/office.recurrence)</span><span class="sxs-lookup"><span data-stu-id="c7a76-531">(nullable) recurrence :[Recurrence](/javascript/api/outlook_1_7/office.recurrence)</span></span>

<span data-ttu-id="c7a76-532">获取或设置约会的定期模式。</span><span class="sxs-lookup"><span data-stu-id="c7a76-532">Gets or sets the location of an appointment.</span></span> <span data-ttu-id="c7a76-533">获取会议请求的定期模式。</span><span class="sxs-lookup"><span data-stu-id="c7a76-533">Gets the recurrence pattern of a meeting request.</span></span> <span data-ttu-id="c7a76-534">约会项目的阅读和撰写模式。</span><span class="sxs-lookup"><span data-stu-id="c7a76-534">Read and compose modes for appointment items.</span></span> <span data-ttu-id="c7a76-535">会议请求项的阅读模式。</span><span class="sxs-lookup"><span data-stu-id="c7a76-535">Read mode for meeting request items.</span></span>

<span data-ttu-id="c7a76-536">如果项目是序列或系列的一个实例，`recurrence` 属性返回定期约会或会议请求的 [recurrence](/javascript/api/outlook_1_7/office.recurrence) 对象。</span><span class="sxs-lookup"><span data-stu-id="c7a76-536">The `recurrence` property returns a [recurrence](/javascript/api/outlook_1_7/office.recurrence) object for recurring appointments or meetings requests if an item is a series or an instance in a series.</span></span> <span data-ttu-id="c7a76-537">`null` 是针对单一约会和单一约会会议请求的返回。</span><span class="sxs-lookup"><span data-stu-id="c7a76-537">`null` is returned for single appointments and meeting requests of single appointments.</span></span> <span data-ttu-id="c7a76-538">`undefined` 是针对非会议请求邮件的返回。</span><span class="sxs-lookup"><span data-stu-id="c7a76-538">`undefined` is returned for messages that are not meeting requests.</span></span>

> <span data-ttu-id="c7a76-539">注释：会议请求具有 IPM.Schedule.Meeting.RequestMeeting 的 `itemClass` 值。</span><span class="sxs-lookup"><span data-stu-id="c7a76-539">Note: Meeting requests have an `itemClass` value of IPM.Schedule.Meeting.Request.</span></span>

> <span data-ttu-id="c7a76-540">注释：如果重复周期对象为 `null` ，这表示此对象是单一约会或者是单一约会的会议请求而非序列的组成部分。</span><span class="sxs-lookup"><span data-stu-id="c7a76-540">Note: If the recurrence object is `null`, this indicates that the object is a single appointment or a meeting request of a single appointment and NOT a part of a series.</span></span>

##### <a name="type"></a><span data-ttu-id="c7a76-541">类型：</span><span class="sxs-lookup"><span data-stu-id="c7a76-541">Type:</span></span>

* [<span data-ttu-id="c7a76-542">重复周期</span><span class="sxs-lookup"><span data-stu-id="c7a76-542">recurrence</span></span>](/javascript/api/outlook_1_7/office.recurrence)

|<span data-ttu-id="c7a76-543">要求</span><span class="sxs-lookup"><span data-stu-id="c7a76-543">Requirement</span></span>|<span data-ttu-id="c7a76-544">值</span><span class="sxs-lookup"><span data-stu-id="c7a76-544">Value</span></span>|
|---|---|
|[<span data-ttu-id="c7a76-545">最低的邮箱要求集版本</span><span class="sxs-lookup"><span data-stu-id="c7a76-545">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c7a76-546">1.7</span><span class="sxs-lookup"><span data-stu-id="c7a76-546">-17</span></span>|
|[<span data-ttu-id="c7a76-547">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="c7a76-547">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c7a76-548">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c7a76-548">ReadItem</span></span>|
|[<span data-ttu-id="c7a76-549">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c7a76-549">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="c7a76-550">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="c7a76-550">Compose or read</span></span>|

####  <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsrecipientsjavascriptapioutlook17officerecipients"></a><span data-ttu-id="c7a76-551">requiredAttendees :数组.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)>|[收件人](/javascript/api/outlook_1_7/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="c7a76-551">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_7/office.recipients)</span></span>

<span data-ttu-id="c7a76-552">提供对事件必需与会者的访问。</span><span class="sxs-lookup"><span data-stu-id="c7a76-552">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="c7a76-553">对象的类型和访问级别取决于当前项的模式。</span><span class="sxs-lookup"><span data-stu-id="c7a76-553">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="c7a76-554">阅读模式</span><span class="sxs-lookup"><span data-stu-id="c7a76-554">Read mode</span></span>

<span data-ttu-id="c7a76-555">`requiredAttendees` 属性返回一个数组，其中包含每个必需与会者的 `EmailAddressDetails` 对象。</span><span class="sxs-lookup"><span data-stu-id="c7a76-555">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="c7a76-556">撰写模式</span><span class="sxs-lookup"><span data-stu-id="c7a76-556">Compose mode</span></span>

<span data-ttu-id="c7a76-557">`requiredAttendees` 属性返回一个 `Recipients` 对象，该对象提供用于获取和设置可选与会者的方法。</span><span class="sxs-lookup"><span data-stu-id="c7a76-557">The `requiredAttendees` property returns a `Recipients` object that provides methods to get and set the required attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="c7a76-558">类型：</span><span class="sxs-lookup"><span data-stu-id="c7a76-558">Type:</span></span>

*   <span data-ttu-id="c7a76-559">数组.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)> |[收件人](/javascript/api/outlook_1_7/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="c7a76-559">Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_7/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="c7a76-560">要求</span><span class="sxs-lookup"><span data-stu-id="c7a76-560">Requirements</span></span>

|<span data-ttu-id="c7a76-561">要求</span><span class="sxs-lookup"><span data-stu-id="c7a76-561">Requirement</span></span>|<span data-ttu-id="c7a76-562">值</span><span class="sxs-lookup"><span data-stu-id="c7a76-562">Value</span></span>|
|---|---|
|[<span data-ttu-id="c7a76-563">最低的邮箱要求集版本</span><span class="sxs-lookup"><span data-stu-id="c7a76-563">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c7a76-564">1.0</span><span class="sxs-lookup"><span data-stu-id="c7a76-564">1.0</span></span>|
|[<span data-ttu-id="c7a76-565">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="c7a76-565">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c7a76-566">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c7a76-566">ReadItem</span></span>|
|[<span data-ttu-id="c7a76-567">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c7a76-567">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="c7a76-568">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="c7a76-568">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="c7a76-569">示例</span><span class="sxs-lookup"><span data-stu-id="c7a76-569">Example</span></span>

```
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
}
```

#### <a name="sender-emailaddressdetailsjavascriptapioutlook17officeemailaddressdetails"></a><span data-ttu-id="c7a76-570">sender :[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="c7a76-570">sender :[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)</span></span>

<span data-ttu-id="c7a76-p126">获取电子邮件发件人的电子邮件地址。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="c7a76-p126">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="c7a76-p127">[`from`](#from-emailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsfromjavascriptapioutlook17officefrom) 和 `sender` 属性表示同一个人，邮件由代理人发送的除外。在此情况下，`from` 属性表示代理程序，而 sender 属性表示代理人。</span><span class="sxs-lookup"><span data-stu-id="c7a76-p127">The [`from`](#from-emailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsfromjavascriptapioutlook17officefrom) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="c7a76-575">`sender` 属性内 `EmailAddressDetails` 对象的 `recipientType` 属性是 `undefined` 。</span><span class="sxs-lookup"><span data-stu-id="c7a76-575">Note: The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="c7a76-576">类型：</span><span class="sxs-lookup"><span data-stu-id="c7a76-576">Type:</span></span>

*   [<span data-ttu-id="c7a76-577">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="c7a76-577">EmailAddressDetails</span></span>](/javascript/api/outlook_1_7/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="c7a76-578">要求</span><span class="sxs-lookup"><span data-stu-id="c7a76-578">Requirements</span></span>

|<span data-ttu-id="c7a76-579">要求</span><span class="sxs-lookup"><span data-stu-id="c7a76-579">Requirement</span></span>|<span data-ttu-id="c7a76-580">值</span><span class="sxs-lookup"><span data-stu-id="c7a76-580">Value</span></span>|
|---|---|
|[<span data-ttu-id="c7a76-581">最低的邮箱要求集版本</span><span class="sxs-lookup"><span data-stu-id="c7a76-581">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c7a76-582">1.0</span><span class="sxs-lookup"><span data-stu-id="c7a76-582">1.0</span></span>|
|[<span data-ttu-id="c7a76-583">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="c7a76-583">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c7a76-584">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c7a76-584">ReadItem</span></span>|
|[<span data-ttu-id="c7a76-585">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c7a76-585">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="c7a76-586">阅读</span><span class="sxs-lookup"><span data-stu-id="c7a76-586">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c7a76-587">示例</span><span class="sxs-lookup"><span data-stu-id="c7a76-587">Example</span></span>

```
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
```

#### <a name="nullable-seriesid-string"></a><span data-ttu-id="c7a76-588">（可为空） seriesId :字符串</span><span class="sxs-lookup"><span data-stu-id="c7a76-588">(nullable) seriesId :String</span></span>

<span data-ttu-id="c7a76-589">获取示例所属序列的 id 。</span><span class="sxs-lookup"><span data-stu-id="c7a76-589">Gets the id of the series that an instance belongs to.</span></span>

<span data-ttu-id="c7a76-590">在 OWA 和 Outlook 中， `seriesId` 返回此项所属父级（序列）项的 Exchange Web 服务 (EWS) ID 。</span><span class="sxs-lookup"><span data-stu-id="c7a76-590">In OWA and Outlook, the `seriesId` returns the Exchange Web Services (EWS) ID of the parent (series) item that this item belongs to.</span></span> <span data-ttu-id="c7a76-591">但在 iOS 和 Android 中，`seriesId` 返回父级项的 REST ID 。</span><span class="sxs-lookup"><span data-stu-id="c7a76-591">However, in iOS and Android, the `seriesId` returns the REST ID of the parent item.</span></span>

> [!NOTE]
> <span data-ttu-id="c7a76-592">`seriesId` 属性返回的标识符与 Exchange Web 服务项目标识符相同。</span><span class="sxs-lookup"><span data-stu-id="c7a76-592">The identifier returned by the `seriesId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="c7a76-593"> `seriesId` 属性与 Outlook REST API  所用的 Outlook ID 不同。</span><span class="sxs-lookup"><span data-stu-id="c7a76-593">The `seriesId` property is not identical to the Outlook IDs used by the Outlook REST API.</span></span> <span data-ttu-id="c7a76-594">使用此值的 REST API 调用之前，应使用 [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string)将其转换。</span><span class="sxs-lookup"><span data-stu-id="c7a76-594">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="c7a76-595">有关详细信息，请参阅 [从 Outlook 外接程序使用 Outlook REST API](https://docs.microsoft.com/outlook/add-ins/use-rest-api)。</span><span class="sxs-lookup"><span data-stu-id="c7a76-595">For more details, see [Use the Outlook REST APIs from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/use-rest-api).</span></span>

<span data-ttu-id="c7a76-596">对于没有父级项的项目，如单一约会、序列项目或会议请求，`seriesId`属性返回 `null`，对于其他非会议请求项目，则返回 `undefined`。</span><span class="sxs-lookup"><span data-stu-id="c7a76-596">The `seriesId` property returns `null` for items that do not have parent items such as single appointments, series items, or meeting requests and returns `undefined` for any other items that are not meeting requests.</span></span>

##### <a name="type"></a><span data-ttu-id="c7a76-597">类型：</span><span class="sxs-lookup"><span data-stu-id="c7a76-597">Type:</span></span>

* <span data-ttu-id="c7a76-598">字符串</span><span class="sxs-lookup"><span data-stu-id="c7a76-598">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="c7a76-599">要求</span><span class="sxs-lookup"><span data-stu-id="c7a76-599">Requirements</span></span>

|<span data-ttu-id="c7a76-600">要求</span><span class="sxs-lookup"><span data-stu-id="c7a76-600">Requirement</span></span>|<span data-ttu-id="c7a76-601">值</span><span class="sxs-lookup"><span data-stu-id="c7a76-601">Value</span></span>|
|---|---|
|[<span data-ttu-id="c7a76-602">最低的邮箱要求集版本</span><span class="sxs-lookup"><span data-stu-id="c7a76-602">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c7a76-603">1.7</span><span class="sxs-lookup"><span data-stu-id="c7a76-603">-17</span></span>|
|[<span data-ttu-id="c7a76-604">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="c7a76-604">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c7a76-605">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c7a76-605">ReadItem</span></span>|
|[<span data-ttu-id="c7a76-606">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c7a76-606">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="c7a76-607">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="c7a76-607">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="c7a76-608">示例</span><span class="sxs-lookup"><span data-stu-id="c7a76-608">Example</span></span>

```
var seriesId = Office.context.mailbox.item.seriesId; 
var isSeries = (seriesId == null);
```

####  <a name="start-datetimejavascriptapioutlook17officetime"></a><span data-ttu-id="c7a76-609">start :日期 |[时间](/javascript/api/outlook_1_7/office.time)</span><span class="sxs-lookup"><span data-stu-id="c7a76-609">start :Date|[Time](/javascript/api/outlook_1_7/office.time)</span></span>

<span data-ttu-id="c7a76-610">获取或设置约会开始的日期和时间。</span><span class="sxs-lookup"><span data-stu-id="c7a76-610">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="c7a76-p130">将 `start` 属性表示为协调世界时 (UTC) 的日期和时间值。可使用 [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook17officelocalclienttime) 方法将该值转换为客户端的本地日期和时间。</span><span class="sxs-lookup"><span data-stu-id="c7a76-p130">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook17officelocalclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="c7a76-613">阅读模式</span><span class="sxs-lookup"><span data-stu-id="c7a76-613">Read mode</span></span>

<span data-ttu-id="c7a76-614">`start` 属性返回 `Date` 对象。</span><span class="sxs-lookup"><span data-stu-id="c7a76-614">The `start` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="c7a76-615">撰写模式</span><span class="sxs-lookup"><span data-stu-id="c7a76-615">Compose mode</span></span>

<span data-ttu-id="c7a76-616">`start` 属性返回 `Time` 对象。</span><span class="sxs-lookup"><span data-stu-id="c7a76-616">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="c7a76-617">使用 [`Time.setAsync`](/javascript/api/outlook_1_7/office.time#setasync-datetime--options--callback-) 方法设置开始时间时，应使用 [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) 方法将客户端的本地时间转换为服务器的 UTC。</span><span class="sxs-lookup"><span data-stu-id="c7a76-617">When you use the [`Time.setAsync`](/javascript/api/outlook_1_7/office.time#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="c7a76-618">类型：</span><span class="sxs-lookup"><span data-stu-id="c7a76-618">Type:</span></span>

*   <span data-ttu-id="c7a76-619">日期 | [时间](/javascript/api/outlook_1_7/office.time)</span><span class="sxs-lookup"><span data-stu-id="c7a76-619">Date | [Time](/javascript/api/outlook_1_7/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="c7a76-620">要求</span><span class="sxs-lookup"><span data-stu-id="c7a76-620">Requirements</span></span>

|<span data-ttu-id="c7a76-621">要求</span><span class="sxs-lookup"><span data-stu-id="c7a76-621">Requirement</span></span>|<span data-ttu-id="c7a76-622">值</span><span class="sxs-lookup"><span data-stu-id="c7a76-622">Value</span></span>|
|---|---|
|[<span data-ttu-id="c7a76-623">最低的邮箱要求集版本</span><span class="sxs-lookup"><span data-stu-id="c7a76-623">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c7a76-624">1.0</span><span class="sxs-lookup"><span data-stu-id="c7a76-624">1.0</span></span>|
|[<span data-ttu-id="c7a76-625">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="c7a76-625">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c7a76-626">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c7a76-626">ReadItem</span></span>|
|[<span data-ttu-id="c7a76-627">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c7a76-627">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="c7a76-628">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="c7a76-628">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="c7a76-629">示例</span><span class="sxs-lookup"><span data-stu-id="c7a76-629">Example</span></span>

<span data-ttu-id="c7a76-630">以下示例通过使用 `Time` 对象的 [`setAsync`](/javascript/api/outlook_1_7/office.time#setasync-datetime--options--callback-) 方法，设置撰写模式下约会的开始时间。</span><span class="sxs-lookup"><span data-stu-id="c7a76-630">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_7/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

```
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

####  <a name="subject-stringsubjectjavascriptapioutlook17officesubject"></a><span data-ttu-id="c7a76-631">subject :字符串 |[主题](/javascript/api/outlook_1_7/office.subject)</span><span class="sxs-lookup"><span data-stu-id="c7a76-631">subject :String|[Subject](/javascript/api/outlook_1_7/office.subject)</span></span>

<span data-ttu-id="c7a76-632">获取或设置显示在项目的主题字段中的说明。</span><span class="sxs-lookup"><span data-stu-id="c7a76-632">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="c7a76-633">`subject` 属性获取或设置由电子邮件服务器发送项目时的整个主题。</span><span class="sxs-lookup"><span data-stu-id="c7a76-633">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="c7a76-634">阅读模式</span><span class="sxs-lookup"><span data-stu-id="c7a76-634">Read mode</span></span>

<span data-ttu-id="c7a76-p131">`subject` 属性返回一个字符串。使用 [`normalizedSubject`](#normalizedsubject-string) 属性获取不带任何前导前缀（如 `RE:` 和 `FW:`）的主题。</span><span class="sxs-lookup"><span data-stu-id="c7a76-p131">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

```
var subject = Office.context.mailbox.item.subject;
```

##### <a name="compose-mode"></a><span data-ttu-id="c7a76-637">撰写模式</span><span class="sxs-lookup"><span data-stu-id="c7a76-637">Compose mode</span></span>

<span data-ttu-id="c7a76-638">`subject` 属性返回一个 `Subject` 对象，该对象提供用于获取和设置主题的方法。</span><span class="sxs-lookup"><span data-stu-id="c7a76-638">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="c7a76-639">类型：</span><span class="sxs-lookup"><span data-stu-id="c7a76-639">Type:</span></span>

*   <span data-ttu-id="c7a76-640">字符串 | [主题](/javascript/api/outlook_1_7/office.subject)</span><span class="sxs-lookup"><span data-stu-id="c7a76-640">String | [Subject](/javascript/api/outlook_1_7/office.subject)</span></span>

##### <a name="requirements"></a><span data-ttu-id="c7a76-641">要求</span><span class="sxs-lookup"><span data-stu-id="c7a76-641">Requirements</span></span>

|<span data-ttu-id="c7a76-642">要求</span><span class="sxs-lookup"><span data-stu-id="c7a76-642">Requirement</span></span>|<span data-ttu-id="c7a76-643">值</span><span class="sxs-lookup"><span data-stu-id="c7a76-643">Value</span></span>|
|---|---|
|[<span data-ttu-id="c7a76-644">最低的邮箱要求集版本</span><span class="sxs-lookup"><span data-stu-id="c7a76-644">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c7a76-645">1.0</span><span class="sxs-lookup"><span data-stu-id="c7a76-645">1.0</span></span>|
|[<span data-ttu-id="c7a76-646">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="c7a76-646">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c7a76-647">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c7a76-647">ReadItem</span></span>|
|[<span data-ttu-id="c7a76-648">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c7a76-648">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="c7a76-649">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="c7a76-649">Compose or read</span></span>|

####  <a name="to-arrayemailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsrecipientsjavascriptapioutlook17officerecipients"></a><span data-ttu-id="c7a76-650">to :数组.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)>|[收件人](/javascript/api/outlook_1_7/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="c7a76-650">to :Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_7/office.recipients)</span></span>

<span data-ttu-id="c7a76-651">提供对邮件的 **发送** 行上收件人的访问。</span><span class="sxs-lookup"><span data-stu-id="c7a76-651">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="c7a76-652">对象的类型和访问级别取决于当前项的模式。</span><span class="sxs-lookup"><span data-stu-id="c7a76-652">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="c7a76-653">阅读模式</span><span class="sxs-lookup"><span data-stu-id="c7a76-653">Read mode</span></span>

<span data-ttu-id="c7a76-p133">`to` 属性返回包含邮件的**收件人**行上所列的每个收件人的 `EmailAddressDetails` 对象的数组。集合上限为 100 个成员。</span><span class="sxs-lookup"><span data-stu-id="c7a76-p133">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="c7a76-656">撰写模式</span><span class="sxs-lookup"><span data-stu-id="c7a76-656">Compose mode</span></span>

<span data-ttu-id="c7a76-657">`to` 属性返回 `Recipients` 对象，该对象提供用于处理邮件**收件人**行上收件人的方法。</span><span class="sxs-lookup"><span data-stu-id="c7a76-657">The to`to` property returns a Recipients`Recipients` object that provides methods for manipulating the recipients on the **To** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="c7a76-658">类型：</span><span class="sxs-lookup"><span data-stu-id="c7a76-658">Type:</span></span>

*   <span data-ttu-id="c7a76-659">数组.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)> |[收件人](/javascript/api/outlook_1_7/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="c7a76-659">Array.<[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_7/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="c7a76-660">要求</span><span class="sxs-lookup"><span data-stu-id="c7a76-660">Requirements</span></span>

|<span data-ttu-id="c7a76-661">要求</span><span class="sxs-lookup"><span data-stu-id="c7a76-661">Requirement</span></span>|<span data-ttu-id="c7a76-662">值</span><span class="sxs-lookup"><span data-stu-id="c7a76-662">Value</span></span>|
|---|---|
|[<span data-ttu-id="c7a76-663">最低的邮箱要求集版本</span><span class="sxs-lookup"><span data-stu-id="c7a76-663">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c7a76-664">1.0</span><span class="sxs-lookup"><span data-stu-id="c7a76-664">1.0</span></span>|
|[<span data-ttu-id="c7a76-665">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="c7a76-665">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c7a76-666">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c7a76-666">ReadItem</span></span>|
|[<span data-ttu-id="c7a76-667">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c7a76-667">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="c7a76-668">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="c7a76-668">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="c7a76-669">示例</span><span class="sxs-lookup"><span data-stu-id="c7a76-669">Example</span></span>

```
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

### <a name="methods"></a><span data-ttu-id="c7a76-670">方法</span><span class="sxs-lookup"><span data-stu-id="c7a76-670">Methods</span></span>

####  <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="c7a76-671">addFileAttachmentAsync(uri, attachmentName, [选项], [回调])</span><span class="sxs-lookup"><span data-stu-id="c7a76-671">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="c7a76-672">将文件作为附件添加到邮件或约会。</span><span class="sxs-lookup"><span data-stu-id="c7a76-672">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="c7a76-673">`addFileAttachmentAsync` 方法在指定的 URI 上载文件并将其附加到撰写窗体中的项目。</span><span class="sxs-lookup"><span data-stu-id="c7a76-673">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="c7a76-674">你随后可以将该标识符与 [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) 方法一同使用，以删除同一个会话中的附件。</span><span class="sxs-lookup"><span data-stu-id="c7a76-674">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c7a76-675">参数：</span><span class="sxs-lookup"><span data-stu-id="c7a76-675">Parameters:</span></span>
|<span data-ttu-id="c7a76-676">名称</span><span class="sxs-lookup"><span data-stu-id="c7a76-676">Name</span></span>|<span data-ttu-id="c7a76-677">类型</span><span class="sxs-lookup"><span data-stu-id="c7a76-677">Type</span></span>|<span data-ttu-id="c7a76-678">属性</span><span class="sxs-lookup"><span data-stu-id="c7a76-678">Attributes</span></span>|<span data-ttu-id="c7a76-679">说明</span><span class="sxs-lookup"><span data-stu-id="c7a76-679">Description</span></span>|
|---|---|---|---|
|`uri`|<span data-ttu-id="c7a76-680">字符串</span><span class="sxs-lookup"><span data-stu-id="c7a76-680">String</span></span>||<span data-ttu-id="c7a76-p134">提供附加到邮件或约会的文件的位置的 URI。最大长度为 2048 个字符。</span><span class="sxs-lookup"><span data-stu-id="c7a76-p134">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`|<span data-ttu-id="c7a76-683">字符串</span><span class="sxs-lookup"><span data-stu-id="c7a76-683">String</span></span>||<span data-ttu-id="c7a76-p135">在附件上载过程中显示的附件名称。最大长度为 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="c7a76-p135">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="c7a76-686">对象</span><span class="sxs-lookup"><span data-stu-id="c7a76-686">Object</span></span>|<span data-ttu-id="c7a76-687">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="c7a76-687">&lt;optional&gt;</span></span>|<span data-ttu-id="c7a76-688">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="c7a76-688">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="c7a76-689">对象</span><span class="sxs-lookup"><span data-stu-id="c7a76-689">Object</span></span>|<span data-ttu-id="c7a76-690">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="c7a76-690">&lt;optional&gt;</span></span>|<span data-ttu-id="c7a76-691">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="c7a76-691">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.isInline`|<span data-ttu-id="c7a76-692">布尔值</span><span class="sxs-lookup"><span data-stu-id="c7a76-692">Boolean</span></span>|<span data-ttu-id="c7a76-693">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="c7a76-693">&lt;optional&gt;</span></span>|<span data-ttu-id="c7a76-694">如果为 `true` ，则表示附件将嵌入在邮件正文中显示，而不应在附件列表中显示。</span><span class="sxs-lookup"><span data-stu-id="c7a76-694">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`callback`|<span data-ttu-id="c7a76-695">函数</span><span class="sxs-lookup"><span data-stu-id="c7a76-695">function</span></span>|<span data-ttu-id="c7a76-696">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="c7a76-696">&lt;optional&gt;</span></span>|<span data-ttu-id="c7a76-697">方法完成后，使用单个参数 `asyncResult` （一个 [`AsyncResult`](/javascript/api/office/office.asyncresult)   对象）调用在 `callback`  参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="c7a76-697">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="c7a76-698">如果成功，附件标识符将在 `asyncResult.value` 属性中提供。</span><span class="sxs-lookup"><span data-stu-id="c7a76-698">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="c7a76-699">如果上传附件失败，`asyncResult` 对象将包含一个提供错误说明的 `Error` 对象。</span><span class="sxs-lookup"><span data-stu-id="c7a76-699">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="c7a76-700">错误</span><span class="sxs-lookup"><span data-stu-id="c7a76-700">Errors</span></span>

|<span data-ttu-id="c7a76-701">错误代码</span><span class="sxs-lookup"><span data-stu-id="c7a76-701">Error code</span></span>|<span data-ttu-id="c7a76-702">说明</span><span class="sxs-lookup"><span data-stu-id="c7a76-702">Description</span></span>|
|------------|-------------|
|`AttachmentSizeExceeded`|<span data-ttu-id="c7a76-703">附件大小超过了允许的大小。</span><span class="sxs-lookup"><span data-stu-id="c7a76-703">The attachment is larger than allowed.</span></span>|
|`FileTypeNotSupported`|<span data-ttu-id="c7a76-704">该附件的扩展名不是允许的扩展名。</span><span class="sxs-lookup"><span data-stu-id="c7a76-704">The attachment has an extension that is not allowed.</span></span>|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="c7a76-705">邮件或约会具有的附件过多。</span><span class="sxs-lookup"><span data-stu-id="c7a76-705">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c7a76-706">要求</span><span class="sxs-lookup"><span data-stu-id="c7a76-706">Requirements</span></span>

|<span data-ttu-id="c7a76-707">要求</span><span class="sxs-lookup"><span data-stu-id="c7a76-707">Requirement</span></span>|<span data-ttu-id="c7a76-708">值</span><span class="sxs-lookup"><span data-stu-id="c7a76-708">Value</span></span>|
|---|---|
|[<span data-ttu-id="c7a76-709">最低的邮箱要求集版本</span><span class="sxs-lookup"><span data-stu-id="c7a76-709">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c7a76-710">1.1</span><span class="sxs-lookup"><span data-stu-id="c7a76-710">1.1</span></span>|
|[<span data-ttu-id="c7a76-711">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="c7a76-711">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c7a76-712">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="c7a76-712">ReadWriteItem</span></span>|
|[<span data-ttu-id="c7a76-713">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c7a76-713">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="c7a76-714">撰写</span><span class="sxs-lookup"><span data-stu-id="c7a76-714">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="c7a76-715">示例</span><span class="sxs-lookup"><span data-stu-id="c7a76-715">Examples</span></span>

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

<span data-ttu-id="c7a76-716">以下示例以嵌入附件方式添加图像文件并在邮件正文中引用此附件。</span><span class="sxs-lookup"><span data-stu-id="c7a76-716">The following example adds an image file as an inline attachment and references the attachment in the message body.</span></span>

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

####  <a name="addhandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="c7a76-717">addHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="c7a76-717">addHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="c7a76-718">添加支持事件的事件处理程序。</span><span class="sxs-lookup"><span data-stu-id="c7a76-718">Adds an event handler for a supported event.</span></span>

<span data-ttu-id="c7a76-719">目前，支持的事件类型是 `Office.EventType.AppointmentTimeChanged` 和 `Office.EventType.RecipientsChanged`。 `Office.EventType.RecurrenceChanged`</span><span class="sxs-lookup"><span data-stu-id="c7a76-719">Currently the supported event types are `Office.EventType.AppointmentTimeChanged`, `Office.EventType.RecipientsChanged`, and `Office.EventType.RecurrenceChanged`</span></span>

##### <a name="parameters"></a><span data-ttu-id="c7a76-720">参数：</span><span class="sxs-lookup"><span data-stu-id="c7a76-720">Parameters:</span></span>

| <span data-ttu-id="c7a76-721">名称</span><span class="sxs-lookup"><span data-stu-id="c7a76-721">Name</span></span> | <span data-ttu-id="c7a76-722">类型</span><span class="sxs-lookup"><span data-stu-id="c7a76-722">Type</span></span> | <span data-ttu-id="c7a76-723">属性</span><span class="sxs-lookup"><span data-stu-id="c7a76-723">Attributes</span></span> | <span data-ttu-id="c7a76-724">说明</span><span class="sxs-lookup"><span data-stu-id="c7a76-724">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="c7a76-725">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="c7a76-725">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="c7a76-726">应调用处理程序的事件。</span><span class="sxs-lookup"><span data-stu-id="c7a76-726">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="c7a76-727">函数</span><span class="sxs-lookup"><span data-stu-id="c7a76-727">Function</span></span> || <span data-ttu-id="c7a76-p136">用于处理事件的函数。此函数必须接受单个参数，即对象文本。参数上的 `type` 属性将匹配传递给 `addHandlerAsync` 的 `eventType` 参数。</span><span class="sxs-lookup"><span data-stu-id="c7a76-p136">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `addHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="c7a76-731">对象</span><span class="sxs-lookup"><span data-stu-id="c7a76-731">Object</span></span> | <span data-ttu-id="c7a76-732">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="c7a76-732">&lt;optional&gt;</span></span> | <span data-ttu-id="c7a76-733">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="c7a76-733">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="c7a76-734">对象</span><span class="sxs-lookup"><span data-stu-id="c7a76-734">Object</span></span> | <span data-ttu-id="c7a76-735">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="c7a76-735">&lt;optional&gt;</span></span> | <span data-ttu-id="c7a76-736">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="c7a76-736">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="c7a76-737">函数</span><span class="sxs-lookup"><span data-stu-id="c7a76-737">function</span></span>| <span data-ttu-id="c7a76-738">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="c7a76-738">&lt;optional&gt;</span></span>|<span data-ttu-id="c7a76-739">方法完成后，使用单个参数 `asyncResult` （一个 [`AsyncResult`](/javascript/api/office/office.asyncresult)   对象）调用在 `callback`  参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="c7a76-739">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c7a76-740">要求</span><span class="sxs-lookup"><span data-stu-id="c7a76-740">Requirements</span></span>

|<span data-ttu-id="c7a76-741">要求</span><span class="sxs-lookup"><span data-stu-id="c7a76-741">Requirement</span></span>| <span data-ttu-id="c7a76-742">值</span><span class="sxs-lookup"><span data-stu-id="c7a76-742">Value</span></span>|
|---|---|
|[<span data-ttu-id="c7a76-743">最低的邮箱要求集版本</span><span class="sxs-lookup"><span data-stu-id="c7a76-743">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c7a76-744">1.7</span><span class="sxs-lookup"><span data-stu-id="c7a76-744">-17</span></span> |
|[<span data-ttu-id="c7a76-745">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="c7a76-745">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c7a76-746">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c7a76-746">ReadItem</span></span> |
|[<span data-ttu-id="c7a76-747">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c7a76-747">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c7a76-748">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="c7a76-748">Compose or read</span></span> |

##### <a name="example"></a><span data-ttu-id="c7a76-749">示例</span><span class="sxs-lookup"><span data-stu-id="c7a76-749">Example</span></span>

```
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

####  <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="c7a76-750">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="c7a76-750">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="c7a76-751">将 Exchange 项目（如邮件）作为附件添加到邮件或约会。</span><span class="sxs-lookup"><span data-stu-id="c7a76-751">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="c7a76-p137">`addItemAttachmentAsync` 方法将包含指定 Exchange 标识符的项目附加到撰写窗体中的项目。如果指定一个回调方法，此方法使用 `asyncResult` 参数调用，该参数包含一个附件标识符或代码，指示附加项目过程中出现的任何错误。可以使用 `options` 参数将状态信息传递给回调方法（如果需要）。</span><span class="sxs-lookup"><span data-stu-id="c7a76-p137">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="c7a76-755">你随后可以将该标识符与 [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) 方法一同使用，以删除同一个会话中的附件。</span><span class="sxs-lookup"><span data-stu-id="c7a76-755">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="c7a76-756">如果 Office 外接程序在 Outlook Web App 中运行，则 `addItemAttachmentAsync` 方法可以将项目附加到项目（正在编辑的项目除外）中；但不支持也不建议这样做。</span><span class="sxs-lookup"><span data-stu-id="c7a76-756">If your Office add-in is running in Outlook Web App, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c7a76-757">参数：</span><span class="sxs-lookup"><span data-stu-id="c7a76-757">Parameters:</span></span>

|<span data-ttu-id="c7a76-758">名称</span><span class="sxs-lookup"><span data-stu-id="c7a76-758">Name</span></span>|<span data-ttu-id="c7a76-759">类型</span><span class="sxs-lookup"><span data-stu-id="c7a76-759">Type</span></span>|<span data-ttu-id="c7a76-760">属性</span><span class="sxs-lookup"><span data-stu-id="c7a76-760">Attributes</span></span>|<span data-ttu-id="c7a76-761">说明</span><span class="sxs-lookup"><span data-stu-id="c7a76-761">Description</span></span>|
|---|---|---|---|
|`itemId`|<span data-ttu-id="c7a76-762">字符串</span><span class="sxs-lookup"><span data-stu-id="c7a76-762">String</span></span>||<span data-ttu-id="c7a76-p138">要附加项目的 Exchange 标识符。最大长度为 100 个字符。</span><span class="sxs-lookup"><span data-stu-id="c7a76-p138">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`|<span data-ttu-id="c7a76-765">字符串</span><span class="sxs-lookup"><span data-stu-id="c7a76-765">String</span></span>||<span data-ttu-id="c7a76-p139">要附加项目的主题。最大长度为 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="c7a76-p139">The sujbect of the item to be attached. The maximum length is 255 characters.</span></span>|
|`options`|<span data-ttu-id="c7a76-768">对象</span><span class="sxs-lookup"><span data-stu-id="c7a76-768">Object</span></span>|<span data-ttu-id="c7a76-769">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="c7a76-769">&lt;optional&gt;</span></span>|<span data-ttu-id="c7a76-770">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="c7a76-770">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="c7a76-771">对象</span><span class="sxs-lookup"><span data-stu-id="c7a76-771">Object</span></span>|<span data-ttu-id="c7a76-772">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="c7a76-772">&lt;optional&gt;</span></span>|<span data-ttu-id="c7a76-773">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="c7a76-773">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="c7a76-774">函数</span><span class="sxs-lookup"><span data-stu-id="c7a76-774">function</span></span>|<span data-ttu-id="c7a76-775">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="c7a76-775">&lt;optional&gt;</span></span>|<span data-ttu-id="c7a76-776">方法完成后，使用单个参数 `asyncResult` （一个 [`AsyncResult`](/javascript/api/office/office.asyncresult)   对象）调用在 `callback`  参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="c7a76-776">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="c7a76-777">如果成功，附件标识符将在 `asyncResult.value` 属性中提供。</span><span class="sxs-lookup"><span data-stu-id="c7a76-777">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="c7a76-778">如果添加附件失败，`asyncResult` 对象将包含一个提供错误说明的 `Error` 对象。</span><span class="sxs-lookup"><span data-stu-id="c7a76-778">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="c7a76-779">错误</span><span class="sxs-lookup"><span data-stu-id="c7a76-779">Errors</span></span>

|<span data-ttu-id="c7a76-780">错误代码</span><span class="sxs-lookup"><span data-stu-id="c7a76-780">Error code</span></span>|<span data-ttu-id="c7a76-781">说明</span><span class="sxs-lookup"><span data-stu-id="c7a76-781">Description</span></span>|
|------------|-------------|
|`NumberOfAttachmentsExceeded`|<span data-ttu-id="c7a76-782">邮件或约会具有的附件过多。</span><span class="sxs-lookup"><span data-stu-id="c7a76-782">The message or appointment has too many attachments.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c7a76-783">要求</span><span class="sxs-lookup"><span data-stu-id="c7a76-783">Requirements</span></span>

|<span data-ttu-id="c7a76-784">要求</span><span class="sxs-lookup"><span data-stu-id="c7a76-784">Requirement</span></span>|<span data-ttu-id="c7a76-785">值</span><span class="sxs-lookup"><span data-stu-id="c7a76-785">Value</span></span>|
|---|---|
|[<span data-ttu-id="c7a76-786">最低的邮箱要求集版本</span><span class="sxs-lookup"><span data-stu-id="c7a76-786">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c7a76-787">1.1</span><span class="sxs-lookup"><span data-stu-id="c7a76-787">1.1</span></span>|
|[<span data-ttu-id="c7a76-788">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="c7a76-788">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c7a76-789">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="c7a76-789">ReadWriteItem</span></span>|
|[<span data-ttu-id="c7a76-790">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c7a76-790">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="c7a76-791">撰写</span><span class="sxs-lookup"><span data-stu-id="c7a76-791">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="c7a76-792">示例</span><span class="sxs-lookup"><span data-stu-id="c7a76-792">Example</span></span>

<span data-ttu-id="c7a76-793">以下示例将现有的 Outlook 项目添加为名为 `My Attachment` 的附件。</span><span class="sxs-lookup"><span data-stu-id="c7a76-793">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

```
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

####  <a name="close"></a><span data-ttu-id="c7a76-794">close()</span><span class="sxs-lookup"><span data-stu-id="c7a76-794">close()</span></span>

<span data-ttu-id="c7a76-795">关闭当前正在撰写的项目。</span><span class="sxs-lookup"><span data-stu-id="c7a76-795">Closes the current item that is being composed.</span></span>

<span data-ttu-id="c7a76-p140">`close` 方法的行为取决于要撰写的项目的当前状态。如果项目具有未保存的更改，客户端将提示用户保存、放弃或取消关闭操作。</span><span class="sxs-lookup"><span data-stu-id="c7a76-p140">The behavior of the `close` method depends on the current state of the item being composed. If the item has unsaved changes, the client prompts the user to save, discard, or cancel the close action.</span></span>

> [!NOTE]
> <span data-ttu-id="c7a76-798">在 Outlook 网页版中，如果是约会项，并之前用`saveAsync` 保存过，会提示用户保存、放弃或取消，即使该项上一次保存后并未有任何更改。</span><span class="sxs-lookup"><span data-stu-id="c7a76-798">In Outlook on the web, if the item is an appointment and it has previously been saved using `saveAsync`, the user is prompted to save, discard, or cancel even if no changes have occurred since the item was last saved.</span></span>

<span data-ttu-id="c7a76-799">在 Outlook 桌面客户端中，如果邮件是内联答复，`close` 方法不起作用。</span><span class="sxs-lookup"><span data-stu-id="c7a76-799">In the Outlook desktop client, if the message is an inline reply, the `close` method has no effect.</span></span>

##### <a name="requirements"></a><span data-ttu-id="c7a76-800">要求</span><span class="sxs-lookup"><span data-stu-id="c7a76-800">Requirements</span></span>

|<span data-ttu-id="c7a76-801">要求</span><span class="sxs-lookup"><span data-stu-id="c7a76-801">Requirement</span></span>|<span data-ttu-id="c7a76-802">值</span><span class="sxs-lookup"><span data-stu-id="c7a76-802">Value</span></span>|
|---|---|
|[<span data-ttu-id="c7a76-803">最低的邮箱要求集版本</span><span class="sxs-lookup"><span data-stu-id="c7a76-803">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c7a76-804">1.3</span><span class="sxs-lookup"><span data-stu-id="c7a76-804">1.3</span></span>|
|[<span data-ttu-id="c7a76-805">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="c7a76-805">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c7a76-806">受限</span><span class="sxs-lookup"><span data-stu-id="c7a76-806">Restricted</span></span>|
|[<span data-ttu-id="c7a76-807">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c7a76-807">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="c7a76-808">撰写</span><span class="sxs-lookup"><span data-stu-id="c7a76-808">Compose</span></span>|

#### <a name="displayreplyallformformdata"></a><span data-ttu-id="c7a76-809">displayReplyAllForm(formData)</span><span class="sxs-lookup"><span data-stu-id="c7a76-809">displayReplyAllForm(formData)</span></span>

<span data-ttu-id="c7a76-810">显示答复窗体，其中包括所选邮件的发件人和所有收件人或所选约会的组织者和所有与会者。</span><span class="sxs-lookup"><span data-stu-id="c7a76-810">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="c7a76-811">在 Outlook for iOS 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="c7a76-811">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="c7a76-812">在 Outlook Web App 中，答复窗体显示为包含 3 列视图的弹出式窗体以及包含 2 列或 1 列视图的弹出式窗体。</span><span class="sxs-lookup"><span data-stu-id="c7a76-812">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="c7a76-813">如果任意字符串参数超出其限制， `displayReplyAllForm` 将引发异常。</span><span class="sxs-lookup"><span data-stu-id="c7a76-813">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="c7a76-p141">当在 `formData.attachments` 参数中指定附件时，Outlook 和 Outlook Web App 尝试下载所有附件并将其附加到答复窗体。如果无法添加任何附件，则在窗体 UI 中显示错误。如果这不可能，则不引发错误消息。</span><span class="sxs-lookup"><span data-stu-id="c7a76-p141">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c7a76-817">参数：</span><span class="sxs-lookup"><span data-stu-id="c7a76-817">Parameters:</span></span>

|<span data-ttu-id="c7a76-818">名称</span><span class="sxs-lookup"><span data-stu-id="c7a76-818">Name</span></span>|<span data-ttu-id="c7a76-819">类型</span><span class="sxs-lookup"><span data-stu-id="c7a76-819">Type</span></span>|<span data-ttu-id="c7a76-820">属性</span><span class="sxs-lookup"><span data-stu-id="c7a76-820">Attributes</span></span>|<span data-ttu-id="c7a76-821">说明</span><span class="sxs-lookup"><span data-stu-id="c7a76-821">Description</span></span>|
|---|---|---|---|
|`formData`|<span data-ttu-id="c7a76-822">字符串 | 对象</span><span class="sxs-lookup"><span data-stu-id="c7a76-822">String &#124; Object</span></span>||<span data-ttu-id="c7a76-p142">一个包含文本和 HTML 且表示答复窗体正文的字符串。字符串限制为 32 KB。</span><span class="sxs-lookup"><span data-stu-id="c7a76-p142">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="c7a76-825">**或**</span><span class="sxs-lookup"><span data-stu-id="c7a76-825">**OR**</span></span><br/><span data-ttu-id="c7a76-p143">包含正文或附件数据和回调函数的对象。对象定义如下。</span><span class="sxs-lookup"><span data-stu-id="c7a76-p143">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span>|
|`formData.htmlBody`|<span data-ttu-id="c7a76-828">字符串</span><span class="sxs-lookup"><span data-stu-id="c7a76-828">String</span></span>|<span data-ttu-id="c7a76-829">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="c7a76-829">&lt;optional&gt;</span></span>|<span data-ttu-id="c7a76-p144">一个包含文本和 HTML 且表示答复窗体正文的字符串。字符串限制为 32 KB。</span><span class="sxs-lookup"><span data-stu-id="c7a76-p144">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
|`formData.attachments`|<span data-ttu-id="c7a76-832">Array.&lt;对象&gt;</span><span class="sxs-lookup"><span data-stu-id="c7a76-832">Array.&lt;Object&gt;</span></span>|<span data-ttu-id="c7a76-833">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="c7a76-833">&lt;optional&gt;</span></span>|<span data-ttu-id="c7a76-834">JSON 对象（是文件或项目附件）的数组。</span><span class="sxs-lookup"><span data-stu-id="c7a76-834">An array of JSON objects that are either file or item attachments.</span></span>|
|`formData.attachments.type`|<span data-ttu-id="c7a76-835">字符串</span><span class="sxs-lookup"><span data-stu-id="c7a76-835">String</span></span>||<span data-ttu-id="c7a76-p145">指示附件的类型。必须是文件附件的 `file` 或项目附件的 `item` 。</span><span class="sxs-lookup"><span data-stu-id="c7a76-p145">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span>|
|`formData.attachments.name`|<span data-ttu-id="c7a76-838">字符串</span><span class="sxs-lookup"><span data-stu-id="c7a76-838">String</span></span>||<span data-ttu-id="c7a76-839">一个包含附件的名称的字符串，最多包含 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="c7a76-839">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
|`formData.attachments.url`|<span data-ttu-id="c7a76-840">字符串</span><span class="sxs-lookup"><span data-stu-id="c7a76-840">String</span></span>||<span data-ttu-id="c7a76-p146">仅在将 `type` 设置为 `file` 时使用。文件位置的 URI。</span><span class="sxs-lookup"><span data-stu-id="c7a76-p146">Only used if `type` is set to `file`. The URI of the location for the file.</span></span>|
|`formData.attachments.isInline`|<span data-ttu-id="c7a76-843">布尔值</span><span class="sxs-lookup"><span data-stu-id="c7a76-843">Boolean</span></span>||<span data-ttu-id="c7a76-p147">仅在将 `type` 设置为 `file` 时使用。如果为 `true`，则表示附件将在邮件正文中内联显示，并且不应显示在附件列表中。</span><span class="sxs-lookup"><span data-stu-id="c7a76-p147">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`formData.attachments.itemId`|<span data-ttu-id="c7a76-846">字符串</span><span class="sxs-lookup"><span data-stu-id="c7a76-846">String</span></span>||<span data-ttu-id="c7a76-p148">仅在将 `type` 设置为 `item` 时使用。附件的 EWS 项目 ID。字符串最长为 100 个字符。</span><span class="sxs-lookup"><span data-stu-id="c7a76-p148">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span>|
|`callback`|<span data-ttu-id="c7a76-850">函数</span><span class="sxs-lookup"><span data-stu-id="c7a76-850">function</span></span>|<span data-ttu-id="c7a76-851">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="c7a76-851">&lt;optional&gt;</span></span>|<span data-ttu-id="c7a76-852">方法完成后，使用单个参数 `asyncResult`（一个 [AsyncResult](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="c7a76-852">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c7a76-853">要求</span><span class="sxs-lookup"><span data-stu-id="c7a76-853">Requirements</span></span>

|<span data-ttu-id="c7a76-854">要求</span><span class="sxs-lookup"><span data-stu-id="c7a76-854">Requirement</span></span>|<span data-ttu-id="c7a76-855">值</span><span class="sxs-lookup"><span data-stu-id="c7a76-855">Value</span></span>|
|---|---|
|[<span data-ttu-id="c7a76-856">最低的邮箱要求集版本</span><span class="sxs-lookup"><span data-stu-id="c7a76-856">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c7a76-857">1.0</span><span class="sxs-lookup"><span data-stu-id="c7a76-857">1.0</span></span>|
|[<span data-ttu-id="c7a76-858">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="c7a76-858">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c7a76-859">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c7a76-859">ReadItem</span></span>|
|[<span data-ttu-id="c7a76-860">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c7a76-860">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="c7a76-861">阅读</span><span class="sxs-lookup"><span data-stu-id="c7a76-861">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="c7a76-862">示例</span><span class="sxs-lookup"><span data-stu-id="c7a76-862">Examples</span></span>

<span data-ttu-id="c7a76-863">以下代码将一个字符串传递到 `displayReplyAllForm` 函数。</span><span class="sxs-lookup"><span data-stu-id="c7a76-863">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="c7a76-864">使用空白正文答复。</span><span class="sxs-lookup"><span data-stu-id="c7a76-864">Reply with an empty body.</span></span>

```
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="c7a76-865">仅使用正文答复。</span><span class="sxs-lookup"><span data-stu-id="c7a76-865">Reply with just a body.</span></span>

```
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="c7a76-866">使用正文和文件附件答复。</span><span class="sxs-lookup"><span data-stu-id="c7a76-866">Reply with a body and a file attachment.</span></span>

```
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

<span data-ttu-id="c7a76-867">使用正文和项目附件答复。</span><span class="sxs-lookup"><span data-stu-id="c7a76-867">Reply with a body and an item attachment.</span></span>

```
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

<span data-ttu-id="c7a76-868">使用正文、文件附件、项目附件和回调答复。</span><span class="sxs-lookup"><span data-stu-id="c7a76-868">Reply with a body, file attachment, item attachment, and a callback.</span></span>

```
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

#### <a name="displayreplyformformdata"></a><span data-ttu-id="c7a76-869">displayReplyForm(formData)</span><span class="sxs-lookup"><span data-stu-id="c7a76-869">displayReplyForm(formData)</span></span>

<span data-ttu-id="c7a76-870">显示答复窗体，其中仅包括所选邮件的发件人或所选约会的组织者。</span><span class="sxs-lookup"><span data-stu-id="c7a76-870">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="c7a76-871">在 Outlook for iOS 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="c7a76-871">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="c7a76-872">在 Outlook Web App 中，答复窗体显示为包含 3 列视图的弹出式窗体以及包含 2 列或 1 列视图的弹出式窗体。</span><span class="sxs-lookup"><span data-stu-id="c7a76-872">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="c7a76-873">如果任意字符串参数超出其限制， `displayReplyForm` 将引发异常。</span><span class="sxs-lookup"><span data-stu-id="c7a76-873">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="c7a76-p149">当在 `formData.attachments` 参数中指定附件时，Outlook 和 Outlook Web App 尝试下载所有附件并将其附加到答复窗体。如果无法添加任何附件，则在窗体 UI 中显示错误。如果这不可能，则不引发错误消息。</span><span class="sxs-lookup"><span data-stu-id="c7a76-p149">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c7a76-877">参数：</span><span class="sxs-lookup"><span data-stu-id="c7a76-877">Parameters:</span></span>

|<span data-ttu-id="c7a76-878">名称</span><span class="sxs-lookup"><span data-stu-id="c7a76-878">Name</span></span>|<span data-ttu-id="c7a76-879">类型</span><span class="sxs-lookup"><span data-stu-id="c7a76-879">Type</span></span>|<span data-ttu-id="c7a76-880">属性</span><span class="sxs-lookup"><span data-stu-id="c7a76-880">Attributes</span></span>|<span data-ttu-id="c7a76-881">说明</span><span class="sxs-lookup"><span data-stu-id="c7a76-881">Description</span></span>|
|---|---|---|---|
|`formData`|<span data-ttu-id="c7a76-882">字符串 | 对象</span><span class="sxs-lookup"><span data-stu-id="c7a76-882">String &#124; Object</span></span>||<span data-ttu-id="c7a76-p150">一个包含文本和 HTML 且表示答复窗体正文的字符串。字符串限制为 32 KB。</span><span class="sxs-lookup"><span data-stu-id="c7a76-p150">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="c7a76-885">**或**</span><span class="sxs-lookup"><span data-stu-id="c7a76-885">**OR**</span></span><br/><span data-ttu-id="c7a76-p151">包含正文或附件数据和回调函数的对象。对象定义如下。</span><span class="sxs-lookup"><span data-stu-id="c7a76-p151">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span>|
|`formData.htmlBody`|<span data-ttu-id="c7a76-888">字符串</span><span class="sxs-lookup"><span data-stu-id="c7a76-888">String</span></span>|<span data-ttu-id="c7a76-889">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="c7a76-889">&lt;optional&gt;</span></span>|<span data-ttu-id="c7a76-p152">一个包含文本和 HTML 且表示答复窗体正文的字符串。字符串限制为 32 KB。</span><span class="sxs-lookup"><span data-stu-id="c7a76-p152">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
|`formData.attachments`|<span data-ttu-id="c7a76-892">Array.&lt;对象&gt;</span><span class="sxs-lookup"><span data-stu-id="c7a76-892">Array.&lt;Object&gt;</span></span>|<span data-ttu-id="c7a76-893">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="c7a76-893">&lt;optional&gt;</span></span>|<span data-ttu-id="c7a76-894">JSON 对象（是文件或项目附件）的数组。</span><span class="sxs-lookup"><span data-stu-id="c7a76-894">An array of JSON objects that are either file or item attachments.</span></span>|
|`formData.attachments.type`|<span data-ttu-id="c7a76-895">字符串</span><span class="sxs-lookup"><span data-stu-id="c7a76-895">String</span></span>||<span data-ttu-id="c7a76-p153">指示附件的类型。必须是文件附件的 `file` 或项目附件的 `item` 。</span><span class="sxs-lookup"><span data-stu-id="c7a76-p153">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span>|
|`formData.attachments.name`|<span data-ttu-id="c7a76-898">字符串</span><span class="sxs-lookup"><span data-stu-id="c7a76-898">String</span></span>||<span data-ttu-id="c7a76-899">一个包含附件的名称的字符串，最多包含 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="c7a76-899">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
|`formData.attachments.url`|<span data-ttu-id="c7a76-900">字符串</span><span class="sxs-lookup"><span data-stu-id="c7a76-900">String</span></span>||<span data-ttu-id="c7a76-p154">仅在将 `type` 设置为 `file` 时使用。文件位置的 URI。</span><span class="sxs-lookup"><span data-stu-id="c7a76-p154">Only used if `type` is set to `file`. The URI of the location for the file.</span></span>|
|`formData.attachments.isInline`|<span data-ttu-id="c7a76-903">布尔值</span><span class="sxs-lookup"><span data-stu-id="c7a76-903">Boolean</span></span>||<span data-ttu-id="c7a76-p155">仅在将 `type` 设置为 `file` 时使用。如果为 `true`，则表示附件将在邮件正文中内联显示，并且不应显示在附件列表中。</span><span class="sxs-lookup"><span data-stu-id="c7a76-p155">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span>|
|`formData.attachments.itemId`|<span data-ttu-id="c7a76-906">字符串</span><span class="sxs-lookup"><span data-stu-id="c7a76-906">String</span></span>||<span data-ttu-id="c7a76-p156">仅在将 `type` 设置为 `item` 时使用。附件的 EWS 项目 ID。字符串最长为 100 个字符。</span><span class="sxs-lookup"><span data-stu-id="c7a76-p156">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span>|
|`callback`|<span data-ttu-id="c7a76-910">函数</span><span class="sxs-lookup"><span data-stu-id="c7a76-910">function</span></span>|<span data-ttu-id="c7a76-911">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="c7a76-911">&lt;optional&gt;</span></span>|<span data-ttu-id="c7a76-912">方法完成后，使用单个参数 `asyncResult`（一个 [AsyncResult](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="c7a76-912">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c7a76-913">要求</span><span class="sxs-lookup"><span data-stu-id="c7a76-913">Requirements</span></span>

|<span data-ttu-id="c7a76-914">要求</span><span class="sxs-lookup"><span data-stu-id="c7a76-914">Requirement</span></span>|<span data-ttu-id="c7a76-915">值</span><span class="sxs-lookup"><span data-stu-id="c7a76-915">Value</span></span>|
|---|---|
|[<span data-ttu-id="c7a76-916">最低的邮箱要求集版本</span><span class="sxs-lookup"><span data-stu-id="c7a76-916">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c7a76-917">1.0</span><span class="sxs-lookup"><span data-stu-id="c7a76-917">1.0</span></span>|
|[<span data-ttu-id="c7a76-918">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="c7a76-918">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c7a76-919">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c7a76-919">ReadItem</span></span>|
|[<span data-ttu-id="c7a76-920">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c7a76-920">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="c7a76-921">阅读</span><span class="sxs-lookup"><span data-stu-id="c7a76-921">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="c7a76-922">示例</span><span class="sxs-lookup"><span data-stu-id="c7a76-922">Examples</span></span>

<span data-ttu-id="c7a76-923">以下代码将一个字符串传递到 `displayReplyForm` 函数。</span><span class="sxs-lookup"><span data-stu-id="c7a76-923">The following code passes a string to the `displayReplyForm` function.</span></span>

```
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="c7a76-924">使用空白正文答复。</span><span class="sxs-lookup"><span data-stu-id="c7a76-924">Reply with an empty body.</span></span>

```
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="c7a76-925">仅使用正文答复。</span><span class="sxs-lookup"><span data-stu-id="c7a76-925">Reply with just a body.</span></span>

```
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="c7a76-926">使用正文和文件附件答复。</span><span class="sxs-lookup"><span data-stu-id="c7a76-926">Reply with a body and a file attachment.</span></span>

```
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

<span data-ttu-id="c7a76-927">使用正文和项目附件答复。</span><span class="sxs-lookup"><span data-stu-id="c7a76-927">Reply with a body and an item attachment.</span></span>

```
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

<span data-ttu-id="c7a76-928">使用正文、文件附件、项目附件和回调答复。</span><span class="sxs-lookup"><span data-stu-id="c7a76-928">Reply with a body, file attachment, item attachment, and a callback.</span></span>

```
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

#### <a name="getentities--entitiesjavascriptapioutlook17officeentities"></a><span data-ttu-id="c7a76-929">getEntities() → {[实体](/javascript/api/outlook_1_7/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="c7a76-929">getEntities() → {[Entities](/javascript/api/outlook_1_7/office.entities)}</span></span>

<span data-ttu-id="c7a76-930">获取在所选项正文中找到的实体。</span><span class="sxs-lookup"><span data-stu-id="c7a76-930">Gets the entities found in the selected item.</span></span>

> [!NOTE]
> <span data-ttu-id="c7a76-931">在 Outlook for iOS 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="c7a76-931">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="c7a76-932">要求</span><span class="sxs-lookup"><span data-stu-id="c7a76-932">Requirements</span></span>

|<span data-ttu-id="c7a76-933">要求</span><span class="sxs-lookup"><span data-stu-id="c7a76-933">Requirement</span></span>|<span data-ttu-id="c7a76-934">值</span><span class="sxs-lookup"><span data-stu-id="c7a76-934">Value</span></span>|
|---|---|
|[<span data-ttu-id="c7a76-935">最低的邮箱要求集版本</span><span class="sxs-lookup"><span data-stu-id="c7a76-935">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c7a76-936">1.0</span><span class="sxs-lookup"><span data-stu-id="c7a76-936">1.0</span></span>|
|[<span data-ttu-id="c7a76-937">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="c7a76-937">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c7a76-938">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c7a76-938">ReadItem</span></span>|
|[<span data-ttu-id="c7a76-939">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c7a76-939">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="c7a76-940">阅读</span><span class="sxs-lookup"><span data-stu-id="c7a76-940">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="c7a76-941">返回：</span><span class="sxs-lookup"><span data-stu-id="c7a76-941">Returns:</span></span>

<span data-ttu-id="c7a76-942">类型：[实体](/javascript/api/outlook_1_7/office.entities)</span><span class="sxs-lookup"><span data-stu-id="c7a76-942">Type: [Entities](/javascript/api/outlook_1_7/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="c7a76-943">示例</span><span class="sxs-lookup"><span data-stu-id="c7a76-943">Example</span></span>

<span data-ttu-id="c7a76-944">以下示例访问当前项正文中的联系人实体。</span><span class="sxs-lookup"><span data-stu-id="c7a76-944">The following example accesses the contacts entities on the current item.</span></span>

```
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlook17officecontactmeetingsuggestionjavascriptapioutlook17officemeetingsuggestionphonenumberjavascriptapioutlook17officephonenumbertasksuggestionjavascriptapioutlook17officetasksuggestion"></a><span data-ttu-id="c7a76-945">getEntitiesByType(entityType) → (可为空)  {数组.<(字符串|[联系人](/javascript/api/outlook_1_7/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_7/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_7/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_7/office.tasksuggestion)) >}</span><span class="sxs-lookup"><span data-stu-id="c7a76-945">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_7/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_7/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_7/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_7/office.tasksuggestion))>}</span></span>

<span data-ttu-id="c7a76-946">获取所选项目中找到的指定实体类型的所有实体的数组。</span><span class="sxs-lookup"><span data-stu-id="c7a76-946">Gets an array of all the entities of the specified entity type found in the selected item.</span></span>

> [!NOTE]
> <span data-ttu-id="c7a76-947">在 Outlook for iOS 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="c7a76-947">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c7a76-948">参数：</span><span class="sxs-lookup"><span data-stu-id="c7a76-948">Parameters:</span></span>

|<span data-ttu-id="c7a76-949">名称</span><span class="sxs-lookup"><span data-stu-id="c7a76-949">Name</span></span>|<span data-ttu-id="c7a76-950">类型</span><span class="sxs-lookup"><span data-stu-id="c7a76-950">Type</span></span>|<span data-ttu-id="c7a76-951">说明</span><span class="sxs-lookup"><span data-stu-id="c7a76-951">Description</span></span>|
|---|---|---|
|`entityType`|[<span data-ttu-id="c7a76-952">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="c7a76-952">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook_1_7/office.mailboxenums.entitytype)|<span data-ttu-id="c7a76-953">EntityType 枚举值之一。</span><span class="sxs-lookup"><span data-stu-id="c7a76-953">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c7a76-954">要求</span><span class="sxs-lookup"><span data-stu-id="c7a76-954">Requirements</span></span>

|<span data-ttu-id="c7a76-955">要求</span><span class="sxs-lookup"><span data-stu-id="c7a76-955">Requirement</span></span>|<span data-ttu-id="c7a76-956">值</span><span class="sxs-lookup"><span data-stu-id="c7a76-956">Value</span></span>|
|---|---|
|[<span data-ttu-id="c7a76-957">最低的邮箱要求集版本</span><span class="sxs-lookup"><span data-stu-id="c7a76-957">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c7a76-958">1.0</span><span class="sxs-lookup"><span data-stu-id="c7a76-958">1.0</span></span>|
|[<span data-ttu-id="c7a76-959">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="c7a76-959">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c7a76-960">受限</span><span class="sxs-lookup"><span data-stu-id="c7a76-960">Restricted</span></span>|
|[<span data-ttu-id="c7a76-961">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c7a76-961">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="c7a76-962">阅读</span><span class="sxs-lookup"><span data-stu-id="c7a76-962">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="c7a76-963">返回：</span><span class="sxs-lookup"><span data-stu-id="c7a76-963">Returns:</span></span>

<span data-ttu-id="c7a76-964">如果在 `entityType` 中传递的值不是 `EntityType` 枚举的有效成员，该方法返回 null。</span><span class="sxs-lookup"><span data-stu-id="c7a76-964">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="c7a76-965">如果指定类型的任何实体都不存在于该项目上，该方法将返回空数组。</span><span class="sxs-lookup"><span data-stu-id="c7a76-965">If no entities of the specified type are present on the item, the method returns an empty array.</span></span> <span data-ttu-id="c7a76-966">否则，返回数组中的对象类型取决于 `entityType` 参数中请求实体的类型。</span><span class="sxs-lookup"><span data-stu-id="c7a76-966">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="c7a76-967">当使用此方法的最低权限级别**受限**时，一些实体类型需要**ReadItem**才能进行访问，如下表中所指定。</span><span class="sxs-lookup"><span data-stu-id="c7a76-967">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

|<span data-ttu-id="c7a76-968">的值 `entityType`</span><span class="sxs-lookup"><span data-stu-id="c7a76-968">Value of `entityType`</span></span>|<span data-ttu-id="c7a76-969">返回数组中对象的类型</span><span class="sxs-lookup"><span data-stu-id="c7a76-969">Type of objects in returned array</span></span>|<span data-ttu-id="c7a76-970">所需权限级别</span><span class="sxs-lookup"><span data-stu-id="c7a76-970">Required Permission Level</span></span>|
|---|---|---|
|`Address`|<span data-ttu-id="c7a76-971">字符串</span><span class="sxs-lookup"><span data-stu-id="c7a76-971">String</span></span>|<span data-ttu-id="c7a76-972">**受限**</span><span class="sxs-lookup"><span data-stu-id="c7a76-972">**Restricted**</span></span>|
|`Contact`|<span data-ttu-id="c7a76-973">联系人</span><span class="sxs-lookup"><span data-stu-id="c7a76-973">Contact</span></span>|<span data-ttu-id="c7a76-974">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="c7a76-974">**ReadItem**</span></span>|
|`EmailAddress`|<span data-ttu-id="c7a76-975">字符串</span><span class="sxs-lookup"><span data-stu-id="c7a76-975">String</span></span>|<span data-ttu-id="c7a76-976">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="c7a76-976">**ReadItem**</span></span>|
|`MeetingSuggestion`|<span data-ttu-id="c7a76-977">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="c7a76-977">MeetingSuggestion</span></span>|<span data-ttu-id="c7a76-978">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="c7a76-978">**ReadItem**</span></span>|
|`PhoneNumber`|<span data-ttu-id="c7a76-979">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="c7a76-979">PhoneNumber</span></span>|<span data-ttu-id="c7a76-980">**受限**</span><span class="sxs-lookup"><span data-stu-id="c7a76-980">**Restricted**</span></span>|
|`TaskSuggestion`|<span data-ttu-id="c7a76-981">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="c7a76-981">TaskSuggestion</span></span>|<span data-ttu-id="c7a76-982">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="c7a76-982">**ReadItem**</span></span>|
|`URL`|<span data-ttu-id="c7a76-983">字符串</span><span class="sxs-lookup"><span data-stu-id="c7a76-983">String</span></span>|<span data-ttu-id="c7a76-984">**受限**</span><span class="sxs-lookup"><span data-stu-id="c7a76-984">**Restricted**</span></span>|

<span data-ttu-id="c7a76-985">类型：数组.<(字符串|[联系人](/javascript/api/outlook_1_7/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_7/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_7/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_7/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="c7a76-985">Type: Array.<(String|[Contact](/javascript/api/outlook_1_7/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_7/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_7/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_7/office.tasksuggestion))></span></span>

##### <a name="example"></a><span data-ttu-id="c7a76-986">示例</span><span class="sxs-lookup"><span data-stu-id="c7a76-986">Example</span></span>

<span data-ttu-id="c7a76-987">以下示例显示了如何访问代表当前项正文中邮政地址的字符串数组。</span><span class="sxs-lookup"><span data-stu-id="c7a76-987">The following example shows how to access an array of strings that represent postal addresses in the subject or body of the current item.</span></span>

```
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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlook17officecontactmeetingsuggestionjavascriptapioutlook17officemeetingsuggestionphonenumberjavascriptapioutlook17officephonenumbertasksuggestionjavascriptapioutlook17officetasksuggestion"></a><span data-ttu-id="c7a76-988">getFilteredEntitiesByName(name) → (可为空) {数组.<(字符串|[联系人](/javascript/api/outlook_1_7/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_7/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_7/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_7/office.tasksuggestion)) >}</span><span class="sxs-lookup"><span data-stu-id="c7a76-988">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_7/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_7/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_7/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_7/office.tasksuggestion))>}</span></span>

<span data-ttu-id="c7a76-989">返回清单 XML 文件所定义的命名筛选器所选项中的已知实体。</span><span class="sxs-lookup"><span data-stu-id="c7a76-989">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="c7a76-990">在 Outlook for iOS 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="c7a76-990">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="c7a76-991">`getFilteredEntitiesByName`方法返回与[ItemHasKnownEntity](/javascript/office/manifest/rule#itemhasknownentity-rule) 规则元素中定义的正则表达式相匹配的实体 ，该规则元素包含于具备特定`FilterName`元素值的清单 XML 文件中。</span><span class="sxs-lookup"><span data-stu-id="c7a76-991">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/javascript/office/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c7a76-992">参数：</span><span class="sxs-lookup"><span data-stu-id="c7a76-992">Parameters:</span></span>

|<span data-ttu-id="c7a76-993">名称</span><span class="sxs-lookup"><span data-stu-id="c7a76-993">Name</span></span>|<span data-ttu-id="c7a76-994">类型</span><span class="sxs-lookup"><span data-stu-id="c7a76-994">Type</span></span>|<span data-ttu-id="c7a76-995">说明</span><span class="sxs-lookup"><span data-stu-id="c7a76-995">Description</span></span>|
|---|---|---|
|`name`|<span data-ttu-id="c7a76-996">字符串</span><span class="sxs-lookup"><span data-stu-id="c7a76-996">String</span></span>|<span data-ttu-id="c7a76-997">定义筛选器匹配的 `ItemHasKnownEntity` 规则元素的名称。</span><span class="sxs-lookup"><span data-stu-id="c7a76-997">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c7a76-998">要求</span><span class="sxs-lookup"><span data-stu-id="c7a76-998">Requirements</span></span>

|<span data-ttu-id="c7a76-999">要求</span><span class="sxs-lookup"><span data-stu-id="c7a76-999">Requirement</span></span>|<span data-ttu-id="c7a76-1000">值</span><span class="sxs-lookup"><span data-stu-id="c7a76-1000">Value</span></span>|
|---|---|
|[<span data-ttu-id="c7a76-1001">最低的邮箱要求集版本</span><span class="sxs-lookup"><span data-stu-id="c7a76-1001">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c7a76-1002">1.0</span><span class="sxs-lookup"><span data-stu-id="c7a76-1002">1.0</span></span>|
|[<span data-ttu-id="c7a76-1003">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="c7a76-1003">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c7a76-1004">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c7a76-1004">ReadItem</span></span>|
|[<span data-ttu-id="c7a76-1005">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c7a76-1005">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="c7a76-1006">阅读</span><span class="sxs-lookup"><span data-stu-id="c7a76-1006">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="c7a76-1007">返回：</span><span class="sxs-lookup"><span data-stu-id="c7a76-1007">Returns:</span></span>

<span data-ttu-id="c7a76-p158">如果清单中 `ItemHasKnownEntity`  元素没有匹配 `name` 参数的 `FilterName`  元素值，则该方法返回 `null` 。如果 `name` 参数匹配清单中的 `ItemHasKnownEntity` 元素，但在当前匹配的项目中没有实体，则该方法返回一个空数组。</span><span class="sxs-lookup"><span data-stu-id="c7a76-p158">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="c7a76-1010">类型：数组.<(字符串|[联系人](/javascript/api/outlook_1_7/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_7/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_7/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_7/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="c7a76-1010">Type: Array.<(String|[Contact](/javascript/api/outlook_1_7/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_7/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_7/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_7/office.tasksuggestion))></span></span>

#### <a name="getregexmatches--object"></a><span data-ttu-id="c7a76-1011">getRegExMatches() → {对象}</span><span class="sxs-lookup"><span data-stu-id="c7a76-1011">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="c7a76-1012">返回匹配清单 XML 文件定义的正则表达式所选项目的字符串值。</span><span class="sxs-lookup"><span data-stu-id="c7a76-1012">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="c7a76-1013">在 Outlook for iOS 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="c7a76-1013">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="c7a76-p159">`getRegExMatches` 方法返回与每个 `ItemHasRegularExpressionMatch` 所定义的正则表达式或 `ItemHasKnownEntity` 清单 XML 文件中的规则元素相匹配的字符串。对于 `ItemHasRegularExpressionMatch` 规则，匹配字符串必须发生在该规则指定的项目属性中。`PropertyName` 简单类型定义所支持的属性。</span><span class="sxs-lookup"><span data-stu-id="c7a76-p159">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="c7a76-1017">例如，考虑一个加载项具有以下 `Rule` 元素：</span><span class="sxs-lookup"><span data-stu-id="c7a76-1017">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="c7a76-1018">从 `getRegExMatches` 返回的对象应有两个属性：`fruits` 和 `veggies`。</span><span class="sxs-lookup"><span data-stu-id="c7a76-1018">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="c7a76-p160">如果在项目的正文属性上指定 `ItemHasRegularExpressionMatch` 规则，则正则表达式应进一步筛选正文，不应尝试返回该项目的整个正文。使用正则表达式（如 `.*`）获取项目的整个正文并不总是返回预期的结果。而该使用 [`Body.getAsync`](/javascript/api/outlook_1_7/office.body#getasync-coerciontype--options--callback-) 方法检索整个正文。</span><span class="sxs-lookup"><span data-stu-id="c7a76-p160">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook_1_7/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="c7a76-1022">要求</span><span class="sxs-lookup"><span data-stu-id="c7a76-1022">Requirements</span></span>

|<span data-ttu-id="c7a76-1023">要求</span><span class="sxs-lookup"><span data-stu-id="c7a76-1023">Requirement</span></span>|<span data-ttu-id="c7a76-1024">值</span><span class="sxs-lookup"><span data-stu-id="c7a76-1024">Value</span></span>|
|---|---|
|[<span data-ttu-id="c7a76-1025">最低的邮箱要求集版本</span><span class="sxs-lookup"><span data-stu-id="c7a76-1025">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c7a76-1026">1.0</span><span class="sxs-lookup"><span data-stu-id="c7a76-1026">1.0</span></span>|
|[<span data-ttu-id="c7a76-1027">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="c7a76-1027">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c7a76-1028">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c7a76-1028">ReadItem</span></span>|
|[<span data-ttu-id="c7a76-1029">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c7a76-1029">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="c7a76-1030">阅读</span><span class="sxs-lookup"><span data-stu-id="c7a76-1030">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="c7a76-1031">返回：</span><span class="sxs-lookup"><span data-stu-id="c7a76-1031">Returns:</span></span>

<span data-ttu-id="c7a76-p161">一个包含与清单 XML 文件中所定义正则表达式的字符串数组相匹配的对象。每个数组的名称等于匹配 `ItemHasRegularExpressionMatch` 规则的 `RegExName` 属性的相应值或匹配 `ItemHasKnownEntity` 规则的 `FilterName` 属性。</span><span class="sxs-lookup"><span data-stu-id="c7a76-p161">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<dl class="param-type"><span data-ttu-id="c7a76-1034">

<dt>类型</dt>

</span><span class="sxs-lookup"><span data-stu-id="c7a76-1034">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="c7a76-1035">对象</span><span class="sxs-lookup"><span data-stu-id="c7a76-1035">Object</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="c7a76-1036">示例</span><span class="sxs-lookup"><span data-stu-id="c7a76-1036">Example</span></span>

<span data-ttu-id="c7a76-1037">以下示例显示了如何访问正则表达式规则元素 `fruits` 和 `veggies` 的匹配项的数组，这些元素在清单中指定。</span><span class="sxs-lookup"><span data-stu-id="c7a76-1037">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veges = allMatches.veggies;
```

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="c7a76-1038">getRegExMatchesByName(name) → (可为空) {数组.< 字符串 >}</span><span class="sxs-lookup"><span data-stu-id="c7a76-1038">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="c7a76-1039">返回匹配清单 XML 文件定义的命名正则表达式所选项目的字符串值。</span><span class="sxs-lookup"><span data-stu-id="c7a76-1039">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="c7a76-1040">在 Outlook for iOS 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="c7a76-1040">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="c7a76-1041">`getRegExMatchesByName` 方法返回与 XML 文件中的 `ItemHasRegularExpressionMatch` 规则元素中定义的正则表达式相匹配的字符串，该文件具有特定 `RegExName` 元素值。</span><span class="sxs-lookup"><span data-stu-id="c7a76-1041">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="c7a76-p162">如果在项目正文属性上指定 `ItemHasRegularExpressionMatch` 规则，则正则表达式应进一步筛选正文，不应尝试返回该项目的整个正文。使用正则表达式（如 `.*`）获取项目的整个正文并不总是返回预期的结果。</span><span class="sxs-lookup"><span data-stu-id="c7a76-p162">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c7a76-1044">参数：</span><span class="sxs-lookup"><span data-stu-id="c7a76-1044">Parameters:</span></span>

|<span data-ttu-id="c7a76-1045">名称</span><span class="sxs-lookup"><span data-stu-id="c7a76-1045">Name</span></span>|<span data-ttu-id="c7a76-1046">类型</span><span class="sxs-lookup"><span data-stu-id="c7a76-1046">Type</span></span>|<span data-ttu-id="c7a76-1047">说明</span><span class="sxs-lookup"><span data-stu-id="c7a76-1047">Description</span></span>|
|---|---|---|
|`name`|<span data-ttu-id="c7a76-1048">字符串</span><span class="sxs-lookup"><span data-stu-id="c7a76-1048">String</span></span>|<span data-ttu-id="c7a76-1049">定义筛选器匹配的 `ItemHasRegularExpressionMatch` 规则元素的名称。</span><span class="sxs-lookup"><span data-stu-id="c7a76-1049">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c7a76-1050">要求</span><span class="sxs-lookup"><span data-stu-id="c7a76-1050">Requirements</span></span>

|<span data-ttu-id="c7a76-1051">要求</span><span class="sxs-lookup"><span data-stu-id="c7a76-1051">Requirement</span></span>|<span data-ttu-id="c7a76-1052">值</span><span class="sxs-lookup"><span data-stu-id="c7a76-1052">Value</span></span>|
|---|---|
|[<span data-ttu-id="c7a76-1053">最低的邮箱要求集版本</span><span class="sxs-lookup"><span data-stu-id="c7a76-1053">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c7a76-1054">1.0</span><span class="sxs-lookup"><span data-stu-id="c7a76-1054">1.0</span></span>|
|[<span data-ttu-id="c7a76-1055">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="c7a76-1055">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c7a76-1056">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c7a76-1056">ReadItem</span></span>|
|[<span data-ttu-id="c7a76-1057">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c7a76-1057">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="c7a76-1058">阅读</span><span class="sxs-lookup"><span data-stu-id="c7a76-1058">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="c7a76-1059">返回：</span><span class="sxs-lookup"><span data-stu-id="c7a76-1059">Returns:</span></span>

<span data-ttu-id="c7a76-1060">一个包含与清单 XML 文件所定正则表达式的字符串相匹配的数组。</span><span class="sxs-lookup"><span data-stu-id="c7a76-1060">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<dl class="param-type"><span data-ttu-id="c7a76-1061">

<dt>类型</dt>

</span><span class="sxs-lookup"><span data-stu-id="c7a76-1061">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="c7a76-1062">数组.< 字符串 ></span><span class="sxs-lookup"><span data-stu-id="c7a76-1062">Array.< String ></span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="c7a76-1063">示例</span><span class="sxs-lookup"><span data-stu-id="c7a76-1063">Example</span></span>

```
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

####  <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="c7a76-1064">getSelectedDataAsync (coercionType，[选项]，回调) → {字符串}</span><span class="sxs-lookup"><span data-stu-id="c7a76-1064">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="c7a76-1065">以异步方式返回邮件的主题或正文中选定的数据。</span><span class="sxs-lookup"><span data-stu-id="c7a76-1065">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="c7a76-p163">如果没有选定内容，但光标位于正文或主题中，此方法将会为所选数据返回 null。如果选定的是字段，而不是正文或主题，则此方法返回 `InvalidSelection` 错误。</span><span class="sxs-lookup"><span data-stu-id="c7a76-p163">If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c7a76-1068">参数：</span><span class="sxs-lookup"><span data-stu-id="c7a76-1068">Parameters:</span></span>

|<span data-ttu-id="c7a76-1069">名称</span><span class="sxs-lookup"><span data-stu-id="c7a76-1069">Name</span></span>|<span data-ttu-id="c7a76-1070">类型</span><span class="sxs-lookup"><span data-stu-id="c7a76-1070">Type</span></span>|<span data-ttu-id="c7a76-1071">属性</span><span class="sxs-lookup"><span data-stu-id="c7a76-1071">Attributes</span></span>|<span data-ttu-id="c7a76-1072">说明</span><span class="sxs-lookup"><span data-stu-id="c7a76-1072">Description</span></span>|
|---|---|---|---|
|`coercionType`|[<span data-ttu-id="c7a76-1073">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="c7a76-1073">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="c7a76-p164">请求数据的格式。如果为文本，则此方法返回纯文本作为字符串，删除任何显示的 HTML 标记。如果为 HTML，则此方法返回所选文本，不论是纯文本还是 HTML。</span><span class="sxs-lookup"><span data-stu-id="c7a76-p164">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`|<span data-ttu-id="c7a76-1077">对象</span><span class="sxs-lookup"><span data-stu-id="c7a76-1077">Object</span></span>|<span data-ttu-id="c7a76-1078">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="c7a76-1078">&lt;optional&gt;</span></span>|<span data-ttu-id="c7a76-1079">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="c7a76-1079">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="c7a76-1080">对象</span><span class="sxs-lookup"><span data-stu-id="c7a76-1080">Object</span></span>|<span data-ttu-id="c7a76-1081">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="c7a76-1081">&lt;optional&gt;</span></span>|<span data-ttu-id="c7a76-1082">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="c7a76-1082">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="c7a76-1083">函数</span><span class="sxs-lookup"><span data-stu-id="c7a76-1083">function</span></span>||<span data-ttu-id="c7a76-1084">方法完成后，使用单个参数 `asyncResult` （一个 [`AsyncResult`](/javascript/api/office/office.asyncresult)   对象）调用在 `callback`  参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="c7a76-1084">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="c7a76-1085">若要从回调方法访问所选数据，请调用 `asyncResult.value.data`。</span><span class="sxs-lookup"><span data-stu-id="c7a76-1085">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="c7a76-1086">若要访问选定内容的源属性，请调用 `asyncResult.value.sourceProperty`，这将为 `body` 或 `subject` 。</span><span class="sxs-lookup"><span data-stu-id="c7a76-1086">To access the source property that the selection comes from, call , which will be either  or .|</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c7a76-1087">要求</span><span class="sxs-lookup"><span data-stu-id="c7a76-1087">Requirements</span></span>

|<span data-ttu-id="c7a76-1088">要求</span><span class="sxs-lookup"><span data-stu-id="c7a76-1088">Requirement</span></span>|<span data-ttu-id="c7a76-1089">值</span><span class="sxs-lookup"><span data-stu-id="c7a76-1089">Value</span></span>|
|---|---|
|[<span data-ttu-id="c7a76-1090">最低的邮箱要求集版本</span><span class="sxs-lookup"><span data-stu-id="c7a76-1090">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c7a76-1091">1.2</span><span class="sxs-lookup"><span data-stu-id="c7a76-1091">1.2</span></span>|
|[<span data-ttu-id="c7a76-1092">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="c7a76-1092">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c7a76-1093">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="c7a76-1093">ReadWriteItem</span></span>|
|[<span data-ttu-id="c7a76-1094">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c7a76-1094">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="c7a76-1095">撰写</span><span class="sxs-lookup"><span data-stu-id="c7a76-1095">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="c7a76-1096">返回：</span><span class="sxs-lookup"><span data-stu-id="c7a76-1096">Returns:</span></span>

<span data-ttu-id="c7a76-1097">作为字符串的所选数据的格式由 `coercionType` 确定。</span><span class="sxs-lookup"><span data-stu-id="c7a76-1097">The selected data as a string with format determined by `coercionType`.</span></span>

<dl class="param-type"><span data-ttu-id="c7a76-1098">

<dt>类型</dt>

</span><span class="sxs-lookup"><span data-stu-id="c7a76-1098">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="c7a76-1099">字符串</span><span class="sxs-lookup"><span data-stu-id="c7a76-1099">String</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="c7a76-1100">示例</span><span class="sxs-lookup"><span data-stu-id="c7a76-1100">Example</span></span>

```
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

#### <a name="getselectedentities--entitiesjavascriptapioutlook17officeentities"></a><span data-ttu-id="c7a76-1101">getSelectedEntities() → {[实体](/javascript/api/outlook_1_7/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="c7a76-1101">getSelectedEntities() → {[Entities](/javascript/api/outlook_1_7/office.entities)}</span></span>

<span data-ttu-id="c7a76-p166">获取在用户已选择的突出显示匹配项中找到的实体。突出显示匹配项适用于[上下文外接程序](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins)。</span><span class="sxs-lookup"><span data-stu-id="c7a76-p166">Gets the entities found in a highlighted match a user has selected. Highlighted matches apply to [contextual add-ins](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="c7a76-1104">在 Outlook for iOS 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="c7a76-1104">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="c7a76-1105">要求</span><span class="sxs-lookup"><span data-stu-id="c7a76-1105">Requirements</span></span>

|<span data-ttu-id="c7a76-1106">要求</span><span class="sxs-lookup"><span data-stu-id="c7a76-1106">Requirement</span></span>|<span data-ttu-id="c7a76-1107">值</span><span class="sxs-lookup"><span data-stu-id="c7a76-1107">Value</span></span>|
|---|---|
|[<span data-ttu-id="c7a76-1108">最低的邮箱要求集版本</span><span class="sxs-lookup"><span data-stu-id="c7a76-1108">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c7a76-1109">1.6</span><span class="sxs-lookup"><span data-stu-id="c7a76-1109">-16</span></span>|
|[<span data-ttu-id="c7a76-1110">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="c7a76-1110">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c7a76-1111">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c7a76-1111">ReadItem</span></span>|
|[<span data-ttu-id="c7a76-1112">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c7a76-1112">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="c7a76-1113">阅读</span><span class="sxs-lookup"><span data-stu-id="c7a76-1113">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="c7a76-1114">返回：</span><span class="sxs-lookup"><span data-stu-id="c7a76-1114">Returns:</span></span>

<span data-ttu-id="c7a76-1115">类型：[实体](/javascript/api/outlook_1_7/office.entities)</span><span class="sxs-lookup"><span data-stu-id="c7a76-1115">Type: [Entities](/javascript/api/outlook_1_7/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="c7a76-1116">示例</span><span class="sxs-lookup"><span data-stu-id="c7a76-1116">Example</span></span>

<span data-ttu-id="c7a76-1117">以下示例访问用户选择的突出显示匹配项中的地址实体。</span><span class="sxs-lookup"><span data-stu-id="c7a76-1117">The following example accesses the addresses entities in the highlighted match selected by the user.</span></span>

```
var contacts = Office.context.mailbox.item.getSelectedEntities().addresses;
```

#### <a name="getselectedregexmatches--object"></a><span data-ttu-id="c7a76-1118">getSelectedRegExMatches() → {对象}</span><span class="sxs-lookup"><span data-stu-id="c7a76-1118">getSelectedRegExMatches() → {Object}</span></span>

<span data-ttu-id="c7a76-p167">返回突出显示匹配项中匹配在清单 XML 文件中定义的正则表达式的字符串值。突出显示匹配项适用于[上下文外接程序](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins)。</span><span class="sxs-lookup"><span data-stu-id="c7a76-p167">Returns string values in a highlighted match that match the regular expressions defined in the manifest XML file. Highlighted matches apply to [contextual add-ins](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="c7a76-1121">在 Outlook for iOS 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="c7a76-1121">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="c7a76-p168">`getSelectedRegExMatches` 方法返回与每个 `ItemHasRegularExpressionMatch` 所定义的正则表达式或 `ItemHasKnownEntity` 清单 XML 文件中的规则元素相匹配的字符串。对于 `ItemHasRegularExpressionMatch` 规则，匹配字符串必须发生在该规则指定的项目属性中。`PropertyName` 简单类型定义所支持的属性。</span><span class="sxs-lookup"><span data-stu-id="c7a76-p168">The `getSelectedRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="c7a76-1125">例如，考虑一个加载项具有以下 `Rule` 元素：</span><span class="sxs-lookup"><span data-stu-id="c7a76-1125">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="c7a76-1126">从 `getRegExMatches` 返回的对象应有两个属性：`fruits` 和 `veggies`。</span><span class="sxs-lookup"><span data-stu-id="c7a76-1126">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="c7a76-p169">如果在项目的正文属性上指定 `ItemHasRegularExpressionMatch` 规则，则正则表达式应进一步筛选正文，不应尝试返回该项目的整个正文。使用正则表达式（如 `.*`）获取项目的整个正文并不总是返回预期的结果。而该使用 [`Body.getAsync`](/javascript/api/outlook_1_7/office.body#getasync-coerciontype--options--callback-) 方法检索整个正文。</span><span class="sxs-lookup"><span data-stu-id="c7a76-p169">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook_1_7/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="c7a76-1130">要求</span><span class="sxs-lookup"><span data-stu-id="c7a76-1130">Requirements</span></span>

|<span data-ttu-id="c7a76-1131">要求</span><span class="sxs-lookup"><span data-stu-id="c7a76-1131">Requirement</span></span>|<span data-ttu-id="c7a76-1132">值</span><span class="sxs-lookup"><span data-stu-id="c7a76-1132">Value</span></span>|
|---|---|
|[<span data-ttu-id="c7a76-1133">最低的邮箱要求集版本</span><span class="sxs-lookup"><span data-stu-id="c7a76-1133">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c7a76-1134">1.6</span><span class="sxs-lookup"><span data-stu-id="c7a76-1134">-16</span></span>|
|[<span data-ttu-id="c7a76-1135">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="c7a76-1135">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c7a76-1136">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c7a76-1136">ReadItem</span></span>|
|[<span data-ttu-id="c7a76-1137">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c7a76-1137">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="c7a76-1138">阅读</span><span class="sxs-lookup"><span data-stu-id="c7a76-1138">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="c7a76-1139">返回：</span><span class="sxs-lookup"><span data-stu-id="c7a76-1139">Returns:</span></span>

<span data-ttu-id="c7a76-p170">一个包含与清单 XML 文件中所定义正则表达式的字符串数组相匹配的对象。每个数组的名称等于匹配 `ItemHasRegularExpressionMatch` 规则的 `RegExName` 属性的相应值或匹配 `ItemHasKnownEntity` 规则的 `FilterName` 属性。</span><span class="sxs-lookup"><span data-stu-id="c7a76-p170">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

##### <a name="example"></a><span data-ttu-id="c7a76-1142">示例</span><span class="sxs-lookup"><span data-stu-id="c7a76-1142">Example</span></span>

<span data-ttu-id="c7a76-1143">以下示例显示了如何访问正则表达式规则元素 `fruits` 和 `veggies` 的匹配项的数组，这些元素在清单中指定。</span><span class="sxs-lookup"><span data-stu-id="c7a76-1143">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```
var selectedMatches = Office.context.mailbox.item.getSelectedRegExMatches();
var fruits = selectedMatches.fruits;
var veggies = selectedMatches.veggies;
```

####  <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="c7a76-1144">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="c7a76-1144">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="c7a76-1145">为所选项目的加载项异步加载自定义属性。</span><span class="sxs-lookup"><span data-stu-id="c7a76-1145">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="c7a76-p171">自定义属性在每个应用、每个项目中储存为键/值对。此方法在回调中返回 `CustomProperties` 对象，该回调提供方法访问当前项目和当前加载项的特定自定义属性。自定义属性未在项目上加密，因此这不应用作安全存储。</span><span class="sxs-lookup"><span data-stu-id="c7a76-p171">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c7a76-1149">参数：</span><span class="sxs-lookup"><span data-stu-id="c7a76-1149">Parameters:</span></span>

|<span data-ttu-id="c7a76-1150">名称</span><span class="sxs-lookup"><span data-stu-id="c7a76-1150">Name</span></span>|<span data-ttu-id="c7a76-1151">类型</span><span class="sxs-lookup"><span data-stu-id="c7a76-1151">Type</span></span>|<span data-ttu-id="c7a76-1152">属性</span><span class="sxs-lookup"><span data-stu-id="c7a76-1152">Attributes</span></span>|<span data-ttu-id="c7a76-1153">说明</span><span class="sxs-lookup"><span data-stu-id="c7a76-1153">Description</span></span>|
|---|---|---|---|
|`callback`|<span data-ttu-id="c7a76-1154">函数</span><span class="sxs-lookup"><span data-stu-id="c7a76-1154">function</span></span>||<span data-ttu-id="c7a76-1155">方法完成后，使用单个参数 `asyncResult` （一个 [`AsyncResult`](/javascript/api/office/office.asyncresult)   对象）调用在 `callback`  参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="c7a76-1155">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="c7a76-1156">自定义属性作为 [`CustomProperties`](/javascript/api/outlook_1_7/office.customproperties) 对象，在 `asyncResult.value` 属性中提供。</span><span class="sxs-lookup"><span data-stu-id="c7a76-1156">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook_1_7/office.customproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="c7a76-1157">该对象可用于获取、 设置和删除项目中的自定义属性，并将针对自定义属性集的更改保存回服务器。</span><span class="sxs-lookup"><span data-stu-id="c7a76-1157">The custom properties are provided as a CustomProperties object in the asyncResult.value property. This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`|<span data-ttu-id="c7a76-1158">对象</span><span class="sxs-lookup"><span data-stu-id="c7a76-1158">Object</span></span>|<span data-ttu-id="c7a76-1159">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="c7a76-1159">&lt;optional&gt;</span></span>|<span data-ttu-id="c7a76-1160">开发人员可以在回调函数中提供他们想要访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="c7a76-1160">Developers can provide any object they wish to access in the callback method.</span></span> <span data-ttu-id="c7a76-1161">可以通过回调函数的 `asyncResult.asyncContext` 属性访问该对象。</span><span class="sxs-lookup"><span data-stu-id="c7a76-1161">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.|</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c7a76-1162">要求</span><span class="sxs-lookup"><span data-stu-id="c7a76-1162">Requirements</span></span>

|<span data-ttu-id="c7a76-1163">要求</span><span class="sxs-lookup"><span data-stu-id="c7a76-1163">Requirement</span></span>|<span data-ttu-id="c7a76-1164">值</span><span class="sxs-lookup"><span data-stu-id="c7a76-1164">Value</span></span>|
|---|---|
|[<span data-ttu-id="c7a76-1165">最低的邮箱要求集版本</span><span class="sxs-lookup"><span data-stu-id="c7a76-1165">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c7a76-1166">1.0</span><span class="sxs-lookup"><span data-stu-id="c7a76-1166">1.0</span></span>|
|[<span data-ttu-id="c7a76-1167">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="c7a76-1167">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c7a76-1168">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c7a76-1168">ReadItem</span></span>|
|[<span data-ttu-id="c7a76-1169">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c7a76-1169">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="c7a76-1170">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="c7a76-1170">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="c7a76-1171">示例</span><span class="sxs-lookup"><span data-stu-id="c7a76-1171">Example</span></span>

<span data-ttu-id="c7a76-p174">以下代码示例显示了如何使用 `loadCustomPropertiesAsync` 方法异步加载特定于当前项目的自定义属性。该示例还显示了如何使用 `CustomProperties.saveAsync` 方法将这些属性重新保存到服务器。加载自定义属性后，该代码示例将使用 `CustomProperties.get` 方法读取自定义属性 `myProp`，使用 `CustomProperties.set` 方法写入自定义属性 `otherProp`，最后调用 `saveAsync` 方法保存这些自定义属性。</span><span class="sxs-lookup"><span data-stu-id="c7a76-p174">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

```
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

####  <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="c7a76-1175">removeAttachmentAsync (attachmentId，[选项]，[回调])</span><span class="sxs-lookup"><span data-stu-id="c7a76-1175">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="c7a76-1176">将附件从邮件或约会中删除。</span><span class="sxs-lookup"><span data-stu-id="c7a76-1176">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="c7a76-p175">`removeAttachmentAsync` 方法删除项目中带指定标识符的附件。最佳做法是，仅当同一个邮件应用程序在同一会话中添加了一个附件时，你才应使用该附件标识符来删除该附件。在 Outlook Web App 和适用于设备的 OWA 中，附件标识符只在同一个会话中才有效。当用户关闭应用程序，或者如果用户开始在内嵌窗体中撰写，并在随后弹出的内嵌窗体中继续在单独的窗口撰写时，会话即结束。</span><span class="sxs-lookup"><span data-stu-id="c7a76-p175">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item. As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session. In Outlook Web App and OWA for Devices, the attachment identifier is valid only within the same session. A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c7a76-1181">参数：</span><span class="sxs-lookup"><span data-stu-id="c7a76-1181">Parameters:</span></span>

|<span data-ttu-id="c7a76-1182">名称</span><span class="sxs-lookup"><span data-stu-id="c7a76-1182">Name</span></span>|<span data-ttu-id="c7a76-1183">类型</span><span class="sxs-lookup"><span data-stu-id="c7a76-1183">Type</span></span>|<span data-ttu-id="c7a76-1184">属性</span><span class="sxs-lookup"><span data-stu-id="c7a76-1184">Attributes</span></span>|<span data-ttu-id="c7a76-1185">说明</span><span class="sxs-lookup"><span data-stu-id="c7a76-1185">Description</span></span>|
|---|---|---|---|
|`attachmentId`|<span data-ttu-id="c7a76-1186">字符串</span><span class="sxs-lookup"><span data-stu-id="c7a76-1186">String</span></span>||<span data-ttu-id="c7a76-p176">要删除的附件的标识符。字符串的最大长度为 100 个字符。</span><span class="sxs-lookup"><span data-stu-id="c7a76-p176">The identifier of the attachment to remove. The maximum length of the string is 100 characters.</span></span>|
|`options`|<span data-ttu-id="c7a76-1189">对象</span><span class="sxs-lookup"><span data-stu-id="c7a76-1189">Object</span></span>|<span data-ttu-id="c7a76-1190">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="c7a76-1190">&lt;optional&gt;</span></span>|<span data-ttu-id="c7a76-1191">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="c7a76-1191">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="c7a76-1192">对象</span><span class="sxs-lookup"><span data-stu-id="c7a76-1192">Object</span></span>|<span data-ttu-id="c7a76-1193">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="c7a76-1193">&lt;optional&gt;</span></span>|<span data-ttu-id="c7a76-1194">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="c7a76-1194">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="c7a76-1195">函数</span><span class="sxs-lookup"><span data-stu-id="c7a76-1195">function</span></span>|<span data-ttu-id="c7a76-1196">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="c7a76-1196">&lt;optional&gt;</span></span>|<span data-ttu-id="c7a76-1197">方法完成后，使用单个参数 `asyncResult` （一个 [`AsyncResult`](/javascript/api/office/office.asyncresult)   对象）调用在 `callback`  参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="c7a76-1197">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="c7a76-1198">如果删除附件失败，`asyncResult.error` 属性将包含一个说明失败原因的错误代码。</span><span class="sxs-lookup"><span data-stu-id="c7a76-1198">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="c7a76-1199">错误</span><span class="sxs-lookup"><span data-stu-id="c7a76-1199">Errors</span></span>

|<span data-ttu-id="c7a76-1200">错误代码</span><span class="sxs-lookup"><span data-stu-id="c7a76-1200">Error code</span></span>|<span data-ttu-id="c7a76-1201">说明</span><span class="sxs-lookup"><span data-stu-id="c7a76-1201">Description</span></span>|
|------------|-------------|
|`InvalidAttachmentId`|<span data-ttu-id="c7a76-1202">附件标识符不存在。</span><span class="sxs-lookup"><span data-stu-id="c7a76-1202">The attachment identifier does not exist.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c7a76-1203">要求</span><span class="sxs-lookup"><span data-stu-id="c7a76-1203">Requirements</span></span>

|<span data-ttu-id="c7a76-1204">要求</span><span class="sxs-lookup"><span data-stu-id="c7a76-1204">Requirement</span></span>|<span data-ttu-id="c7a76-1205">值</span><span class="sxs-lookup"><span data-stu-id="c7a76-1205">Value</span></span>|
|---|---|
|[<span data-ttu-id="c7a76-1206">最低的邮箱要求集版本</span><span class="sxs-lookup"><span data-stu-id="c7a76-1206">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c7a76-1207">1.1</span><span class="sxs-lookup"><span data-stu-id="c7a76-1207">1.1</span></span>|
|[<span data-ttu-id="c7a76-1208">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="c7a76-1208">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c7a76-1209">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="c7a76-1209">ReadWriteItem</span></span>|
|[<span data-ttu-id="c7a76-1210">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c7a76-1210">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="c7a76-1211">撰写</span><span class="sxs-lookup"><span data-stu-id="c7a76-1211">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="c7a76-1212">示例</span><span class="sxs-lookup"><span data-stu-id="c7a76-1212">Example</span></span>

<span data-ttu-id="c7a76-1213">以下代码删除包含标识符 '0' 的附件。</span><span class="sxs-lookup"><span data-stu-id="c7a76-1213">The following code removes an attachment with an identifier of '0'.</span></span>

```
Office.context.mailbox.item.removeAttachmentAsync(
  '0',
  { asyncContext : null },
  function (asyncResult)
  {
    console.log(asyncResult.status);
  }
);
```

####  <a name="removehandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="c7a76-1214">removeHandlerAsync (eventType，处理程序，[选项]，[回调])</span><span class="sxs-lookup"><span data-stu-id="c7a76-1214">addHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="c7a76-1215">删除受支持事件的事件处理程序。</span><span class="sxs-lookup"><span data-stu-id="c7a76-1215">Removes an event handler for a</span></span>

<span data-ttu-id="c7a76-1216">目前，支持的事件类型是 `Office.EventType.AppointmentTimeChanged` 和 `Office.EventType.RecipientsChanged`。 `Office.EventType.RecurrenceChanged`</span><span class="sxs-lookup"><span data-stu-id="c7a76-1216">Currently the supported event types are `Office.EventType.AppointmentTimeChanged`, `Office.EventType.RecipientsChanged`, and `Office.EventType.RecurrenceChanged`</span></span>

##### <a name="parameters"></a><span data-ttu-id="c7a76-1217">参数：</span><span class="sxs-lookup"><span data-stu-id="c7a76-1217">Parameters:</span></span>

| <span data-ttu-id="c7a76-1218">名称</span><span class="sxs-lookup"><span data-stu-id="c7a76-1218">Name</span></span> | <span data-ttu-id="c7a76-1219">类型</span><span class="sxs-lookup"><span data-stu-id="c7a76-1219">Type</span></span> | <span data-ttu-id="c7a76-1220">属性</span><span class="sxs-lookup"><span data-stu-id="c7a76-1220">Attributes</span></span> | <span data-ttu-id="c7a76-1221">说明</span><span class="sxs-lookup"><span data-stu-id="c7a76-1221">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="c7a76-1222">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="c7a76-1222">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="c7a76-1223">应调用处理程序的事件。</span><span class="sxs-lookup"><span data-stu-id="c7a76-1223">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="c7a76-1224">函数</span><span class="sxs-lookup"><span data-stu-id="c7a76-1224">Function</span></span> || <span data-ttu-id="c7a76-p177">用于处理事件的函数。此函数必须接受单个参数，即对象文本。参数上的 `type` 属性将匹配传递给 `removeHandlerAsync` 的 `eventType` 参数。</span><span class="sxs-lookup"><span data-stu-id="c7a76-p177">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `removeHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="c7a76-1228">对象</span><span class="sxs-lookup"><span data-stu-id="c7a76-1228">Object</span></span> | <span data-ttu-id="c7a76-1229">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="c7a76-1229">&lt;optional&gt;</span></span> | <span data-ttu-id="c7a76-1230">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="c7a76-1230">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="c7a76-1231">对象</span><span class="sxs-lookup"><span data-stu-id="c7a76-1231">Object</span></span> | <span data-ttu-id="c7a76-1232">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="c7a76-1232">&lt;optional&gt;</span></span> | <span data-ttu-id="c7a76-1233">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="c7a76-1233">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="c7a76-1234">函数</span><span class="sxs-lookup"><span data-stu-id="c7a76-1234">function</span></span>| <span data-ttu-id="c7a76-1235">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="c7a76-1235">&lt;optional&gt;</span></span>|<span data-ttu-id="c7a76-1236">方法完成后，使用单个参数 `asyncResult` （一个 [`AsyncResult`](/javascript/api/office/office.asyncresult)   对象）调用在 `callback`  参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="c7a76-1236">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c7a76-1237">要求</span><span class="sxs-lookup"><span data-stu-id="c7a76-1237">Requirements</span></span>

|<span data-ttu-id="c7a76-1238">要求</span><span class="sxs-lookup"><span data-stu-id="c7a76-1238">Requirement</span></span>| <span data-ttu-id="c7a76-1239">值</span><span class="sxs-lookup"><span data-stu-id="c7a76-1239">Value</span></span>|
|---|---|
|[<span data-ttu-id="c7a76-1240">最低的邮箱要求集版本</span><span class="sxs-lookup"><span data-stu-id="c7a76-1240">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c7a76-1241">1.7</span><span class="sxs-lookup"><span data-stu-id="c7a76-1241">-17</span></span> |
|[<span data-ttu-id="c7a76-1242">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="c7a76-1242">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c7a76-1243">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c7a76-1243">ReadItem</span></span> |
|[<span data-ttu-id="c7a76-1244">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c7a76-1244">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c7a76-1245">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="c7a76-1245">Compose or read</span></span> |

##### <a name="example"></a><span data-ttu-id="c7a76-1246">示例</span><span class="sxs-lookup"><span data-stu-id="c7a76-1246">Example</span></span>

```
Office.initialize = function (reason) {
  $(document).ready(function () {
    Office.context.mailbox.item.removeHandlerAsync(Office.EventType.RecurrenceChanged, loadNewItem, function (result) {
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

####  <a name="saveasyncoptions-callback"></a><span data-ttu-id="c7a76-1247">saveAsync ([选项] 回调)</span><span class="sxs-lookup"><span data-stu-id="c7a76-1247">saveAsync([options], callback)</span></span>

<span data-ttu-id="c7a76-1248">异步保存项目。</span><span class="sxs-lookup"><span data-stu-id="c7a76-1248">Asynchronously saves an item.</span></span>

<span data-ttu-id="c7a76-p178">调用时，此方法将当前邮件保存为草稿，并通过回调方法返回项目 ID。在 Outlook Web App 或 Outlook 联机模式下，该项目被保存到服务器中。在 Outlook 缓存模式下，该项目被保存到本地缓存中。</span><span class="sxs-lookup"><span data-stu-id="c7a76-p178">When invoked, this method saves the current message as a draft and returns the item id via the callback method. In Outlook Web App or Outlook in online mode, the item is saved to the server. In Outlook in cached mode, the item is saved to the local cache.</span></span>

> [!NOTE]
> <span data-ttu-id="c7a76-1252">如果加载项调用 `saveAsync` 中的项目在撰写模式下才能获取 `itemId` 若要使用 EWS 或 REST API，请注意，缓存模式 Outlook 时，可能需要一些时间才能将项目实际同步到服务器。</span><span class="sxs-lookup"><span data-stu-id="c7a76-1252">Note: If your add-in calls `saveAsync` on an item in compose mode in order to get an `itemId` to use with EWS or the REST API, be aware that when Outlook is in cached mode, it may take some time before the item is actually synced to the server. Until the item is synced, using the  will return an error.</span></span> <span data-ttu-id="c7a76-1253">直到该项目同步，使用 `itemId` 将返回错误。</span><span class="sxs-lookup"><span data-stu-id="c7a76-1253">Until the item is synced, using the `itemId` will return an error.</span></span>

<span data-ttu-id="c7a76-p180">由于约会没有草稿状态，如果以撰写模式在约会中调用 `saveAsync`，则该项将被保存为用户日历中的正常约会。对于之前未保存过的新约会，则不会发送邀请。保存现有约会将向添加或删除的与会者发送更新。</span><span class="sxs-lookup"><span data-stu-id="c7a76-p180">Since appointments have no draft state, if `saveAsync` is called on an appointment in compose mode, the item will be saved as a normal appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent. Saving an existing appointment will send an update to added or removed attendees.</span></span>

> [!NOTE]
> <span data-ttu-id="c7a76-1257">以下客户端在约会上的撰写模式下具有 `saveAsync` 的不同行为：</span><span class="sxs-lookup"><span data-stu-id="c7a76-1257">The following clients have different behavior for `saveAsync` on appointments in compose mode:</span></span>
>
> - <span data-ttu-id="c7a76-1258">Mac Outlook 在会议的撰写模式中不支持 `saveAsync` 。</span><span class="sxs-lookup"><span data-stu-id="c7a76-1258">Note: Mac Outlook does not support `saveAsync` on a meeting in compose mode. Calling  on a meeting in Mac Outlook will return an error.</span></span> <span data-ttu-id="c7a76-1259">在 Mac Outlook 中的会议上调用 `saveAsync` ，则将返回错误。</span><span class="sxs-lookup"><span data-stu-id="c7a76-1259">Note: Mac Outlook does not support  on a meeting in compose mode. Calling `saveAsync` on a meeting in Mac Outlook will return an error.</span></span>
> - <span data-ttu-id="c7a76-1260">当 `saveAsync` 在撰写模式调用约会时，Outlook 网页版总会发送一个邀请或更新。</span><span class="sxs-lookup"><span data-stu-id="c7a76-1260">Outlook on the web always sends an invitation or update when `saveAsync` is called on an appointment in compose mode.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c7a76-1261">参数：</span><span class="sxs-lookup"><span data-stu-id="c7a76-1261">Parameters:</span></span>

|<span data-ttu-id="c7a76-1262">名称</span><span class="sxs-lookup"><span data-stu-id="c7a76-1262">Name</span></span>|<span data-ttu-id="c7a76-1263">类型</span><span class="sxs-lookup"><span data-stu-id="c7a76-1263">Type</span></span>|<span data-ttu-id="c7a76-1264">属性</span><span class="sxs-lookup"><span data-stu-id="c7a76-1264">Attributes</span></span>|<span data-ttu-id="c7a76-1265">说明</span><span class="sxs-lookup"><span data-stu-id="c7a76-1265">Description</span></span>|
|---|---|---|---|
|`options`|<span data-ttu-id="c7a76-1266">对象</span><span class="sxs-lookup"><span data-stu-id="c7a76-1266">Object</span></span>|<span data-ttu-id="c7a76-1267">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="c7a76-1267">&lt;optional&gt;</span></span>|<span data-ttu-id="c7a76-1268">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="c7a76-1268">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="c7a76-1269">对象</span><span class="sxs-lookup"><span data-stu-id="c7a76-1269">Object</span></span>|<span data-ttu-id="c7a76-1270">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="c7a76-1270">&lt;optional&gt;</span></span>|<span data-ttu-id="c7a76-1271">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="c7a76-1271">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`|<span data-ttu-id="c7a76-1272">函数</span><span class="sxs-lookup"><span data-stu-id="c7a76-1272">function</span></span>||<span data-ttu-id="c7a76-1273">方法完成后，使用单个参数 `asyncResult` （一个 [`AsyncResult`](/javascript/api/office/office.asyncresult)   对象）调用在 `callback`  参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="c7a76-1273">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="c7a76-1274">如果成功，该项目标识符在 `asyncResult.value` 属性中提供。</span><span class="sxs-lookup"><span data-stu-id="c7a76-1274">On success, the item identifier is provided in the `asyncResult.value` property.|</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c7a76-1275">要求</span><span class="sxs-lookup"><span data-stu-id="c7a76-1275">Requirements</span></span>

|<span data-ttu-id="c7a76-1276">要求</span><span class="sxs-lookup"><span data-stu-id="c7a76-1276">Requirement</span></span>|<span data-ttu-id="c7a76-1277">值</span><span class="sxs-lookup"><span data-stu-id="c7a76-1277">Value</span></span>|
|---|---|
|[<span data-ttu-id="c7a76-1278">最低的邮箱要求集版本</span><span class="sxs-lookup"><span data-stu-id="c7a76-1278">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c7a76-1279">1.3</span><span class="sxs-lookup"><span data-stu-id="c7a76-1279">1.3</span></span>|
|[<span data-ttu-id="c7a76-1280">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="c7a76-1280">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c7a76-1281">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="c7a76-1281">ReadWriteItem</span></span>|
|[<span data-ttu-id="c7a76-1282">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c7a76-1282">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="c7a76-1283">撰写</span><span class="sxs-lookup"><span data-stu-id="c7a76-1283">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="c7a76-1284">示例</span><span class="sxs-lookup"><span data-stu-id="c7a76-1284">Examples</span></span>

```
Office.context.mailbox.item.saveAsync(
  function callback(result) {
    // Process the result
  });
```

<span data-ttu-id="c7a76-p182">下面是传递给回调函数的 `result` 参数示例。`value` 属性包含的该项的项目 ID。</span><span class="sxs-lookup"><span data-stu-id="c7a76-p182">The following is an example of the `result` parameter passed to the callback function. The `value` property contains the item ID of the item.</span></span>

```
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

####  <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="c7a76-1287">setSelectedDataAsync (数据，[选项]，回调)</span><span class="sxs-lookup"><span data-stu-id="c7a76-1287">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="c7a76-1288">以异步方式将数据插入到邮件的正文或主题中。</span><span class="sxs-lookup"><span data-stu-id="c7a76-1288">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="c7a76-p183">`setSelectedDataAsync` 方法将指定字符串插入到项目主题或正文的光标位置，或者，如果在编辑器中已选择文本，则该方法将替换选择的文本。如果光标不在正文或主题字段中，则返回错误。插入之后，光标会位于插入内容的末尾。</span><span class="sxs-lookup"><span data-stu-id="c7a76-p183">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c7a76-1292">参数：</span><span class="sxs-lookup"><span data-stu-id="c7a76-1292">Parameters:</span></span>

|<span data-ttu-id="c7a76-1293">名称</span><span class="sxs-lookup"><span data-stu-id="c7a76-1293">Name</span></span>|<span data-ttu-id="c7a76-1294">类型</span><span class="sxs-lookup"><span data-stu-id="c7a76-1294">Type</span></span>|<span data-ttu-id="c7a76-1295">属性</span><span class="sxs-lookup"><span data-stu-id="c7a76-1295">Attributes</span></span>|<span data-ttu-id="c7a76-1296">说明</span><span class="sxs-lookup"><span data-stu-id="c7a76-1296">Description</span></span>|
|---|---|---|---|
|`data`|<span data-ttu-id="c7a76-1297">字符串</span><span class="sxs-lookup"><span data-stu-id="c7a76-1297">String</span></span>||<span data-ttu-id="c7a76-p184">要插入的数据。数据不得超过 1,000,000 个字符。如果传入的数据超过 1,000,000 个字符，则会引发 `ArgumentOutOfRange` 异常。</span><span class="sxs-lookup"><span data-stu-id="c7a76-p184">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`|<span data-ttu-id="c7a76-1301">对象</span><span class="sxs-lookup"><span data-stu-id="c7a76-1301">Object</span></span>|<span data-ttu-id="c7a76-1302">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="c7a76-1302">&lt;optional&gt;</span></span>|<span data-ttu-id="c7a76-1303">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="c7a76-1303">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`|<span data-ttu-id="c7a76-1304">对象</span><span class="sxs-lookup"><span data-stu-id="c7a76-1304">Object</span></span>|<span data-ttu-id="c7a76-1305">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="c7a76-1305">&lt;optional&gt;</span></span>|<span data-ttu-id="c7a76-1306">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="c7a76-1306">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`|[<span data-ttu-id="c7a76-1307">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="c7a76-1307">Office.CoercionType</span></span>](office.md#coerciontype-string)|<span data-ttu-id="c7a76-1308">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="c7a76-1308">&lt;optional&gt;</span></span>|<span data-ttu-id="c7a76-p185">如果是 `text`，则在 Outlook Web App 和 Outlook 中应用当前样式。如果该字段是 HTML 编辑器，则仅插入文本数据，即使数据为 HTML。</span><span class="sxs-lookup"><span data-stu-id="c7a76-p185">If `text`, the current style is applied in Outlook Web App and Outlook. If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="c7a76-p186">如果 `html` 和该字段支持 HTML（主题不支持），则在 Outlook Web App 中应用当前样式，而在 Outlook 中应用默认样式。如果该字段是文本字段，则返回 `InvalidDataFormat` 错误。</span><span class="sxs-lookup"><span data-stu-id="c7a76-p186">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook Web App and the default style is applied in Outlook. If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="c7a76-1313">如果未设置 `coercionType`，则结果取决于该字段：如果该字段是 HTML，则使用 HTML；如果该字段是文本，则使用纯文本。</span><span class="sxs-lookup"><span data-stu-id="c7a76-1313">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`|<span data-ttu-id="c7a76-1314">函数</span><span class="sxs-lookup"><span data-stu-id="c7a76-1314">function</span></span>||<span data-ttu-id="c7a76-1315">方法完成后，使用单个参数 `asyncResult` （一个 [`AsyncResult`](/javascript/api/office/office.asyncresult)   对象）调用在 `callback`  参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="c7a76-1315">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c7a76-1316">要求</span><span class="sxs-lookup"><span data-stu-id="c7a76-1316">Requirements</span></span>

|<span data-ttu-id="c7a76-1317">要求</span><span class="sxs-lookup"><span data-stu-id="c7a76-1317">Requirement</span></span>|<span data-ttu-id="c7a76-1318">值</span><span class="sxs-lookup"><span data-stu-id="c7a76-1318">Value</span></span>|
|---|---|
|[<span data-ttu-id="c7a76-1319">最低的邮箱要求集版本</span><span class="sxs-lookup"><span data-stu-id="c7a76-1319">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)|<span data-ttu-id="c7a76-1320">1.2</span><span class="sxs-lookup"><span data-stu-id="c7a76-1320">1.2</span></span>|
|[<span data-ttu-id="c7a76-1321">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="c7a76-1321">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)|<span data-ttu-id="c7a76-1322">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="c7a76-1322">ReadWriteItem</span></span>|
|[<span data-ttu-id="c7a76-1323">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c7a76-1323">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)|<span data-ttu-id="c7a76-1324">撰写</span><span class="sxs-lookup"><span data-stu-id="c7a76-1324">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="c7a76-1325">示例</span><span class="sxs-lookup"><span data-stu-id="c7a76-1325">Example</span></span>

```
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```