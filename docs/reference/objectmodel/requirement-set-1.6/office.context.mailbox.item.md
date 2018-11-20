
# <a name="item"></a><span data-ttu-id="1daf4-101">item</span><span class="sxs-lookup"><span data-stu-id="1daf4-101">item</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmditem"></a><span data-ttu-id="1daf4-102">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span><span class="sxs-lookup"><span data-stu-id="1daf4-102">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span></span>

<span data-ttu-id="1daf4-p101">`item` 命名空间用于访问当前选定的邮件、会议请求或约会。可以通过使用 [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlook16officemailboxenumsitemtype) 属性确定 `item` 的类型。</span><span class="sxs-lookup"><span data-stu-id="1daf4-p101">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlook16officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="1daf4-105">Requirements</span><span class="sxs-lookup"><span data-stu-id="1daf4-105">Requirements</span></span>

|<span data-ttu-id="1daf4-106">要求</span><span class="sxs-lookup"><span data-stu-id="1daf4-106">Requirement</span></span>| <span data-ttu-id="1daf4-107">值</span><span class="sxs-lookup"><span data-stu-id="1daf4-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="1daf4-108">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="1daf4-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1daf4-109">1.0</span><span class="sxs-lookup"><span data-stu-id="1daf4-109">1.0</span></span>|
|[<span data-ttu-id="1daf4-110">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="1daf4-110">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1daf4-111">受限</span><span class="sxs-lookup"><span data-stu-id="1daf4-111">Restricted</span></span>|
|[<span data-ttu-id="1daf4-112">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="1daf4-112">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1daf4-113">Compose 或 Read</span><span class="sxs-lookup"><span data-stu-id="1daf4-113">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="1daf4-114">成员和方法</span><span class="sxs-lookup"><span data-stu-id="1daf4-114">Members and methods</span></span>

| <span data-ttu-id="1daf4-115">成员</span><span class="sxs-lookup"><span data-stu-id="1daf4-115">Member</span></span> | <span data-ttu-id="1daf4-116">类型</span><span class="sxs-lookup"><span data-stu-id="1daf4-116">Type</span></span> |
|--------|------|
| [<span data-ttu-id="1daf4-117">attachments</span><span class="sxs-lookup"><span data-stu-id="1daf4-117">attachments</span></span>](#attachments-arrayattachmentdetailsjavascriptapioutlook16officeattachmentdetails) | <span data-ttu-id="1daf4-118">成员</span><span class="sxs-lookup"><span data-stu-id="1daf4-118">Member</span></span> |
| [<span data-ttu-id="1daf4-119">bcc</span><span class="sxs-lookup"><span data-stu-id="1daf4-119">bcc</span></span>](#bcc-recipientsjavascriptapioutlook16officerecipients) | <span data-ttu-id="1daf4-120">成员</span><span class="sxs-lookup"><span data-stu-id="1daf4-120">Member</span></span> |
| [<span data-ttu-id="1daf4-121">body</span><span class="sxs-lookup"><span data-stu-id="1daf4-121">body</span></span>](#body-bodyjavascriptapioutlook16officebody) | <span data-ttu-id="1daf4-122">成员</span><span class="sxs-lookup"><span data-stu-id="1daf4-122">Member</span></span> |
| [<span data-ttu-id="1daf4-123">cc</span><span class="sxs-lookup"><span data-stu-id="1daf4-123">cc</span></span>](#cc-arrayemailaddressdetailsjavascriptapioutlook16officeemailaddressdetailsrecipientsjavascriptapioutlook16officerecipients) | <span data-ttu-id="1daf4-124">成员</span><span class="sxs-lookup"><span data-stu-id="1daf4-124">Member</span></span> |
| [<span data-ttu-id="1daf4-125">conversationId</span><span class="sxs-lookup"><span data-stu-id="1daf4-125">conversationId</span></span>](#nullable-conversationid-string) | <span data-ttu-id="1daf4-126">成员</span><span class="sxs-lookup"><span data-stu-id="1daf4-126">Member</span></span> |
| [<span data-ttu-id="1daf4-127">dateTimeCreated</span><span class="sxs-lookup"><span data-stu-id="1daf4-127">dateTimeCreated</span></span>](#datetimecreated-date) | <span data-ttu-id="1daf4-128">成员</span><span class="sxs-lookup"><span data-stu-id="1daf4-128">Member</span></span> |
| [<span data-ttu-id="1daf4-129">dateTimeModified</span><span class="sxs-lookup"><span data-stu-id="1daf4-129">dateTimeModified</span></span>](#datetimemodified-date) | <span data-ttu-id="1daf4-130">成员</span><span class="sxs-lookup"><span data-stu-id="1daf4-130">Member</span></span> |
| [<span data-ttu-id="1daf4-131">end</span><span class="sxs-lookup"><span data-stu-id="1daf4-131">end</span></span>](#end-datetimejavascriptapioutlook16officetime) | <span data-ttu-id="1daf4-132">成员</span><span class="sxs-lookup"><span data-stu-id="1daf4-132">Member</span></span> |
| [<span data-ttu-id="1daf4-133">from</span><span class="sxs-lookup"><span data-stu-id="1daf4-133">from</span></span>](#from-emailaddressdetailsjavascriptapioutlook16officeemailaddressdetails) | <span data-ttu-id="1daf4-134">成员</span><span class="sxs-lookup"><span data-stu-id="1daf4-134">Member</span></span> |
| [<span data-ttu-id="1daf4-135">internetMessageId</span><span class="sxs-lookup"><span data-stu-id="1daf4-135">internetMessageId</span></span>](#internetmessageid-string) | <span data-ttu-id="1daf4-136">成员</span><span class="sxs-lookup"><span data-stu-id="1daf4-136">Member</span></span> |
| [<span data-ttu-id="1daf4-137">itemClass</span><span class="sxs-lookup"><span data-stu-id="1daf4-137">itemClass</span></span>](#itemclass-string) | <span data-ttu-id="1daf4-138">成员</span><span class="sxs-lookup"><span data-stu-id="1daf4-138">Member</span></span> |
| [<span data-ttu-id="1daf4-139">itemId</span><span class="sxs-lookup"><span data-stu-id="1daf4-139">itemId</span></span>](#nullable-itemid-string) | <span data-ttu-id="1daf4-140">成员</span><span class="sxs-lookup"><span data-stu-id="1daf4-140">Member</span></span> |
| [<span data-ttu-id="1daf4-141">itemType</span><span class="sxs-lookup"><span data-stu-id="1daf4-141">itemType</span></span>](#itemtype-officemailboxenumsitemtypejavascriptapioutlook16officemailboxenumsitemtype) | <span data-ttu-id="1daf4-142">成员</span><span class="sxs-lookup"><span data-stu-id="1daf4-142">Member</span></span> |
| [<span data-ttu-id="1daf4-143">location</span><span class="sxs-lookup"><span data-stu-id="1daf4-143">location</span></span>](#location-stringlocationjavascriptapioutlook16officelocation) | <span data-ttu-id="1daf4-144">成员</span><span class="sxs-lookup"><span data-stu-id="1daf4-144">Member</span></span> |
| [<span data-ttu-id="1daf4-145">normalizedSubject</span><span class="sxs-lookup"><span data-stu-id="1daf4-145">normalizedSubject</span></span>](#normalizedsubject-string) | <span data-ttu-id="1daf4-146">成员</span><span class="sxs-lookup"><span data-stu-id="1daf4-146">Member</span></span> |
| [<span data-ttu-id="1daf4-147">notificationMessages</span><span class="sxs-lookup"><span data-stu-id="1daf4-147">notificationMessages</span></span>](#notificationmessages-notificationmessagesjavascriptapioutlook16officenotificationmessages) | <span data-ttu-id="1daf4-148">成员</span><span class="sxs-lookup"><span data-stu-id="1daf4-148">Member</span></span> |
| [<span data-ttu-id="1daf4-149">optionalAttendees</span><span class="sxs-lookup"><span data-stu-id="1daf4-149">optionalAttendees</span></span>](#optionalattendees-arrayemailaddressdetailsjavascriptapioutlook16officeemailaddressdetailsrecipientsjavascriptapioutlook16officerecipients) | <span data-ttu-id="1daf4-150">成员</span><span class="sxs-lookup"><span data-stu-id="1daf4-150">Member</span></span> |
| [<span data-ttu-id="1daf4-151">organizer</span><span class="sxs-lookup"><span data-stu-id="1daf4-151">organizer</span></span>](#organizer-emailaddressdetailsjavascriptapioutlook16officeemailaddressdetails) | <span data-ttu-id="1daf4-152">成员</span><span class="sxs-lookup"><span data-stu-id="1daf4-152">Member</span></span> |
| [<span data-ttu-id="1daf4-153">requiredAttendees</span><span class="sxs-lookup"><span data-stu-id="1daf4-153">requiredAttendees</span></span>](#requiredattendees-arrayemailaddressdetailsjavascriptapioutlook16officeemailaddressdetailsrecipientsjavascriptapioutlook16officerecipients) | <span data-ttu-id="1daf4-154">Member</span><span class="sxs-lookup"><span data-stu-id="1daf4-154">Member</span></span> |
| [<span data-ttu-id="1daf4-155">sender</span><span class="sxs-lookup"><span data-stu-id="1daf4-155">sender</span></span>](#sender-emailaddressdetailsjavascriptapioutlook16officeemailaddressdetails) | <span data-ttu-id="1daf4-156">成员</span><span class="sxs-lookup"><span data-stu-id="1daf4-156">Member</span></span> |
| [<span data-ttu-id="1daf4-157">start</span><span class="sxs-lookup"><span data-stu-id="1daf4-157">start</span></span>](#start-datetimejavascriptapioutlook16officetime) | <span data-ttu-id="1daf4-158">成员</span><span class="sxs-lookup"><span data-stu-id="1daf4-158">Member</span></span> |
| [<span data-ttu-id="1daf4-159">subject</span><span class="sxs-lookup"><span data-stu-id="1daf4-159">subject</span></span>](#subject-stringsubjectjavascriptapioutlook16officesubject) | <span data-ttu-id="1daf4-160">成员</span><span class="sxs-lookup"><span data-stu-id="1daf4-160">Member</span></span> |
| [<span data-ttu-id="1daf4-161">to</span><span class="sxs-lookup"><span data-stu-id="1daf4-161">to</span></span>](#to-arrayemailaddressdetailsjavascriptapioutlook16officeemailaddressdetailsrecipientsjavascriptapioutlook16officerecipients) | <span data-ttu-id="1daf4-162">成员</span><span class="sxs-lookup"><span data-stu-id="1daf4-162">Member</span></span> |
| [<span data-ttu-id="1daf4-163">addFileAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="1daf4-163">addFileAttachmentAsync</span></span>](#addfileattachmentasyncuri-attachmentname-options-callback) | <span data-ttu-id="1daf4-164">方法</span><span class="sxs-lookup"><span data-stu-id="1daf4-164">Method</span></span> |
| [<span data-ttu-id="1daf4-165">addItemAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="1daf4-165">addItemAttachmentAsync</span></span>](#additemattachmentasyncitemid-attachmentname-options-callback) | <span data-ttu-id="1daf4-166">方法</span><span class="sxs-lookup"><span data-stu-id="1daf4-166">Method</span></span> |
| [<span data-ttu-id="1daf4-167">close</span><span class="sxs-lookup"><span data-stu-id="1daf4-167">close</span></span>](#close) | <span data-ttu-id="1daf4-168">方法</span><span class="sxs-lookup"><span data-stu-id="1daf4-168">Method</span></span> |
| [<span data-ttu-id="1daf4-169">displayReplyAllForm</span><span class="sxs-lookup"><span data-stu-id="1daf4-169">displayReplyAllForm</span></span>](#displayreplyallformformdata) | <span data-ttu-id="1daf4-170">方法</span><span class="sxs-lookup"><span data-stu-id="1daf4-170">Method</span></span> |
| [<span data-ttu-id="1daf4-171">displayReplyForm</span><span class="sxs-lookup"><span data-stu-id="1daf4-171">displayReplyForm</span></span>](#displayreplyformformdata) | <span data-ttu-id="1daf4-172">方法</span><span class="sxs-lookup"><span data-stu-id="1daf4-172">Method</span></span> |
| [<span data-ttu-id="1daf4-173">getEntities</span><span class="sxs-lookup"><span data-stu-id="1daf4-173">getEntities</span></span>](#getentities--entitiesjavascriptapioutlook16officeentities) | <span data-ttu-id="1daf4-174">方法</span><span class="sxs-lookup"><span data-stu-id="1daf4-174">Method</span></span> |
| [<span data-ttu-id="1daf4-175">getEntitiesByType</span><span class="sxs-lookup"><span data-stu-id="1daf4-175">getEntitiesByType</span></span>](#getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlook16officecontactmeetingsuggestionjavascriptapioutlook16officemeetingsuggestionphonenumberjavascriptapioutlook16officephonenumbertasksuggestionjavascriptapioutlook16officetasksuggestion) | <span data-ttu-id="1daf4-176">方法</span><span class="sxs-lookup"><span data-stu-id="1daf4-176">Method</span></span> |
| [<span data-ttu-id="1daf4-177">getFilteredEntitiesByName</span><span class="sxs-lookup"><span data-stu-id="1daf4-177">getFilteredEntitiesByName</span></span>](#getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlook16officecontactmeetingsuggestionjavascriptapioutlook16officemeetingsuggestionphonenumberjavascriptapioutlook16officephonenumbertasksuggestionjavascriptapioutlook16officetasksuggestion) | <span data-ttu-id="1daf4-178">方法</span><span class="sxs-lookup"><span data-stu-id="1daf4-178">Method</span></span> |
| [<span data-ttu-id="1daf4-179">getRegExMatches</span><span class="sxs-lookup"><span data-stu-id="1daf4-179">getRegExMatches</span></span>](#getregexmatches--object) | <span data-ttu-id="1daf4-180">方法</span><span class="sxs-lookup"><span data-stu-id="1daf4-180">Method</span></span> |
| [<span data-ttu-id="1daf4-181">getRegExMatchesByName</span><span class="sxs-lookup"><span data-stu-id="1daf4-181">getRegExMatchesByName</span></span>](#getregexmatchesbynamename--nullable-array-string-) | <span data-ttu-id="1daf4-182">方法</span><span class="sxs-lookup"><span data-stu-id="1daf4-182">Method</span></span> |
| [<span data-ttu-id="1daf4-183">getSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="1daf4-183">getSelectedDataAsync</span></span>](#getselecteddataasynccoerciontype-options-callback--string) | <span data-ttu-id="1daf4-184">方法</span><span class="sxs-lookup"><span data-stu-id="1daf4-184">Method</span></span> |
| [<span data-ttu-id="1daf4-185">getSelectedEntities</span><span class="sxs-lookup"><span data-stu-id="1daf4-185">getSelectedEntities</span></span>](#getselectedentities--entitiesjavascriptapioutlook16officeentities) | <span data-ttu-id="1daf4-186">方法</span><span class="sxs-lookup"><span data-stu-id="1daf4-186">Method</span></span> |
| [<span data-ttu-id="1daf4-187">getSelectedRegExMatches</span><span class="sxs-lookup"><span data-stu-id="1daf4-187">getSelectedRegExMatches</span></span>](#getselectedregexmatches--object) | <span data-ttu-id="1daf4-188">方法</span><span class="sxs-lookup"><span data-stu-id="1daf4-188">Method</span></span> |
| [<span data-ttu-id="1daf4-189">loadCustomPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="1daf4-189">loadCustomPropertiesAsync</span></span>](#loadcustompropertiesasynccallback-usercontext) | <span data-ttu-id="1daf4-190">方法</span><span class="sxs-lookup"><span data-stu-id="1daf4-190">Method</span></span> |
| [<span data-ttu-id="1daf4-191">removeAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="1daf4-191">removeAttachmentAsync</span></span>](#removeattachmentasyncattachmentid-options-callback) | <span data-ttu-id="1daf4-192">方法</span><span class="sxs-lookup"><span data-stu-id="1daf4-192">Method</span></span> |
| [<span data-ttu-id="1daf4-193">saveAsync</span><span class="sxs-lookup"><span data-stu-id="1daf4-193">saveAsync</span></span>](#saveasyncoptions-callback) | <span data-ttu-id="1daf4-194">方法</span><span class="sxs-lookup"><span data-stu-id="1daf4-194">Method</span></span> |
| [<span data-ttu-id="1daf4-195">setSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="1daf4-195">setSelectedDataAsync</span></span>](#setselecteddataasyncdata-options-callback) | <span data-ttu-id="1daf4-196">方法</span><span class="sxs-lookup"><span data-stu-id="1daf4-196">Method</span></span> |

### <a name="example"></a><span data-ttu-id="1daf4-197">示例</span><span class="sxs-lookup"><span data-stu-id="1daf4-197">Example</span></span>

<span data-ttu-id="1daf4-198">以下 JavaScript 代码示例显示了如何访问 Outlook 中当前项目的 `subject` 属性。</span><span class="sxs-lookup"><span data-stu-id="1daf4-198">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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

### <a name="members"></a><span data-ttu-id="1daf4-199">成员</span><span class="sxs-lookup"><span data-stu-id="1daf4-199">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlook16officeattachmentdetails"></a><span data-ttu-id="1daf4-200">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_6/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="1daf4-200">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_6/office.attachmentdetails)></span></span>

<span data-ttu-id="1daf4-p102">获取项目的附件数组。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="1daf4-p102">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="1daf4-203">由于潜在的安全问题，某些类型的文件会受到 Outlook 阻止，并且不会返回。</span><span class="sxs-lookup"><span data-stu-id="1daf4-203">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="1daf4-204">如需了解更多信息，请参阅 [Outlook 中阻止的附件](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519)。</span><span class="sxs-lookup"><span data-stu-id="1daf4-204">For more information see [Payments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="1daf4-205">类型：</span><span class="sxs-lookup"><span data-stu-id="1daf4-205">Type:</span></span>

*   <span data-ttu-id="1daf4-206">Array.<[AttachmentDetails](/javascript/api/outlook_1_6/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="1daf4-206">Array.<[AttachmentDetails](/javascript/api/outlook_1_6/office.attachmentdetails)></span></span>

##### <a name="requirements"></a><span data-ttu-id="1daf4-207">要求</span><span class="sxs-lookup"><span data-stu-id="1daf4-207">Requirements</span></span>

|<span data-ttu-id="1daf4-208">要求</span><span class="sxs-lookup"><span data-stu-id="1daf4-208">Requirement</span></span>| <span data-ttu-id="1daf4-209">值</span><span class="sxs-lookup"><span data-stu-id="1daf4-209">Value</span></span>|
|---|---|
|[<span data-ttu-id="1daf4-210">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="1daf4-210">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1daf4-211">1.0</span><span class="sxs-lookup"><span data-stu-id="1daf4-211">1.0</span></span>|
|[<span data-ttu-id="1daf4-212">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="1daf4-212">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1daf4-213">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1daf4-213">ReadItem</span></span>|
|[<span data-ttu-id="1daf4-214">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="1daf4-214">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1daf4-215">阅读</span><span class="sxs-lookup"><span data-stu-id="1daf4-215">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="1daf4-216">示例</span><span class="sxs-lookup"><span data-stu-id="1daf4-216">Example</span></span>

<span data-ttu-id="1daf4-217">以下代码使用当前项目上所有附件的详细信息构成 HTML 字符串。</span><span class="sxs-lookup"><span data-stu-id="1daf4-217">The following code builds an HTML string with details of all attachments on the current item.</span></span>

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

####  <a name="bcc-recipientsjavascriptapioutlook16officerecipients"></a><span data-ttu-id="1daf4-218">bcc :[Recipients](/javascript/api/outlook_1_6/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="1daf4-218">bcc :[Recipients](/javascript/api/outlook_1_6/office.recipients)</span></span>

<span data-ttu-id="1daf4-219">获取对象，该对象提供用于获取或更新邮件的密件抄送 (Bcc) 行上的收件人的方法。</span><span class="sxs-lookup"><span data-stu-id="1daf4-219">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="1daf4-220">仅限撰写模式。</span><span class="sxs-lookup"><span data-stu-id="1daf4-220">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="1daf4-221">类型：</span><span class="sxs-lookup"><span data-stu-id="1daf4-221">Type:</span></span>

*   [<span data-ttu-id="1daf4-222">收件人</span><span class="sxs-lookup"><span data-stu-id="1daf4-222">Recipients</span></span>](/javascript/api/outlook_1_6/office.recipients)

##### <a name="requirements"></a><span data-ttu-id="1daf4-223">要求</span><span class="sxs-lookup"><span data-stu-id="1daf4-223">Requirements</span></span>

|<span data-ttu-id="1daf4-224">要求</span><span class="sxs-lookup"><span data-stu-id="1daf4-224">Requirement</span></span>| <span data-ttu-id="1daf4-225">值</span><span class="sxs-lookup"><span data-stu-id="1daf4-225">Value</span></span>|
|---|---|
|[<span data-ttu-id="1daf4-226">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="1daf4-226">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1daf4-227">1.1</span><span class="sxs-lookup"><span data-stu-id="1daf4-227">1.1</span></span>|
|[<span data-ttu-id="1daf4-228">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="1daf4-228">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1daf4-229">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1daf4-229">ReadItem</span></span>|
|[<span data-ttu-id="1daf4-230">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="1daf4-230">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1daf4-231">撰写</span><span class="sxs-lookup"><span data-stu-id="1daf4-231">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="1daf4-232">示例</span><span class="sxs-lookup"><span data-stu-id="1daf4-232">Example</span></span>

```js
Office.context.mailbox.item.bcc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.bcc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.bcc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfBccRecipients = asyncResult.value;
}
```

####  <a name="body-bodyjavascriptapioutlook16officebody"></a><span data-ttu-id="1daf4-233">body :[Body](/javascript/api/outlook_1_6/office.body)</span><span class="sxs-lookup"><span data-stu-id="1daf4-233">body :[Body](/javascript/api/outlook_1_6/office.body)</span></span>

<span data-ttu-id="1daf4-234">获取一个提供用于处理项目正文的方法的对象。</span><span class="sxs-lookup"><span data-stu-id="1daf4-234">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="1daf4-235">类型：</span><span class="sxs-lookup"><span data-stu-id="1daf4-235">Type:</span></span>

*   [<span data-ttu-id="1daf4-236">Body</span><span class="sxs-lookup"><span data-stu-id="1daf4-236">Body</span></span>](/javascript/api/outlook_1_6/office.body)

##### <a name="requirements"></a><span data-ttu-id="1daf4-237">要求</span><span class="sxs-lookup"><span data-stu-id="1daf4-237">Requirements</span></span>

|<span data-ttu-id="1daf4-238">要求</span><span class="sxs-lookup"><span data-stu-id="1daf4-238">Requirement</span></span>| <span data-ttu-id="1daf4-239">值</span><span class="sxs-lookup"><span data-stu-id="1daf4-239">Value</span></span>|
|---|---|
|[<span data-ttu-id="1daf4-240">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="1daf4-240">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1daf4-241">1.1</span><span class="sxs-lookup"><span data-stu-id="1daf4-241">1.1</span></span>|
|[<span data-ttu-id="1daf4-242">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="1daf4-242">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1daf4-243">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1daf4-243">ReadItem</span></span>|
|[<span data-ttu-id="1daf4-244">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="1daf4-244">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1daf4-245">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="1daf4-245">Compose or read</span></span>|

####  <a name="cc-arrayemailaddressdetailsjavascriptapioutlook16officeemailaddressdetailsrecipientsjavascriptapioutlook16officerecipients"></a><span data-ttu-id="1daf4-246">cc :Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_6/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="1daf4-246">cc :Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_6/office.recipients)</span></span>

<span data-ttu-id="1daf4-247">提供对邮件的抄送 (Cc) 收件人的访问权限。</span><span class="sxs-lookup"><span data-stu-id="1daf4-247">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="1daf4-248">对象的类型和访问级别取决于当前项目的模式。</span><span class="sxs-lookup"><span data-stu-id="1daf4-248">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="1daf4-249">阅读模式</span><span class="sxs-lookup"><span data-stu-id="1daf4-249">Read mode</span></span>

<span data-ttu-id="1daf4-p106">`cc` 属性返回包含邮件的**抄送**行上所列的每个收件人的 `EmailAddressDetails` 对象的数组。集合上限为 100 个成员。</span><span class="sxs-lookup"><span data-stu-id="1daf4-p106">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="1daf4-252">撰写模式</span><span class="sxs-lookup"><span data-stu-id="1daf4-252">Compose mode</span></span>

<span data-ttu-id="1daf4-253">`cc` 属性返回 `Recipients`一个对象，该对象提供用于获取或更新邮件的**抄送**行上收件人的方法。</span><span class="sxs-lookup"><span data-stu-id="1daf4-253">The `cc` property returns a `Recipients` object that provides methods for manipulating the recipients on the **Cc** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="1daf4-254">类型：</span><span class="sxs-lookup"><span data-stu-id="1daf4-254">Type:</span></span>

*   <span data-ttu-id="1daf4-255">Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_6/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="1daf4-255">Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_6/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="1daf4-256">要求</span><span class="sxs-lookup"><span data-stu-id="1daf4-256">Requirements</span></span>

|<span data-ttu-id="1daf4-257">要求</span><span class="sxs-lookup"><span data-stu-id="1daf4-257">Requirement</span></span>| <span data-ttu-id="1daf4-258">值</span><span class="sxs-lookup"><span data-stu-id="1daf4-258">Value</span></span>|
|---|---|
|[<span data-ttu-id="1daf4-259">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="1daf4-259">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1daf4-260">1.0</span><span class="sxs-lookup"><span data-stu-id="1daf4-260">1.0</span></span>|
|[<span data-ttu-id="1daf4-261">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="1daf4-261">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1daf4-262">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1daf4-262">ReadItem</span></span>|
|[<span data-ttu-id="1daf4-263">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="1daf4-263">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1daf4-264">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="1daf4-264">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="1daf4-265">示例</span><span class="sxs-lookup"><span data-stu-id="1daf4-265">Example</span></span>

```js
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

####  <a name="nullable-conversationid-string"></a><span data-ttu-id="1daf4-266">(nullable) conversationId :String</span><span class="sxs-lookup"><span data-stu-id="1daf4-266">(nullable) conversationId :String</span></span>

<span data-ttu-id="1daf4-267">获取包含特定消息的电子邮件会话的标识符。</span><span class="sxs-lookup"><span data-stu-id="1daf4-267">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="1daf4-p107">如果在阅读窗体或撰写窗体的回复中激活邮件应用程序，则此属性可以获得一个整数值。如果用户随后更改了回复邮件的主题（若发送回复），则该邮件的对话 ID 将改变且之前获取的值将不适用。</span><span class="sxs-lookup"><span data-stu-id="1daf4-p107">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="1daf4-p108">对于撰写窗体的新项目，此属性获得一个 null 值。如果用户设置一个主题并保存该项目，`conversationId` 属性将返回一个值。</span><span class="sxs-lookup"><span data-stu-id="1daf4-p108">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="1daf4-272">类型:</span><span class="sxs-lookup"><span data-stu-id="1daf4-272">Type:</span></span>

*   <span data-ttu-id="1daf4-273">String</span><span class="sxs-lookup"><span data-stu-id="1daf4-273">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="1daf4-274">要求</span><span class="sxs-lookup"><span data-stu-id="1daf4-274">Requirements</span></span>

|<span data-ttu-id="1daf4-275">要求</span><span class="sxs-lookup"><span data-stu-id="1daf4-275">Requirement</span></span>| <span data-ttu-id="1daf4-276">值</span><span class="sxs-lookup"><span data-stu-id="1daf4-276">Value</span></span>|
|---|---|
|[<span data-ttu-id="1daf4-277">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="1daf4-277">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1daf4-278">1.0</span><span class="sxs-lookup"><span data-stu-id="1daf4-278">1.0</span></span>|
|[<span data-ttu-id="1daf4-279">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="1daf4-279">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1daf4-280">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1daf4-280">ReadItem</span></span>|
|[<span data-ttu-id="1daf4-281">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="1daf4-281">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1daf4-282">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="1daf4-282">Compose or read</span></span>|

#### <a name="datetimecreated-date"></a><span data-ttu-id="1daf4-283">dateTimeCreated :Date</span><span class="sxs-lookup"><span data-stu-id="1daf4-283">dateTimeCreated :Date</span></span>

<span data-ttu-id="1daf4-p109">获取项目创建的日期和时间。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="1daf4-p109">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="1daf4-286">类型：</span><span class="sxs-lookup"><span data-stu-id="1daf4-286">Type:</span></span>

*   <span data-ttu-id="1daf4-287">日期</span><span class="sxs-lookup"><span data-stu-id="1daf4-287">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="1daf4-288">要求</span><span class="sxs-lookup"><span data-stu-id="1daf4-288">Requirements</span></span>

|<span data-ttu-id="1daf4-289">要求</span><span class="sxs-lookup"><span data-stu-id="1daf4-289">Requirement</span></span>| <span data-ttu-id="1daf4-290">值</span><span class="sxs-lookup"><span data-stu-id="1daf4-290">Value</span></span>|
|---|---|
|[<span data-ttu-id="1daf4-291">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="1daf4-291">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1daf4-292">1.0</span><span class="sxs-lookup"><span data-stu-id="1daf4-292">1.0</span></span>|
|[<span data-ttu-id="1daf4-293">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="1daf4-293">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1daf4-294">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1daf4-294">ReadItem</span></span>|
|[<span data-ttu-id="1daf4-295">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="1daf4-295">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1daf4-296">阅读</span><span class="sxs-lookup"><span data-stu-id="1daf4-296">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="1daf4-297">示例</span><span class="sxs-lookup"><span data-stu-id="1daf4-297">Example</span></span>

```js
var created = Office.context.mailbox.item.dateTimeCreated;
```

#### <a name="datetimemodified-date"></a><span data-ttu-id="1daf4-298">dateTimeModified :Date</span><span class="sxs-lookup"><span data-stu-id="1daf4-298">dateTimeModified :Date</span></span>

<span data-ttu-id="1daf4-p110">获取项目最近一次修改的日期和时间。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="1daf4-p110">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="1daf4-301">在 Outlook for iOS 或 Outlook for Android 中不支持此成员。</span><span class="sxs-lookup"><span data-stu-id="1daf4-301">Note: This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="type"></a><span data-ttu-id="1daf4-302">类型:</span><span class="sxs-lookup"><span data-stu-id="1daf4-302">Type:</span></span>

*   <span data-ttu-id="1daf4-303">日期</span><span class="sxs-lookup"><span data-stu-id="1daf4-303">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="1daf4-304">要求</span><span class="sxs-lookup"><span data-stu-id="1daf4-304">Requirements</span></span>

|<span data-ttu-id="1daf4-305">要求</span><span class="sxs-lookup"><span data-stu-id="1daf4-305">Requirement</span></span>| <span data-ttu-id="1daf4-306">值</span><span class="sxs-lookup"><span data-stu-id="1daf4-306">Value</span></span>|
|---|---|
|[<span data-ttu-id="1daf4-307">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="1daf4-307">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1daf4-308">1.0</span><span class="sxs-lookup"><span data-stu-id="1daf4-308">1.0</span></span>|
|[<span data-ttu-id="1daf4-309">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="1daf4-309">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1daf4-310">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1daf4-310">ReadItem</span></span>|
|[<span data-ttu-id="1daf4-311">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="1daf4-311">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1daf4-312">阅读</span><span class="sxs-lookup"><span data-stu-id="1daf4-312">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="1daf4-313">示例</span><span class="sxs-lookup"><span data-stu-id="1daf4-313">Example</span></span>

```js
var modified = Office.context.mailbox.item.dateTimeModified;
```

####  <a name="end-datetimejavascriptapioutlook16officetime"></a><span data-ttu-id="1daf4-314">end :Date|[Time](/javascript/api/outlook_1_6/office.time)</span><span class="sxs-lookup"><span data-stu-id="1daf4-314">end :Date|[Time](/javascript/api/outlook_1_6/office.time)</span></span>

<span data-ttu-id="1daf4-315">获取或设置约会结束的日期和时间。</span><span class="sxs-lookup"><span data-stu-id="1daf4-315">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="1daf4-p111">将 `end` 属性表示为协调世界时 (UTC) 的日期和时间值。可使用 [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook16officelocalclienttime) 方法将 end 属性值转换为客户端的本地日期和时间。</span><span class="sxs-lookup"><span data-stu-id="1daf4-p111">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook16officelocalclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="1daf4-318">阅读模式</span><span class="sxs-lookup"><span data-stu-id="1daf4-318">Read mode</span></span>

<span data-ttu-id="1daf4-319">`end` 属性返回 `Date` 对象。</span><span class="sxs-lookup"><span data-stu-id="1daf4-319">The `end` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="1daf4-320">撰写模式</span><span class="sxs-lookup"><span data-stu-id="1daf4-320">Compose mode</span></span>

<span data-ttu-id="1daf4-321">`end` 属性返回 `Time` 对象。</span><span class="sxs-lookup"><span data-stu-id="1daf4-321">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="1daf4-322">使用 [`Time.setAsync`](/javascript/api/outlook_1_6/office.time#setasync-datetime--options--callback-) 方法设置结束时间时，应使用 [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) 方法将客户端的本地时间转换为服务器的 UTC。</span><span class="sxs-lookup"><span data-stu-id="1daf4-322">When you use the [`Time.setAsync`](/javascript/api/outlook_1_6/office.time#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="1daf4-323">类型：</span><span class="sxs-lookup"><span data-stu-id="1daf4-323">Type:</span></span>

*   <span data-ttu-id="1daf4-324">Date | [Time](/javascript/api/outlook_1_6/office.time)</span><span class="sxs-lookup"><span data-stu-id="1daf4-324">Date | [Time](/javascript/api/outlook_1_6/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="1daf4-325">要求</span><span class="sxs-lookup"><span data-stu-id="1daf4-325">Requirements</span></span>

|<span data-ttu-id="1daf4-326">要求</span><span class="sxs-lookup"><span data-stu-id="1daf4-326">Requirement</span></span>| <span data-ttu-id="1daf4-327">值</span><span class="sxs-lookup"><span data-stu-id="1daf4-327">Value</span></span>|
|---|---|
|[<span data-ttu-id="1daf4-328">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="1daf4-328">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1daf4-329">1.0</span><span class="sxs-lookup"><span data-stu-id="1daf4-329">1.0</span></span>|
|[<span data-ttu-id="1daf4-330">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="1daf4-330">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1daf4-331">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1daf4-331">ReadItem</span></span>|
|[<span data-ttu-id="1daf4-332">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="1daf4-332">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1daf4-333">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="1daf4-333">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="1daf4-334">示例</span><span class="sxs-lookup"><span data-stu-id="1daf4-334">Example</span></span>

<span data-ttu-id="1daf4-335">以下示例通过使用 `Time` 对象的 [`setAsync`](/javascript/api/outlook_1_6/office.time#setasync-datetime--options--callback-) 方法，设置撰写模式下约会的结束时间。</span><span class="sxs-lookup"><span data-stu-id="1daf4-335">The following example sets the end time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_6/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

#### <a name="from-emailaddressdetailsjavascriptapioutlook16officeemailaddressdetails"></a><span data-ttu-id="1daf4-336">from :[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="1daf4-336">from :[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)</span></span>

<span data-ttu-id="1daf4-p112">获取邮件发件人的电子邮件地址。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="1daf4-p112">Gets the email address of the sender of a message. Read mode only.</span></span>

<span data-ttu-id="1daf4-p113">`from` 和 [`sender`](#sender-emailaddressdetailsjavascriptapioutlook16officeemailaddressdetails) 属性表示同一个人，邮件由代理人发送的除外。在此情况下，`from` 属性表示代理程序，而 sender 属性表示代理人。</span><span class="sxs-lookup"><span data-stu-id="1daf4-p113">The `from` and [`sender`](#sender-emailaddressdetailsjavascriptapioutlook16officeemailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="1daf4-341">`from` 属性中 `EmailAddressDetails` 对象的 `recipientType` 属性为 `undefined`。</span><span class="sxs-lookup"><span data-stu-id="1daf4-341">Note: The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="1daf4-342">类型：</span><span class="sxs-lookup"><span data-stu-id="1daf4-342">Type:</span></span>

*   [<span data-ttu-id="1daf4-343">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="1daf4-343">EmailAddressDetails</span></span>](/javascript/api/outlook_1_6/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="1daf4-344">要求</span><span class="sxs-lookup"><span data-stu-id="1daf4-344">Requirements</span></span>

|<span data-ttu-id="1daf4-345">要求</span><span class="sxs-lookup"><span data-stu-id="1daf4-345">Requirement</span></span>| <span data-ttu-id="1daf4-346">值</span><span class="sxs-lookup"><span data-stu-id="1daf4-346">Value</span></span>|
|---|---|
|[<span data-ttu-id="1daf4-347">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="1daf4-347">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1daf4-348">1.0</span><span class="sxs-lookup"><span data-stu-id="1daf4-348">1.0</span></span>|
|[<span data-ttu-id="1daf4-349">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="1daf4-349">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1daf4-350">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1daf4-350">ReadItem</span></span>|
|[<span data-ttu-id="1daf4-351">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="1daf4-351">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1daf4-352">阅读</span><span class="sxs-lookup"><span data-stu-id="1daf4-352">Read</span></span>|

#### <a name="internetmessageid-string"></a><span data-ttu-id="1daf4-353">internetMessageId :String</span><span class="sxs-lookup"><span data-stu-id="1daf4-353">internetMessageId :String</span></span>

<span data-ttu-id="1daf4-p114">获取电子邮件的 Internet 消息标识符。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="1daf4-p114">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="1daf4-356">类型：</span><span class="sxs-lookup"><span data-stu-id="1daf4-356">Type:</span></span>

*   <span data-ttu-id="1daf4-357">String</span><span class="sxs-lookup"><span data-stu-id="1daf4-357">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="1daf4-358">要求</span><span class="sxs-lookup"><span data-stu-id="1daf4-358">Requirements</span></span>

|<span data-ttu-id="1daf4-359">要求</span><span class="sxs-lookup"><span data-stu-id="1daf4-359">Requirement</span></span>| <span data-ttu-id="1daf4-360">值</span><span class="sxs-lookup"><span data-stu-id="1daf4-360">Value</span></span>|
|---|---|
|[<span data-ttu-id="1daf4-361">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="1daf4-361">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1daf4-362">1.0</span><span class="sxs-lookup"><span data-stu-id="1daf4-362">1.0</span></span>|
|[<span data-ttu-id="1daf4-363">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="1daf4-363">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1daf4-364">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1daf4-364">ReadItem</span></span>|
|[<span data-ttu-id="1daf4-365">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="1daf4-365">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1daf4-366">阅读</span><span class="sxs-lookup"><span data-stu-id="1daf4-366">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="1daf4-367">示例</span><span class="sxs-lookup"><span data-stu-id="1daf4-367">Example</span></span>

```js
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

#### <a name="itemclass-string"></a><span data-ttu-id="1daf4-368">itemClass :String</span><span class="sxs-lookup"><span data-stu-id="1daf4-368">itemClass :String</span></span>

<span data-ttu-id="1daf4-p115">获取选定项目的 Exchange Web 服务项目类。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="1daf4-p115">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="1daf4-p116">`itemClass` 属性指定所选项目的邮件类别。以下是邮件或约会项目的默认邮件类别。</span><span class="sxs-lookup"><span data-stu-id="1daf4-p116">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

| <span data-ttu-id="1daf4-373">类型</span><span class="sxs-lookup"><span data-stu-id="1daf4-373">Type</span></span> | <span data-ttu-id="1daf4-374">说明</span><span class="sxs-lookup"><span data-stu-id="1daf4-374">Description</span></span> | <span data-ttu-id="1daf4-375">项目类</span><span class="sxs-lookup"><span data-stu-id="1daf4-375">item class</span></span> |
| --- | --- | --- |
| <span data-ttu-id="1daf4-376">约会项目</span><span class="sxs-lookup"><span data-stu-id="1daf4-376">Appointment items</span></span> | <span data-ttu-id="1daf4-377">这些是项目类 `IPM.Appointment` 或 `IPM.Appointment.Occurence` 的日历项目。</span><span class="sxs-lookup"><span data-stu-id="1daf4-377">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurence`.</span></span> | `IPM.Appointment`<br />`IPM.Appointment.Occurence` |
| <span data-ttu-id="1daf4-378">邮件项目</span><span class="sxs-lookup"><span data-stu-id="1daf4-378">Message items</span></span> | <span data-ttu-id="1daf4-379">这些项目包括具有默认邮件类别 `IPM.Note` 的电子邮件，以及使用 `IPM.Schedule.Meeting` 作为基础邮件类别的会议请求、响应和取消。</span><span class="sxs-lookup"><span data-stu-id="1daf4-379">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span> | `IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled` |

<span data-ttu-id="1daf4-380">你可以创建用于扩展默认邮件类别的自定义邮件类别，例如，自定义约会邮件类别 `IPM.Appointment.Contoso`。</span><span class="sxs-lookup"><span data-stu-id="1daf4-380">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="1daf4-381">类型：</span><span class="sxs-lookup"><span data-stu-id="1daf4-381">Type:</span></span>

*   <span data-ttu-id="1daf4-382">String</span><span class="sxs-lookup"><span data-stu-id="1daf4-382">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="1daf4-383">要求</span><span class="sxs-lookup"><span data-stu-id="1daf4-383">Requirements</span></span>

|<span data-ttu-id="1daf4-384">要求</span><span class="sxs-lookup"><span data-stu-id="1daf4-384">Requirement</span></span>| <span data-ttu-id="1daf4-385">值</span><span class="sxs-lookup"><span data-stu-id="1daf4-385">Value</span></span>|
|---|---|
|[<span data-ttu-id="1daf4-386">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="1daf4-386">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1daf4-387">1.0</span><span class="sxs-lookup"><span data-stu-id="1daf4-387">1.0</span></span>|
|[<span data-ttu-id="1daf4-388">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="1daf4-388">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1daf4-389">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1daf4-389">ReadItem</span></span>|
|[<span data-ttu-id="1daf4-390">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="1daf4-390">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1daf4-391">阅读</span><span class="sxs-lookup"><span data-stu-id="1daf4-391">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="1daf4-392">示例</span><span class="sxs-lookup"><span data-stu-id="1daf4-392">Example</span></span>

```js
var itemClass = Office.context.mailbox.item.itemClass;
```

#### <a name="nullable-itemid-string"></a><span data-ttu-id="1daf4-393">(nullable) itemId :String</span><span class="sxs-lookup"><span data-stu-id="1daf4-393">(nullable) itemId :String</span></span>

<span data-ttu-id="1daf4-p117">获取当前项目的 Exchange Web 服务项目标识符。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="1daf4-p117">Gets the Exchange Web Services item identifier for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="1daf4-396">`itemId` 属性返回的标识符与 Exchange Web 服务项目标识符相同。</span><span class="sxs-lookup"><span data-stu-id="1daf4-396">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="1daf4-397">`itemId` 属性与 Outlook 条目 ID 或 Outlook REST API 使用的 ID 不同。</span><span class="sxs-lookup"><span data-stu-id="1daf4-397">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="1daf4-398">使用此值进行 REST API 调用前，应使用 [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string) 对它进行转换。</span><span class="sxs-lookup"><span data-stu-id="1daf4-398">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="1daf4-399">有关详细信息，请参阅[从 Outlook 加载项使用 Outlook REST API](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id)。</span><span class="sxs-lookup"><span data-stu-id="1daf4-399">For more details, see [Use the Outlook REST APIs from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

<span data-ttu-id="1daf4-p119">`itemId` 属性在撰写模式下不可用。如果需要项目标识符，[`saveAsync`](#saveasyncoptions-callback) 方法可用于将项目保存到存储，这将在回调函数的 [`AsyncResult.value`](/javascript/api/office/office.asyncresult) 参数中返回项目标识符。</span><span class="sxs-lookup"><span data-stu-id="1daf4-p119">The `itemId` property is not available in compose mode. If an item identifier is required, the [`saveAsync`](#saveasyncoptions-callback) method can be used to save the item to the store, which will return the item identifier in the [`AsyncResult.value`](/javascript/api/office/office.asyncresult) parameter in the callback function.</span></span>

##### <a name="type"></a><span data-ttu-id="1daf4-402">类型:</span><span class="sxs-lookup"><span data-stu-id="1daf4-402">Type:</span></span>

*   <span data-ttu-id="1daf4-403">String</span><span class="sxs-lookup"><span data-stu-id="1daf4-403">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="1daf4-404">要求</span><span class="sxs-lookup"><span data-stu-id="1daf4-404">Requirements</span></span>

|<span data-ttu-id="1daf4-405">要求</span><span class="sxs-lookup"><span data-stu-id="1daf4-405">Requirement</span></span>| <span data-ttu-id="1daf4-406">值</span><span class="sxs-lookup"><span data-stu-id="1daf4-406">Value</span></span>|
|---|---|
|[<span data-ttu-id="1daf4-407">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="1daf4-407">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1daf4-408">1.0</span><span class="sxs-lookup"><span data-stu-id="1daf4-408">1.0</span></span>|
|[<span data-ttu-id="1daf4-409">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="1daf4-409">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1daf4-410">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1daf4-410">ReadItem</span></span>|
|[<span data-ttu-id="1daf4-411">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="1daf4-411">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1daf4-412">阅读</span><span class="sxs-lookup"><span data-stu-id="1daf4-412">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="1daf4-413">示例</span><span class="sxs-lookup"><span data-stu-id="1daf4-413">Example</span></span>

<span data-ttu-id="1daf4-p120">以下代码检查项目标识符是否存在。如果 `itemId` 属性返回 `null` 或 `undefined`，则将项目保存到存储，并从异步结果中获取项目标识符。</span><span class="sxs-lookup"><span data-stu-id="1daf4-p120">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

```js
var itemId = Office.context.mailbox.item.itemId;
if (itemId === null || itemId == undefined) {
  Office.context.mailbox.item.saveAsync(function(result){
    itemId = result.value;
  });
}
```

####  <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlook16officemailboxenumsitemtype"></a><span data-ttu-id="1daf4-416">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_6/office.mailboxenums.itemtype)</span><span class="sxs-lookup"><span data-stu-id="1daf4-416">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_6/office.mailboxenums.itemtype)</span></span>

<span data-ttu-id="1daf4-417">获取实例表示的项的类型。</span><span class="sxs-lookup"><span data-stu-id="1daf4-417">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="1daf4-418">`itemType` 属性返回其中一个 `ItemType` 枚举值，指示 `item` 对象实例是邮件还是约会。</span><span class="sxs-lookup"><span data-stu-id="1daf4-418">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="1daf4-419">类型：</span><span class="sxs-lookup"><span data-stu-id="1daf4-419">Type:</span></span>

*   [<span data-ttu-id="1daf4-420">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="1daf4-420">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook_1_6/office.mailboxenums.itemtype)

##### <a name="requirements"></a><span data-ttu-id="1daf4-421">要求</span><span class="sxs-lookup"><span data-stu-id="1daf4-421">Requirements</span></span>

|<span data-ttu-id="1daf4-422">要求</span><span class="sxs-lookup"><span data-stu-id="1daf4-422">Requirement</span></span>| <span data-ttu-id="1daf4-423">值</span><span class="sxs-lookup"><span data-stu-id="1daf4-423">Value</span></span>|
|---|---|
|[<span data-ttu-id="1daf4-424">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="1daf4-424">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1daf4-425">1.0</span><span class="sxs-lookup"><span data-stu-id="1daf4-425">1.0</span></span>|
|[<span data-ttu-id="1daf4-426">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="1daf4-426">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1daf4-427">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1daf4-427">ReadItem</span></span>|
|[<span data-ttu-id="1daf4-428">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="1daf4-428">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1daf4-429">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="1daf4-429">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="1daf4-430">示例</span><span class="sxs-lookup"><span data-stu-id="1daf4-430">Example</span></span>

```js
if (Office.context.mailbox.item.itemType == Office.MailboxEnums.ItemType.Message)
  // do something
else
  // do something else
```

####  <a name="location-stringlocationjavascriptapioutlook16officelocation"></a><span data-ttu-id="1daf4-431">location :String|[Location](/javascript/api/outlook_1_6/office.location)</span><span class="sxs-lookup"><span data-stu-id="1daf4-431">location :String|[Location](/javascript/api/outlook_1_6/office.location)</span></span>

<span data-ttu-id="1daf4-432">获取或设置约会的位置。</span><span class="sxs-lookup"><span data-stu-id="1daf4-432">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="1daf4-433">阅读模式</span><span class="sxs-lookup"><span data-stu-id="1daf4-433">Read mode</span></span>

<span data-ttu-id="1daf4-434">`location` 属性返回一个包含约会位置的字符串。</span><span class="sxs-lookup"><span data-stu-id="1daf4-434">The `location` property returns a string that contains the location of the appointment.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="1daf4-435">撰写模式</span><span class="sxs-lookup"><span data-stu-id="1daf4-435">Compose mode</span></span>

<span data-ttu-id="1daf4-436">`location` 属性返回一个 `Location` 对象，该对象提供用于获取和设置约会位置的方法。</span><span class="sxs-lookup"><span data-stu-id="1daf4-436">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="1daf4-437">类型：</span><span class="sxs-lookup"><span data-stu-id="1daf4-437">Type:</span></span>

*   <span data-ttu-id="1daf4-438">String | [Location](/javascript/api/outlook_1_6/office.location)</span><span class="sxs-lookup"><span data-stu-id="1daf4-438">String | [Location](/javascript/api/outlook_1_6/office.location)</span></span>

##### <a name="requirements"></a><span data-ttu-id="1daf4-439">要求</span><span class="sxs-lookup"><span data-stu-id="1daf4-439">Requirements</span></span>

|<span data-ttu-id="1daf4-440">要求</span><span class="sxs-lookup"><span data-stu-id="1daf4-440">Requirement</span></span>| <span data-ttu-id="1daf4-441">值</span><span class="sxs-lookup"><span data-stu-id="1daf4-441">Value</span></span>|
|---|---|
|[<span data-ttu-id="1daf4-442">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="1daf4-442">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1daf4-443">1.0</span><span class="sxs-lookup"><span data-stu-id="1daf4-443">1.0</span></span>|
|[<span data-ttu-id="1daf4-444">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="1daf4-444">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1daf4-445">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1daf4-445">ReadItem</span></span>|
|[<span data-ttu-id="1daf4-446">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="1daf4-446">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1daf4-447">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="1daf4-447">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="1daf4-448">示例</span><span class="sxs-lookup"><span data-stu-id="1daf4-448">Example</span></span>

```js
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

#### <a name="normalizedsubject-string"></a><span data-ttu-id="1daf4-449">normalizedSubject :String</span><span class="sxs-lookup"><span data-stu-id="1daf4-449">normalizedSubject :String</span></span>

<span data-ttu-id="1daf4-p121">获取删除了所有前缀（包括 `RE:` 和 `FWD:`）的项目主题。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="1daf4-p121">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="1daf4-p122">normalizedSubject 属性获取包含由电子邮件程序添加的任何标准前缀（如 `RE:` 和 `FW:`）的项目主题。若要获取包含完整前缀的项目主题，请使用 [`subject`](#subject-stringsubjectjavascriptapioutlook16officesubject) 属性。</span><span class="sxs-lookup"><span data-stu-id="1daf4-p122">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubjectjavascriptapioutlook16officesubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="1daf4-454">类型:</span><span class="sxs-lookup"><span data-stu-id="1daf4-454">Type:</span></span>

*   <span data-ttu-id="1daf4-455">String</span><span class="sxs-lookup"><span data-stu-id="1daf4-455">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="1daf4-456">要求</span><span class="sxs-lookup"><span data-stu-id="1daf4-456">Requirements</span></span>

|<span data-ttu-id="1daf4-457">要求</span><span class="sxs-lookup"><span data-stu-id="1daf4-457">Requirement</span></span>| <span data-ttu-id="1daf4-458">值</span><span class="sxs-lookup"><span data-stu-id="1daf4-458">Value</span></span>|
|---|---|
|[<span data-ttu-id="1daf4-459">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="1daf4-459">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1daf4-460">1.0</span><span class="sxs-lookup"><span data-stu-id="1daf4-460">1.0</span></span>|
|[<span data-ttu-id="1daf4-461">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="1daf4-461">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1daf4-462">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1daf4-462">ReadItem</span></span>|
|[<span data-ttu-id="1daf4-463">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="1daf4-463">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1daf4-464">阅读</span><span class="sxs-lookup"><span data-stu-id="1daf4-464">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="1daf4-465">示例</span><span class="sxs-lookup"><span data-stu-id="1daf4-465">Example</span></span>

```js
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
```

####  <a name="notificationmessages-notificationmessagesjavascriptapioutlook16officenotificationmessages"></a><span data-ttu-id="1daf4-466">notificationMessages :[NotificationMessages](/javascript/api/outlook_1_6/office.notificationmessages)</span><span class="sxs-lookup"><span data-stu-id="1daf4-466">notificationMessages :[NotificationMessages](/javascript/api/outlook_1_6/office.notificationmessages)</span></span>

<span data-ttu-id="1daf4-467">获取项目的通知邮件。</span><span class="sxs-lookup"><span data-stu-id="1daf4-467">Gets the notification messages for an item.</span></span>

##### <a name="type"></a><span data-ttu-id="1daf4-468">类型：</span><span class="sxs-lookup"><span data-stu-id="1daf4-468">Type:</span></span>

*   [<span data-ttu-id="1daf4-469">NotificationMessages</span><span class="sxs-lookup"><span data-stu-id="1daf4-469">NotificationMessages</span></span>](/javascript/api/outlook_1_6/office.notificationmessages)

##### <a name="requirements"></a><span data-ttu-id="1daf4-470">要求</span><span class="sxs-lookup"><span data-stu-id="1daf4-470">Requirements</span></span>

|<span data-ttu-id="1daf4-471">要求</span><span class="sxs-lookup"><span data-stu-id="1daf4-471">Requirement</span></span>| <span data-ttu-id="1daf4-472">值</span><span class="sxs-lookup"><span data-stu-id="1daf4-472">Value</span></span>|
|---|---|
|[<span data-ttu-id="1daf4-473">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="1daf4-473">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1daf4-474">1.3</span><span class="sxs-lookup"><span data-stu-id="1daf4-474">1.3</span></span>|
|[<span data-ttu-id="1daf4-475">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="1daf4-475">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1daf4-476">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1daf4-476">ReadItem</span></span>|
|[<span data-ttu-id="1daf4-477">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="1daf4-477">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1daf4-478">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="1daf4-478">Compose or read</span></span>|

####  <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlook16officeemailaddressdetailsrecipientsjavascriptapioutlook16officerecipients"></a><span data-ttu-id="1daf4-479">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_6/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="1daf4-479">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_6/office.recipients)</span></span>

<span data-ttu-id="1daf4-480">提供对事件的可选与会者的访问权限。</span><span class="sxs-lookup"><span data-stu-id="1daf4-480">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="1daf4-481">对象的类型和访问级别取决于当前项目的模式。</span><span class="sxs-lookup"><span data-stu-id="1daf4-481">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="1daf4-482">阅读模式</span><span class="sxs-lookup"><span data-stu-id="1daf4-482">Read mode</span></span>

<span data-ttu-id="1daf4-483">`optionalAttendees` 属性返回一个数组，其中包含每个可选与会者的 `EmailAddressDetails` 对象。</span><span class="sxs-lookup"><span data-stu-id="1daf4-483">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="1daf4-484">撰写模式</span><span class="sxs-lookup"><span data-stu-id="1daf4-484">Compose mode</span></span>

<span data-ttu-id="1daf4-485">`optionalAttendees` 属性返回一个 `Recipients` 对象，该对象提供用于获取或更新可选与会者的方法。</span><span class="sxs-lookup"><span data-stu-id="1daf4-485">The `optionalAttendees` property returns a `Recipients` object that provides methods to get and set the optional attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="1daf4-486">类型：</span><span class="sxs-lookup"><span data-stu-id="1daf4-486">Type:</span></span>

*   <span data-ttu-id="1daf4-487">Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="1daf4-487">Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="1daf4-488">要求</span><span class="sxs-lookup"><span data-stu-id="1daf4-488">Requirements</span></span>

|<span data-ttu-id="1daf4-489">要求</span><span class="sxs-lookup"><span data-stu-id="1daf4-489">Requirement</span></span>| <span data-ttu-id="1daf4-490">值</span><span class="sxs-lookup"><span data-stu-id="1daf4-490">Value</span></span>|
|---|---|
|[<span data-ttu-id="1daf4-491">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="1daf4-491">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1daf4-492">1.0</span><span class="sxs-lookup"><span data-stu-id="1daf4-492">1.0</span></span>|
|[<span data-ttu-id="1daf4-493">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="1daf4-493">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1daf4-494">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1daf4-494">ReadItem</span></span>|
|[<span data-ttu-id="1daf4-495">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="1daf4-495">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1daf4-496">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="1daf4-496">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="1daf4-497">示例</span><span class="sxs-lookup"><span data-stu-id="1daf4-497">Example</span></span>

```js
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

#### <a name="organizer-emailaddressdetailsjavascriptapioutlook16officeemailaddressdetails"></a><span data-ttu-id="1daf4-498">organizer :[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="1daf4-498">organizer :[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)</span></span>

<span data-ttu-id="1daf4-p124">获取指定会议的会议组织者的电子邮件地址。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="1daf4-p124">Gets the email address of the meeting organizer for a specified meeting. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="1daf4-501">类型:</span><span class="sxs-lookup"><span data-stu-id="1daf4-501">Type:</span></span>

*   [<span data-ttu-id="1daf4-502">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="1daf4-502">EmailAddressDetails</span></span>](/javascript/api/outlook_1_6/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="1daf4-503">要求</span><span class="sxs-lookup"><span data-stu-id="1daf4-503">Requirements</span></span>

|<span data-ttu-id="1daf4-504">要求</span><span class="sxs-lookup"><span data-stu-id="1daf4-504">Requirement</span></span>| <span data-ttu-id="1daf4-505">值</span><span class="sxs-lookup"><span data-stu-id="1daf4-505">Value</span></span>|
|---|---|
|[<span data-ttu-id="1daf4-506">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="1daf4-506">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1daf4-507">1.0</span><span class="sxs-lookup"><span data-stu-id="1daf4-507">1.0</span></span>|
|[<span data-ttu-id="1daf4-508">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="1daf4-508">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1daf4-509">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1daf4-509">ReadItem</span></span>|
|[<span data-ttu-id="1daf4-510">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="1daf4-510">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1daf4-511">阅读</span><span class="sxs-lookup"><span data-stu-id="1daf4-511">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="1daf4-512">示例</span><span class="sxs-lookup"><span data-stu-id="1daf4-512">Example</span></span>

```js
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
```

####  <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlook16officeemailaddressdetailsrecipientsjavascriptapioutlook16officerecipients"></a><span data-ttu-id="1daf4-513">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_6/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="1daf4-513">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_6/office.recipients)</span></span>

<span data-ttu-id="1daf4-514">提供对事件的必需与会者的访问权限。</span><span class="sxs-lookup"><span data-stu-id="1daf4-514">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="1daf4-515">对象的类型和访问级别取决于当前项目的模式。</span><span class="sxs-lookup"><span data-stu-id="1daf4-515">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="1daf4-516">阅读模式</span><span class="sxs-lookup"><span data-stu-id="1daf4-516">Read mode</span></span>

<span data-ttu-id="1daf4-517">`requiredAttendees` 属性返回一个数组，其中包含每个必需与会者的 `EmailAddressDetails` 对象。</span><span class="sxs-lookup"><span data-stu-id="1daf4-517">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="1daf4-518">撰写模式</span><span class="sxs-lookup"><span data-stu-id="1daf4-518">Compose mode</span></span>

<span data-ttu-id="1daf4-519">`requiredAttendees` 属性返回一个 `Recipients` 对象，该对象提供用于获取或更新必需与会者的方法。</span><span class="sxs-lookup"><span data-stu-id="1daf4-519">The `requiredAttendees` property returns a `Recipients` object that provides methods to get and set the required attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="1daf4-520">类型：</span><span class="sxs-lookup"><span data-stu-id="1daf4-520">Type:</span></span>

*   <span data-ttu-id="1daf4-521">Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_6/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="1daf4-521">Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_6/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="1daf4-522">要求</span><span class="sxs-lookup"><span data-stu-id="1daf4-522">Requirements</span></span>

|<span data-ttu-id="1daf4-523">要求</span><span class="sxs-lookup"><span data-stu-id="1daf4-523">Requirement</span></span>| <span data-ttu-id="1daf4-524">值</span><span class="sxs-lookup"><span data-stu-id="1daf4-524">Value</span></span>|
|---|---|
|[<span data-ttu-id="1daf4-525">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="1daf4-525">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1daf4-526">1.0</span><span class="sxs-lookup"><span data-stu-id="1daf4-526">1.0</span></span>|
|[<span data-ttu-id="1daf4-527">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="1daf4-527">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1daf4-528">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1daf4-528">ReadItem</span></span>|
|[<span data-ttu-id="1daf4-529">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="1daf4-529">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1daf4-530">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="1daf4-530">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="1daf4-531">示例</span><span class="sxs-lookup"><span data-stu-id="1daf4-531">Example</span></span>

```js
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
}
```

#### <a name="sender-emailaddressdetailsjavascriptapioutlook16officeemailaddressdetails"></a><span data-ttu-id="1daf4-532">发件人：[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="1daf4-532">sender :[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)</span></span>

<span data-ttu-id="1daf4-p126">获取电子邮件发件人的电子邮件地址。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="1daf4-p126">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="1daf4-p127">[`from`](#from-emailaddressdetailsjavascriptapioutlook16officeemailaddressdetails) 和 `sender` 属性表示同一个人，邮件由代理人发送的除外。在此情况下，`from` 属性表示代理程序，而 sender 属性表示代理人。</span><span class="sxs-lookup"><span data-stu-id="1daf4-p127">The [`from`](#from-emailaddressdetailsjavascriptapioutlook16officeemailaddressdetails) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="1daf4-537">`sender` 属性中 `EmailAddressDetails` 对象的 `recipientType` 属性为 `undefined`。</span><span class="sxs-lookup"><span data-stu-id="1daf4-537">Note: The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="1daf4-538">类型：</span><span class="sxs-lookup"><span data-stu-id="1daf4-538">Type:</span></span>

*   [<span data-ttu-id="1daf4-539">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="1daf4-539">EmailAddressDetails</span></span>](/javascript/api/outlook_1_6/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="1daf4-540">要求</span><span class="sxs-lookup"><span data-stu-id="1daf4-540">Requirements</span></span>

|<span data-ttu-id="1daf4-541">要求</span><span class="sxs-lookup"><span data-stu-id="1daf4-541">Requirement</span></span>| <span data-ttu-id="1daf4-542">值</span><span class="sxs-lookup"><span data-stu-id="1daf4-542">Value</span></span>|
|---|---|
|[<span data-ttu-id="1daf4-543">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="1daf4-543">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1daf4-544">1.0</span><span class="sxs-lookup"><span data-stu-id="1daf4-544">1.0</span></span>|
|[<span data-ttu-id="1daf4-545">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="1daf4-545">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1daf4-546">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1daf4-546">ReadItem</span></span>|
|[<span data-ttu-id="1daf4-547">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="1daf4-547">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1daf4-548">阅读</span><span class="sxs-lookup"><span data-stu-id="1daf4-548">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="1daf4-549">示例</span><span class="sxs-lookup"><span data-stu-id="1daf4-549">Example</span></span>

```js
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
```

####  <a name="start-datetimejavascriptapioutlook16officetime"></a><span data-ttu-id="1daf4-550">start :Date|[Time](/javascript/api/outlook_1_6/office.time)</span><span class="sxs-lookup"><span data-stu-id="1daf4-550">start :Date|[Time](/javascript/api/outlook_1_6/office.time)</span></span>

<span data-ttu-id="1daf4-551">获取或设置约会开始的日期和时间。</span><span class="sxs-lookup"><span data-stu-id="1daf4-551">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="1daf4-p128">将 `start` 属性表示为协调世界时 (UTC) 的日期和时间值。可使用 [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook16officelocalclienttime) 方法将该值转换为客户端的本地日期和时间。</span><span class="sxs-lookup"><span data-stu-id="1daf4-p128">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook16officelocalclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="1daf4-554">阅读模式</span><span class="sxs-lookup"><span data-stu-id="1daf4-554">Read mode</span></span>

<span data-ttu-id="1daf4-555">`start` 属性返回 `Date` 对象。</span><span class="sxs-lookup"><span data-stu-id="1daf4-555">The `start` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="1daf4-556">撰写模式</span><span class="sxs-lookup"><span data-stu-id="1daf4-556">Compose mode</span></span>

<span data-ttu-id="1daf4-557">`start` 属性返回 `Time` 对象。</span><span class="sxs-lookup"><span data-stu-id="1daf4-557">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="1daf4-558">使用 [`Time.setAsync`](/javascript/api/outlook_1_6/office.time#setasync-datetime--options--callback-) 方法设置开始时间时，应使用 [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) 方法将客户端的本地时间转换为服务器的 UTC。</span><span class="sxs-lookup"><span data-stu-id="1daf4-558">When you use the [`Time.setAsync`](/javascript/api/outlook_1_6/office.time#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="1daf4-559">类型：</span><span class="sxs-lookup"><span data-stu-id="1daf4-559">Type:</span></span>

*   <span data-ttu-id="1daf4-560">Date | [Time](/javascript/api/outlook_1_6/office.time)</span><span class="sxs-lookup"><span data-stu-id="1daf4-560">Date | [Time](/javascript/api/outlook_1_6/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="1daf4-561">要求</span><span class="sxs-lookup"><span data-stu-id="1daf4-561">Requirements</span></span>

|<span data-ttu-id="1daf4-562">要求</span><span class="sxs-lookup"><span data-stu-id="1daf4-562">Requirement</span></span>| <span data-ttu-id="1daf4-563">值</span><span class="sxs-lookup"><span data-stu-id="1daf4-563">Value</span></span>|
|---|---|
|[<span data-ttu-id="1daf4-564">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="1daf4-564">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1daf4-565">1.0</span><span class="sxs-lookup"><span data-stu-id="1daf4-565">1.0</span></span>|
|[<span data-ttu-id="1daf4-566">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="1daf4-566">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1daf4-567">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1daf4-567">ReadItem</span></span>|
|[<span data-ttu-id="1daf4-568">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="1daf4-568">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1daf4-569">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="1daf4-569">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="1daf4-570">示例</span><span class="sxs-lookup"><span data-stu-id="1daf4-570">Example</span></span>

<span data-ttu-id="1daf4-571">以下示例通过使用 `Time` 对象的 [`setAsync`](/javascript/api/outlook_1_6/office.time#setasync-datetime--options--callback-) 方法，设置撰写模式下约会的开始时间。</span><span class="sxs-lookup"><span data-stu-id="1daf4-571">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_6/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

####  <a name="subject-stringsubjectjavascriptapioutlook16officesubject"></a><span data-ttu-id="1daf4-572">subject :String|[Subject](/javascript/api/outlook_1_6/office.subject)</span><span class="sxs-lookup"><span data-stu-id="1daf4-572">subject :String|[Subject](/javascript/api/outlook_1_6/office.subject)</span></span>

<span data-ttu-id="1daf4-573">获取或设置显示在项目的主题字段中的说明。</span><span class="sxs-lookup"><span data-stu-id="1daf4-573">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="1daf4-574">`subject` 属性获取或设置由电子邮件服务器发送项目时的整个主题。</span><span class="sxs-lookup"><span data-stu-id="1daf4-574">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="1daf4-575">阅读模式</span><span class="sxs-lookup"><span data-stu-id="1daf4-575">Read mode</span></span>

<span data-ttu-id="1daf4-p129">`subject` 属性返回一个字符串。使用 [`normalizedSubject`](#normalizedsubject-string) 属性获取不带任何前导前缀（如 `RE:` 和 `FW:`）的主题。</span><span class="sxs-lookup"><span data-stu-id="1daf4-p129">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

```js
var subject = Office.context.mailbox.item.subject;
```

##### <a name="compose-mode"></a><span data-ttu-id="1daf4-578">撰写模式</span><span class="sxs-lookup"><span data-stu-id="1daf4-578">Compose mode</span></span>

<span data-ttu-id="1daf4-579">`subject` 属性返回一个 `Subject` 对象，该对象提供用于获取和设置主题的方法。</span><span class="sxs-lookup"><span data-stu-id="1daf4-579">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```js
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="1daf4-580">类型：</span><span class="sxs-lookup"><span data-stu-id="1daf4-580">Type:</span></span>

*   <span data-ttu-id="1daf4-581">String | [Subject](/javascript/api/outlook_1_6/office.subject)</span><span class="sxs-lookup"><span data-stu-id="1daf4-581">String | [Subject](/javascript/api/outlook_1_6/office.subject)</span></span>

##### <a name="requirements"></a><span data-ttu-id="1daf4-582">要求</span><span class="sxs-lookup"><span data-stu-id="1daf4-582">Requirements</span></span>

|<span data-ttu-id="1daf4-583">要求</span><span class="sxs-lookup"><span data-stu-id="1daf4-583">Requirement</span></span>| <span data-ttu-id="1daf4-584">值</span><span class="sxs-lookup"><span data-stu-id="1daf4-584">Value</span></span>|
|---|---|
|[<span data-ttu-id="1daf4-585">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="1daf4-585">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1daf4-586">1.0</span><span class="sxs-lookup"><span data-stu-id="1daf4-586">1.0</span></span>|
|[<span data-ttu-id="1daf4-587">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="1daf4-587">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1daf4-588">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1daf4-588">ReadItem</span></span>|
|[<span data-ttu-id="1daf4-589">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="1daf4-589">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1daf4-590">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="1daf4-590">Compose or read</span></span>|

####  <a name="to-arrayemailaddressdetailsjavascriptapioutlook16officeemailaddressdetailsrecipientsjavascriptapioutlook16officerecipients"></a><span data-ttu-id="1daf4-591">to :Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_6/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="1daf4-591">to :Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_6/office.recipients)</span></span>

<span data-ttu-id="1daf4-592">提供对邮件的**发送至**行上的收件人的访问权限。</span><span class="sxs-lookup"><span data-stu-id="1daf4-592">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="1daf4-593">对象的类型和访问级别取决于当前项目的模式。</span><span class="sxs-lookup"><span data-stu-id="1daf4-593">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="1daf4-594">阅读模式</span><span class="sxs-lookup"><span data-stu-id="1daf4-594">Read mode</span></span>

<span data-ttu-id="1daf4-p131">`to` 属性返回包含邮件的**收件人**行上所列的每个收件人的 `EmailAddressDetails` 对象的数组。集合上限为 100 个成员。</span><span class="sxs-lookup"><span data-stu-id="1daf4-p131">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="1daf4-597">撰写模式</span><span class="sxs-lookup"><span data-stu-id="1daf4-597">Compose mode</span></span>

<span data-ttu-id="1daf4-598">`to` 属性返回 `Recipients`一个对象，该对象提供用于获取或更新邮件的**发送至**行上收件人的方法。</span><span class="sxs-lookup"><span data-stu-id="1daf4-598">The to`to` property returns a Recipients`Recipients` object that provides methods for manipulating the recipients on the **To** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="1daf4-599">类型：</span><span class="sxs-lookup"><span data-stu-id="1daf4-599">Type:</span></span>

*   <span data-ttu-id="1daf4-600">Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_6/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="1daf4-600">Array.<[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_6/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="1daf4-601">要求</span><span class="sxs-lookup"><span data-stu-id="1daf4-601">Requirements</span></span>

|<span data-ttu-id="1daf4-602">要求</span><span class="sxs-lookup"><span data-stu-id="1daf4-602">Requirement</span></span>| <span data-ttu-id="1daf4-603">值</span><span class="sxs-lookup"><span data-stu-id="1daf4-603">Value</span></span>|
|---|---|
|[<span data-ttu-id="1daf4-604">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="1daf4-604">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1daf4-605">1.0</span><span class="sxs-lookup"><span data-stu-id="1daf4-605">1.0</span></span>|
|[<span data-ttu-id="1daf4-606">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="1daf4-606">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1daf4-607">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1daf4-607">ReadItem</span></span>|
|[<span data-ttu-id="1daf4-608">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="1daf4-608">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1daf4-609">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="1daf4-609">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="1daf4-610">示例</span><span class="sxs-lookup"><span data-stu-id="1daf4-610">Example</span></span>

```js
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

### <a name="methods"></a><span data-ttu-id="1daf4-611">方法</span><span class="sxs-lookup"><span data-stu-id="1daf4-611">Methods</span></span>

####  <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="1daf4-612">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="1daf4-612">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="1daf4-613">将文件作为附件添加到邮件或约会。</span><span class="sxs-lookup"><span data-stu-id="1daf4-613">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="1daf4-614">`addFileAttachmentAsync` 方法在指定的 URI 上载文件并将其附加到撰写窗体中的项目。</span><span class="sxs-lookup"><span data-stu-id="1daf4-614">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="1daf4-615">你随后可以将该标识符与 [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) 方法一同使用，以删除同一个会话中的附件。</span><span class="sxs-lookup"><span data-stu-id="1daf4-615">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="1daf4-616">参数：</span><span class="sxs-lookup"><span data-stu-id="1daf4-616">Parameters:</span></span>

|<span data-ttu-id="1daf4-617">名称</span><span class="sxs-lookup"><span data-stu-id="1daf4-617">Name</span></span>| <span data-ttu-id="1daf4-618">类型</span><span class="sxs-lookup"><span data-stu-id="1daf4-618">Type</span></span>| <span data-ttu-id="1daf4-619">属性</span><span class="sxs-lookup"><span data-stu-id="1daf4-619">Attributes</span></span>| <span data-ttu-id="1daf4-620">说明</span><span class="sxs-lookup"><span data-stu-id="1daf4-620">Description</span></span>|
|---|---|---|---|
|`uri`| <span data-ttu-id="1daf4-621">字符串</span><span class="sxs-lookup"><span data-stu-id="1daf4-621">String</span></span>||<span data-ttu-id="1daf4-p132">提供附加到邮件或约会的文件的位置的 URI。最大长度为 2048 个字符。</span><span class="sxs-lookup"><span data-stu-id="1daf4-p132">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="1daf4-624">字符串</span><span class="sxs-lookup"><span data-stu-id="1daf4-624">String</span></span>||<span data-ttu-id="1daf4-p133">在附件上载过程中显示的附件名称。最大长度为 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="1daf4-p133">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="1daf4-627">Object</span><span class="sxs-lookup"><span data-stu-id="1daf4-627">Object</span></span>| <span data-ttu-id="1daf4-628">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="1daf4-628">&lt;optional&gt;</span></span>|<span data-ttu-id="1daf4-629">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="1daf4-629">An object literal that contains one or more of the following properties.</span></span>|
| `options.asyncContext` | <span data-ttu-id="1daf4-630">对象</span><span class="sxs-lookup"><span data-stu-id="1daf4-630">Object</span></span> | <span data-ttu-id="1daf4-631">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="1daf4-631">&lt;optional&gt;</span></span> | <span data-ttu-id="1daf4-632">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="1daf4-632">Developers can provide any object they wish to access in the callback method.</span></span> |
| `options.isInline` | <span data-ttu-id="1daf4-633">布尔值</span><span class="sxs-lookup"><span data-stu-id="1daf4-633">Boolean</span></span> | <span data-ttu-id="1daf4-634">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="1daf4-634">&lt;optional&gt;</span></span> | <span data-ttu-id="1daf4-635">如果为 `true`，则表示附件将在邮件正文中内联显示，并且不应显示在附件列表中。</span><span class="sxs-lookup"><span data-stu-id="1daf4-635">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
|`callback`| <span data-ttu-id="1daf4-636">函数</span><span class="sxs-lookup"><span data-stu-id="1daf4-636">function</span></span>| <span data-ttu-id="1daf4-637">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="1daf4-637">&lt;optional&gt;</span></span>|<span data-ttu-id="1daf4-638">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="1daf4-638">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="1daf4-639">如果成功，附件标识符将在 `asyncResult.value` 属性中提供。</span><span class="sxs-lookup"><span data-stu-id="1daf4-639">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="1daf4-640">如果上传附件失败，`asyncResult` 对象将包含一个提供错误说明的 `Error` 对象。</span><span class="sxs-lookup"><span data-stu-id="1daf4-640">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="1daf4-641">错误</span><span class="sxs-lookup"><span data-stu-id="1daf4-641">Errors</span></span>

| <span data-ttu-id="1daf4-642">错误代码</span><span class="sxs-lookup"><span data-stu-id="1daf4-642">Error code</span></span> | <span data-ttu-id="1daf4-643">说明</span><span class="sxs-lookup"><span data-stu-id="1daf4-643">Description</span></span> |
|------------|-------------|
| `AttachmentSizeExceeded` | <span data-ttu-id="1daf4-644">附件大小超过了允许的大小。</span><span class="sxs-lookup"><span data-stu-id="1daf4-644">The attachment is larger than allowed.</span></span> |
| `FileTypeNotSupported` | <span data-ttu-id="1daf4-645">该附件的扩展名不是允许的扩展名。</span><span class="sxs-lookup"><span data-stu-id="1daf4-645">The attachment has an extension that is not allowed.</span></span> |
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="1daf4-646">邮件或约会具有的附件过多。</span><span class="sxs-lookup"><span data-stu-id="1daf4-646">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="1daf4-647">要求</span><span class="sxs-lookup"><span data-stu-id="1daf4-647">Requirements</span></span>

|<span data-ttu-id="1daf4-648">要求</span><span class="sxs-lookup"><span data-stu-id="1daf4-648">Requirement</span></span>| <span data-ttu-id="1daf4-649">值</span><span class="sxs-lookup"><span data-stu-id="1daf4-649">Value</span></span>|
|---|---|
|[<span data-ttu-id="1daf4-650">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="1daf4-650">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1daf4-651">1.1</span><span class="sxs-lookup"><span data-stu-id="1daf4-651">1.1</span></span>|
|[<span data-ttu-id="1daf4-652">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="1daf4-652">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1daf4-653">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="1daf4-653">ReadWriteItem</span></span>|
|[<span data-ttu-id="1daf4-654">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="1daf4-654">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1daf4-655">撰写</span><span class="sxs-lookup"><span data-stu-id="1daf4-655">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="1daf4-656">示例</span><span class="sxs-lookup"><span data-stu-id="1daf4-656">Examples</span></span>

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

<span data-ttu-id="1daf4-657">以下示例将图像文件添加为内联附件，并在邮件正文中引用该附件。</span><span class="sxs-lookup"><span data-stu-id="1daf4-657">The following example adds an image file as an inline attachment and references the attachment in the message body.</span></span>

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

####  <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="1daf4-658">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="1daf4-658">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="1daf4-659">将 Exchange 项目（如邮件）作为附件添加到邮件或约会。</span><span class="sxs-lookup"><span data-stu-id="1daf4-659">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="1daf4-p134">`addItemAttachmentAsync` 方法将包含指定 Exchange 标识符的项目附加到撰写窗体中的项目。如果指定一个回调方法，此方法使用 `asyncResult` 参数调用，该参数包含一个附件标识符或代码，指示附加项目过程中出现的任何错误。可以使用 `options` 参数将状态信息传递给回调方法（如果需要）。</span><span class="sxs-lookup"><span data-stu-id="1daf4-p134">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="1daf4-663">随后可以将该标识符与 [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) 方法一同使用，以删除同一个会话中的附件。</span><span class="sxs-lookup"><span data-stu-id="1daf4-663">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="1daf4-664">如果 Office 加载项在 Outlook Web App 中运行，则 `addItemAttachmentAsync` 方法可以将项目附加到项目（正在编辑的项目除外）中；然而，不支持也不建议这样做。</span><span class="sxs-lookup"><span data-stu-id="1daf4-664">If your Office add-in is running in Outlook Web App, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="1daf4-665">参数：</span><span class="sxs-lookup"><span data-stu-id="1daf4-665">Parameters:</span></span>

|<span data-ttu-id="1daf4-666">名称</span><span class="sxs-lookup"><span data-stu-id="1daf4-666">Name</span></span>| <span data-ttu-id="1daf4-667">类型</span><span class="sxs-lookup"><span data-stu-id="1daf4-667">Type</span></span>| <span data-ttu-id="1daf4-668">属性</span><span class="sxs-lookup"><span data-stu-id="1daf4-668">Attributes</span></span>| <span data-ttu-id="1daf4-669">说明</span><span class="sxs-lookup"><span data-stu-id="1daf4-669">Description</span></span>|
|---|---|---|---|
|`itemId`| <span data-ttu-id="1daf4-670">字符串</span><span class="sxs-lookup"><span data-stu-id="1daf4-670">String</span></span>||<span data-ttu-id="1daf4-p135">要附加的项目的 Exchange 标识符。最大长度为 100 个字符。</span><span class="sxs-lookup"><span data-stu-id="1daf4-p135">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="1daf4-673">字符串</span><span class="sxs-lookup"><span data-stu-id="1daf4-673">String</span></span>||<span data-ttu-id="1daf4-p136">要附加的项目的主题。最大长度为 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="1daf4-p136">The sujbect of the item to be attached. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="1daf4-676">Object</span><span class="sxs-lookup"><span data-stu-id="1daf4-676">Object</span></span>| <span data-ttu-id="1daf4-677">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="1daf4-677">&lt;optional&gt;</span></span>|<span data-ttu-id="1daf4-678">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="1daf4-678">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="1daf4-679">对象</span><span class="sxs-lookup"><span data-stu-id="1daf4-679">Object</span></span>| <span data-ttu-id="1daf4-680">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="1daf4-680">&lt;optional&gt;</span></span>|<span data-ttu-id="1daf4-681">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="1daf4-681">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="1daf4-682">函数</span><span class="sxs-lookup"><span data-stu-id="1daf4-682">function</span></span>| <span data-ttu-id="1daf4-683">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="1daf4-683">&lt;optional&gt;</span></span>|<span data-ttu-id="1daf4-684">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="1daf4-684">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="1daf4-685">如果成功，附件标识符将在 `asyncResult.value` 属性中提供。</span><span class="sxs-lookup"><span data-stu-id="1daf4-685">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="1daf4-686">如果添加附件失败，`asyncResult` 对象将包含一个提供错误说明的 `Error` 对象。</span><span class="sxs-lookup"><span data-stu-id="1daf4-686">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="1daf4-687">错误</span><span class="sxs-lookup"><span data-stu-id="1daf4-687">Errors</span></span>

| <span data-ttu-id="1daf4-688">错误代码</span><span class="sxs-lookup"><span data-stu-id="1daf4-688">Error code</span></span> | <span data-ttu-id="1daf4-689">说明</span><span class="sxs-lookup"><span data-stu-id="1daf4-689">Description</span></span> |
|------------|-------------|
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="1daf4-690">邮件或约会具有的附件过多。</span><span class="sxs-lookup"><span data-stu-id="1daf4-690">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="1daf4-691">要求</span><span class="sxs-lookup"><span data-stu-id="1daf4-691">Requirements</span></span>

|<span data-ttu-id="1daf4-692">要求</span><span class="sxs-lookup"><span data-stu-id="1daf4-692">Requirement</span></span>| <span data-ttu-id="1daf4-693">值</span><span class="sxs-lookup"><span data-stu-id="1daf4-693">Value</span></span>|
|---|---|
|[<span data-ttu-id="1daf4-694">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="1daf4-694">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1daf4-695">1.1</span><span class="sxs-lookup"><span data-stu-id="1daf4-695">1.1</span></span>|
|[<span data-ttu-id="1daf4-696">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="1daf4-696">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1daf4-697">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="1daf4-697">ReadWriteItem</span></span>|
|[<span data-ttu-id="1daf4-698">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="1daf4-698">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1daf4-699">撰写</span><span class="sxs-lookup"><span data-stu-id="1daf4-699">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="1daf4-700">示例</span><span class="sxs-lookup"><span data-stu-id="1daf4-700">Example</span></span>

<span data-ttu-id="1daf4-701">以下示例将现有的 Outlook 项目添加为名为 `My Attachment` 的附件。</span><span class="sxs-lookup"><span data-stu-id="1daf4-701">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

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

####  <a name="close"></a><span data-ttu-id="1daf4-702">close()</span><span class="sxs-lookup"><span data-stu-id="1daf4-702">close()</span></span>

<span data-ttu-id="1daf4-703">关闭当前正在撰写的项目。</span><span class="sxs-lookup"><span data-stu-id="1daf4-703">Closes the current item that is being composed.</span></span>

<span data-ttu-id="1daf4-p137">
            \`close\` 方法的行为取决于要撰写的项目的当前状态。如果项目具有未保存的更改，客户端将提示用户保存、放弃或取消关闭操作。</span><span class="sxs-lookup"><span data-stu-id="1daf4-p137">The behavior of the `close` method depends on the current state of the item being composed. If the item has unsaved changes, the client prompts the user to save, discard, or cancel the close action.</span></span>

> [!NOTE]
> <span data-ttu-id="1daf4-706">在 Outlook 网页版中，如果该项目是约会并且之前已使用 `saveAsync` 保存，则即使自上次保存项目后未发生任何更改，也会提示用户保存、放弃或取消。</span><span class="sxs-lookup"><span data-stu-id="1daf4-706">In Outlook on the web, if the item is an appointment and it has previously been saved using `saveAsync`, the user is prompted to save, discard, or cancel even if no changes have occurred since the item was last saved.</span></span>

<span data-ttu-id="1daf4-707">在 Outlook 桌面客户端中，如果邮件是内联答复，`close` 方法不起作用。</span><span class="sxs-lookup"><span data-stu-id="1daf4-707">In the Outlook desktop client, if the message is an inline reply, the `close` method has no effect.</span></span>

##### <a name="requirements"></a><span data-ttu-id="1daf4-708">要求</span><span class="sxs-lookup"><span data-stu-id="1daf4-708">Requirements</span></span>

|<span data-ttu-id="1daf4-709">要求</span><span class="sxs-lookup"><span data-stu-id="1daf4-709">Requirement</span></span>| <span data-ttu-id="1daf4-710">值</span><span class="sxs-lookup"><span data-stu-id="1daf4-710">Value</span></span>|
|---|---|
|[<span data-ttu-id="1daf4-711">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="1daf4-711">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1daf4-712">1.3</span><span class="sxs-lookup"><span data-stu-id="1daf4-712">1.3</span></span>|
|[<span data-ttu-id="1daf4-713">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="1daf4-713">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1daf4-714">受限</span><span class="sxs-lookup"><span data-stu-id="1daf4-714">Restricted</span></span>|
|[<span data-ttu-id="1daf4-715">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="1daf4-715">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1daf4-716">撰写</span><span class="sxs-lookup"><span data-stu-id="1daf4-716">Compose</span></span>|

#### <a name="displayreplyallformformdata"></a><span data-ttu-id="1daf4-717">displayReplyAllForm(formData)</span><span class="sxs-lookup"><span data-stu-id="1daf4-717">displayReplyAllForm(formData)</span></span>

<span data-ttu-id="1daf4-718">显示答复窗体，其中包括所选邮件的发件人和所有收件人或所选约会的组织者和所有与会者。</span><span class="sxs-lookup"><span data-stu-id="1daf4-718">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="1daf4-719">在 Outlook for iOS 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="1daf4-719">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="1daf4-720">在 Outlook Web App 中，答复窗体显示为包含 3 列视图的弹出式窗体以及包含 2 列或 1 列视图的弹出式窗体。</span><span class="sxs-lookup"><span data-stu-id="1daf4-720">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="1daf4-721">如果任意字符串参数超出其限制，`displayReplyAllForm` 将引发异常。</span><span class="sxs-lookup"><span data-stu-id="1daf4-721">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="1daf4-p138">当在 `formData.attachments` 参数中指定附件时，Outlook 和 Outlook Web App 尝试下载所有附件并将其附加到答复窗体。如果无法添加任何附件，则在窗体 UI 中显示错误。如果这不可能，则不引发错误消息。</span><span class="sxs-lookup"><span data-stu-id="1daf4-p138">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="1daf4-725">参数：</span><span class="sxs-lookup"><span data-stu-id="1daf4-725">Parameters:</span></span>

| <span data-ttu-id="1daf4-726">名称</span><span class="sxs-lookup"><span data-stu-id="1daf4-726">Name</span></span> | <span data-ttu-id="1daf4-727">类型</span><span class="sxs-lookup"><span data-stu-id="1daf4-727">Type</span></span> | <span data-ttu-id="1daf4-728">属性</span><span class="sxs-lookup"><span data-stu-id="1daf4-728">Attributes</span></span> | <span data-ttu-id="1daf4-729">说明</span><span class="sxs-lookup"><span data-stu-id="1daf4-729">Description</span></span> |
|---|---|---|---|
|`formData`| <span data-ttu-id="1daf4-730">字符串 &#124; 对象</span><span class="sxs-lookup"><span data-stu-id="1daf4-730">String &#124; Object</span></span>| |<span data-ttu-id="1daf4-p139">一个包含文本和 HTML 且表示答复窗体的正文的字符串。字符串限制为 32 KB。</span><span class="sxs-lookup"><span data-stu-id="1daf4-p139">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="1daf4-733">**或**</span><span class="sxs-lookup"><span data-stu-id="1daf4-733">**OR**</span></span><br/><span data-ttu-id="1daf4-p140">包含正文或附件数据和回调函数的对象。对象定义如下。</span><span class="sxs-lookup"><span data-stu-id="1daf4-p140">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="1daf4-736">字符串</span><span class="sxs-lookup"><span data-stu-id="1daf4-736">String</span></span> | <span data-ttu-id="1daf4-737">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="1daf4-737">&lt;optional&gt;</span></span> | <span data-ttu-id="1daf4-p141">一个包含文本和 HTML 且表示答复窗体的正文的字符串。字符串限制为 32 KB。</span><span class="sxs-lookup"><span data-stu-id="1daf4-p141">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="1daf4-740">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="1daf4-740">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="1daf4-741">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="1daf4-741">&lt;optional&gt;</span></span> | <span data-ttu-id="1daf4-742">JSON 对象（是文件或项目附件）的数组。</span><span class="sxs-lookup"><span data-stu-id="1daf4-742">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="1daf4-743">字符串</span><span class="sxs-lookup"><span data-stu-id="1daf4-743">String</span></span> | | <span data-ttu-id="1daf4-p142">指示附件的类型。必须是文件附件的 `file` 或项目附件的 `item`。</span><span class="sxs-lookup"><span data-stu-id="1daf4-p142">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="1daf4-746">字符串</span><span class="sxs-lookup"><span data-stu-id="1daf4-746">String</span></span> | | <span data-ttu-id="1daf4-747">一个包含附件的名称的字符串，最多包含 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="1daf4-747">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="1daf4-748">String</span><span class="sxs-lookup"><span data-stu-id="1daf4-748">String</span></span> | | <span data-ttu-id="1daf4-p143">仅在将 `type` 设置为 `file` 时使用。文件的位置的 URI。</span><span class="sxs-lookup"><span data-stu-id="1daf4-p143">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.isInline` | <span data-ttu-id="1daf4-751">布尔</span><span class="sxs-lookup"><span data-stu-id="1daf4-751">Boolean</span></span> | | <span data-ttu-id="1daf4-p144">仅在将 `type` 设置为 `file` 时使用。如果为 `true`，则表示附件将在邮件正文中内联显示，并且不应显示在附件列表中。</span><span class="sxs-lookup"><span data-stu-id="1daf4-p144">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="1daf4-754">String</span><span class="sxs-lookup"><span data-stu-id="1daf4-754">String</span></span> | | <span data-ttu-id="1daf4-p145">仅在将 `type` 设置为 `item` 时使用。附件的 EWS 项目 ID。字符串最长为 100 个字符。</span><span class="sxs-lookup"><span data-stu-id="1daf4-p145">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="1daf4-758">函数</span><span class="sxs-lookup"><span data-stu-id="1daf4-758">function</span></span> | <span data-ttu-id="1daf4-759">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="1daf4-759">&lt;optional&gt;</span></span> | <span data-ttu-id="1daf4-760">方法完成后，使用单个参数 `asyncResult`（一个 [AsyncResult](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="1daf4-760">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="1daf4-761">要求</span><span class="sxs-lookup"><span data-stu-id="1daf4-761">Requirements</span></span>

|<span data-ttu-id="1daf4-762">要求</span><span class="sxs-lookup"><span data-stu-id="1daf4-762">Requirement</span></span>| <span data-ttu-id="1daf4-763">值</span><span class="sxs-lookup"><span data-stu-id="1daf4-763">Value</span></span>|
|---|---|
|[<span data-ttu-id="1daf4-764">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="1daf4-764">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1daf4-765">1.0</span><span class="sxs-lookup"><span data-stu-id="1daf4-765">1.0</span></span>|
|[<span data-ttu-id="1daf4-766">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="1daf4-766">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1daf4-767">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1daf4-767">ReadItem</span></span>|
|[<span data-ttu-id="1daf4-768">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="1daf4-768">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1daf4-769">阅读</span><span class="sxs-lookup"><span data-stu-id="1daf4-769">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="1daf4-770">示例</span><span class="sxs-lookup"><span data-stu-id="1daf4-770">Examples</span></span>

<span data-ttu-id="1daf4-771">以下代码将一个字符串传递到 `displayReplyAllForm` 函数。</span><span class="sxs-lookup"><span data-stu-id="1daf4-771">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="1daf4-772">使用空白正文答复。</span><span class="sxs-lookup"><span data-stu-id="1daf4-772">Reply with an empty body.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="1daf4-773">仅使用正文答复。</span><span class="sxs-lookup"><span data-stu-id="1daf4-773">Reply with just a body.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="1daf4-774">使用正文和文件附件答复。</span><span class="sxs-lookup"><span data-stu-id="1daf4-774">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="1daf4-775">使用正文和项目附件答复。</span><span class="sxs-lookup"><span data-stu-id="1daf4-775">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="1daf4-776">使用正文、文件附件、项目附件和回调答复。</span><span class="sxs-lookup"><span data-stu-id="1daf4-776">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="displayreplyformformdata"></a><span data-ttu-id="1daf4-777">displayReplyForm(formData)</span><span class="sxs-lookup"><span data-stu-id="1daf4-777">displayReplyForm(formData)</span></span>

<span data-ttu-id="1daf4-778">显示答复窗体，其中仅包括所选邮件的发件人或所选约会的组织者。</span><span class="sxs-lookup"><span data-stu-id="1daf4-778">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="1daf4-779">在 Outlook for iOS 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="1daf4-779">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="1daf4-780">在 Outlook Web App 中，答复窗体显示为包含 3 列视图的弹出式窗体以及包含 2 列或 1 列视图的弹出式窗体。</span><span class="sxs-lookup"><span data-stu-id="1daf4-780">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="1daf4-781">如果任意字符串参数超出其限制，`displayReplyForm` 将引发异常。</span><span class="sxs-lookup"><span data-stu-id="1daf4-781">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="1daf4-p146">当在 `formData.attachments` 参数中指定附件时，Outlook 和 Outlook Web App 尝试下载所有附件并将其附加到答复窗体。如果无法添加任何附件，则在窗体 UI 中显示错误。如果这不可能，则不引发错误消息。</span><span class="sxs-lookup"><span data-stu-id="1daf4-p146">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="1daf4-785">参数：</span><span class="sxs-lookup"><span data-stu-id="1daf4-785">Parameters:</span></span>

| <span data-ttu-id="1daf4-786">名称</span><span class="sxs-lookup"><span data-stu-id="1daf4-786">Name</span></span> | <span data-ttu-id="1daf4-787">类型</span><span class="sxs-lookup"><span data-stu-id="1daf4-787">Type</span></span> | <span data-ttu-id="1daf4-788">属性</span><span class="sxs-lookup"><span data-stu-id="1daf4-788">Attributes</span></span> | <span data-ttu-id="1daf4-789">说明</span><span class="sxs-lookup"><span data-stu-id="1daf4-789">Description</span></span> |
|---|---|---|---|
|`formData`| <span data-ttu-id="1daf4-790">字符串 &#124; 对象</span><span class="sxs-lookup"><span data-stu-id="1daf4-790">String &#124; Object</span></span>| | <span data-ttu-id="1daf4-p147">一个包含文本和 HTML 且表示答复窗体的正文的字符串。字符串限制为 32 KB。</span><span class="sxs-lookup"><span data-stu-id="1daf4-p147">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="1daf4-793">**或**</span><span class="sxs-lookup"><span data-stu-id="1daf4-793">**OR**</span></span><br/><span data-ttu-id="1daf4-p148">包含正文或附件数据和回调函数的对象。对象定义如下。</span><span class="sxs-lookup"><span data-stu-id="1daf4-p148">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="1daf4-796">字符串</span><span class="sxs-lookup"><span data-stu-id="1daf4-796">String</span></span> | <span data-ttu-id="1daf4-797">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="1daf4-797">&lt;optional&gt;</span></span> | <span data-ttu-id="1daf4-p149">一个包含文本和 HTML 且表示答复窗体的正文的字符串。字符串限制为 32 KB。</span><span class="sxs-lookup"><span data-stu-id="1daf4-p149">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="1daf4-800">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="1daf4-800">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="1daf4-801">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="1daf4-801">&lt;optional&gt;</span></span> | <span data-ttu-id="1daf4-802">JSON 对象（是文件或项目附件）的数组。</span><span class="sxs-lookup"><span data-stu-id="1daf4-802">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="1daf4-803">字符串</span><span class="sxs-lookup"><span data-stu-id="1daf4-803">String</span></span> | | <span data-ttu-id="1daf4-p150">指示附件的类型。必须是文件附件的 `file` 或项目附件的 `item`。</span><span class="sxs-lookup"><span data-stu-id="1daf4-p150">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="1daf4-806">字符串</span><span class="sxs-lookup"><span data-stu-id="1daf4-806">String</span></span> | | <span data-ttu-id="1daf4-807">一个包含附件的名称的字符串，最多包含 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="1daf4-807">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="1daf4-808">String</span><span class="sxs-lookup"><span data-stu-id="1daf4-808">String</span></span> | | <span data-ttu-id="1daf4-p151">仅在将 `type` 设置为 `file` 时使用。文件的位置的 URI。</span><span class="sxs-lookup"><span data-stu-id="1daf4-p151">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.isInline` | <span data-ttu-id="1daf4-811">布尔</span><span class="sxs-lookup"><span data-stu-id="1daf4-811">Boolean</span></span> | | <span data-ttu-id="1daf4-p152">仅在将 `type` 设置为 `file` 时使用。如果为 `true`，则表示附件将在邮件正文中内联显示，并且不应显示在附件列表中。</span><span class="sxs-lookup"><span data-stu-id="1daf4-p152">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="1daf4-814">String</span><span class="sxs-lookup"><span data-stu-id="1daf4-814">String</span></span> | | <span data-ttu-id="1daf4-p153">仅在将 `type` 设置为 `item` 时使用。附件的 EWS 项目 ID。字符串最长为 100 个字符。</span><span class="sxs-lookup"><span data-stu-id="1daf4-p153">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="1daf4-818">函数</span><span class="sxs-lookup"><span data-stu-id="1daf4-818">function</span></span> | <span data-ttu-id="1daf4-819">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="1daf4-819">&lt;optional&gt;</span></span> | <span data-ttu-id="1daf4-820">方法完成后，使用单个参数 `asyncResult`（一个 [AsyncResult](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="1daf4-820">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="1daf4-821">要求</span><span class="sxs-lookup"><span data-stu-id="1daf4-821">Requirements</span></span>

|<span data-ttu-id="1daf4-822">要求</span><span class="sxs-lookup"><span data-stu-id="1daf4-822">Requirement</span></span>| <span data-ttu-id="1daf4-823">值</span><span class="sxs-lookup"><span data-stu-id="1daf4-823">Value</span></span>|
|---|---|
|[<span data-ttu-id="1daf4-824">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="1daf4-824">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1daf4-825">1.0</span><span class="sxs-lookup"><span data-stu-id="1daf4-825">1.0</span></span>|
|[<span data-ttu-id="1daf4-826">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="1daf4-826">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1daf4-827">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1daf4-827">ReadItem</span></span>|
|[<span data-ttu-id="1daf4-828">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="1daf4-828">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1daf4-829">阅读</span><span class="sxs-lookup"><span data-stu-id="1daf4-829">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="1daf4-830">示例</span><span class="sxs-lookup"><span data-stu-id="1daf4-830">Examples</span></span>

<span data-ttu-id="1daf4-831">以下代码将一个字符串传递到 `displayReplyForm` 函数。</span><span class="sxs-lookup"><span data-stu-id="1daf4-831">The following code passes a string to the `displayReplyForm` function.</span></span>

```js
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="1daf4-832">使用空白正文答复。</span><span class="sxs-lookup"><span data-stu-id="1daf4-832">Reply with an empty body.</span></span>

```js
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="1daf4-833">仅使用正文答复。</span><span class="sxs-lookup"><span data-stu-id="1daf4-833">Reply with just a body.</span></span>

```js
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="1daf4-834">使用正文和文件附件答复。</span><span class="sxs-lookup"><span data-stu-id="1daf4-834">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="1daf4-835">使用正文和项目附件答复。</span><span class="sxs-lookup"><span data-stu-id="1daf4-835">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="1daf4-836">使用正文、文件附件、项目附件和回调答复。</span><span class="sxs-lookup"><span data-stu-id="1daf4-836">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="getentities--entitiesjavascriptapioutlook16officeentities"></a><span data-ttu-id="1daf4-837">getEntities() → {[Entities](/javascript/api/outlook_1_6/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="1daf4-837">getEntities() → {[Entities](/javascript/api/outlook_1_6/office.entities)}</span></span>

<span data-ttu-id="1daf4-838">获取在所选项目的正文中找到的实体。</span><span class="sxs-lookup"><span data-stu-id="1daf4-838">Gets the entities found in the selected item.</span></span>

> [!NOTE]
> <span data-ttu-id="1daf4-839">在 Outlook for iOS 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="1daf4-839">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="1daf4-840">要求</span><span class="sxs-lookup"><span data-stu-id="1daf4-840">Requirements</span></span>

|<span data-ttu-id="1daf4-841">要求</span><span class="sxs-lookup"><span data-stu-id="1daf4-841">Requirement</span></span>| <span data-ttu-id="1daf4-842">值</span><span class="sxs-lookup"><span data-stu-id="1daf4-842">Value</span></span>|
|---|---|
|[<span data-ttu-id="1daf4-843">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="1daf4-843">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1daf4-844">1.0</span><span class="sxs-lookup"><span data-stu-id="1daf4-844">1.0</span></span>|
|[<span data-ttu-id="1daf4-845">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="1daf4-845">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1daf4-846">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1daf4-846">ReadItem</span></span>|
|[<span data-ttu-id="1daf4-847">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="1daf4-847">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1daf4-848">阅读</span><span class="sxs-lookup"><span data-stu-id="1daf4-848">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="1daf4-849">返回：</span><span class="sxs-lookup"><span data-stu-id="1daf4-849">Returns:</span></span>

<span data-ttu-id="1daf4-850">类型：[Entities](/javascript/api/outlook_1_6/office.entities)</span><span class="sxs-lookup"><span data-stu-id="1daf4-850">Type: [Entities](/javascript/api/outlook_1_6/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="1daf4-851">示例</span><span class="sxs-lookup"><span data-stu-id="1daf4-851">Example</span></span>

<span data-ttu-id="1daf4-852">以下示例访问当前项目的正文中的联系人实体。</span><span class="sxs-lookup"><span data-stu-id="1daf4-852">The following example accesses the contacts entities on the current item.</span></span>

```js
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlook16officecontactmeetingsuggestionjavascriptapioutlook16officemeetingsuggestionphonenumberjavascriptapioutlook16officephonenumbertasksuggestionjavascriptapioutlook16officetasksuggestion"></a><span data-ttu-id="1daf4-853">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_6/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_6/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_6/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_6/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="1daf4-853">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_6/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_6/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_6/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_6/office.tasksuggestion))>}</span></span>

<span data-ttu-id="1daf4-854">获取所选项目的正文中找到的指定实体类型的所有实体的数组。</span><span class="sxs-lookup"><span data-stu-id="1daf4-854">Gets an array of all the entities of the specified entity type found in the selected item.</span></span>

> [!NOTE]
> <span data-ttu-id="1daf4-855">在 Outlook for iOS 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="1daf4-855">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="1daf4-856">参数：</span><span class="sxs-lookup"><span data-stu-id="1daf4-856">Parameters:</span></span>

|<span data-ttu-id="1daf4-857">名称</span><span class="sxs-lookup"><span data-stu-id="1daf4-857">Name</span></span>| <span data-ttu-id="1daf4-858">类型</span><span class="sxs-lookup"><span data-stu-id="1daf4-858">Type</span></span>| <span data-ttu-id="1daf4-859">说明</span><span class="sxs-lookup"><span data-stu-id="1daf4-859">Description</span></span>|
|---|---|---|
|`entityType`| [<span data-ttu-id="1daf4-860">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="1daf4-860">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook_1_6/office.mailboxenums.entitytype)|<span data-ttu-id="1daf4-861">EntityType 枚举值之一。</span><span class="sxs-lookup"><span data-stu-id="1daf4-861">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="1daf4-862">要求</span><span class="sxs-lookup"><span data-stu-id="1daf4-862">Requirements</span></span>

|<span data-ttu-id="1daf4-863">要求</span><span class="sxs-lookup"><span data-stu-id="1daf4-863">Requirement</span></span>| <span data-ttu-id="1daf4-864">值</span><span class="sxs-lookup"><span data-stu-id="1daf4-864">Value</span></span>|
|---|---|
|[<span data-ttu-id="1daf4-865">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="1daf4-865">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1daf4-866">1.0</span><span class="sxs-lookup"><span data-stu-id="1daf4-866">1.0</span></span>|
|[<span data-ttu-id="1daf4-867">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="1daf4-867">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1daf4-868">受限</span><span class="sxs-lookup"><span data-stu-id="1daf4-868">Restricted</span></span>|
|[<span data-ttu-id="1daf4-869">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="1daf4-869">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1daf4-870">阅读</span><span class="sxs-lookup"><span data-stu-id="1daf4-870">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="1daf4-871">返回：</span><span class="sxs-lookup"><span data-stu-id="1daf4-871">Returns:</span></span>

<span data-ttu-id="1daf4-872">如果在 `entityType` 中传递的值不是 `EntityType` 枚举的有效成员，该方法返回 null。</span><span class="sxs-lookup"><span data-stu-id="1daf4-872">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="1daf4-873">如果指定类型的任何实体都不存在于该项目的正文中，该方法将返回空数组。</span><span class="sxs-lookup"><span data-stu-id="1daf4-873">If no entities of the specified type are present on the item, the method returns an empty array.</span></span> <span data-ttu-id="1daf4-874">否则，返回的数组中对象的类型取决于 `entityType` 参数中请求实体的类型。</span><span class="sxs-lookup"><span data-stu-id="1daf4-874">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="1daf4-875">当使用此方法的最低权限级别**受限**时，某些实体类型需要 **ReadItem** 才能进行访问，如下表中所指定。</span><span class="sxs-lookup"><span data-stu-id="1daf4-875">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

| <span data-ttu-id="1daf4-876">`entityType` 的值</span><span class="sxs-lookup"><span data-stu-id="1daf4-876">Value of `entityType`</span></span> | <span data-ttu-id="1daf4-877">返回的数组中对象的类型</span><span class="sxs-lookup"><span data-stu-id="1daf4-877">Type of objects in returned array</span></span> | <span data-ttu-id="1daf4-878">所需权限级别</span><span class="sxs-lookup"><span data-stu-id="1daf4-878">Required Permission Level</span></span> |
| --- | --- | --- |
| `Address` | <span data-ttu-id="1daf4-879">String</span><span class="sxs-lookup"><span data-stu-id="1daf4-879">String</span></span> | <span data-ttu-id="1daf4-880">**受限**</span><span class="sxs-lookup"><span data-stu-id="1daf4-880">**Restricted**</span></span> |
| `Contact` | <span data-ttu-id="1daf4-881">Contact</span><span class="sxs-lookup"><span data-stu-id="1daf4-881">Contact</span></span> | <span data-ttu-id="1daf4-882">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="1daf4-882">**ReadItem**</span></span> |
| `EmailAddress` | <span data-ttu-id="1daf4-883">String</span><span class="sxs-lookup"><span data-stu-id="1daf4-883">String</span></span> | <span data-ttu-id="1daf4-884">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="1daf4-884">**ReadItem**</span></span> |
| `MeetingSuggestion` | <span data-ttu-id="1daf4-885">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="1daf4-885">MeetingSuggestion</span></span> | <span data-ttu-id="1daf4-886">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="1daf4-886">**ReadItem**</span></span> |
| `PhoneNumber` | <span data-ttu-id="1daf4-887">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="1daf4-887">PhoneNumber</span></span> | <span data-ttu-id="1daf4-888">**受限**</span><span class="sxs-lookup"><span data-stu-id="1daf4-888">**Restricted**</span></span> |
| `TaskSuggestion` | <span data-ttu-id="1daf4-889">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="1daf4-889">TaskSuggestion</span></span> | <span data-ttu-id="1daf4-890">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="1daf4-890">**ReadItem**</span></span> |
| `URL` | <span data-ttu-id="1daf4-891">String</span><span class="sxs-lookup"><span data-stu-id="1daf4-891">String</span></span> | <span data-ttu-id="1daf4-892">**受限**</span><span class="sxs-lookup"><span data-stu-id="1daf4-892">**Restricted**</span></span> |

<span data-ttu-id="1daf4-893">类型：Array.<(String|[Contact](/javascript/api/outlook_1_6/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_6/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_6/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_6/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="1daf4-893">Type: Array.<(String|[Contact](/javascript/api/outlook_1_6/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_6/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_6/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_6/office.tasksuggestion))></span></span>

##### <a name="example"></a><span data-ttu-id="1daf4-894">示例</span><span class="sxs-lookup"><span data-stu-id="1daf4-894">Example</span></span>

<span data-ttu-id="1daf4-895">以下示例显示了如何访问表示当前项目的正文中的邮政地址的字符串数组。</span><span class="sxs-lookup"><span data-stu-id="1daf4-895">The following example shows how to access an array of strings that represent postal addresses in the subject or body of the current item.</span></span>

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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlook16officecontactmeetingsuggestionjavascriptapioutlook16officemeetingsuggestionphonenumberjavascriptapioutlook16officephonenumbertasksuggestionjavascriptapioutlook16officetasksuggestion"></a><span data-ttu-id="1daf4-896">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_6/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_6/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_6/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_6/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="1daf4-896">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_6/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_6/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_6/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_6/office.tasksuggestion))>}</span></span>

<span data-ttu-id="1daf4-897">返回传递清单 XML 文件中定义的命名筛选器的所选项目中的已知实体。</span><span class="sxs-lookup"><span data-stu-id="1daf4-897">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="1daf4-898">在 Outlook for iOS 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="1daf4-898">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="1daf4-899">`getFilteredEntitiesByName` 方法返回匹配在具有指定 `FilterName` 元素值的清单 XML 文件中的 [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) 规则元素中定义的正则表达式的实体。</span><span class="sxs-lookup"><span data-stu-id="1daf4-899">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="1daf4-900">参数：</span><span class="sxs-lookup"><span data-stu-id="1daf4-900">Parameters:</span></span>

|<span data-ttu-id="1daf4-901">名称</span><span class="sxs-lookup"><span data-stu-id="1daf4-901">Name</span></span>| <span data-ttu-id="1daf4-902">类型</span><span class="sxs-lookup"><span data-stu-id="1daf4-902">Type</span></span>| <span data-ttu-id="1daf4-903">描述</span><span class="sxs-lookup"><span data-stu-id="1daf4-903">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="1daf4-904">字符串</span><span class="sxs-lookup"><span data-stu-id="1daf4-904">String</span></span>|<span data-ttu-id="1daf4-905">定义筛选器匹配的 `ItemHasKnownEntity` 规则元素的名称。</span><span class="sxs-lookup"><span data-stu-id="1daf4-905">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="1daf4-906">要求</span><span class="sxs-lookup"><span data-stu-id="1daf4-906">Requirements</span></span>

|<span data-ttu-id="1daf4-907">要求</span><span class="sxs-lookup"><span data-stu-id="1daf4-907">Requirement</span></span>| <span data-ttu-id="1daf4-908">值</span><span class="sxs-lookup"><span data-stu-id="1daf4-908">Value</span></span>|
|---|---|
|[<span data-ttu-id="1daf4-909">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="1daf4-909">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1daf4-910">1.0</span><span class="sxs-lookup"><span data-stu-id="1daf4-910">1.0</span></span>|
|[<span data-ttu-id="1daf4-911">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="1daf4-911">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1daf4-912">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1daf4-912">ReadItem</span></span>|
|[<span data-ttu-id="1daf4-913">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="1daf4-913">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1daf4-914">阅读</span><span class="sxs-lookup"><span data-stu-id="1daf4-914">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="1daf4-915">返回：</span><span class="sxs-lookup"><span data-stu-id="1daf4-915">Returns:</span></span>

<span data-ttu-id="1daf4-p155">如果具有匹配 `name` 参数的 `FilterName` 元素值的清单中没有任何 `ItemHasKnownEntity` 元素，则该方法返回 `null`。如果 `name` 参数匹配清单中的 `ItemHasKnownEntity` 元素，但在匹配的当前项目中没有实体，则该方法返回一个空数组。</span><span class="sxs-lookup"><span data-stu-id="1daf4-p155">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="1daf4-918">类型：Array.<(String|[Contact](/javascript/api/outlook_1_6/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_6/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_6/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_6/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="1daf4-918">Type: Array.<(String|[Contact](/javascript/api/outlook_1_6/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_6/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_6/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_6/office.tasksuggestion))></span></span>

#### <a name="getregexmatches--object"></a><span data-ttu-id="1daf4-919">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="1daf4-919">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="1daf4-920">返回所选项目中匹配在清单 XML 文件中定义的正则表达式的字符串值。</span><span class="sxs-lookup"><span data-stu-id="1daf4-920">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="1daf4-921">在 Outlook for iOS 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="1daf4-921">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="1daf4-p156">`getRegExMatches` 方法返回匹配在清单 XML 文件中的每个 `ItemHasRegularExpressionMatch` 或 `ItemHasKnownEntity` 规则元素中定义的正则表达式的字符串。对于 `ItemHasRegularExpressionMatch` 规则，匹配字符串必须发生在该规则指定的项目的属性中。`PropertyName` 简单类型定义支持的属性。</span><span class="sxs-lookup"><span data-stu-id="1daf4-p156">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="1daf4-925">例如，考虑一个外接程序清单具有以下 `Rule` 元素：</span><span class="sxs-lookup"><span data-stu-id="1daf4-925">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="1daf4-926">从 `getRegExMatches` 返回的对象应有两个属性：`fruits` 和 `veggies`。</span><span class="sxs-lookup"><span data-stu-id="1daf4-926">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="1daf4-p157">如果在项目的正文属性上指定 `ItemHasRegularExpressionMatch` 规则，则正则表达式应进一步筛选正文，不应尝试返回该项目的整个正文。使用正则表达式（如 `.*`）获取项目的整个正文并不总是返回预期的结果。而是使用 [`Body.getAsync`](/javascript/api/outlook_1_6/office.body#getasync-coerciontype--options--callback-) 方法检索整个正文。</span><span class="sxs-lookup"><span data-stu-id="1daf4-p157">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook_1_6/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="1daf4-930">Requirements</span><span class="sxs-lookup"><span data-stu-id="1daf4-930">Requirements</span></span>

|<span data-ttu-id="1daf4-931">要求</span><span class="sxs-lookup"><span data-stu-id="1daf4-931">Requirement</span></span>| <span data-ttu-id="1daf4-932">值</span><span class="sxs-lookup"><span data-stu-id="1daf4-932">Value</span></span>|
|---|---|
|[<span data-ttu-id="1daf4-933">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="1daf4-933">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1daf4-934">1.0</span><span class="sxs-lookup"><span data-stu-id="1daf4-934">1.0</span></span>|
|[<span data-ttu-id="1daf4-935">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="1daf4-935">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1daf4-936">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1daf4-936">ReadItem</span></span>|
|[<span data-ttu-id="1daf4-937">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="1daf4-937">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1daf4-938">阅读</span><span class="sxs-lookup"><span data-stu-id="1daf4-938">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="1daf4-939">返回：</span><span class="sxs-lookup"><span data-stu-id="1daf4-939">Returns:</span></span>

<span data-ttu-id="1daf4-p158">一个包含匹配在清单 XML 文件中定义的正则表达式的字符串数组的对象。每个数组的名称等于匹配 `ItemHasRegularExpressionMatch` 规则的 `RegExName` 属性或匹配 `ItemHasKnownEntity` 规则的 `FilterName` 属性的相应值。</span><span class="sxs-lookup"><span data-stu-id="1daf4-p158">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<dl class="param-type"><span data-ttu-id="1daf4-942">

<dt>
类型</dt>


</span><span class="sxs-lookup"><span data-stu-id="1daf4-942">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="1daf4-943">对象</span><span class="sxs-lookup"><span data-stu-id="1daf4-943">Object</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="1daf4-944">示例</span><span class="sxs-lookup"><span data-stu-id="1daf4-944">Example</span></span>

<span data-ttu-id="1daf4-945">以下示例显示了如何访问正则表达式规则元素 `fruits` 和 `veggies` 的匹配项的数组，这些元素在清单中指定。</span><span class="sxs-lookup"><span data-stu-id="1daf4-945">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```js
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veges = allMatches.veggies;
```

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="1daf4-946">getRegExMatchesByName(name) → (nullable) {Array.<String>}</span><span class="sxs-lookup"><span data-stu-id="1daf4-946">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="1daf4-947">返回匹配在清单 XML 文件中定义的命名正则表达式的所选项目中的字符串值。</span><span class="sxs-lookup"><span data-stu-id="1daf4-947">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="1daf4-948">在 Outlook for iOS 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="1daf4-948">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="1daf4-949">`getRegExMatchesByName` 方法返回匹配在具有指定 `RegExName` 元素值的清单 XML 文件中的 `ItemHasRegularExpressionMatch` 规则元素中定义的正则表达式的字符串。</span><span class="sxs-lookup"><span data-stu-id="1daf4-949">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="1daf4-p159">如果在项目的正文属性上指定 `ItemHasRegularExpressionMatch` 规则，则正则表达式应进一步筛选正文，不应尝试返回该项目的整个正文。使用正则表达式（如 `.*`）获取项目的整个正文并不总是返回预期的结果。</span><span class="sxs-lookup"><span data-stu-id="1daf4-p159">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="1daf4-952">参数：</span><span class="sxs-lookup"><span data-stu-id="1daf4-952">Parameters:</span></span>

|<span data-ttu-id="1daf4-953">名称</span><span class="sxs-lookup"><span data-stu-id="1daf4-953">Name</span></span>| <span data-ttu-id="1daf4-954">类型</span><span class="sxs-lookup"><span data-stu-id="1daf4-954">Type</span></span>| <span data-ttu-id="1daf4-955">描述</span><span class="sxs-lookup"><span data-stu-id="1daf4-955">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="1daf4-956">字符串</span><span class="sxs-lookup"><span data-stu-id="1daf4-956">String</span></span>|<span data-ttu-id="1daf4-957">定义筛选器匹配的 `ItemHasRegularExpressionMatch` 规则元素的名称。</span><span class="sxs-lookup"><span data-stu-id="1daf4-957">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="1daf4-958">要求</span><span class="sxs-lookup"><span data-stu-id="1daf4-958">Requirements</span></span>

|<span data-ttu-id="1daf4-959">要求</span><span class="sxs-lookup"><span data-stu-id="1daf4-959">Requirement</span></span>| <span data-ttu-id="1daf4-960">值</span><span class="sxs-lookup"><span data-stu-id="1daf4-960">Value</span></span>|
|---|---|
|[<span data-ttu-id="1daf4-961">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="1daf4-961">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1daf4-962">1.0</span><span class="sxs-lookup"><span data-stu-id="1daf4-962">1.0</span></span>|
|[<span data-ttu-id="1daf4-963">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="1daf4-963">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1daf4-964">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1daf4-964">ReadItem</span></span>|
|[<span data-ttu-id="1daf4-965">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="1daf4-965">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1daf4-966">阅读</span><span class="sxs-lookup"><span data-stu-id="1daf4-966">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="1daf4-967">返回：</span><span class="sxs-lookup"><span data-stu-id="1daf4-967">Returns:</span></span>

<span data-ttu-id="1daf4-968">一个包含匹配在清单 XML 文件中定义的正则表达式的字符串的数组。</span><span class="sxs-lookup"><span data-stu-id="1daf4-968">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<dl class="param-type"><span data-ttu-id="1daf4-969">

<dt>类型</dt>

</span><span class="sxs-lookup"><span data-stu-id="1daf4-969">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="1daf4-970">Array.< String ></span><span class="sxs-lookup"><span data-stu-id="1daf4-970">Array.< String ></span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="1daf4-971">示例</span><span class="sxs-lookup"><span data-stu-id="1daf4-971">Example</span></span>

```js
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

####  <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="1daf4-972">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="1daf4-972">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="1daf4-973">以异步方式返回邮件的主题或正文中选定的数据。</span><span class="sxs-lookup"><span data-stu-id="1daf4-973">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="1daf4-p160">如果没有选定内容，但光标位于正文或主题中，此方法将会为所选数据返回 null。如果选定的是字段，而不是正文或主题，则此方法返回 `InvalidSelection` 错误。</span><span class="sxs-lookup"><span data-stu-id="1daf4-p160">If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="1daf4-976">参数：</span><span class="sxs-lookup"><span data-stu-id="1daf4-976">Parameters:</span></span>

|<span data-ttu-id="1daf4-977">名称</span><span class="sxs-lookup"><span data-stu-id="1daf4-977">Name</span></span>| <span data-ttu-id="1daf4-978">类型</span><span class="sxs-lookup"><span data-stu-id="1daf4-978">Type</span></span>| <span data-ttu-id="1daf4-979">属性</span><span class="sxs-lookup"><span data-stu-id="1daf4-979">Attributes</span></span>| <span data-ttu-id="1daf4-980">说明</span><span class="sxs-lookup"><span data-stu-id="1daf4-980">Description</span></span>|
|---|---|---|---|
|`coercionType`| [<span data-ttu-id="1daf4-981">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="1daf4-981">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="1daf4-p161">请求数据的格式。如果为文本，则此方法返回纯文本作为字符串，删除任何显示的 HTML 标记。如果为 HTML，则此方法返回所选文本，不论是纯文本还是 HTML。</span><span class="sxs-lookup"><span data-stu-id="1daf4-p161">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`| <span data-ttu-id="1daf4-985">对象</span><span class="sxs-lookup"><span data-stu-id="1daf4-985">Object</span></span>| <span data-ttu-id="1daf4-986">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="1daf4-986">&lt;optional&gt;</span></span>|<span data-ttu-id="1daf4-987">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="1daf4-987">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="1daf4-988">对象</span><span class="sxs-lookup"><span data-stu-id="1daf4-988">Object</span></span>| <span data-ttu-id="1daf4-989">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="1daf4-989">&lt;optional&gt;</span></span>|<span data-ttu-id="1daf4-990">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="1daf4-990">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="1daf4-991">函数</span><span class="sxs-lookup"><span data-stu-id="1daf4-991">function</span></span>||<span data-ttu-id="1daf4-992">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="1daf4-992">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="1daf4-993">若要从回调方法访问所选数据，请调用 `asyncResult.value.data`。</span><span class="sxs-lookup"><span data-stu-id="1daf4-993">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="1daf4-994">若要访问选定内容的源属性，请调用 `asyncResult.value.sourceProperty`，这将为 `body` 或 `subject`。</span><span class="sxs-lookup"><span data-stu-id="1daf4-994">To access the source property that the selection comes from, call , which will be either  or .|</span></span>|

##### <a name="requirements"></a><span data-ttu-id="1daf4-995">要求</span><span class="sxs-lookup"><span data-stu-id="1daf4-995">Requirements</span></span>

|<span data-ttu-id="1daf4-996">要求</span><span class="sxs-lookup"><span data-stu-id="1daf4-996">Requirement</span></span>| <span data-ttu-id="1daf4-997">值</span><span class="sxs-lookup"><span data-stu-id="1daf4-997">Value</span></span>|
|---|---|
|[<span data-ttu-id="1daf4-998">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="1daf4-998">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1daf4-999">1.2</span><span class="sxs-lookup"><span data-stu-id="1daf4-999">1.2</span></span>|
|[<span data-ttu-id="1daf4-1000">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="1daf4-1000">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1daf4-1001">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="1daf4-1001">ReadWriteItem</span></span>|
|[<span data-ttu-id="1daf4-1002">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="1daf4-1002">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1daf4-1003">撰写</span><span class="sxs-lookup"><span data-stu-id="1daf4-1003">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="1daf4-1004">返回：</span><span class="sxs-lookup"><span data-stu-id="1daf4-1004">Returns:</span></span>

<span data-ttu-id="1daf4-1005">作为字符串的所选数据的格式由 `coercionType` 确定。</span><span class="sxs-lookup"><span data-stu-id="1daf4-1005">The selected data as a string with format determined by `coercionType`.</span></span>

<dl class="param-type"><span data-ttu-id="1daf4-1006">

<dt>
类型</dt>


</span><span class="sxs-lookup"><span data-stu-id="1daf4-1006">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="1daf4-1007">字符串</span><span class="sxs-lookup"><span data-stu-id="1daf4-1007">String</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="1daf4-1008">示例</span><span class="sxs-lookup"><span data-stu-id="1daf4-1008">Example</span></span>

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

#### <a name="getselectedentities--entitiesjavascriptapioutlook16officeentities"></a><span data-ttu-id="1daf4-1009">getSelectedEntities() → {[Entities](/javascript/api/outlook_1_6/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="1daf4-1009">getSelectedEntities() → {[Entities](/javascript/api/outlook_1_6/office.entities)}</span></span>

<span data-ttu-id="1daf4-p163">获取在用户已选择的突出显示匹配项中找到的实体。突出显示匹配项适用于[上下文加载项](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins)。</span><span class="sxs-lookup"><span data-stu-id="1daf4-p163">Gets the entities found in a highlighted match a user has selected. Highlighted matches apply to [contextual add-ins](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="1daf4-1012">在 Outlook for iOS 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="1daf4-1012">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="1daf4-1013">要求</span><span class="sxs-lookup"><span data-stu-id="1daf4-1013">Requirements</span></span>

|<span data-ttu-id="1daf4-1014">要求</span><span class="sxs-lookup"><span data-stu-id="1daf4-1014">Requirement</span></span>| <span data-ttu-id="1daf4-1015">值</span><span class="sxs-lookup"><span data-stu-id="1daf4-1015">Value</span></span>|
|---|---|
|[<span data-ttu-id="1daf4-1016">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="1daf4-1016">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1daf4-1017">1.6</span><span class="sxs-lookup"><span data-stu-id="1daf4-1017">-16</span></span> |
|[<span data-ttu-id="1daf4-1018">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="1daf4-1018">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1daf4-1019">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1daf4-1019">ReadItem</span></span>|
|[<span data-ttu-id="1daf4-1020">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="1daf4-1020">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1daf4-1021">阅读</span><span class="sxs-lookup"><span data-stu-id="1daf4-1021">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="1daf4-1022">返回：</span><span class="sxs-lookup"><span data-stu-id="1daf4-1022">Returns:</span></span>

<span data-ttu-id="1daf4-1023">类型：[Entities](/javascript/api/outlook_1_6/office.entities)</span><span class="sxs-lookup"><span data-stu-id="1daf4-1023">Type: [Entities](/javascript/api/outlook_1_6/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="1daf4-1024">示例</span><span class="sxs-lookup"><span data-stu-id="1daf4-1024">Example</span></span>

<span data-ttu-id="1daf4-1025">以下示例访问用户选择的突出显示匹配项中的地址实体。</span><span class="sxs-lookup"><span data-stu-id="1daf4-1025">The following example accesses the addresses entities in the highlighted match selected by the user.</span></span>

```js
var contacts = Office.context.mailbox.item.getSelectedEntities().addresses;
```

#### <a name="getselectedregexmatches--object"></a><span data-ttu-id="1daf4-1026">getSelectedRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="1daf4-1026">getSelectedRegExMatches() → {Object}</span></span>

<span data-ttu-id="1daf4-p164">返回突出显示匹配项中匹配在清单 XML 文件中定义的正则表达式的字符串值。突出显示匹配项适用于[上下文外接程序](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins)。</span><span class="sxs-lookup"><span data-stu-id="1daf4-p164">Returns string values in a highlighted match that match the regular expressions defined in the manifest XML file. Highlighted matches apply to [contextual add-ins](https://docs.microsoft.com/outlook/add-ins/contextual-outlook-add-ins).</span></span>

> [!NOTE]
> <span data-ttu-id="1daf4-1029">在 Outlook for iOS 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="1daf4-1029">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="1daf4-p165">`getSelectedRegExMatches` 方法返回匹配在清单 XML 文件中的每个 `ItemHasRegularExpressionMatch` 或 `ItemHasKnownEntity` 规则元素中定义的正则表达式的字符串。对于 `ItemHasRegularExpressionMatch` 规则，匹配字符串必须发生在该规则指定的项目的属性中。`PropertyName` 简单类型定义支持的属性。</span><span class="sxs-lookup"><span data-stu-id="1daf4-p165">The `getSelectedRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="1daf4-1033">例如，考虑一个外接程序清单具有以下 `Rule` 元素：</span><span class="sxs-lookup"><span data-stu-id="1daf4-1033">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="1daf4-1034">从 `getRegExMatches` 返回的对象应有两个属性：`fruits` 和 `veggies`。</span><span class="sxs-lookup"><span data-stu-id="1daf4-1034">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="1daf4-p166">如果在项目的正文属性上指定 `ItemHasRegularExpressionMatch` 规则，则正则表达式应进一步筛选正文，不应尝试返回该项目的整个正文。使用正则表达式（如 `.*`）获取项目的整个正文并不总是返回预期的结果。而是使用 [`Body.getAsync`](/javascript/api/outlook_1_6/office.body#getasync-coerciontype--options--callback-) 方法检索整个正文。</span><span class="sxs-lookup"><span data-stu-id="1daf4-p166">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook_1_6/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="1daf4-1038">Requirements</span><span class="sxs-lookup"><span data-stu-id="1daf4-1038">Requirements</span></span>

|<span data-ttu-id="1daf4-1039">要求</span><span class="sxs-lookup"><span data-stu-id="1daf4-1039">Requirement</span></span>| <span data-ttu-id="1daf4-1040">值</span><span class="sxs-lookup"><span data-stu-id="1daf4-1040">Value</span></span>|
|---|---|
|[<span data-ttu-id="1daf4-1041">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="1daf4-1041">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1daf4-1042">1.6</span><span class="sxs-lookup"><span data-stu-id="1daf4-1042">-16</span></span> |
|[<span data-ttu-id="1daf4-1043">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="1daf4-1043">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1daf4-1044">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1daf4-1044">ReadItem</span></span>|
|[<span data-ttu-id="1daf4-1045">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="1daf4-1045">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1daf4-1046">阅读</span><span class="sxs-lookup"><span data-stu-id="1daf4-1046">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="1daf4-1047">返回：</span><span class="sxs-lookup"><span data-stu-id="1daf4-1047">Returns:</span></span>

<span data-ttu-id="1daf4-p167">一个包含匹配在清单 XML 文件中定义的正则表达式的字符串数组的对象。每个数组的名称等于匹配 `ItemHasRegularExpressionMatch` 规则的 `RegExName` 属性或匹配 `ItemHasKnownEntity` 规则的 `FilterName` 属性的相应值。</span><span class="sxs-lookup"><span data-stu-id="1daf4-p167">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

##### <a name="example"></a><span data-ttu-id="1daf4-1050">示例</span><span class="sxs-lookup"><span data-stu-id="1daf4-1050">Example</span></span>

<span data-ttu-id="1daf4-1051">以下示例显示了如何访问正则表达式规则元素 `fruits` 和 `veggies` 的匹配项的数组，这些元素在清单中指定。</span><span class="sxs-lookup"><span data-stu-id="1daf4-1051">The following example shows how to access the array of matches for the regular expression rule elements `fruits` and `veggies`, which are specified in the manifest.</span></span>

```
var selectedMatches = Office.context.mailbox.item.getSelectedRegExMatches();
var fruits = selectedMatches.fruits;
var veggies = selectedMatches.veggies;
```

####  <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="1daf4-1052">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="1daf4-1052">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="1daf4-1053">异步加载所选项目上此外接程序的自定义属性。</span><span class="sxs-lookup"><span data-stu-id="1daf4-1053">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="1daf4-p168">自定义属性基于每个应用、每个项目存储为键/值对。此方法在回调中返回 `CustomProperties` 对象，该回调提供访问特定于当前项目和当前外接程序的自定义属性的方法。自定义属性未在项目上加密，因此这不应用作安全存储。</span><span class="sxs-lookup"><span data-stu-id="1daf4-p168">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="1daf4-1057">参数：</span><span class="sxs-lookup"><span data-stu-id="1daf4-1057">Parameters:</span></span>

|<span data-ttu-id="1daf4-1058">名称</span><span class="sxs-lookup"><span data-stu-id="1daf4-1058">Name</span></span>| <span data-ttu-id="1daf4-1059">类型</span><span class="sxs-lookup"><span data-stu-id="1daf4-1059">Type</span></span>| <span data-ttu-id="1daf4-1060">属性</span><span class="sxs-lookup"><span data-stu-id="1daf4-1060">Attributes</span></span>| <span data-ttu-id="1daf4-1061">说明</span><span class="sxs-lookup"><span data-stu-id="1daf4-1061">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="1daf4-1062">函数</span><span class="sxs-lookup"><span data-stu-id="1daf4-1062">function</span></span>||<span data-ttu-id="1daf4-1063">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="1daf4-1063">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="1daf4-1064">自定义属性作为 `asyncResult.value` 属性中的 [`CustomProperties`](/javascript/api/outlook_1_6/office.customproperties) 对象提供。</span><span class="sxs-lookup"><span data-stu-id="1daf4-1064">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook_1_6/office.customproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="1daf4-1065">此对象可用于获取、设置以及从项目中删除自定义属性，并将自定义属性集的更改重新保存到服务器。</span><span class="sxs-lookup"><span data-stu-id="1daf4-1065">The custom properties are provided as a CustomProperties object in the asyncResult.value property. This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`| <span data-ttu-id="1daf4-1066">对象</span><span class="sxs-lookup"><span data-stu-id="1daf4-1066">Object</span></span>| <span data-ttu-id="1daf4-1067">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="1daf4-1067">&lt;optional&gt;</span></span>|<span data-ttu-id="1daf4-1068">开发人员可以提供他们想要在回调函数中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="1daf4-1068">Developers can provide any object they wish to access in the callback method.</span></span> <span data-ttu-id="1daf4-1069">此对象可以通过回调函数中的 `asyncResult.asyncContext` 属性进行访问。</span><span class="sxs-lookup"><span data-stu-id="1daf4-1069">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.|</span></span>|

##### <a name="requirements"></a><span data-ttu-id="1daf4-1070">要求</span><span class="sxs-lookup"><span data-stu-id="1daf4-1070">Requirements</span></span>

|<span data-ttu-id="1daf4-1071">要求</span><span class="sxs-lookup"><span data-stu-id="1daf4-1071">Requirement</span></span>| <span data-ttu-id="1daf4-1072">值</span><span class="sxs-lookup"><span data-stu-id="1daf4-1072">Value</span></span>|
|---|---|
|[<span data-ttu-id="1daf4-1073">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="1daf4-1073">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1daf4-1074">1.0</span><span class="sxs-lookup"><span data-stu-id="1daf4-1074">1.0</span></span>|
|[<span data-ttu-id="1daf4-1075">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="1daf4-1075">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1daf4-1076">ReadItem</span><span class="sxs-lookup"><span data-stu-id="1daf4-1076">ReadItem</span></span>|
|[<span data-ttu-id="1daf4-1077">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="1daf4-1077">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1daf4-1078">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="1daf4-1078">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="1daf4-1079">示例</span><span class="sxs-lookup"><span data-stu-id="1daf4-1079">Example</span></span>

<span data-ttu-id="1daf4-p171">以下代码示例显示了如何使用 `loadCustomPropertiesAsync` 方法异步加载特定于当前项目的自定义属性。该示例还显示了如何使用 `CustomProperties.saveAsync` 方法将这些属性重新保存到服务器。加载自定义属性后，该代码示例将使用 `CustomProperties.get` 方法读取自定义属性 `myProp`，使用 `CustomProperties.set` 方法写入自定义属性 `otherProp`，最后调用 `saveAsync` 方法保存这些自定义属性。</span><span class="sxs-lookup"><span data-stu-id="1daf4-p171">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

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

####  <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="1daf4-1083">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="1daf4-1083">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="1daf4-1084">将附件从邮件或约会中删除。</span><span class="sxs-lookup"><span data-stu-id="1daf4-1084">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="1daf4-p172">`removeAttachmentAsync` 方法删除项目中带指定标识符的附件。最佳做法是，仅当同一个邮件应用程序在同一会话中添加了一个附件时，你才应使用该附件标识符来删除该附件。在 Outlook Web App 和适用于设备的 OWA 中，附件标识符只在同一个会话中才有效。当用户关闭应用程序，或者如果用户开始在内嵌窗体中撰写，并在随后弹出的内嵌窗体中继续在单独的窗口撰写时，会话即结束。</span><span class="sxs-lookup"><span data-stu-id="1daf4-p172">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item. As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session. In Outlook Web App and OWA for Devices, the attachment identifier is valid only within the same session. A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="1daf4-1089">参数：</span><span class="sxs-lookup"><span data-stu-id="1daf4-1089">Parameters:</span></span>

|<span data-ttu-id="1daf4-1090">名称</span><span class="sxs-lookup"><span data-stu-id="1daf4-1090">Name</span></span>| <span data-ttu-id="1daf4-1091">类型</span><span class="sxs-lookup"><span data-stu-id="1daf4-1091">Type</span></span>| <span data-ttu-id="1daf4-1092">属性</span><span class="sxs-lookup"><span data-stu-id="1daf4-1092">Attributes</span></span>| <span data-ttu-id="1daf4-1093">说明</span><span class="sxs-lookup"><span data-stu-id="1daf4-1093">Description</span></span>|
|---|---|---|---|
|`attachmentId`| <span data-ttu-id="1daf4-1094">字符串</span><span class="sxs-lookup"><span data-stu-id="1daf4-1094">String</span></span>||<span data-ttu-id="1daf4-p173">要删除的附件的标识符。字符串的最大长度为 100 个字符。</span><span class="sxs-lookup"><span data-stu-id="1daf4-p173">The identifier of the attachment to remove. The maximum length of the string is 100 characters.</span></span>|
|`options`| <span data-ttu-id="1daf4-1097">Object</span><span class="sxs-lookup"><span data-stu-id="1daf4-1097">Object</span></span>| <span data-ttu-id="1daf4-1098">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="1daf4-1098">&lt;optional&gt;</span></span>|<span data-ttu-id="1daf4-1099">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="1daf4-1099">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="1daf4-1100">对象</span><span class="sxs-lookup"><span data-stu-id="1daf4-1100">Object</span></span>| <span data-ttu-id="1daf4-1101">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="1daf4-1101">&lt;optional&gt;</span></span>|<span data-ttu-id="1daf4-1102">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="1daf4-1102">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="1daf4-1103">函数</span><span class="sxs-lookup"><span data-stu-id="1daf4-1103">function</span></span>| <span data-ttu-id="1daf4-1104">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="1daf4-1104">&lt;optional&gt;</span></span>|<span data-ttu-id="1daf4-1105">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="1daf4-1105">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="1daf4-1106">如果删除附件失败，`asyncResult.error` 属性将包含一个说明失败原因的错误代码。</span><span class="sxs-lookup"><span data-stu-id="1daf4-1106">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="1daf4-1107">错误</span><span class="sxs-lookup"><span data-stu-id="1daf4-1107">Errors</span></span>

| <span data-ttu-id="1daf4-1108">错误代码</span><span class="sxs-lookup"><span data-stu-id="1daf4-1108">Error code</span></span> | <span data-ttu-id="1daf4-1109">说明</span><span class="sxs-lookup"><span data-stu-id="1daf4-1109">Description</span></span> |
|------------|-------------|
| `InvalidAttachmentId` | <span data-ttu-id="1daf4-1110">附件标识符不存在。</span><span class="sxs-lookup"><span data-stu-id="1daf4-1110">The attachment identifier does not exist.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="1daf4-1111">要求</span><span class="sxs-lookup"><span data-stu-id="1daf4-1111">Requirements</span></span>

|<span data-ttu-id="1daf4-1112">要求</span><span class="sxs-lookup"><span data-stu-id="1daf4-1112">Requirement</span></span>| <span data-ttu-id="1daf4-1113">值</span><span class="sxs-lookup"><span data-stu-id="1daf4-1113">Value</span></span>|
|---|---|
|[<span data-ttu-id="1daf4-1114">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="1daf4-1114">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1daf4-1115">1.1</span><span class="sxs-lookup"><span data-stu-id="1daf4-1115">1.1</span></span>|
|[<span data-ttu-id="1daf4-1116">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="1daf4-1116">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1daf4-1117">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="1daf4-1117">ReadWriteItem</span></span>|
|[<span data-ttu-id="1daf4-1118">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="1daf4-1118">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1daf4-1119">撰写</span><span class="sxs-lookup"><span data-stu-id="1daf4-1119">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="1daf4-1120">示例</span><span class="sxs-lookup"><span data-stu-id="1daf4-1120">Example</span></span>

<span data-ttu-id="1daf4-1121">以下代码删除包含标识符 '0' 的附件。</span><span class="sxs-lookup"><span data-stu-id="1daf4-1121">The following code removes an attachment with an identifier of '0'.</span></span>

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

####  <a name="saveasyncoptions-callback"></a><span data-ttu-id="1daf4-1122">saveAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="1daf4-1122">saveAsync([options], callback)</span></span>

<span data-ttu-id="1daf4-1123">异步保存项目。</span><span class="sxs-lookup"><span data-stu-id="1daf4-1123">Asynchronously saves an item.</span></span>

<span data-ttu-id="1daf4-p174">调用时，此方法将当前邮件保存为草稿，并通过回调方法返回项目 ID。在 Outlook Web App 或 Outlook 联机模式下，该项目被保存到服务器中。在 Outlook 缓存模式下，该项目被保存到本地缓存中。</span><span class="sxs-lookup"><span data-stu-id="1daf4-p174">When invoked, this method saves the current message as a draft and returns the item id via the callback method. In Outlook Web App or Outlook in online mode, the item is saved to the server. In Outlook in cached mode, the item is saved to the local cache.</span></span>

> [!NOTE]
> <span data-ttu-id="1daf4-1127">如果加载项在撰写模式下对某个项目调用 `saveAsync` 来获得 `itemId`，以便与 EWS 或 REST API 一同使用，请注意，当 Outlook 处于高速缓存模式时，可能需要一段时间项目才能真正同步到服务器。</span><span class="sxs-lookup"><span data-stu-id="1daf4-1127">Note: If your add-in calls `saveAsync` on an item in compose mode in order to get an `itemId` to use with EWS or the REST API, be aware that when Outlook is in cached mode, it may take some time before the item is actually synced to the server. Until the item is synced, using the  will return an error.</span></span> <span data-ttu-id="1daf4-1128">在项目同步前，使用 `itemId` 将返回一个错误。</span><span class="sxs-lookup"><span data-stu-id="1daf4-1128">Until the item is synced, using the `itemId` will return an error.</span></span>

<span data-ttu-id="1daf4-p176">由于约会没有草稿状态，如果以撰写模式在约会中调用 `saveAsync`，则该项将被保存为用户日历中的正常约会。对于之前未保存过的新约会，则不会发送邀请。保存现有约会将向添加或删除的与会者发送更新。</span><span class="sxs-lookup"><span data-stu-id="1daf4-p176">Since appointments have no draft state, if `saveAsync` is called on an appointment in compose mode, the item will be saved as a normal appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent. Saving an existing appointment will send an update to added or removed attendees.</span></span>

> [!NOTE]
> <span data-ttu-id="1daf4-1132">以下客户端在撰写模式下对约会上的 `saveAsync` 具有不同的行为：</span><span class="sxs-lookup"><span data-stu-id="1daf4-1132">The following clients have different behavior for `saveAsync` on appointments in compose mode:</span></span>
>
> - <span data-ttu-id="1daf4-1133">Mac Outlook 不支持在撰写模式下对会议执行 `saveAsync` 操作。</span><span class="sxs-lookup"><span data-stu-id="1daf4-1133">Note: Mac Outlook does not support `saveAsync` on a meeting in compose mode. Calling  on a meeting in Mac Outlook will return an error.</span></span> <span data-ttu-id="1daf4-1134">对 Mac Outlook 中的会议调用 `saveAsync` 将会返回一个错误。</span><span class="sxs-lookup"><span data-stu-id="1daf4-1134">Note: Mac Outlook does not support  on a meeting in compose mode. Calling `saveAsync` on a meeting in Mac Outlook will return an error.</span></span>
> - <span data-ttu-id="1daf4-1135">撰写模式下在约会上调用 `saveAsync` 时，Outlook 网页版始终发送邀请或更新。</span><span class="sxs-lookup"><span data-stu-id="1daf4-1135">Outlook on the web always sends an invitation or update when `saveAsync` is called on an appointment in compose mode.</span></span>

##### <a name="parameters"></a><span data-ttu-id="1daf4-1136">参数：</span><span class="sxs-lookup"><span data-stu-id="1daf4-1136">Parameters:</span></span>

|<span data-ttu-id="1daf4-1137">名称</span><span class="sxs-lookup"><span data-stu-id="1daf4-1137">Name</span></span>| <span data-ttu-id="1daf4-1138">类型</span><span class="sxs-lookup"><span data-stu-id="1daf4-1138">Type</span></span>| <span data-ttu-id="1daf4-1139">属性</span><span class="sxs-lookup"><span data-stu-id="1daf4-1139">Attributes</span></span>| <span data-ttu-id="1daf4-1140">说明</span><span class="sxs-lookup"><span data-stu-id="1daf4-1140">Description</span></span>|
|---|---|---|---|
|`options`| <span data-ttu-id="1daf4-1141">Object</span><span class="sxs-lookup"><span data-stu-id="1daf4-1141">Object</span></span>| <span data-ttu-id="1daf4-1142">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="1daf4-1142">&lt;optional&gt;</span></span>|<span data-ttu-id="1daf4-1143">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="1daf4-1143">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="1daf4-1144">对象</span><span class="sxs-lookup"><span data-stu-id="1daf4-1144">Object</span></span>| <span data-ttu-id="1daf4-1145">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="1daf4-1145">&lt;optional&gt;</span></span>|<span data-ttu-id="1daf4-1146">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="1daf4-1146">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="1daf4-1147">函数</span><span class="sxs-lookup"><span data-stu-id="1daf4-1147">function</span></span>||<span data-ttu-id="1daf4-1148">方法完成后，使用一个参数 `asyncResult`（[`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="1daf4-1148">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="1daf4-1149">如果成功，该项目标识符在 `asyncResult.value` 属性中提供。</span><span class="sxs-lookup"><span data-stu-id="1daf4-1149">On success, the item identifier is provided in the `asyncResult.value` property.|</span></span>|

##### <a name="requirements"></a><span data-ttu-id="1daf4-1150">要求</span><span class="sxs-lookup"><span data-stu-id="1daf4-1150">Requirements</span></span>

|<span data-ttu-id="1daf4-1151">要求</span><span class="sxs-lookup"><span data-stu-id="1daf4-1151">Requirement</span></span>| <span data-ttu-id="1daf4-1152">值</span><span class="sxs-lookup"><span data-stu-id="1daf4-1152">Value</span></span>|
|---|---|
|[<span data-ttu-id="1daf4-1153">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="1daf4-1153">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1daf4-1154">1.3</span><span class="sxs-lookup"><span data-stu-id="1daf4-1154">1.3</span></span>|
|[<span data-ttu-id="1daf4-1155">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="1daf4-1155">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1daf4-1156">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="1daf4-1156">ReadWriteItem</span></span>|
|[<span data-ttu-id="1daf4-1157">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="1daf4-1157">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1daf4-1158">撰写</span><span class="sxs-lookup"><span data-stu-id="1daf4-1158">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="1daf4-1159">示例</span><span class="sxs-lookup"><span data-stu-id="1daf4-1159">Examples</span></span>

```js
Office.context.mailbox.item.saveAsync(
  function callback(result) {
    // Process the result
  });
```

<span data-ttu-id="1daf4-p178">下面是传递给回调函数的 `result` 参数的示例。`value` 属性包含的项目的项目 ID。</span><span class="sxs-lookup"><span data-stu-id="1daf4-p178">The following is an example of the `result` parameter passed to the callback function. The `value` property contains the item ID of the item.</span></span>

```json
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

####  <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="1daf4-1162">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="1daf4-1162">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="1daf4-1163">以异步方式将数据插入到邮件的正文或主题中。</span><span class="sxs-lookup"><span data-stu-id="1daf4-1163">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="1daf4-p179">`setSelectedDataAsync` 方法将指定字符串插入到项目主题或正文的光标位置，或者，如果在编辑器中已选择文本，则该方法将替换选择的文本。如果光标不在正文或主题字段中，则返回错误。插入之后，光标会位于插入内容的末尾。</span><span class="sxs-lookup"><span data-stu-id="1daf4-p179">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="1daf4-1167">参数：</span><span class="sxs-lookup"><span data-stu-id="1daf4-1167">Parameters:</span></span>

|<span data-ttu-id="1daf4-1168">名称</span><span class="sxs-lookup"><span data-stu-id="1daf4-1168">Name</span></span>| <span data-ttu-id="1daf4-1169">类型</span><span class="sxs-lookup"><span data-stu-id="1daf4-1169">Type</span></span>| <span data-ttu-id="1daf4-1170">属性</span><span class="sxs-lookup"><span data-stu-id="1daf4-1170">Attributes</span></span>| <span data-ttu-id="1daf4-1171">说明</span><span class="sxs-lookup"><span data-stu-id="1daf4-1171">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="1daf4-1172">字符串</span><span class="sxs-lookup"><span data-stu-id="1daf4-1172">String</span></span>||<span data-ttu-id="1daf4-p180">要插入的数据。数据不得超过 1,000,000 个字符。如果传入的数据超过 1,000,000 个字符，则会引发 `ArgumentOutOfRange` 异常。</span><span class="sxs-lookup"><span data-stu-id="1daf4-p180">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`| <span data-ttu-id="1daf4-1176">Object</span><span class="sxs-lookup"><span data-stu-id="1daf4-1176">Object</span></span>| <span data-ttu-id="1daf4-1177">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="1daf4-1177">&lt;optional&gt;</span></span>|<span data-ttu-id="1daf4-1178">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="1daf4-1178">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="1daf4-1179">对象</span><span class="sxs-lookup"><span data-stu-id="1daf4-1179">Object</span></span>| <span data-ttu-id="1daf4-1180">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="1daf4-1180">&lt;optional&gt;</span></span>|<span data-ttu-id="1daf4-1181">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="1daf4-1181">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`| [<span data-ttu-id="1daf4-1182">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="1daf4-1182">Office.CoercionType</span></span>](office.md#coerciontype-string)| <span data-ttu-id="1daf4-1183">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="1daf4-1183">&lt;optional&gt;</span></span>|<span data-ttu-id="1daf4-p181">如果为 `text`，则在 Outlook Web App 和 Outlook 中应用当前样式。如果该字段是 HTML 编辑器，则仅插入文本数据，即使数据为 HTML。</span><span class="sxs-lookup"><span data-stu-id="1daf4-p181">If `text`, the current style is applied in Outlook Web App and Outlook. If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="1daf4-p182">如果 `html` 和该字段支持 HTML（主题不支持），则在 Outlook Web App 中应用当前样式，而在 Outlook 中应用默认样式。如果该字段是文本字段，则返回 `InvalidDataFormat` 错误。</span><span class="sxs-lookup"><span data-stu-id="1daf4-p182">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook Web App and the default style is applied in Outlook. If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="1daf4-1188">如果未设置 `coercionType`，则结果取决于该字段：如果该字段是 HTML，则使用 HTML；如果该字段是文本，则使用纯文本。</span><span class="sxs-lookup"><span data-stu-id="1daf4-1188">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`| <span data-ttu-id="1daf4-1189">function</span><span class="sxs-lookup"><span data-stu-id="1daf4-1189">function</span></span>||<span data-ttu-id="1daf4-1190">方法完成后，使用单个参数 `callback`（一个 [`asyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `AsyncResult` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="1daf4-1190">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="1daf4-1191">Requirements</span><span class="sxs-lookup"><span data-stu-id="1daf4-1191">Requirements</span></span>

|<span data-ttu-id="1daf4-1192">要求</span><span class="sxs-lookup"><span data-stu-id="1daf4-1192">Requirement</span></span>| <span data-ttu-id="1daf4-1193">值</span><span class="sxs-lookup"><span data-stu-id="1daf4-1193">Value</span></span>|
|---|---|
|[<span data-ttu-id="1daf4-1194">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="1daf4-1194">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1daf4-1195">1.2</span><span class="sxs-lookup"><span data-stu-id="1daf4-1195">1.2</span></span>|
|[<span data-ttu-id="1daf4-1196">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="1daf4-1196">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="1daf4-1197">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="1daf4-1197">ReadWriteItem</span></span>|
|[<span data-ttu-id="1daf4-1198">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="1daf4-1198">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1daf4-1199">撰写</span><span class="sxs-lookup"><span data-stu-id="1daf4-1199">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="1daf4-1200">示例</span><span class="sxs-lookup"><span data-stu-id="1daf4-1200">Example</span></span>

```js
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```