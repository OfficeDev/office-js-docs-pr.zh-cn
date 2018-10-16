
# <a name="item"></a><span data-ttu-id="34657-101">item</span><span class="sxs-lookup"><span data-stu-id="34657-101">item</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmditem"></a><span data-ttu-id="34657-102">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span><span class="sxs-lookup"><span data-stu-id="34657-102">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span></span>

<span data-ttu-id="34657-p101">`item`命名空间用于访问当前选定的邮件、会议请求或安排。可以通过使用[itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlook15officemailboxenumsitemtype) 属性确定`item`的类型。</span><span class="sxs-lookup"><span data-stu-id="34657-p101">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlook15officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="34657-105">要求</span><span class="sxs-lookup"><span data-stu-id="34657-105">Requirements</span></span>

|<span data-ttu-id="34657-106">要求</span><span class="sxs-lookup"><span data-stu-id="34657-106">Requirement</span></span>| <span data-ttu-id="34657-107">值</span><span class="sxs-lookup"><span data-stu-id="34657-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="34657-108">最低邮箱要求集版本</span><span class="sxs-lookup"><span data-stu-id="34657-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="34657-109">1.0</span><span class="sxs-lookup"><span data-stu-id="34657-109">1.0</span></span>|
|[<span data-ttu-id="34657-110">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="34657-110">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="34657-111">Restricted</span><span class="sxs-lookup"><span data-stu-id="34657-111">Restricted</span></span>|
|[<span data-ttu-id="34657-112">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="34657-112">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="34657-113">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="34657-113">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="34657-114">成员和方法</span><span class="sxs-lookup"><span data-stu-id="34657-114">Members and methods</span></span>

| <span data-ttu-id="34657-115">成员</span><span class="sxs-lookup"><span data-stu-id="34657-115">Member</span></span> | <span data-ttu-id="34657-116">类型</span><span class="sxs-lookup"><span data-stu-id="34657-116">Type</span></span> |
|--------|------|
| [<span data-ttu-id="34657-117">attachments</span><span class="sxs-lookup"><span data-stu-id="34657-117">attachments</span></span>](#attachments-arrayattachmentdetailsjavascriptapioutlook15officeattachmentdetails) | <span data-ttu-id="34657-118">成员</span><span class="sxs-lookup"><span data-stu-id="34657-118">Member</span></span> |
| [<span data-ttu-id="34657-119">bcc</span><span class="sxs-lookup"><span data-stu-id="34657-119">bcc</span></span>](#bcc-recipientsjavascriptapioutlook15officerecipients) | <span data-ttu-id="34657-120">成员</span><span class="sxs-lookup"><span data-stu-id="34657-120">Member</span></span> |
| [<span data-ttu-id="34657-121">body</span><span class="sxs-lookup"><span data-stu-id="34657-121">body</span></span>](#body-bodyjavascriptapioutlook15officebody) | <span data-ttu-id="34657-122">成员</span><span class="sxs-lookup"><span data-stu-id="34657-122">Member</span></span> |
| [<span data-ttu-id="34657-123">cc</span><span class="sxs-lookup"><span data-stu-id="34657-123">cc</span></span>](#cc-arrayemailaddressdetailsjavascriptapioutlook15officeemailaddressdetailsrecipientsjavascriptapioutlook15officerecipients) | <span data-ttu-id="34657-124">成员</span><span class="sxs-lookup"><span data-stu-id="34657-124">Member</span></span> |
| [<span data-ttu-id="34657-125">conversationId</span><span class="sxs-lookup"><span data-stu-id="34657-125">conversationId</span></span>](#nullable-conversationid-string) | <span data-ttu-id="34657-126">成员</span><span class="sxs-lookup"><span data-stu-id="34657-126">Member</span></span> |
| [<span data-ttu-id="34657-127">dateTimeCreated</span><span class="sxs-lookup"><span data-stu-id="34657-127">dateTimeCreated</span></span>](#datetimecreated-date) | <span data-ttu-id="34657-128">成员</span><span class="sxs-lookup"><span data-stu-id="34657-128">Member</span></span> |
| [<span data-ttu-id="34657-129">dateTimeModified</span><span class="sxs-lookup"><span data-stu-id="34657-129">dateTimeModified</span></span>](#datetimemodified-date) | <span data-ttu-id="34657-130">成员</span><span class="sxs-lookup"><span data-stu-id="34657-130">Member</span></span> |
| [<span data-ttu-id="34657-131">end</span><span class="sxs-lookup"><span data-stu-id="34657-131">end</span></span>](#end-datetimejavascriptapioutlook15officetime) | <span data-ttu-id="34657-132">成员</span><span class="sxs-lookup"><span data-stu-id="34657-132">Member</span></span> |
| [<span data-ttu-id="34657-133">from</span><span class="sxs-lookup"><span data-stu-id="34657-133">from</span></span>](#from-emailaddressdetailsjavascriptapioutlook15officeemailaddressdetails) | <span data-ttu-id="34657-134">成员</span><span class="sxs-lookup"><span data-stu-id="34657-134">Member</span></span> |
| [<span data-ttu-id="34657-135">internetMessageId</span><span class="sxs-lookup"><span data-stu-id="34657-135">internetMessageId</span></span>](#internetmessageid-string) | <span data-ttu-id="34657-136">成员</span><span class="sxs-lookup"><span data-stu-id="34657-136">Member</span></span> |
| [<span data-ttu-id="34657-137">itemClass</span><span class="sxs-lookup"><span data-stu-id="34657-137">itemClass</span></span>](#itemclass-string) | <span data-ttu-id="34657-138">成员</span><span class="sxs-lookup"><span data-stu-id="34657-138">Member</span></span> |
| [<span data-ttu-id="34657-139">itemId</span><span class="sxs-lookup"><span data-stu-id="34657-139">itemId</span></span>](#nullable-itemid-string) | <span data-ttu-id="34657-140">成员</span><span class="sxs-lookup"><span data-stu-id="34657-140">Member</span></span> |
| [<span data-ttu-id="34657-141">itemType</span><span class="sxs-lookup"><span data-stu-id="34657-141">itemType</span></span>](#itemtype-officemailboxenumsitemtypejavascriptapioutlook15officemailboxenumsitemtype) | <span data-ttu-id="34657-142">成员</span><span class="sxs-lookup"><span data-stu-id="34657-142">Member</span></span> |
| [<span data-ttu-id="34657-143">location</span><span class="sxs-lookup"><span data-stu-id="34657-143">location</span></span>](#location-stringlocationjavascriptapioutlook15officelocation) | <span data-ttu-id="34657-144">成员</span><span class="sxs-lookup"><span data-stu-id="34657-144">Member</span></span> |
| [<span data-ttu-id="34657-145">normalizedSubject</span><span class="sxs-lookup"><span data-stu-id="34657-145">normalizedSubject</span></span>](#normalizedsubject-string) | <span data-ttu-id="34657-146">成员</span><span class="sxs-lookup"><span data-stu-id="34657-146">Member</span></span> |
| [<span data-ttu-id="34657-147">notificationMessages</span><span class="sxs-lookup"><span data-stu-id="34657-147">notificationMessages</span></span>](#notificationmessages-notificationmessagesjavascriptapioutlook15officenotificationmessages) | <span data-ttu-id="34657-148">成员</span><span class="sxs-lookup"><span data-stu-id="34657-148">Member</span></span> |
| [<span data-ttu-id="34657-149">optionalAttendees</span><span class="sxs-lookup"><span data-stu-id="34657-149">optionalAttendees</span></span>](#optionalattendees-arrayemailaddressdetailsjavascriptapioutlook15officeemailaddressdetailsrecipientsjavascriptapioutlook15officerecipients) | <span data-ttu-id="34657-150">成员</span><span class="sxs-lookup"><span data-stu-id="34657-150">Member</span></span> |
| [<span data-ttu-id="34657-151">organizer</span><span class="sxs-lookup"><span data-stu-id="34657-151">organizer</span></span>](#organizer-emailaddressdetailsjavascriptapioutlook15officeemailaddressdetails) | <span data-ttu-id="34657-152">成员</span><span class="sxs-lookup"><span data-stu-id="34657-152">Member</span></span> |
| [<span data-ttu-id="34657-153">requiredAttendees</span><span class="sxs-lookup"><span data-stu-id="34657-153">requiredAttendees</span></span>](#requiredattendees-arrayemailaddressdetailsjavascriptapioutlook15officeemailaddressdetailsrecipientsjavascriptapioutlook15officerecipients) | <span data-ttu-id="34657-154">成员</span><span class="sxs-lookup"><span data-stu-id="34657-154">Member</span></span> |
| [<span data-ttu-id="34657-155">sender</span><span class="sxs-lookup"><span data-stu-id="34657-155">sender</span></span>](#sender-emailaddressdetailsjavascriptapioutlook15officeemailaddressdetails) | <span data-ttu-id="34657-156">成员</span><span class="sxs-lookup"><span data-stu-id="34657-156">Member</span></span> |
| [<span data-ttu-id="34657-157">start</span><span class="sxs-lookup"><span data-stu-id="34657-157">start</span></span>](#start-datetimejavascriptapioutlook15officetime) | <span data-ttu-id="34657-158">成员</span><span class="sxs-lookup"><span data-stu-id="34657-158">Member</span></span> |
| [<span data-ttu-id="34657-159">subject</span><span class="sxs-lookup"><span data-stu-id="34657-159">subject</span></span>](#subject-stringsubjectjavascriptapioutlook15officesubject) | <span data-ttu-id="34657-160">成员</span><span class="sxs-lookup"><span data-stu-id="34657-160">Member</span></span> |
| [<span data-ttu-id="34657-161">to</span><span class="sxs-lookup"><span data-stu-id="34657-161">to</span></span>](#to-arrayemailaddressdetailsjavascriptapioutlook15officeemailaddressdetailsrecipientsjavascriptapioutlook15officerecipients) | <span data-ttu-id="34657-162">成员</span><span class="sxs-lookup"><span data-stu-id="34657-162">Member</span></span> |
| [<span data-ttu-id="34657-163">addFileAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="34657-163">addFileAttachmentAsync</span></span>](#addfileattachmentasyncuri-attachmentname-options-callback) | <span data-ttu-id="34657-164">方法</span><span class="sxs-lookup"><span data-stu-id="34657-164">Method</span></span> |
| [<span data-ttu-id="34657-165">addItemAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="34657-165">addItemAttachmentAsync</span></span>](#additemattachmentasyncitemid-attachmentname-options-callback) | <span data-ttu-id="34657-166">方法</span><span class="sxs-lookup"><span data-stu-id="34657-166">Method</span></span> |
| [<span data-ttu-id="34657-167">close</span><span class="sxs-lookup"><span data-stu-id="34657-167">close</span></span>](#close) | <span data-ttu-id="34657-168">方法</span><span class="sxs-lookup"><span data-stu-id="34657-168">Method</span></span> |
| [<span data-ttu-id="34657-169">displayReplyAllForm</span><span class="sxs-lookup"><span data-stu-id="34657-169">displayReplyAllForm</span></span>](#displayreplyallformformdata) | <span data-ttu-id="34657-170">方法</span><span class="sxs-lookup"><span data-stu-id="34657-170">Method</span></span> |
| [<span data-ttu-id="34657-171">displayReplyForm</span><span class="sxs-lookup"><span data-stu-id="34657-171">displayReplyForm</span></span>](#displayreplyformformdata) | <span data-ttu-id="34657-172">方法</span><span class="sxs-lookup"><span data-stu-id="34657-172">Method</span></span> |
| [<span data-ttu-id="34657-173">getEntities</span><span class="sxs-lookup"><span data-stu-id="34657-173">getEntities</span></span>](#getentities--entitiesjavascriptapioutlook15officeentities) | <span data-ttu-id="34657-174">方法</span><span class="sxs-lookup"><span data-stu-id="34657-174">Method</span></span> |
| [<span data-ttu-id="34657-175">getEntitiesByType</span><span class="sxs-lookup"><span data-stu-id="34657-175">getEntitiesByType</span></span>](#getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlook15officecontactmeetingsuggestionjavascriptapioutlook15officemeetingsuggestionphonenumberjavascriptapioutlook15officephonenumbertasksuggestionjavascriptapioutlook15officetasksuggestion) | <span data-ttu-id="34657-176">方法</span><span class="sxs-lookup"><span data-stu-id="34657-176">Method</span></span> |
| [<span data-ttu-id="34657-177">getFilteredEntitiesByName</span><span class="sxs-lookup"><span data-stu-id="34657-177">getFilteredEntitiesByName</span></span>](#getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlook15officecontactmeetingsuggestionjavascriptapioutlook15officemeetingsuggestionphonenumberjavascriptapioutlook15officephonenumbertasksuggestionjavascriptapioutlook15officetasksuggestion) | <span data-ttu-id="34657-178">方法</span><span class="sxs-lookup"><span data-stu-id="34657-178">Method</span></span> |
| [<span data-ttu-id="34657-179">getRegExMatches</span><span class="sxs-lookup"><span data-stu-id="34657-179">getRegExMatches</span></span>](#getregexmatches--object) | <span data-ttu-id="34657-180">方法</span><span class="sxs-lookup"><span data-stu-id="34657-180">Method</span></span> |
| [<span data-ttu-id="34657-181">getRegExMatchesByName</span><span class="sxs-lookup"><span data-stu-id="34657-181">getRegExMatchesByName</span></span>](#getregexmatchesbynamename--nullable-array-string-) | <span data-ttu-id="34657-182">方法</span><span class="sxs-lookup"><span data-stu-id="34657-182">Method</span></span> |
| [<span data-ttu-id="34657-183">getSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="34657-183">getSelectedDataAsync</span></span>](#getselecteddataasynccoerciontype-options-callback--string) | <span data-ttu-id="34657-184">方法</span><span class="sxs-lookup"><span data-stu-id="34657-184">Method</span></span> |
| [<span data-ttu-id="34657-185">loadCustomPropertiesAsync</span><span class="sxs-lookup"><span data-stu-id="34657-185">loadCustomPropertiesAsync</span></span>](#loadcustompropertiesasynccallback-usercontext) | <span data-ttu-id="34657-186">方法</span><span class="sxs-lookup"><span data-stu-id="34657-186">Method</span></span> |
| [<span data-ttu-id="34657-187">removeAttachmentAsync</span><span class="sxs-lookup"><span data-stu-id="34657-187">removeAttachmentAsync</span></span>](#removeattachmentasyncattachmentid-options-callback) | <span data-ttu-id="34657-188">方法</span><span class="sxs-lookup"><span data-stu-id="34657-188">Method</span></span> |
| [<span data-ttu-id="34657-189">saveAsync</span><span class="sxs-lookup"><span data-stu-id="34657-189">saveAsync</span></span>](#saveasyncoptions-callback) | <span data-ttu-id="34657-190">方法</span><span class="sxs-lookup"><span data-stu-id="34657-190">Method</span></span> |
| [<span data-ttu-id="34657-191">setSelectedDataAsync</span><span class="sxs-lookup"><span data-stu-id="34657-191">setSelectedDataAsync</span></span>](#setselecteddataasyncdata-options-callback) | <span data-ttu-id="34657-192">方法</span><span class="sxs-lookup"><span data-stu-id="34657-192">Method</span></span> |

### <a name="example"></a><span data-ttu-id="34657-193">示例</span><span class="sxs-lookup"><span data-stu-id="34657-193">Example</span></span>

<span data-ttu-id="34657-194">以下 JavaScript 代码示例显示了如何访问 Outlook 中当前项目的 `subject` 属性。</span><span class="sxs-lookup"><span data-stu-id="34657-194">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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

### <a name="members"></a><span data-ttu-id="34657-195">成员</span><span class="sxs-lookup"><span data-stu-id="34657-195">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlook15officeattachmentdetails"></a><span data-ttu-id="34657-196">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_5/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="34657-196">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_5/office.attachmentdetails)></span></span>

<span data-ttu-id="34657-p102">获取项目的附件数组。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="34657-p102">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="34657-199">某些类型的文件因潜在的安全问题被 Outlook 阻止，因此没有返回。</span><span class="sxs-lookup"><span data-stu-id="34657-199">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="34657-200">有关详细信息，请参阅 [在 Outlook 中被阻止的附件](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519)。</span><span class="sxs-lookup"><span data-stu-id="34657-200">For more information see [Payments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="34657-201">类型：</span><span class="sxs-lookup"><span data-stu-id="34657-201">Type:</span></span>

*   <span data-ttu-id="34657-202">数组。 <[AttachmentDetails](/javascript/api/outlook_1_5/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="34657-202">Array.<[AttachmentDetails](/javascript/api/outlook_1_5/office.attachmentdetails)></span></span>

##### <a name="requirements"></a><span data-ttu-id="34657-203">要求</span><span class="sxs-lookup"><span data-stu-id="34657-203">Requirements</span></span>

|<span data-ttu-id="34657-204">要求</span><span class="sxs-lookup"><span data-stu-id="34657-204">Requirement</span></span>| <span data-ttu-id="34657-205">值</span><span class="sxs-lookup"><span data-stu-id="34657-205">Value</span></span>|
|---|---|
|[<span data-ttu-id="34657-206">最低邮箱要求集版本</span><span class="sxs-lookup"><span data-stu-id="34657-206">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="34657-207">1.0</span><span class="sxs-lookup"><span data-stu-id="34657-207">1.0</span></span>|
|[<span data-ttu-id="34657-208">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="34657-208">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="34657-209">ReadItem</span><span class="sxs-lookup"><span data-stu-id="34657-209">ReadItem</span></span>|
|[<span data-ttu-id="34657-210">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="34657-210">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="34657-211">阅读</span><span class="sxs-lookup"><span data-stu-id="34657-211">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="34657-212">示例</span><span class="sxs-lookup"><span data-stu-id="34657-212">Example</span></span>

<span data-ttu-id="34657-213">以下代码使用当前项目上所有附件的详细信息构成 HTML 字符串。</span><span class="sxs-lookup"><span data-stu-id="34657-213">The following code builds an HTML string with details of all attachments on the current item.</span></span>

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

####  <a name="bcc-recipientsjavascriptapioutlook15officerecipients"></a><span data-ttu-id="34657-214">密件抄送：[收件人](/javascript/api/outlook_1_5/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="34657-214">bcc :[Recipients](/javascript/api/outlook_1_5/office.recipients)</span></span>

<span data-ttu-id="34657-215">获取一个对象，提供用于获取或更新邮件的密件抄送 （密件抄送副本） 行上收件人的方法。</span><span class="sxs-lookup"><span data-stu-id="34657-215">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="34657-216">仅限撰写模式。</span><span class="sxs-lookup"><span data-stu-id="34657-216">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="34657-217">类型：</span><span class="sxs-lookup"><span data-stu-id="34657-217">Type:</span></span>

*   [<span data-ttu-id="34657-218">收件人</span><span class="sxs-lookup"><span data-stu-id="34657-218">Recipients</span></span>](/javascript/api/outlook_1_5/office.recipients)

##### <a name="requirements"></a><span data-ttu-id="34657-219">要求</span><span class="sxs-lookup"><span data-stu-id="34657-219">Requirements</span></span>

|<span data-ttu-id="34657-220">要求</span><span class="sxs-lookup"><span data-stu-id="34657-220">Requirement</span></span>| <span data-ttu-id="34657-221">值</span><span class="sxs-lookup"><span data-stu-id="34657-221">Value</span></span>|
|---|---|
|[<span data-ttu-id="34657-222">最低邮箱要求集版本</span><span class="sxs-lookup"><span data-stu-id="34657-222">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="34657-223">1.1</span><span class="sxs-lookup"><span data-stu-id="34657-223">1.1</span></span>|
|[<span data-ttu-id="34657-224">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="34657-224">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="34657-225">ReadItem</span><span class="sxs-lookup"><span data-stu-id="34657-225">ReadItem</span></span>|
|[<span data-ttu-id="34657-226">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="34657-226">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="34657-227">撰写</span><span class="sxs-lookup"><span data-stu-id="34657-227">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="34657-228">示例</span><span class="sxs-lookup"><span data-stu-id="34657-228">Example</span></span>

```
Office.context.mailbox.item.bcc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.bcc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.bcc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfBccRecipients = asyncResult.value;
}
```

####  <a name="body-bodyjavascriptapioutlook15officebody"></a><span data-ttu-id="34657-229">body :[Body](/javascript/api/outlook_1_5/office.body)</span><span class="sxs-lookup"><span data-stu-id="34657-229">body :[Body](/javascript/api/outlook_1_5/office.body)</span></span>

<span data-ttu-id="34657-230">获取一个提供用于处理项目正文的方法的对象。</span><span class="sxs-lookup"><span data-stu-id="34657-230">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="34657-231">类型：</span><span class="sxs-lookup"><span data-stu-id="34657-231">Type:</span></span>

*   [<span data-ttu-id="34657-232">Body</span><span class="sxs-lookup"><span data-stu-id="34657-232">Body</span></span>](/javascript/api/outlook_1_5/office.body)

##### <a name="requirements"></a><span data-ttu-id="34657-233">要求</span><span class="sxs-lookup"><span data-stu-id="34657-233">Requirements</span></span>

|<span data-ttu-id="34657-234">要求</span><span class="sxs-lookup"><span data-stu-id="34657-234">Requirement</span></span>| <span data-ttu-id="34657-235">值</span><span class="sxs-lookup"><span data-stu-id="34657-235">Value</span></span>|
|---|---|
|[<span data-ttu-id="34657-236">最低的邮箱要求集版本</span><span class="sxs-lookup"><span data-stu-id="34657-236">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="34657-237">1.1</span><span class="sxs-lookup"><span data-stu-id="34657-237">1.1</span></span>|
|[<span data-ttu-id="34657-238">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="34657-238">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="34657-239">ReadItem</span><span class="sxs-lookup"><span data-stu-id="34657-239">ReadItem</span></span>|
|[<span data-ttu-id="34657-240">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="34657-240">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="34657-241">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="34657-241">Compose or read</span></span>|

####  <a name="cc-arrayemailaddressdetailsjavascriptapioutlook15officeemailaddressdetailsrecipientsjavascriptapioutlook15officerecipients"></a><span data-ttu-id="34657-242">抄送： 数组。 <[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)>|[收件人](/javascript/api/outlook_1_5/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="34657-242">cc :Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_5/office.recipients)</span></span>

<span data-ttu-id="34657-243">提供对邮件抄送 (cc) 收件人的访问。</span><span class="sxs-lookup"><span data-stu-id="34657-243">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="34657-244">对象的类型和访问级别取决于当前项的模式。</span><span class="sxs-lookup"><span data-stu-id="34657-244">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="34657-245">阅读模式</span><span class="sxs-lookup"><span data-stu-id="34657-245">Read mode</span></span>

<span data-ttu-id="34657-p106">`cc`属性返回包含邮件的**抄送**行上所列每个收件人的 `EmailAddressDetails` 对象数组。集合上限为 100 个成员。</span><span class="sxs-lookup"><span data-stu-id="34657-p106">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="34657-248">撰写模式</span><span class="sxs-lookup"><span data-stu-id="34657-248">Compose mode</span></span>

<span data-ttu-id="34657-249">`cc` 属性返回 `Recipients` 对象，该对象提供用于处理邮件**抄送**行上收件人的方法。</span><span class="sxs-lookup"><span data-stu-id="34657-249">The `cc` property returns a `Recipients` object that provides methods for manipulating the recipients on the **Cc** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="34657-250">类型：</span><span class="sxs-lookup"><span data-stu-id="34657-250">Type:</span></span>

*   <span data-ttu-id="34657-251">数组。 <[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)> |[收件人](/javascript/api/outlook_1_5/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="34657-251">Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_5/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="34657-252">要求</span><span class="sxs-lookup"><span data-stu-id="34657-252">Requirements</span></span>

|<span data-ttu-id="34657-253">要求</span><span class="sxs-lookup"><span data-stu-id="34657-253">Requirement</span></span>| <span data-ttu-id="34657-254">值</span><span class="sxs-lookup"><span data-stu-id="34657-254">Value</span></span>|
|---|---|
|[<span data-ttu-id="34657-255">最低邮箱要求集版本</span><span class="sxs-lookup"><span data-stu-id="34657-255">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="34657-256">1.0</span><span class="sxs-lookup"><span data-stu-id="34657-256">1.0</span></span>|
|[<span data-ttu-id="34657-257">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="34657-257">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="34657-258">ReadItem</span><span class="sxs-lookup"><span data-stu-id="34657-258">ReadItem</span></span>|
|[<span data-ttu-id="34657-259">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="34657-259">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="34657-260">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="34657-260">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="34657-261">示例</span><span class="sxs-lookup"><span data-stu-id="34657-261">Example</span></span>

```
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

####  <a name="nullable-conversationid-string"></a><span data-ttu-id="34657-262">(nullable) conversationId :String</span><span class="sxs-lookup"><span data-stu-id="34657-262">(nullable) conversationId :String</span></span>

<span data-ttu-id="34657-263">获取包含特定消息的电子邮件会话的标识符。</span><span class="sxs-lookup"><span data-stu-id="34657-263">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="34657-p107">如果在阅读窗体或撰写窗体的回复中激活邮件应用程序，则此属性可以获得一个整数值。如果用户随后更改了回复邮件的主题（若发送回复），则该邮件的对话 ID 将改变且之前获取的值将不适用。</span><span class="sxs-lookup"><span data-stu-id="34657-p107">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="34657-p108">对于撰写窗体的新项目，此属性获得一个 null 值。如果用户设置一个主题并保存该项目，`conversationId` 属性将返回一个值。</span><span class="sxs-lookup"><span data-stu-id="34657-p108">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="34657-268">类型：</span><span class="sxs-lookup"><span data-stu-id="34657-268">Type:</span></span>

*   <span data-ttu-id="34657-269">String</span><span class="sxs-lookup"><span data-stu-id="34657-269">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="34657-270">要求</span><span class="sxs-lookup"><span data-stu-id="34657-270">Requirements</span></span>

|<span data-ttu-id="34657-271">要求</span><span class="sxs-lookup"><span data-stu-id="34657-271">Requirement</span></span>| <span data-ttu-id="34657-272">值</span><span class="sxs-lookup"><span data-stu-id="34657-272">Value</span></span>|
|---|---|
|[<span data-ttu-id="34657-273">最低邮箱要求集版本</span><span class="sxs-lookup"><span data-stu-id="34657-273">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="34657-274">1.0</span><span class="sxs-lookup"><span data-stu-id="34657-274">1.0</span></span>|
|[<span data-ttu-id="34657-275">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="34657-275">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="34657-276">ReadItem</span><span class="sxs-lookup"><span data-stu-id="34657-276">ReadItem</span></span>|
|[<span data-ttu-id="34657-277">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="34657-277">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="34657-278">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="34657-278">Compose or read</span></span>|

#### <a name="datetimecreated-date"></a><span data-ttu-id="34657-279">dateTimeCreated：日期</span><span class="sxs-lookup"><span data-stu-id="34657-279">dateTimeCreated :Date</span></span>

<span data-ttu-id="34657-p109">获取项目创建的日期和时间。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="34657-p109">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="34657-282">类型：</span><span class="sxs-lookup"><span data-stu-id="34657-282">Type:</span></span>

*   <span data-ttu-id="34657-283">Date</span><span class="sxs-lookup"><span data-stu-id="34657-283">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="34657-284">要求</span><span class="sxs-lookup"><span data-stu-id="34657-284">Requirements</span></span>

|<span data-ttu-id="34657-285">要求</span><span class="sxs-lookup"><span data-stu-id="34657-285">Requirement</span></span>| <span data-ttu-id="34657-286">值</span><span class="sxs-lookup"><span data-stu-id="34657-286">Value</span></span>|
|---|---|
|[<span data-ttu-id="34657-287">最低邮箱要求集版本</span><span class="sxs-lookup"><span data-stu-id="34657-287">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="34657-288">1.0</span><span class="sxs-lookup"><span data-stu-id="34657-288">1.0</span></span>|
|[<span data-ttu-id="34657-289">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="34657-289">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="34657-290">ReadItem</span><span class="sxs-lookup"><span data-stu-id="34657-290">ReadItem</span></span>|
|[<span data-ttu-id="34657-291">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="34657-291">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="34657-292">阅读</span><span class="sxs-lookup"><span data-stu-id="34657-292">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="34657-293">示例</span><span class="sxs-lookup"><span data-stu-id="34657-293">Example</span></span>

```
var created = Office.context.mailbox.item.dateTimeCreated;
```

#### <a name="datetimemodified-date"></a><span data-ttu-id="34657-294">dateTimeModified： 日期</span><span class="sxs-lookup"><span data-stu-id="34657-294">dateTimeModified :Date</span></span>

<span data-ttu-id="34657-p110">获取项目最近一次修改的日期和时间。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="34657-p110">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="34657-297">注意：在 iOS 版 Outlook 或  Android 版 Outlook 中不支持此成员。</span><span class="sxs-lookup"><span data-stu-id="34657-297">Note: This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="type"></a><span data-ttu-id="34657-298">类型：</span><span class="sxs-lookup"><span data-stu-id="34657-298">Type:</span></span>

*   <span data-ttu-id="34657-299">Date</span><span class="sxs-lookup"><span data-stu-id="34657-299">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="34657-300">要求</span><span class="sxs-lookup"><span data-stu-id="34657-300">Requirements</span></span>

|<span data-ttu-id="34657-301">要求</span><span class="sxs-lookup"><span data-stu-id="34657-301">Requirement</span></span>| <span data-ttu-id="34657-302">值</span><span class="sxs-lookup"><span data-stu-id="34657-302">Value</span></span>|
|---|---|
|[<span data-ttu-id="34657-303">最低邮箱要求集版本</span><span class="sxs-lookup"><span data-stu-id="34657-303">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="34657-304">1.0</span><span class="sxs-lookup"><span data-stu-id="34657-304">1.0</span></span>|
|[<span data-ttu-id="34657-305">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="34657-305">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="34657-306">ReadItem</span><span class="sxs-lookup"><span data-stu-id="34657-306">ReadItem</span></span>|
|[<span data-ttu-id="34657-307">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="34657-307">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="34657-308">阅读</span><span class="sxs-lookup"><span data-stu-id="34657-308">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="34657-309">示例</span><span class="sxs-lookup"><span data-stu-id="34657-309">Example</span></span>

```
var modified = Office.context.mailbox.item.dateTimeModified;
```

####  <a name="end-datetimejavascriptapioutlook15officetime"></a><span data-ttu-id="34657-310">end :Date|[Time](/javascript/api/outlook_1_5/office.time)</span><span class="sxs-lookup"><span data-stu-id="34657-310">end :Date|[Time](/javascript/api/outlook_1_5/office.time)</span></span>

<span data-ttu-id="34657-311">获取或设置约会结束的日期和时间。</span><span class="sxs-lookup"><span data-stu-id="34657-311">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="34657-p111">将 `end` 属性表示为协调世界时 (UTC) 的日期和时间值。可使用 [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook15officelocalclienttime) 方法将 end 属性值转换为客户端的本地日期和时间。</span><span class="sxs-lookup"><span data-stu-id="34657-p111">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook15officelocalclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="34657-314">阅读模式</span><span class="sxs-lookup"><span data-stu-id="34657-314">Read mode</span></span>

<span data-ttu-id="34657-315">`end` 属性返回 `Date` 对象。</span><span class="sxs-lookup"><span data-stu-id="34657-315">The `end` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="34657-316">撰写模式</span><span class="sxs-lookup"><span data-stu-id="34657-316">Compose mode</span></span>

<span data-ttu-id="34657-317">`end` 属性返回 `Time` 对象。</span><span class="sxs-lookup"><span data-stu-id="34657-317">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="34657-318">使用  方法设置结束时间时，应使用  方法将客户端的本地时间转换为服务器的 UTC。 [`Time.setAsync`](/javascript/api/outlook_1_5/office.time#setasync-datetime--options--callback-) [ `convertToUtcClientTime`  ](office.context.mailbox.md#converttoutcclienttimeinput--date)</span><span class="sxs-lookup"><span data-stu-id="34657-318">When you use the [`Time.setAsync`](/javascript/api/outlook_1_5/office.time#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="34657-319">类型：</span><span class="sxs-lookup"><span data-stu-id="34657-319">Type:</span></span>

*   <span data-ttu-id="34657-320">日期 | [时间](/javascript/api/outlook_1_5/office.time)</span><span class="sxs-lookup"><span data-stu-id="34657-320">Date | [Time](/javascript/api/outlook_1_5/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="34657-321">要求</span><span class="sxs-lookup"><span data-stu-id="34657-321">Requirements</span></span>

|<span data-ttu-id="34657-322">要求</span><span class="sxs-lookup"><span data-stu-id="34657-322">Requirement</span></span>| <span data-ttu-id="34657-323">值</span><span class="sxs-lookup"><span data-stu-id="34657-323">Value</span></span>|
|---|---|
|[<span data-ttu-id="34657-324">最低邮箱要求集版本</span><span class="sxs-lookup"><span data-stu-id="34657-324">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="34657-325">1.0</span><span class="sxs-lookup"><span data-stu-id="34657-325">1.0</span></span>|
|[<span data-ttu-id="34657-326">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="34657-326">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="34657-327">ReadItem</span><span class="sxs-lookup"><span data-stu-id="34657-327">ReadItem</span></span>|
|[<span data-ttu-id="34657-328">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="34657-328">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="34657-329">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="34657-329">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="34657-330">示例</span><span class="sxs-lookup"><span data-stu-id="34657-330">Example</span></span>

<span data-ttu-id="34657-331">以下示例通过使用 `Time` 对象的 [`setAsync`](/javascript/api/outlook_1_5/office.time#setasync-datetime--options--callback-) 方法，设置撰写模式下约会的结束时间。</span><span class="sxs-lookup"><span data-stu-id="34657-331">The following example sets the end time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_5/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

#### <a name="from-emailaddressdetailsjavascriptapioutlook15officeemailaddressdetails"></a><span data-ttu-id="34657-332">从：[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="34657-332">from :[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)</span></span>

<span data-ttu-id="34657-p112">获取邮件发件人的电子邮件地址。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="34657-p112">Gets the email address of the sender of a message. Read mode only.</span></span>

<span data-ttu-id="34657-p113">`from` 和 [`sender`](#sender-emailaddressdetailsjavascriptapioutlook15officeemailaddressdetails) 属性表示同一个人，邮件由代理人发送的除外。在此情况下，`from` 属性表示代理程序，而 sender 属性表示代理人。</span><span class="sxs-lookup"><span data-stu-id="34657-p113">The `from` and [`sender`](#sender-emailaddressdetailsjavascriptapioutlook15officeemailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="34657-337">`EmailAddressDetails` 对象的 `recipientType` 属性 在 `from` 属性是 `undefined`。</span><span class="sxs-lookup"><span data-stu-id="34657-337">Note: The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="34657-338">类型：</span><span class="sxs-lookup"><span data-stu-id="34657-338">Type:</span></span>

*   [<span data-ttu-id="34657-339">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="34657-339">EmailAddressDetails</span></span>](/javascript/api/outlook_1_5/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="34657-340">要求</span><span class="sxs-lookup"><span data-stu-id="34657-340">Requirements</span></span>

|<span data-ttu-id="34657-341">要求</span><span class="sxs-lookup"><span data-stu-id="34657-341">Requirement</span></span>| <span data-ttu-id="34657-342">值</span><span class="sxs-lookup"><span data-stu-id="34657-342">Value</span></span>|
|---|---|
|[<span data-ttu-id="34657-343">最低邮箱要求集版本</span><span class="sxs-lookup"><span data-stu-id="34657-343">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="34657-344">1.0</span><span class="sxs-lookup"><span data-stu-id="34657-344">1.0</span></span>|
|[<span data-ttu-id="34657-345">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="34657-345">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="34657-346">ReadItem</span><span class="sxs-lookup"><span data-stu-id="34657-346">ReadItem</span></span>|
|[<span data-ttu-id="34657-347">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="34657-347">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="34657-348">阅读</span><span class="sxs-lookup"><span data-stu-id="34657-348">Read</span></span>|

#### <a name="internetmessageid-string"></a><span data-ttu-id="34657-349">internetMessageId :String</span><span class="sxs-lookup"><span data-stu-id="34657-349">internetMessageId :String</span></span>

<span data-ttu-id="34657-p114">获取电子邮件的 Internet 消息标识符。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="34657-p114">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="34657-352">类型：</span><span class="sxs-lookup"><span data-stu-id="34657-352">Type:</span></span>

*   <span data-ttu-id="34657-353">String</span><span class="sxs-lookup"><span data-stu-id="34657-353">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="34657-354">要求</span><span class="sxs-lookup"><span data-stu-id="34657-354">Requirements</span></span>

|<span data-ttu-id="34657-355">要求</span><span class="sxs-lookup"><span data-stu-id="34657-355">Requirement</span></span>| <span data-ttu-id="34657-356">值</span><span class="sxs-lookup"><span data-stu-id="34657-356">Value</span></span>|
|---|---|
|[<span data-ttu-id="34657-357">最低邮箱要求集版本</span><span class="sxs-lookup"><span data-stu-id="34657-357">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="34657-358">1.0</span><span class="sxs-lookup"><span data-stu-id="34657-358">1.0</span></span>|
|[<span data-ttu-id="34657-359">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="34657-359">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="34657-360">ReadItem</span><span class="sxs-lookup"><span data-stu-id="34657-360">ReadItem</span></span>|
|[<span data-ttu-id="34657-361">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="34657-361">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="34657-362">阅读</span><span class="sxs-lookup"><span data-stu-id="34657-362">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="34657-363">示例</span><span class="sxs-lookup"><span data-stu-id="34657-363">Example</span></span>

```
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

#### <a name="itemclass-string"></a><span data-ttu-id="34657-364">itemClass :String</span><span class="sxs-lookup"><span data-stu-id="34657-364">itemClass :String</span></span>

<span data-ttu-id="34657-p115">获取选定项目的 Exchange Web 服务项目类。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="34657-p115">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="34657-p116">`itemClass` 属性指定所选项目的邮件类别。以下是邮件或约会项目的默认邮件类别。</span><span class="sxs-lookup"><span data-stu-id="34657-p116">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

| <span data-ttu-id="34657-369">类型</span><span class="sxs-lookup"><span data-stu-id="34657-369">Type</span></span> | <span data-ttu-id="34657-370">说明</span><span class="sxs-lookup"><span data-stu-id="34657-370">Description</span></span> | <span data-ttu-id="34657-371">项目类</span><span class="sxs-lookup"><span data-stu-id="34657-371">item class</span></span> |
| --- | --- | --- |
| <span data-ttu-id="34657-372">约会项目</span><span class="sxs-lookup"><span data-stu-id="34657-372">Appointment items</span></span> | <span data-ttu-id="34657-373">这些是项目类 `IPM.Appointment` 或 `IPM.Appointment.Occurence` 的日历项目。</span><span class="sxs-lookup"><span data-stu-id="34657-373">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurence`.</span></span> | `IPM.Appointment`<br />`IPM.Appointment.Occurence` |
| <span data-ttu-id="34657-374">邮件项目</span><span class="sxs-lookup"><span data-stu-id="34657-374">Message items</span></span> | <span data-ttu-id="34657-375">这些项目包括具有默认邮件类别 `IPM.Note` 的电子邮件，以及使用 `IPM.Schedule.Meeting` 作为基础邮件类别的会议请求、响应和取消。</span><span class="sxs-lookup"><span data-stu-id="34657-375">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span> | `IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled` |

<span data-ttu-id="34657-376">你可以创建用于扩展默认邮件类别的自定义邮件类别，例如，自定义约会邮件类别 `IPM.Appointment.Contoso`。</span><span class="sxs-lookup"><span data-stu-id="34657-376">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="34657-377">类型：</span><span class="sxs-lookup"><span data-stu-id="34657-377">Type:</span></span>

*   <span data-ttu-id="34657-378">String</span><span class="sxs-lookup"><span data-stu-id="34657-378">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="34657-379">要求</span><span class="sxs-lookup"><span data-stu-id="34657-379">Requirements</span></span>

|<span data-ttu-id="34657-380">要求</span><span class="sxs-lookup"><span data-stu-id="34657-380">Requirement</span></span>| <span data-ttu-id="34657-381">值</span><span class="sxs-lookup"><span data-stu-id="34657-381">Value</span></span>|
|---|---|
|[<span data-ttu-id="34657-382">最低邮箱要求集版本</span><span class="sxs-lookup"><span data-stu-id="34657-382">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="34657-383">1.0</span><span class="sxs-lookup"><span data-stu-id="34657-383">1.0</span></span>|
|[<span data-ttu-id="34657-384">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="34657-384">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="34657-385">ReadItem</span><span class="sxs-lookup"><span data-stu-id="34657-385">ReadItem</span></span>|
|[<span data-ttu-id="34657-386">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="34657-386">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="34657-387">阅读</span><span class="sxs-lookup"><span data-stu-id="34657-387">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="34657-388">示例</span><span class="sxs-lookup"><span data-stu-id="34657-388">Example</span></span>

```
var itemClass = Office.context.mailbox.item.itemClass;
```

#### <a name="nullable-itemid-string"></a><span data-ttu-id="34657-389">（可空类型）itemId ：字符串</span><span class="sxs-lookup"><span data-stu-id="34657-389">(nullable) itemId :String</span></span>

<span data-ttu-id="34657-p117">获取当前项目的 Exchange Web 服务项目标识符。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="34657-p117">Gets the Exchange Web Services item identifier for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="34657-392">`itemId` 属性返回的标识符与 Exchange Web 服务项目标识符相同。</span><span class="sxs-lookup"><span data-stu-id="34657-392">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="34657-393">`itemId` 属性与 Outlook 条目 ID 或使用 Outlook REST API 的 ID不同。</span><span class="sxs-lookup"><span data-stu-id="34657-393">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="34657-394">使用此值的 REST API 调用之前，应使用 [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string)将其转换。</span><span class="sxs-lookup"><span data-stu-id="34657-394">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="34657-395">有关详细信息，请参阅 [从 Outlook 外接程序使用 Outlook REST API](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id)。</span><span class="sxs-lookup"><span data-stu-id="34657-395">For more details, see [Use the Outlook REST APIs from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

<span data-ttu-id="34657-p119">`itemId` 属性在撰写模式下不可用。如果需要项目标识符，[`saveAsync`](#saveasyncoptions-callback) 方法可用于将项目保存到存储，这将在回调函数的 [`AsyncResult.value`](/javascript/api/office/office.asyncresult) 参数中返回项目标识符。</span><span class="sxs-lookup"><span data-stu-id="34657-p119">The `itemId` property is not available in compose mode. If an item identifier is required, the [`saveAsync`](#saveasyncoptions-callback) method can be used to save the item to the store, which will return the item identifier in the [`AsyncResult.value`](/javascript/api/office/office.asyncresult) parameter in the callback function.</span></span>

##### <a name="type"></a><span data-ttu-id="34657-398">类型：</span><span class="sxs-lookup"><span data-stu-id="34657-398">Type:</span></span>

*   <span data-ttu-id="34657-399">String</span><span class="sxs-lookup"><span data-stu-id="34657-399">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="34657-400">要求</span><span class="sxs-lookup"><span data-stu-id="34657-400">Requirements</span></span>

|<span data-ttu-id="34657-401">要求</span><span class="sxs-lookup"><span data-stu-id="34657-401">Requirement</span></span>| <span data-ttu-id="34657-402">值</span><span class="sxs-lookup"><span data-stu-id="34657-402">Value</span></span>|
|---|---|
|[<span data-ttu-id="34657-403">最低邮箱要求集版本</span><span class="sxs-lookup"><span data-stu-id="34657-403">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="34657-404">1.0</span><span class="sxs-lookup"><span data-stu-id="34657-404">1.0</span></span>|
|[<span data-ttu-id="34657-405">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="34657-405">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="34657-406">ReadItem</span><span class="sxs-lookup"><span data-stu-id="34657-406">ReadItem</span></span>|
|[<span data-ttu-id="34657-407">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="34657-407">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="34657-408">阅读</span><span class="sxs-lookup"><span data-stu-id="34657-408">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="34657-409">示例</span><span class="sxs-lookup"><span data-stu-id="34657-409">Example</span></span>

<span data-ttu-id="34657-p120">以下代码检查项目标识符是否存在。如果 `itemId` 属性返回 `null` 或 `undefined`，则将项目保存到存储，并从异步结果中获取项目标识符。</span><span class="sxs-lookup"><span data-stu-id="34657-p120">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

```
var itemId = Office.context.mailbox.item.itemId;
if (itemId === null || itemId == undefined) {
  Office.context.mailbox.item.saveAsync(function(result){
    itemId = result.value;
  });
}
```

####  <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlook15officemailboxenumsitemtype"></a><span data-ttu-id="34657-412">itemType:[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_5/office.mailboxenums.itemtype)</span><span class="sxs-lookup"><span data-stu-id="34657-412">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_5/office.mailboxenums.itemtype)</span></span>

<span data-ttu-id="34657-413">获取实例代表项的类型。</span><span class="sxs-lookup"><span data-stu-id="34657-413">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="34657-414">`itemType` 属性返回其中一个 `ItemType` 枚举值，指示 `item` 对象实例是邮件还是约会。</span><span class="sxs-lookup"><span data-stu-id="34657-414">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="34657-415">类型：</span><span class="sxs-lookup"><span data-stu-id="34657-415">Type:</span></span>

*   [<span data-ttu-id="34657-416">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="34657-416">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook_1_5/office.mailboxenums.itemtype)

##### <a name="requirements"></a><span data-ttu-id="34657-417">要求</span><span class="sxs-lookup"><span data-stu-id="34657-417">Requirements</span></span>

|<span data-ttu-id="34657-418">要求</span><span class="sxs-lookup"><span data-stu-id="34657-418">Requirement</span></span>| <span data-ttu-id="34657-419">值</span><span class="sxs-lookup"><span data-stu-id="34657-419">Value</span></span>|
|---|---|
|[<span data-ttu-id="34657-420">最低邮箱要求集版本</span><span class="sxs-lookup"><span data-stu-id="34657-420">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="34657-421">1.0</span><span class="sxs-lookup"><span data-stu-id="34657-421">1.0</span></span>|
|[<span data-ttu-id="34657-422">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="34657-422">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="34657-423">ReadItem</span><span class="sxs-lookup"><span data-stu-id="34657-423">ReadItem</span></span>|
|[<span data-ttu-id="34657-424">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="34657-424">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="34657-425">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="34657-425">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="34657-426">示例</span><span class="sxs-lookup"><span data-stu-id="34657-426">Example</span></span>

```
if (Office.context.mailbox.item.itemType == Office.MailboxEnums.ItemType.Message)
  // do something
else
  // do something else
```

####  <a name="location-stringlocationjavascriptapioutlook15officelocation"></a><span data-ttu-id="34657-427">location :String|[Location](/javascript/api/outlook_1_5/office.location)</span><span class="sxs-lookup"><span data-stu-id="34657-427">location :String|[Location](/javascript/api/outlook_1_5/office.location)</span></span>

<span data-ttu-id="34657-428">获取或设置约会的位置。</span><span class="sxs-lookup"><span data-stu-id="34657-428">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="34657-429">阅读模式</span><span class="sxs-lookup"><span data-stu-id="34657-429">Read mode</span></span>

<span data-ttu-id="34657-430">`location` 属性返回一个包含约会位置的字符串。</span><span class="sxs-lookup"><span data-stu-id="34657-430">The `location` property returns a string that contains the location of the appointment.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="34657-431">撰写模式</span><span class="sxs-lookup"><span data-stu-id="34657-431">Compose mode</span></span>

<span data-ttu-id="34657-432">`location` 属性返回一个 `Location` 对象，该对象提供用于获取和设置约会位置的方法。</span><span class="sxs-lookup"><span data-stu-id="34657-432">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="34657-433">类型：</span><span class="sxs-lookup"><span data-stu-id="34657-433">Type:</span></span>

*   <span data-ttu-id="34657-434">字符串 | [位置](/javascript/api/outlook_1_5/office.location)</span><span class="sxs-lookup"><span data-stu-id="34657-434">String | [Location](/javascript/api/outlook_1_5/office.location)</span></span>

##### <a name="requirements"></a><span data-ttu-id="34657-435">要求</span><span class="sxs-lookup"><span data-stu-id="34657-435">Requirements</span></span>

|<span data-ttu-id="34657-436">要求</span><span class="sxs-lookup"><span data-stu-id="34657-436">Requirement</span></span>| <span data-ttu-id="34657-437">值</span><span class="sxs-lookup"><span data-stu-id="34657-437">Value</span></span>|
|---|---|
|[<span data-ttu-id="34657-438">最低邮箱要求集版本</span><span class="sxs-lookup"><span data-stu-id="34657-438">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="34657-439">1.0</span><span class="sxs-lookup"><span data-stu-id="34657-439">1.0</span></span>|
|[<span data-ttu-id="34657-440">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="34657-440">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="34657-441">ReadItem</span><span class="sxs-lookup"><span data-stu-id="34657-441">ReadItem</span></span>|
|[<span data-ttu-id="34657-442">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="34657-442">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="34657-443">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="34657-443">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="34657-444">示例</span><span class="sxs-lookup"><span data-stu-id="34657-444">Example</span></span>

```
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

#### <a name="normalizedsubject-string"></a><span data-ttu-id="34657-445">normalizedSubject :String</span><span class="sxs-lookup"><span data-stu-id="34657-445">normalizedSubject :String</span></span>

<span data-ttu-id="34657-p121">获取删除了所有前缀（包括 `RE:` 和 `FWD:`）的项目主题。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="34657-p121">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="34657-p122">normalizedSubject 属性获取包含由电子邮件程序添加的任何标准前缀（如 `RE:` 和 `FW:`）的项目主题。若要获取包含完整前缀的项目主题，请使用 [`subject`](#subject-stringsubjectjavascriptapioutlook15officesubject) 属性。</span><span class="sxs-lookup"><span data-stu-id="34657-p122">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubjectjavascriptapioutlook15officesubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="34657-450">类型：</span><span class="sxs-lookup"><span data-stu-id="34657-450">Type:</span></span>

*   <span data-ttu-id="34657-451">String</span><span class="sxs-lookup"><span data-stu-id="34657-451">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="34657-452">要求</span><span class="sxs-lookup"><span data-stu-id="34657-452">Requirements</span></span>

|<span data-ttu-id="34657-453">要求</span><span class="sxs-lookup"><span data-stu-id="34657-453">Requirement</span></span>| <span data-ttu-id="34657-454">值</span><span class="sxs-lookup"><span data-stu-id="34657-454">Value</span></span>|
|---|---|
|[<span data-ttu-id="34657-455">最低邮箱要求集版本</span><span class="sxs-lookup"><span data-stu-id="34657-455">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="34657-456">1.0</span><span class="sxs-lookup"><span data-stu-id="34657-456">1.0</span></span>|
|[<span data-ttu-id="34657-457">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="34657-457">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="34657-458">ReadItem</span><span class="sxs-lookup"><span data-stu-id="34657-458">ReadItem</span></span>|
|[<span data-ttu-id="34657-459">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="34657-459">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="34657-460">阅读</span><span class="sxs-lookup"><span data-stu-id="34657-460">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="34657-461">示例</span><span class="sxs-lookup"><span data-stu-id="34657-461">Example</span></span>

```
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
```

####  <a name="notificationmessages-notificationmessagesjavascriptapioutlook15officenotificationmessages"></a><span data-ttu-id="34657-462">notificationMessages:[NotificationMessages](/javascript/api/outlook_1_5/office.notificationmessages)</span><span class="sxs-lookup"><span data-stu-id="34657-462">notificationMessages :[NotificationMessages](/javascript/api/outlook_1_5/office.notificationmessages)</span></span>

<span data-ttu-id="34657-463">获取一个项目的通知邮件。</span><span class="sxs-lookup"><span data-stu-id="34657-463">Gets the notification messages for an item.</span></span>

##### <a name="type"></a><span data-ttu-id="34657-464">类型：</span><span class="sxs-lookup"><span data-stu-id="34657-464">Type:</span></span>

*   [<span data-ttu-id="34657-465">NotificationMessages</span><span class="sxs-lookup"><span data-stu-id="34657-465">NotificationMessages</span></span>](/javascript/api/outlook_1_5/office.notificationmessages)

##### <a name="requirements"></a><span data-ttu-id="34657-466">要求</span><span class="sxs-lookup"><span data-stu-id="34657-466">Requirements</span></span>

|<span data-ttu-id="34657-467">要求</span><span class="sxs-lookup"><span data-stu-id="34657-467">Requirement</span></span>| <span data-ttu-id="34657-468">值</span><span class="sxs-lookup"><span data-stu-id="34657-468">Value</span></span>|
|---|---|
|[<span data-ttu-id="34657-469">最低邮箱要求集版本</span><span class="sxs-lookup"><span data-stu-id="34657-469">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="34657-470">1.3</span><span class="sxs-lookup"><span data-stu-id="34657-470">1.3</span></span>|
|[<span data-ttu-id="34657-471">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="34657-471">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="34657-472">ReadItem</span><span class="sxs-lookup"><span data-stu-id="34657-472">ReadItem</span></span>|
|[<span data-ttu-id="34657-473">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="34657-473">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="34657-474">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="34657-474">Compose or read</span></span>|

####  <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlook15officeemailaddressdetailsrecipientsjavascriptapioutlook15officerecipients"></a><span data-ttu-id="34657-475">optionalAttendees :数组.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)>|[收件人](/javascript/api/outlook_1_5/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="34657-475">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_5/office.recipients)</span></span>

<span data-ttu-id="34657-476">提供对事件可选与会者的访问。</span><span class="sxs-lookup"><span data-stu-id="34657-476">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="34657-477">对象的类型和访问级别取决于当前项的模式。</span><span class="sxs-lookup"><span data-stu-id="34657-477">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="34657-478">阅读模式</span><span class="sxs-lookup"><span data-stu-id="34657-478">Read mode</span></span>

<span data-ttu-id="34657-479">`optionalAttendees` 属性返回一个数组，其中包含每个可选与会者的 `EmailAddressDetails` 对象。</span><span class="sxs-lookup"><span data-stu-id="34657-479">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="34657-480">撰写模式</span><span class="sxs-lookup"><span data-stu-id="34657-480">Compose mode</span></span>

<span data-ttu-id="34657-481">`optionalAttendees` 属性返回一个 `Recipients` 对象，该对象提供用于获取和设置可选与会者的方法。</span><span class="sxs-lookup"><span data-stu-id="34657-481">The `optionalAttendees` property returns a `Recipients` object that provides methods to get and set the optional attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="34657-482">类型：</span><span class="sxs-lookup"><span data-stu-id="34657-482">Type:</span></span>

*   <span data-ttu-id="34657-483">数组。 <[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)> |[收件人](/javascript/api/outlook_1_5/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="34657-483">Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_5/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="34657-484">要求</span><span class="sxs-lookup"><span data-stu-id="34657-484">Requirements</span></span>

|<span data-ttu-id="34657-485">要求</span><span class="sxs-lookup"><span data-stu-id="34657-485">Requirement</span></span>| <span data-ttu-id="34657-486">值</span><span class="sxs-lookup"><span data-stu-id="34657-486">Value</span></span>|
|---|---|
|[<span data-ttu-id="34657-487">最低邮箱要求集版本</span><span class="sxs-lookup"><span data-stu-id="34657-487">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="34657-488">1.0</span><span class="sxs-lookup"><span data-stu-id="34657-488">1.0</span></span>|
|[<span data-ttu-id="34657-489">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="34657-489">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="34657-490">ReadItem</span><span class="sxs-lookup"><span data-stu-id="34657-490">ReadItem</span></span>|
|[<span data-ttu-id="34657-491">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="34657-491">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="34657-492">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="34657-492">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="34657-493">示例</span><span class="sxs-lookup"><span data-stu-id="34657-493">Example</span></span>

```
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

#### <a name="organizer-emailaddressdetailsjavascriptapioutlook15officeemailaddressdetails"></a><span data-ttu-id="34657-494">组织者：[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="34657-494">organizer :[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)</span></span>

<span data-ttu-id="34657-p124">获取指定会议的会议组织者的电子邮件地址。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="34657-p124">Gets the email address of the meeting organizer for a specified meeting. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="34657-497">类型：</span><span class="sxs-lookup"><span data-stu-id="34657-497">Type:</span></span>

*   [<span data-ttu-id="34657-498">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="34657-498">EmailAddressDetails</span></span>](/javascript/api/outlook_1_5/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="34657-499">要求</span><span class="sxs-lookup"><span data-stu-id="34657-499">Requirements</span></span>

|<span data-ttu-id="34657-500">要求</span><span class="sxs-lookup"><span data-stu-id="34657-500">Requirement</span></span>| <span data-ttu-id="34657-501">值</span><span class="sxs-lookup"><span data-stu-id="34657-501">Value</span></span>|
|---|---|
|[<span data-ttu-id="34657-502">最低邮箱要求集版本</span><span class="sxs-lookup"><span data-stu-id="34657-502">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="34657-503">1.0</span><span class="sxs-lookup"><span data-stu-id="34657-503">1.0</span></span>|
|[<span data-ttu-id="34657-504">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="34657-504">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="34657-505">ReadItem</span><span class="sxs-lookup"><span data-stu-id="34657-505">ReadItem</span></span>|
|[<span data-ttu-id="34657-506">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="34657-506">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="34657-507">阅读</span><span class="sxs-lookup"><span data-stu-id="34657-507">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="34657-508">示例</span><span class="sxs-lookup"><span data-stu-id="34657-508">Example</span></span>

```
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
```

####  <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlook15officeemailaddressdetailsrecipientsjavascriptapioutlook15officerecipients"></a><span data-ttu-id="34657-509">requiredAttendees： 数组。 <[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)>|[收件人](/javascript/api/outlook_1_5/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="34657-509">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_5/office.recipients)</span></span>

<span data-ttu-id="34657-510">提供对事件必需与会者的访问。</span><span class="sxs-lookup"><span data-stu-id="34657-510">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="34657-511">对象的类型和访问级别取决于当前项的模式。</span><span class="sxs-lookup"><span data-stu-id="34657-511">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="34657-512">阅读模式</span><span class="sxs-lookup"><span data-stu-id="34657-512">Read mode</span></span>

<span data-ttu-id="34657-513">`requiredAttendees` 属性返回一个数组，其中包含每个必需与会者的 `EmailAddressDetails` 对象。</span><span class="sxs-lookup"><span data-stu-id="34657-513">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="34657-514">撰写模式</span><span class="sxs-lookup"><span data-stu-id="34657-514">Compose mode</span></span>

<span data-ttu-id="34657-515">`requiredAttendees` 属性返回一个 `Recipients` 对象，该对象提供用于获取和设置可选与会者的方法。</span><span class="sxs-lookup"><span data-stu-id="34657-515">The `requiredAttendees` property returns a `Recipients` object that provides methods to get and set the required attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="34657-516">类型：</span><span class="sxs-lookup"><span data-stu-id="34657-516">Type:</span></span>

*   <span data-ttu-id="34657-517">数组。 <[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)> |[收件人](/javascript/api/outlook_1_5/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="34657-517">Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_5/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="34657-518">要求</span><span class="sxs-lookup"><span data-stu-id="34657-518">Requirements</span></span>

|<span data-ttu-id="34657-519">要求</span><span class="sxs-lookup"><span data-stu-id="34657-519">Requirement</span></span>| <span data-ttu-id="34657-520">值</span><span class="sxs-lookup"><span data-stu-id="34657-520">Value</span></span>|
|---|---|
|[<span data-ttu-id="34657-521">最低邮箱要求集版本</span><span class="sxs-lookup"><span data-stu-id="34657-521">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="34657-522">1.0</span><span class="sxs-lookup"><span data-stu-id="34657-522">1.0</span></span>|
|[<span data-ttu-id="34657-523">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="34657-523">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="34657-524">ReadItem</span><span class="sxs-lookup"><span data-stu-id="34657-524">ReadItem</span></span>|
|[<span data-ttu-id="34657-525">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="34657-525">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="34657-526">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="34657-526">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="34657-527">示例</span><span class="sxs-lookup"><span data-stu-id="34657-527">Example</span></span>

```
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
}
```

#### <a name="sender-emailaddressdetailsjavascriptapioutlook15officeemailaddressdetails"></a><span data-ttu-id="34657-528">发件人：[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="34657-528">sender :[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)</span></span>

<span data-ttu-id="34657-p126">获取电子邮件发件人的电子邮件地址。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="34657-p126">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="34657-p127">[`from`](#from-emailaddressdetailsjavascriptapioutlook15officeemailaddressdetails) 和 `sender` 属性表示同一个人，邮件由代理人发送的除外。在此情况下，`from` 属性表示代理程序，而 sender 属性表示代理人。</span><span class="sxs-lookup"><span data-stu-id="34657-p127">The [`from`](#from-emailaddressdetailsjavascriptapioutlook15officeemailaddressdetails) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="34657-533">`EmailAddressDetails` 对象的 `recipientType` 属性在 `sender` 属性是 `undefined`。</span><span class="sxs-lookup"><span data-stu-id="34657-533">Note: The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="34657-534">类型：</span><span class="sxs-lookup"><span data-stu-id="34657-534">Type:</span></span>

*   [<span data-ttu-id="34657-535">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="34657-535">EmailAddressDetails</span></span>](/javascript/api/outlook_1_5/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="34657-536">要求</span><span class="sxs-lookup"><span data-stu-id="34657-536">Requirements</span></span>

|<span data-ttu-id="34657-537">要求</span><span class="sxs-lookup"><span data-stu-id="34657-537">Requirement</span></span>| <span data-ttu-id="34657-538">值</span><span class="sxs-lookup"><span data-stu-id="34657-538">Value</span></span>|
|---|---|
|[<span data-ttu-id="34657-539">最低邮箱要求集版本</span><span class="sxs-lookup"><span data-stu-id="34657-539">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="34657-540">1.0</span><span class="sxs-lookup"><span data-stu-id="34657-540">1.0</span></span>|
|[<span data-ttu-id="34657-541">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="34657-541">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="34657-542">ReadItem</span><span class="sxs-lookup"><span data-stu-id="34657-542">ReadItem</span></span>|
|[<span data-ttu-id="34657-543">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="34657-543">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="34657-544">阅读</span><span class="sxs-lookup"><span data-stu-id="34657-544">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="34657-545">示例</span><span class="sxs-lookup"><span data-stu-id="34657-545">Example</span></span>

```
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
```

####  <a name="start-datetimejavascriptapioutlook15officetime"></a><span data-ttu-id="34657-546">start :Date|[Time](/javascript/api/outlook_1_5/office.time)</span><span class="sxs-lookup"><span data-stu-id="34657-546">start :Date|[Time](/javascript/api/outlook_1_5/office.time)</span></span>

<span data-ttu-id="34657-547">获取或设置约会开始的日期和时间。</span><span class="sxs-lookup"><span data-stu-id="34657-547">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="34657-p128">将 `start` 属性表示为协调世界时 (UTC) 的日期和时间值。可使用 [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook15officelocalclienttime) 方法将该值转换为客户端的本地日期和时间。</span><span class="sxs-lookup"><span data-stu-id="34657-p128">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook15officelocalclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="34657-550">阅读模式</span><span class="sxs-lookup"><span data-stu-id="34657-550">Read mode</span></span>

<span data-ttu-id="34657-551">`start` 属性返回 `Date` 对象。</span><span class="sxs-lookup"><span data-stu-id="34657-551">The `start` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="34657-552">撰写模式</span><span class="sxs-lookup"><span data-stu-id="34657-552">Compose mode</span></span>

<span data-ttu-id="34657-553">`start` 属性返回 `Time` 对象。</span><span class="sxs-lookup"><span data-stu-id="34657-553">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="34657-554">使用 [`Time.setAsync`](/javascript/api/outlook_1_5/office.time#setasync-datetime--options--callback-) 方法设置开始时间时，应使用 [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) 方法将客户端的本地时间转换为服务器的 UTC。</span><span class="sxs-lookup"><span data-stu-id="34657-554">When you use the [`Time.setAsync`](/javascript/api/outlook_1_5/office.time#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="34657-555">类型：</span><span class="sxs-lookup"><span data-stu-id="34657-555">Type:</span></span>

*   <span data-ttu-id="34657-556">日期 | [时间](/javascript/api/outlook_1_5/office.time)</span><span class="sxs-lookup"><span data-stu-id="34657-556">Date | [Time](/javascript/api/outlook_1_5/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="34657-557">要求</span><span class="sxs-lookup"><span data-stu-id="34657-557">Requirements</span></span>

|<span data-ttu-id="34657-558">要求</span><span class="sxs-lookup"><span data-stu-id="34657-558">Requirement</span></span>| <span data-ttu-id="34657-559">值</span><span class="sxs-lookup"><span data-stu-id="34657-559">Value</span></span>|
|---|---|
|[<span data-ttu-id="34657-560">最低邮箱要求集版本</span><span class="sxs-lookup"><span data-stu-id="34657-560">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="34657-561">1.0</span><span class="sxs-lookup"><span data-stu-id="34657-561">1.0</span></span>|
|[<span data-ttu-id="34657-562">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="34657-562">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="34657-563">ReadItem</span><span class="sxs-lookup"><span data-stu-id="34657-563">ReadItem</span></span>|
|[<span data-ttu-id="34657-564">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="34657-564">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="34657-565">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="34657-565">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="34657-566">示例</span><span class="sxs-lookup"><span data-stu-id="34657-566">Example</span></span>

<span data-ttu-id="34657-567">以下示例通过使用 `Time` 对象的 [`setAsync`](/javascript/api/outlook_1_5/office.time#setasync-datetime--options--callback-) 方法，设置撰写模式下约会的开始时间。</span><span class="sxs-lookup"><span data-stu-id="34657-567">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_5/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

####  <a name="subject-stringsubjectjavascriptapioutlook15officesubject"></a><span data-ttu-id="34657-568">主题： 字符串 |[主题](/javascript/api/outlook_1_5/office.subject)</span><span class="sxs-lookup"><span data-stu-id="34657-568">subject :String|[Subject](/javascript/api/outlook_1_5/office.subject)</span></span>

<span data-ttu-id="34657-569">获取或设置显示在项目的主题字段中的说明。</span><span class="sxs-lookup"><span data-stu-id="34657-569">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="34657-570">`subject` 属性获取或设置由电子邮件服务器发送项目时的整个主题。</span><span class="sxs-lookup"><span data-stu-id="34657-570">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="34657-571">阅读模式</span><span class="sxs-lookup"><span data-stu-id="34657-571">Read mode</span></span>

<span data-ttu-id="34657-p129">`subject` 属性返回一个字符串。使用 [`normalizedSubject`](#normalizedsubject-string) 属性获取不带任何前导前缀（如 `RE:` 和 `FW:`）的主题。</span><span class="sxs-lookup"><span data-stu-id="34657-p129">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

```
var subject = Office.context.mailbox.item.subject;
```

##### <a name="compose-mode"></a><span data-ttu-id="34657-574">撰写模式</span><span class="sxs-lookup"><span data-stu-id="34657-574">Compose mode</span></span>

<span data-ttu-id="34657-575">`subject` 属性返回一个 `Subject` 对象，该对象提供用于获取和设置主题的方法。</span><span class="sxs-lookup"><span data-stu-id="34657-575">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="34657-576">类型：</span><span class="sxs-lookup"><span data-stu-id="34657-576">Type:</span></span>

*   <span data-ttu-id="34657-577">字符串 | [主题](/javascript/api/outlook_1_5/office.subject)</span><span class="sxs-lookup"><span data-stu-id="34657-577">String | [Subject](/javascript/api/outlook_1_5/office.subject)</span></span>

##### <a name="requirements"></a><span data-ttu-id="34657-578">要求</span><span class="sxs-lookup"><span data-stu-id="34657-578">Requirements</span></span>

|<span data-ttu-id="34657-579">要求</span><span class="sxs-lookup"><span data-stu-id="34657-579">Requirement</span></span>| <span data-ttu-id="34657-580">值</span><span class="sxs-lookup"><span data-stu-id="34657-580">Value</span></span>|
|---|---|
|[<span data-ttu-id="34657-581">最低邮箱要求集版本</span><span class="sxs-lookup"><span data-stu-id="34657-581">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="34657-582">1.0</span><span class="sxs-lookup"><span data-stu-id="34657-582">1.0</span></span>|
|[<span data-ttu-id="34657-583">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="34657-583">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="34657-584">ReadItem</span><span class="sxs-lookup"><span data-stu-id="34657-584">ReadItem</span></span>|
|[<span data-ttu-id="34657-585">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="34657-585">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="34657-586">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="34657-586">Compose or read</span></span>|

####  <a name="to-arrayemailaddressdetailsjavascriptapioutlook15officeemailaddressdetailsrecipientsjavascriptapioutlook15officerecipients"></a><span data-ttu-id="34657-587">发送： 数组。 <[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)>|[收件人](/javascript/api/outlook_1_5/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="34657-587">to :Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_5/office.recipients)</span></span>

<span data-ttu-id="34657-588">提供对邮件的 **发送** 行上收件人的访问。</span><span class="sxs-lookup"><span data-stu-id="34657-588">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="34657-589">对象的类型和访问级别取决于当前项的模式。</span><span class="sxs-lookup"><span data-stu-id="34657-589">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="34657-590">阅读模式</span><span class="sxs-lookup"><span data-stu-id="34657-590">Read mode</span></span>

<span data-ttu-id="34657-p131">`to` 属性返回包含邮件的**收件人**行上所列的每个收件人的 `EmailAddressDetails` 对象的数组。集合上限为 100 个成员。</span><span class="sxs-lookup"><span data-stu-id="34657-p131">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="34657-593">撰写模式</span><span class="sxs-lookup"><span data-stu-id="34657-593">Compose mode</span></span>

<span data-ttu-id="34657-594">`to` 属性返回 `Recipients` 对象，该对象提供用于处理邮件**收件人**行上收件人的方法。</span><span class="sxs-lookup"><span data-stu-id="34657-594">The to`to` property returns a Recipients`Recipients` object that provides methods for manipulating the recipients on the **To** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="34657-595">类型：</span><span class="sxs-lookup"><span data-stu-id="34657-595">Type:</span></span>

*   <span data-ttu-id="34657-596">数组。 <[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)> |[收件人](/javascript/api/outlook_1_5/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="34657-596">Array.<[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_5/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="34657-597">要求</span><span class="sxs-lookup"><span data-stu-id="34657-597">Requirements</span></span>

|<span data-ttu-id="34657-598">要求</span><span class="sxs-lookup"><span data-stu-id="34657-598">Requirement</span></span>| <span data-ttu-id="34657-599">值</span><span class="sxs-lookup"><span data-stu-id="34657-599">Value</span></span>|
|---|---|
|[<span data-ttu-id="34657-600">最低邮箱要求集版本</span><span class="sxs-lookup"><span data-stu-id="34657-600">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="34657-601">1.0</span><span class="sxs-lookup"><span data-stu-id="34657-601">1.0</span></span>|
|[<span data-ttu-id="34657-602">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="34657-602">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="34657-603">ReadItem</span><span class="sxs-lookup"><span data-stu-id="34657-603">ReadItem</span></span>|
|[<span data-ttu-id="34657-604">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="34657-604">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="34657-605">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="34657-605">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="34657-606">示例</span><span class="sxs-lookup"><span data-stu-id="34657-606">Example</span></span>

```
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

### <a name="methods"></a><span data-ttu-id="34657-607">方法</span><span class="sxs-lookup"><span data-stu-id="34657-607">Methods</span></span>

####  <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="34657-608">addFileAttachmentAsync (uri，attachmentName，[选项] [回调])</span><span class="sxs-lookup"><span data-stu-id="34657-608">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="34657-609">将文件作为附件添加到邮件或约会。</span><span class="sxs-lookup"><span data-stu-id="34657-609">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="34657-610">`addFileAttachmentAsync` 方法在指定的 URI 上载文件并将其附加到撰写窗体中的项目。</span><span class="sxs-lookup"><span data-stu-id="34657-610">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="34657-611">你随后可以将该标识符与 [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) 方法一同使用，以删除同一个会话中的附件。</span><span class="sxs-lookup"><span data-stu-id="34657-611">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="34657-612">参数：</span><span class="sxs-lookup"><span data-stu-id="34657-612">Parameters:</span></span>

|<span data-ttu-id="34657-613">名称</span><span class="sxs-lookup"><span data-stu-id="34657-613">Name</span></span>| <span data-ttu-id="34657-614">类型</span><span class="sxs-lookup"><span data-stu-id="34657-614">Type</span></span>| <span data-ttu-id="34657-615">属性</span><span class="sxs-lookup"><span data-stu-id="34657-615">Attributes</span></span>| <span data-ttu-id="34657-616">说明</span><span class="sxs-lookup"><span data-stu-id="34657-616">Description</span></span>|
|---|---|---|---|
|`uri`| <span data-ttu-id="34657-617">String</span><span class="sxs-lookup"><span data-stu-id="34657-617">String</span></span>||<span data-ttu-id="34657-p132">提供附加到邮件或约会的文件的位置的 URI。最大长度为 2048 个字符。</span><span class="sxs-lookup"><span data-stu-id="34657-p132">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="34657-620">String</span><span class="sxs-lookup"><span data-stu-id="34657-620">String</span></span>||<span data-ttu-id="34657-p133">在附件上载过程中显示的附件名称。最大长度为 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="34657-p133">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="34657-623">Object</span><span class="sxs-lookup"><span data-stu-id="34657-623">Object</span></span>| <span data-ttu-id="34657-624">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="34657-624">&lt;optional&gt;</span></span>|<span data-ttu-id="34657-625">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="34657-625">An object literal that contains one or more of the following properties.</span></span>|
| `options.asyncContext` | <span data-ttu-id="34657-626">Object</span><span class="sxs-lookup"><span data-stu-id="34657-626">Object</span></span> | <span data-ttu-id="34657-627">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="34657-627">&lt;optional&gt;</span></span> | <span data-ttu-id="34657-628">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="34657-628">Developers can provide any object they wish to access in the callback method.</span></span> |
| `options.isInline` | <span data-ttu-id="34657-629">布尔值</span><span class="sxs-lookup"><span data-stu-id="34657-629">Boolean</span></span> | <span data-ttu-id="34657-630">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="34657-630">&lt;optional&gt;</span></span> | <span data-ttu-id="34657-631">如果 `true` ，指示附件将嵌入在邮件正文中显示，而不应在附件列表中显示。</span><span class="sxs-lookup"><span data-stu-id="34657-631">If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
|`callback`| <span data-ttu-id="34657-632">function</span><span class="sxs-lookup"><span data-stu-id="34657-632">function</span></span>| <span data-ttu-id="34657-633">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="34657-633">&lt;optional&gt;</span></span>|<span data-ttu-id="34657-634">方法完成后，使用单个参数 `asyncResult` （一个 [`AsyncResult`](/javascript/api/office/office.asyncresult)   对象）调用在 `callback`  参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="34657-634">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="34657-635">如果成功，附件标识符将在 `asyncResult.value` 属性中提供。</span><span class="sxs-lookup"><span data-stu-id="34657-635">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="34657-636">如果上传附件失败，`asyncResult` 对象将包含一个提供错误说明的 `Error` 对象。</span><span class="sxs-lookup"><span data-stu-id="34657-636">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="34657-637">错误</span><span class="sxs-lookup"><span data-stu-id="34657-637">Errors</span></span>

| <span data-ttu-id="34657-638">错误代码</span><span class="sxs-lookup"><span data-stu-id="34657-638">Error code</span></span> | <span data-ttu-id="34657-639">说明</span><span class="sxs-lookup"><span data-stu-id="34657-639">Description</span></span> |
|------------|-------------|
| `AttachmentSizeExceeded` | <span data-ttu-id="34657-640">附件大小超过了允许的大小。</span><span class="sxs-lookup"><span data-stu-id="34657-640">The attachment is larger than allowed.</span></span> |
| `FileTypeNotSupported` | <span data-ttu-id="34657-641">该附件的扩展名不是允许的扩展名。</span><span class="sxs-lookup"><span data-stu-id="34657-641">The attachment has an extension that is not allowed.</span></span> |
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="34657-642">邮件或约会具有的附件过多。</span><span class="sxs-lookup"><span data-stu-id="34657-642">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="34657-643">要求</span><span class="sxs-lookup"><span data-stu-id="34657-643">Requirements</span></span>

|<span data-ttu-id="34657-644">要求</span><span class="sxs-lookup"><span data-stu-id="34657-644">Requirement</span></span>| <span data-ttu-id="34657-645">值</span><span class="sxs-lookup"><span data-stu-id="34657-645">Value</span></span>|
|---|---|
|[<span data-ttu-id="34657-646">最低的邮箱要求集版本</span><span class="sxs-lookup"><span data-stu-id="34657-646">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="34657-647">1.1</span><span class="sxs-lookup"><span data-stu-id="34657-647">1.1</span></span>|
|[<span data-ttu-id="34657-648">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="34657-648">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="34657-649">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="34657-649">ReadWriteItem</span></span>|
|[<span data-ttu-id="34657-650">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="34657-650">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="34657-651">撰写</span><span class="sxs-lookup"><span data-stu-id="34657-651">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="34657-652">示例</span><span class="sxs-lookup"><span data-stu-id="34657-652">Examples</span></span>

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

<span data-ttu-id="34657-653">以下示例以嵌入附件方式添加图像文件并在邮件正文中引用此附件。</span><span class="sxs-lookup"><span data-stu-id="34657-653">The following example adds an image file as an inline attachment and references the attachment in the message body.</span></span>

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

####  <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="34657-654">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="34657-654">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="34657-655">将 Exchange 项目（如邮件）作为附件添加到邮件或约会。</span><span class="sxs-lookup"><span data-stu-id="34657-655">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="34657-p134">`addItemAttachmentAsync` 方法将包含指定 Exchange 标识符的项目附加到撰写窗体中的项目。如果指定一个回调方法，此方法使用 `asyncResult` 参数调用，该参数包含一个附件标识符或代码，指示附加项目过程中出现的任何错误。可以使用 `options` 参数将状态信息传递给回调方法（如果需要）。</span><span class="sxs-lookup"><span data-stu-id="34657-p134">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="34657-659">你随后可以将该标识符与 [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) 方法一同使用，以删除同一个会话中的附件。</span><span class="sxs-lookup"><span data-stu-id="34657-659">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="34657-660">如果 Office 外接程序在 Outlook Web App 中运行，则 `addItemAttachmentAsync` 方法可以将项目附加到项目（正在编辑的项目除外）中；然而，不支持也不建议这样做。</span><span class="sxs-lookup"><span data-stu-id="34657-660">If your Office add-in is running in Outlook Web App, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="34657-661">参数：</span><span class="sxs-lookup"><span data-stu-id="34657-661">Parameters:</span></span>

|<span data-ttu-id="34657-662">名称</span><span class="sxs-lookup"><span data-stu-id="34657-662">Name</span></span>| <span data-ttu-id="34657-663">类型</span><span class="sxs-lookup"><span data-stu-id="34657-663">Type</span></span>| <span data-ttu-id="34657-664">属性</span><span class="sxs-lookup"><span data-stu-id="34657-664">Attributes</span></span>| <span data-ttu-id="34657-665">说明</span><span class="sxs-lookup"><span data-stu-id="34657-665">Description</span></span>|
|---|---|---|---|
|`itemId`| <span data-ttu-id="34657-666">String</span><span class="sxs-lookup"><span data-stu-id="34657-666">String</span></span>||<span data-ttu-id="34657-p135">要附加项目的 Exchange 标识符。最大长度为 100 个字符。</span><span class="sxs-lookup"><span data-stu-id="34657-p135">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="34657-669">String</span><span class="sxs-lookup"><span data-stu-id="34657-669">String</span></span>||<span data-ttu-id="34657-p136">要附加项目的主题。最大长度为 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="34657-p136">The sujbect of the item to be attached. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="34657-672">Object</span><span class="sxs-lookup"><span data-stu-id="34657-672">Object</span></span>| <span data-ttu-id="34657-673">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="34657-673">&lt;optional&gt;</span></span>|<span data-ttu-id="34657-674">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="34657-674">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="34657-675">Object</span><span class="sxs-lookup"><span data-stu-id="34657-675">Object</span></span>| <span data-ttu-id="34657-676">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="34657-676">&lt;optional&gt;</span></span>|<span data-ttu-id="34657-677">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="34657-677">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="34657-678">function</span><span class="sxs-lookup"><span data-stu-id="34657-678">function</span></span>| <span data-ttu-id="34657-679">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="34657-679">&lt;optional&gt;</span></span>|<span data-ttu-id="34657-680">方法完成后，使用单个参数 `asyncResult` （一个 [`AsyncResult`](/javascript/api/office/office.asyncresult)   对象）调用在 `callback`  参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="34657-680">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="34657-681">如果成功，附件标识符将在 `asyncResult.value` 属性中提供。</span><span class="sxs-lookup"><span data-stu-id="34657-681">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="34657-682">如果添加附件失败，`asyncResult` 对象将包含一个提供错误说明的 `Error` 对象。</span><span class="sxs-lookup"><span data-stu-id="34657-682">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="34657-683">错误</span><span class="sxs-lookup"><span data-stu-id="34657-683">Errors</span></span>

| <span data-ttu-id="34657-684">错误代码</span><span class="sxs-lookup"><span data-stu-id="34657-684">Error code</span></span> | <span data-ttu-id="34657-685">说明</span><span class="sxs-lookup"><span data-stu-id="34657-685">Description</span></span> |
|------------|-------------|
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="34657-686">邮件或约会具有的附件过多。</span><span class="sxs-lookup"><span data-stu-id="34657-686">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="34657-687">要求</span><span class="sxs-lookup"><span data-stu-id="34657-687">Requirements</span></span>

|<span data-ttu-id="34657-688">要求</span><span class="sxs-lookup"><span data-stu-id="34657-688">Requirement</span></span>| <span data-ttu-id="34657-689">值</span><span class="sxs-lookup"><span data-stu-id="34657-689">Value</span></span>|
|---|---|
|[<span data-ttu-id="34657-690">最低的邮箱要求集版本</span><span class="sxs-lookup"><span data-stu-id="34657-690">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="34657-691">1.1</span><span class="sxs-lookup"><span data-stu-id="34657-691">1.1</span></span>|
|[<span data-ttu-id="34657-692">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="34657-692">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="34657-693">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="34657-693">ReadWriteItem</span></span>|
|[<span data-ttu-id="34657-694">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="34657-694">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="34657-695">撰写</span><span class="sxs-lookup"><span data-stu-id="34657-695">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="34657-696">示例</span><span class="sxs-lookup"><span data-stu-id="34657-696">Example</span></span>

<span data-ttu-id="34657-697">以下示例将现有的 Outlook 项目添加为名为 `My Attachment` 的附件。</span><span class="sxs-lookup"><span data-stu-id="34657-697">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

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

####  <a name="close"></a><span data-ttu-id="34657-698">close()</span><span class="sxs-lookup"><span data-stu-id="34657-698">close()</span></span>

<span data-ttu-id="34657-699">关闭当前正在撰写的项目。</span><span class="sxs-lookup"><span data-stu-id="34657-699">Closes the current item that is being composed.</span></span>

<span data-ttu-id="34657-p137">`close` 方法的行为取决于要撰写的项目的当前状态。如果项目具有未保存的更改，客户端将提示用户保存、放弃或取消关闭操作。</span><span class="sxs-lookup"><span data-stu-id="34657-p137">The behavior of the `close` method depends on the current state of the item being composed. If the item has unsaved changes, the client prompts the user to save, discard, or cancel the close action.</span></span>

> [!NOTE]
> <span data-ttu-id="34657-702">在 Outlook 网页版中，如果是约会项，并之前用`saveAsync` 保存过，会提示用户保存、放弃或取消，即使该项上一次保存后并未有任何更改。</span><span class="sxs-lookup"><span data-stu-id="34657-702">In Outlook on the web, if the item is an appointment and it has previously been saved using `saveAsync`, the user is prompted to save, discard, or cancel even if no changes have occurred since the item was last saved.</span></span>

<span data-ttu-id="34657-703">在 Outlook 桌面客户端中，如果邮件是内联答复，`close` 方法不起作用。</span><span class="sxs-lookup"><span data-stu-id="34657-703">In the Outlook desktop client, if the message is an inline reply, the `close` method has no effect.</span></span>

##### <a name="requirements"></a><span data-ttu-id="34657-704">要求</span><span class="sxs-lookup"><span data-stu-id="34657-704">Requirements</span></span>

|<span data-ttu-id="34657-705">要求</span><span class="sxs-lookup"><span data-stu-id="34657-705">Requirement</span></span>| <span data-ttu-id="34657-706">值</span><span class="sxs-lookup"><span data-stu-id="34657-706">Value</span></span>|
|---|---|
|[<span data-ttu-id="34657-707">最低邮箱要求集版本</span><span class="sxs-lookup"><span data-stu-id="34657-707">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="34657-708">1.3</span><span class="sxs-lookup"><span data-stu-id="34657-708">1.3</span></span>|
|[<span data-ttu-id="34657-709">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="34657-709">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="34657-710">Restricted</span><span class="sxs-lookup"><span data-stu-id="34657-710">Restricted</span></span>|
|[<span data-ttu-id="34657-711">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="34657-711">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="34657-712">撰写</span><span class="sxs-lookup"><span data-stu-id="34657-712">Compose</span></span>|

#### <a name="displayreplyallformformdata"></a><span data-ttu-id="34657-713">displayReplyAllForm(formData)</span><span class="sxs-lookup"><span data-stu-id="34657-713">displayReplyAllForm(formData)</span></span>

<span data-ttu-id="34657-714">显示答复窗体，其中包括所选邮件的发件人和所有收件人或所选约会的组织者和所有与会者。</span><span class="sxs-lookup"><span data-stu-id="34657-714">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="34657-715">在 Outlook for iOS 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="34657-715">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="34657-716">在 Outlook Web App 中，答复窗体显示为包含 3 列视图的弹出式窗体以及包含 2 列或 1 列视图的弹出式窗体。</span><span class="sxs-lookup"><span data-stu-id="34657-716">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="34657-717">如果任意字符串参数超出其限制， `displayReplyAllForm` 将引发异常。</span><span class="sxs-lookup"><span data-stu-id="34657-717">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="34657-p138">当在 `formData.attachments` 参数中指定附件时，Outlook 和 Outlook Web App 尝试下载所有附件并将其附加到答复窗体。如果无法添加任何附件，则在窗体 UI 中显示错误。如果这不可能，则不引发错误消息。</span><span class="sxs-lookup"><span data-stu-id="34657-p138">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="34657-721">参数：</span><span class="sxs-lookup"><span data-stu-id="34657-721">Parameters:</span></span>

| <span data-ttu-id="34657-722">名称</span><span class="sxs-lookup"><span data-stu-id="34657-722">Name</span></span> | <span data-ttu-id="34657-723">类型</span><span class="sxs-lookup"><span data-stu-id="34657-723">Type</span></span> | <span data-ttu-id="34657-724">属性</span><span class="sxs-lookup"><span data-stu-id="34657-724">Attributes</span></span> | <span data-ttu-id="34657-725">说明</span><span class="sxs-lookup"><span data-stu-id="34657-725">Description</span></span> |
|---|---|---|---|
|`formData`| <span data-ttu-id="34657-726">字符串 | 对象</span><span class="sxs-lookup"><span data-stu-id="34657-726">String &#124; Object</span></span>| |<span data-ttu-id="34657-p139">一个包含文本和 HTML 且代表答复窗体正文的字符串。字符串限制为 32 KB。</span><span class="sxs-lookup"><span data-stu-id="34657-p139">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="34657-729">**OR**</span><span class="sxs-lookup"><span data-stu-id="34657-729">**OR**</span></span><br/><span data-ttu-id="34657-p140">包含正文或附件数据和回调函数的对象。对象定义如下。</span><span class="sxs-lookup"><span data-stu-id="34657-p140">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="34657-732">String</span><span class="sxs-lookup"><span data-stu-id="34657-732">String</span></span> | <span data-ttu-id="34657-733">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="34657-733">&lt;optional&gt;</span></span> | <span data-ttu-id="34657-p141">一个包含文本和 HTML 且表示答复窗体正文的字符串。字符串限制为 32 KB。</span><span class="sxs-lookup"><span data-stu-id="34657-p141">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="34657-736">数组。&lt;对象&gt;</span><span class="sxs-lookup"><span data-stu-id="34657-736">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="34657-737">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="34657-737">&lt;optional&gt;</span></span> | <span data-ttu-id="34657-738">JSON 对象（是文件或项目附件）的数组。</span><span class="sxs-lookup"><span data-stu-id="34657-738">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="34657-739">String</span><span class="sxs-lookup"><span data-stu-id="34657-739">String</span></span> | | <span data-ttu-id="34657-p142">指示附件的类型。必须是文件附件的 `file` 或项目附件的 `item` 。</span><span class="sxs-lookup"><span data-stu-id="34657-p142">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="34657-742">String</span><span class="sxs-lookup"><span data-stu-id="34657-742">String</span></span> | | <span data-ttu-id="34657-743">一个包含附件的名称的字符串，最多包含 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="34657-743">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="34657-744">String</span><span class="sxs-lookup"><span data-stu-id="34657-744">String</span></span> | | <span data-ttu-id="34657-p143">仅在 `type` 设置为 `file` 时才使用。文件位置的 URI。</span><span class="sxs-lookup"><span data-stu-id="34657-p143">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.isInline` | <span data-ttu-id="34657-747">Boolean</span><span class="sxs-lookup"><span data-stu-id="34657-747">Boolean</span></span> | | <span data-ttu-id="34657-p144">仅在 `type` 设置为 `file` 时才使用。如果为 `true`，表示将在邮件正文中嵌入显示附件，并且不应在附件列表中显示。</span><span class="sxs-lookup"><span data-stu-id="34657-p144">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="34657-750">String</span><span class="sxs-lookup"><span data-stu-id="34657-750">String</span></span> | | <span data-ttu-id="34657-p145">仅在将 `type` 设置为 `item` 时使用。附件的 EWS 项目 ID。字符串最长为 100 个字符。</span><span class="sxs-lookup"><span data-stu-id="34657-p145">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="34657-754">function</span><span class="sxs-lookup"><span data-stu-id="34657-754">function</span></span> | <span data-ttu-id="34657-755">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="34657-755">&lt;optional&gt;</span></span> | <span data-ttu-id="34657-756">方法完成后，使用单个参数 `asyncResult`（一个 [AsyncResult](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="34657-756">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="34657-757">要求</span><span class="sxs-lookup"><span data-stu-id="34657-757">Requirements</span></span>

|<span data-ttu-id="34657-758">要求</span><span class="sxs-lookup"><span data-stu-id="34657-758">Requirement</span></span>| <span data-ttu-id="34657-759">值</span><span class="sxs-lookup"><span data-stu-id="34657-759">Value</span></span>|
|---|---|
|[<span data-ttu-id="34657-760">最低邮箱要求集版本</span><span class="sxs-lookup"><span data-stu-id="34657-760">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="34657-761">1.0</span><span class="sxs-lookup"><span data-stu-id="34657-761">1.0</span></span>|
|[<span data-ttu-id="34657-762">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="34657-762">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="34657-763">ReadItem</span><span class="sxs-lookup"><span data-stu-id="34657-763">ReadItem</span></span>|
|[<span data-ttu-id="34657-764">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="34657-764">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="34657-765">阅读</span><span class="sxs-lookup"><span data-stu-id="34657-765">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="34657-766">示例</span><span class="sxs-lookup"><span data-stu-id="34657-766">Examples</span></span>

<span data-ttu-id="34657-767">以下代码将一个字符串传递到 `displayReplyAllForm` 函数。</span><span class="sxs-lookup"><span data-stu-id="34657-767">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="34657-768">使用空白正文答复。</span><span class="sxs-lookup"><span data-stu-id="34657-768">Reply with an empty body.</span></span>

```
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="34657-769">仅使用正文答复。</span><span class="sxs-lookup"><span data-stu-id="34657-769">Reply with just a body.</span></span>

```
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="34657-770">使用正文和文件附件答复。</span><span class="sxs-lookup"><span data-stu-id="34657-770">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="34657-771">使用正文和项目附件答复。</span><span class="sxs-lookup"><span data-stu-id="34657-771">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="34657-772">使用正文、文件附件、项目附件和回调答复。</span><span class="sxs-lookup"><span data-stu-id="34657-772">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="displayreplyformformdata"></a><span data-ttu-id="34657-773">displayReplyForm(formData)</span><span class="sxs-lookup"><span data-stu-id="34657-773">displayReplyForm(formData)</span></span>

<span data-ttu-id="34657-774">显示答复窗体，其中仅包括所选邮件的发件人或所选约会的组织者。</span><span class="sxs-lookup"><span data-stu-id="34657-774">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="34657-775">在 Outlook for iOS 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="34657-775">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="34657-776">在 Outlook Web App 中，答复窗体显示为包含 3 列视图的弹出式窗体以及包含 2 列或 1 列视图的弹出式窗体。</span><span class="sxs-lookup"><span data-stu-id="34657-776">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="34657-777">如果任意字符串参数超出其限制， `displayReplyForm` 将引发异常。</span><span class="sxs-lookup"><span data-stu-id="34657-777">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="34657-p146">当在 `formData.attachments` 参数中指定附件时，Outlook 和 Outlook Web App 尝试下载所有附件并将其附加到答复窗体。如果无法添加任何附件，则在窗体 UI 中显示错误。如果这不可能，则不引发错误消息。</span><span class="sxs-lookup"><span data-stu-id="34657-p146">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="34657-781">参数：</span><span class="sxs-lookup"><span data-stu-id="34657-781">Parameters:</span></span>

| <span data-ttu-id="34657-782">名称</span><span class="sxs-lookup"><span data-stu-id="34657-782">Name</span></span> | <span data-ttu-id="34657-783">类型</span><span class="sxs-lookup"><span data-stu-id="34657-783">Type</span></span> | <span data-ttu-id="34657-784">属性</span><span class="sxs-lookup"><span data-stu-id="34657-784">Attributes</span></span> | <span data-ttu-id="34657-785">说明</span><span class="sxs-lookup"><span data-stu-id="34657-785">Description</span></span> |
|---|---|---|---|
|`formData`| <span data-ttu-id="34657-786">字符串 | 对象</span><span class="sxs-lookup"><span data-stu-id="34657-786">String &#124; Object</span></span>| | <span data-ttu-id="34657-p147">一个包含文本和 HTML 且代表答复窗体正文的字符串。字符串限制为 32 KB。</span><span class="sxs-lookup"><span data-stu-id="34657-p147">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="34657-789">**OR**</span><span class="sxs-lookup"><span data-stu-id="34657-789">**OR**</span></span><br/><span data-ttu-id="34657-p148">包含正文或附件数据和回调函数的对象。对象定义如下。</span><span class="sxs-lookup"><span data-stu-id="34657-p148">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="34657-792">String</span><span class="sxs-lookup"><span data-stu-id="34657-792">String</span></span> | <span data-ttu-id="34657-793">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="34657-793">&lt;optional&gt;</span></span> | <span data-ttu-id="34657-p149">一个包含文本和 HTML 且表示答复窗体正文的字符串。字符串限制为 32 KB。</span><span class="sxs-lookup"><span data-stu-id="34657-p149">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="34657-796">数组。&lt;对象&gt;</span><span class="sxs-lookup"><span data-stu-id="34657-796">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="34657-797">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="34657-797">&lt;optional&gt;</span></span> | <span data-ttu-id="34657-798">JSON 对象（是文件或项目附件）的数组。</span><span class="sxs-lookup"><span data-stu-id="34657-798">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="34657-799">String</span><span class="sxs-lookup"><span data-stu-id="34657-799">String</span></span> | | <span data-ttu-id="34657-p150">指示附件的类型。必须是文件附件的 `file` 或项目附件的 `item` 。</span><span class="sxs-lookup"><span data-stu-id="34657-p150">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="34657-802">String</span><span class="sxs-lookup"><span data-stu-id="34657-802">String</span></span> | | <span data-ttu-id="34657-803">一个包含附件的名称的字符串，最多包含 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="34657-803">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="34657-804">String</span><span class="sxs-lookup"><span data-stu-id="34657-804">String</span></span> | | <span data-ttu-id="34657-p151">仅在 `type` 设置为 `file` 时才使用。文件位置的 URI。</span><span class="sxs-lookup"><span data-stu-id="34657-p151">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.isInline` | <span data-ttu-id="34657-807">布尔值</span><span class="sxs-lookup"><span data-stu-id="34657-807">Boolean</span></span> | | <span data-ttu-id="34657-p152">仅在 `type` 设置为 `file` 时才使用。如果为 `true`，表示将在邮件正文中嵌入显示附件，并且不应在附件列表中显示。</span><span class="sxs-lookup"><span data-stu-id="34657-p152">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="34657-810">String</span><span class="sxs-lookup"><span data-stu-id="34657-810">String</span></span> | | <span data-ttu-id="34657-p153">仅在将 `type` 设置为 `item` 时使用。附件的 EWS 项目 ID。字符串最长为 100 个字符。</span><span class="sxs-lookup"><span data-stu-id="34657-p153">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="34657-814">function</span><span class="sxs-lookup"><span data-stu-id="34657-814">function</span></span> | <span data-ttu-id="34657-815">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="34657-815">&lt;optional&gt;</span></span> | <span data-ttu-id="34657-816">方法完成后，使用单个参数 `asyncResult`（一个 [AsyncResult](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="34657-816">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="34657-817">要求</span><span class="sxs-lookup"><span data-stu-id="34657-817">Requirements</span></span>

|<span data-ttu-id="34657-818">要求</span><span class="sxs-lookup"><span data-stu-id="34657-818">Requirement</span></span>| <span data-ttu-id="34657-819">值</span><span class="sxs-lookup"><span data-stu-id="34657-819">Value</span></span>|
|---|---|
|[<span data-ttu-id="34657-820">最低邮箱要求集版本</span><span class="sxs-lookup"><span data-stu-id="34657-820">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="34657-821">1.0</span><span class="sxs-lookup"><span data-stu-id="34657-821">1.0</span></span>|
|[<span data-ttu-id="34657-822">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="34657-822">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="34657-823">ReadItem</span><span class="sxs-lookup"><span data-stu-id="34657-823">ReadItem</span></span>|
|[<span data-ttu-id="34657-824">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="34657-824">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="34657-825">阅读</span><span class="sxs-lookup"><span data-stu-id="34657-825">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="34657-826">示例</span><span class="sxs-lookup"><span data-stu-id="34657-826">Examples</span></span>

<span data-ttu-id="34657-827">以下代码将一个字符串传递到 `displayReplyForm` 函数。</span><span class="sxs-lookup"><span data-stu-id="34657-827">The following code passes a string to the `displayReplyForm` function.</span></span>

```
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="34657-828">使用空白正文答复。</span><span class="sxs-lookup"><span data-stu-id="34657-828">Reply with an empty body.</span></span>

```
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="34657-829">仅使用正文答复。</span><span class="sxs-lookup"><span data-stu-id="34657-829">Reply with just a body.</span></span>

```
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="34657-830">使用正文和文件附件答复。</span><span class="sxs-lookup"><span data-stu-id="34657-830">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="34657-831">使用正文和项目附件答复。</span><span class="sxs-lookup"><span data-stu-id="34657-831">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="34657-832">使用正文、文件附件、项目附件和回调答复。</span><span class="sxs-lookup"><span data-stu-id="34657-832">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="getentities--entitiesjavascriptapioutlook15officeentities"></a><span data-ttu-id="34657-833">getEntities() → {[实体](/javascript/api/outlook_1_5/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="34657-833">getEntities() → {[Entities](/javascript/api/outlook_1_5/office.entities)}</span></span>

<span data-ttu-id="34657-834">获取在所选项正文中找到的实体。</span><span class="sxs-lookup"><span data-stu-id="34657-834">Gets the entities found in the selected item.</span></span>

> [!NOTE]
> <span data-ttu-id="34657-835">在 Outlook for iOS 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="34657-835">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="34657-836">要求</span><span class="sxs-lookup"><span data-stu-id="34657-836">Requirements</span></span>

|<span data-ttu-id="34657-837">要求</span><span class="sxs-lookup"><span data-stu-id="34657-837">Requirement</span></span>| <span data-ttu-id="34657-838">值</span><span class="sxs-lookup"><span data-stu-id="34657-838">Value</span></span>|
|---|---|
|[<span data-ttu-id="34657-839">最低邮箱要求集版本</span><span class="sxs-lookup"><span data-stu-id="34657-839">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="34657-840">1.0</span><span class="sxs-lookup"><span data-stu-id="34657-840">1.0</span></span>|
|[<span data-ttu-id="34657-841">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="34657-841">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="34657-842">ReadItem</span><span class="sxs-lookup"><span data-stu-id="34657-842">ReadItem</span></span>|
|[<span data-ttu-id="34657-843">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="34657-843">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="34657-844">阅读</span><span class="sxs-lookup"><span data-stu-id="34657-844">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="34657-845">返回：</span><span class="sxs-lookup"><span data-stu-id="34657-845">Returns:</span></span>

<span data-ttu-id="34657-846">类型：[Entities](/javascript/api/outlook_1_5/office.entities)</span><span class="sxs-lookup"><span data-stu-id="34657-846">Type: [Entities](/javascript/api/outlook_1_5/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="34657-847">示例</span><span class="sxs-lookup"><span data-stu-id="34657-847">Example</span></span>

<span data-ttu-id="34657-848">以下示例访问当前项正文中的联系人实体。</span><span class="sxs-lookup"><span data-stu-id="34657-848">The following example accesses the contacts entities on the current item.</span></span>

```
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlook15officecontactmeetingsuggestionjavascriptapioutlook15officemeetingsuggestionphonenumberjavascriptapioutlook15officephonenumbertasksuggestionjavascriptapioutlook15officetasksuggestion"></a><span data-ttu-id="34657-849">getEntitiesByType(entityType) → (nullable)  {数组。 <(String|[联系人](/javascript/api/outlook_1_5/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_5/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_5/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_5/office.tasksuggestion)) >}</span><span class="sxs-lookup"><span data-stu-id="34657-849">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_5/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_5/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_5/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_5/office.tasksuggestion))>}</span></span>

<span data-ttu-id="34657-850">获取所选项目正文中找到的指定实体类型的所有实体的数组。</span><span class="sxs-lookup"><span data-stu-id="34657-850">Gets an array of all the entities of the specified entity type found in the selected item.</span></span>

> [!NOTE]
> <span data-ttu-id="34657-851">在 Outlook for iOS 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="34657-851">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="34657-852">参数：</span><span class="sxs-lookup"><span data-stu-id="34657-852">Parameters:</span></span>

|<span data-ttu-id="34657-853">名称</span><span class="sxs-lookup"><span data-stu-id="34657-853">Name</span></span>| <span data-ttu-id="34657-854">类型</span><span class="sxs-lookup"><span data-stu-id="34657-854">Type</span></span>| <span data-ttu-id="34657-855">说明</span><span class="sxs-lookup"><span data-stu-id="34657-855">Description</span></span>|
|---|---|---|
|`entityType`| [<span data-ttu-id="34657-856">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="34657-856">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook_1_5/office.mailboxenums.entitytype)|<span data-ttu-id="34657-857">EntityType 枚举值之一。</span><span class="sxs-lookup"><span data-stu-id="34657-857">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="34657-858">要求</span><span class="sxs-lookup"><span data-stu-id="34657-858">Requirements</span></span>

|<span data-ttu-id="34657-859">要求</span><span class="sxs-lookup"><span data-stu-id="34657-859">Requirement</span></span>| <span data-ttu-id="34657-860">值</span><span class="sxs-lookup"><span data-stu-id="34657-860">Value</span></span>|
|---|---|
|[<span data-ttu-id="34657-861">最低邮箱要求集版本</span><span class="sxs-lookup"><span data-stu-id="34657-861">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="34657-862">1.0</span><span class="sxs-lookup"><span data-stu-id="34657-862">1.0</span></span>|
|[<span data-ttu-id="34657-863">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="34657-863">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="34657-864">Restricted</span><span class="sxs-lookup"><span data-stu-id="34657-864">Restricted</span></span>|
|[<span data-ttu-id="34657-865">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="34657-865">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="34657-866">阅读</span><span class="sxs-lookup"><span data-stu-id="34657-866">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="34657-867">返回：</span><span class="sxs-lookup"><span data-stu-id="34657-867">Returns:</span></span>

<span data-ttu-id="34657-868">如果在 `entityType` 中传递的值不是 `EntityType` 枚举的有效成员，该方法返回 null。</span><span class="sxs-lookup"><span data-stu-id="34657-868">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="34657-869">如果指定类型的任何实体都不存在于该项目上，该方法将返回空数组。</span><span class="sxs-lookup"><span data-stu-id="34657-869">If no entities of the specified type are present on the item, the method returns an empty array.</span></span> <span data-ttu-id="34657-870">否则，返回数组中的对象类型取决于 `entityType` 参数中请求实体的类型。</span><span class="sxs-lookup"><span data-stu-id="34657-870">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="34657-871">当使用此方法的最低权限级别**受限**时，某些实体类型需要 **ReadItem** 才能进行访问，如下表中所指定。</span><span class="sxs-lookup"><span data-stu-id="34657-871">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

| <span data-ttu-id="34657-872">值对应于 `entityType`</span><span class="sxs-lookup"><span data-stu-id="34657-872">Value of `entityType`</span></span> | <span data-ttu-id="34657-873">返回的数组中对象的类型</span><span class="sxs-lookup"><span data-stu-id="34657-873">Type of objects in returned array</span></span> | <span data-ttu-id="34657-874">所需权限级别</span><span class="sxs-lookup"><span data-stu-id="34657-874">Required Permission Level</span></span> |
| --- | --- | --- |
| `Address` | <span data-ttu-id="34657-875">String</span><span class="sxs-lookup"><span data-stu-id="34657-875">String</span></span> | <span data-ttu-id="34657-876">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="34657-876">**Restricted**</span></span> |
| `Contact` | <span data-ttu-id="34657-877">联系人</span><span class="sxs-lookup"><span data-stu-id="34657-877">Contact</span></span> | <span data-ttu-id="34657-878">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="34657-878">**ReadItem**</span></span> |
| `EmailAddress` | <span data-ttu-id="34657-879">String</span><span class="sxs-lookup"><span data-stu-id="34657-879">String</span></span> | <span data-ttu-id="34657-880">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="34657-880">**ReadItem**</span></span> |
| `MeetingSuggestion` | <span data-ttu-id="34657-881">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="34657-881">MeetingSuggestion</span></span> | <span data-ttu-id="34657-882">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="34657-882">**ReadItem**</span></span> |
| `PhoneNumber` | <span data-ttu-id="34657-883">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="34657-883">PhoneNumber</span></span> | <span data-ttu-id="34657-884">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="34657-884">**Restricted**</span></span> |
| `TaskSuggestion` | <span data-ttu-id="34657-885">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="34657-885">TaskSuggestion</span></span> | <span data-ttu-id="34657-886">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="34657-886">**ReadItem**</span></span> |
| `URL` | <span data-ttu-id="34657-887">String</span><span class="sxs-lookup"><span data-stu-id="34657-887">String</span></span> | <span data-ttu-id="34657-888">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="34657-888">**Restricted**</span></span> |

<span data-ttu-id="34657-889">类型：数组.<(字符串|[Contact](/javascript/api/outlook_1_5/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_5/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_5/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_5/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="34657-889">Type: Array.<(String|[Contact](/javascript/api/outlook_1_5/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_5/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_5/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_5/office.tasksuggestion))></span></span>

##### <a name="example"></a><span data-ttu-id="34657-890">示例</span><span class="sxs-lookup"><span data-stu-id="34657-890">Example</span></span>

<span data-ttu-id="34657-891">以下示例显示了如何访问代表当前项正文中邮政地址的字符串数组。</span><span class="sxs-lookup"><span data-stu-id="34657-891">The following example shows how to access an array of strings that represent postal addresses in the subject or body of the current item.</span></span>

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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlook15officecontactmeetingsuggestionjavascriptapioutlook15officemeetingsuggestionphonenumberjavascriptapioutlook15officephonenumbertasksuggestionjavascriptapioutlook15officetasksuggestion"></a><span data-ttu-id="34657-892">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_5/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_5/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_5/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_5/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="34657-892">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_5/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_5/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_5/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_5/office.tasksuggestion))>}</span></span>

<span data-ttu-id="34657-893">返回传递清单 XML 文件中定义的命名筛选器所选项中的已知实体。</span><span class="sxs-lookup"><span data-stu-id="34657-893">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="34657-894">在 Outlook for iOS 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="34657-894">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="34657-895">`getFilteredEntitiesByName` 方法返回与具有指定 `FilterName` 元素值的清单 XML 文件中的 [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) 规则元素中定义的规则表达式相匹配的实体。</span><span class="sxs-lookup"><span data-stu-id="34657-895">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="34657-896">参数：</span><span class="sxs-lookup"><span data-stu-id="34657-896">Parameters:</span></span>

|<span data-ttu-id="34657-897">名称</span><span class="sxs-lookup"><span data-stu-id="34657-897">Name</span></span>| <span data-ttu-id="34657-898">类型</span><span class="sxs-lookup"><span data-stu-id="34657-898">Type</span></span>| <span data-ttu-id="34657-899">说明</span><span class="sxs-lookup"><span data-stu-id="34657-899">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="34657-900">String</span><span class="sxs-lookup"><span data-stu-id="34657-900">String</span></span>|<span data-ttu-id="34657-901">定义筛选器匹配的 `ItemHasKnownEntity` 规则元素的名称。</span><span class="sxs-lookup"><span data-stu-id="34657-901">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="34657-902">要求</span><span class="sxs-lookup"><span data-stu-id="34657-902">Requirements</span></span>

|<span data-ttu-id="34657-903">要求</span><span class="sxs-lookup"><span data-stu-id="34657-903">Requirement</span></span>| <span data-ttu-id="34657-904">值</span><span class="sxs-lookup"><span data-stu-id="34657-904">Value</span></span>|
|---|---|
|[<span data-ttu-id="34657-905">最低邮箱要求集版本</span><span class="sxs-lookup"><span data-stu-id="34657-905">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="34657-906">1.0</span><span class="sxs-lookup"><span data-stu-id="34657-906">1.0</span></span>|
|[<span data-ttu-id="34657-907">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="34657-907">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="34657-908">ReadItem</span><span class="sxs-lookup"><span data-stu-id="34657-908">ReadItem</span></span>|
|[<span data-ttu-id="34657-909">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="34657-909">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="34657-910">阅读</span><span class="sxs-lookup"><span data-stu-id="34657-910">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="34657-911">返回：</span><span class="sxs-lookup"><span data-stu-id="34657-911">Returns:</span></span>

<span data-ttu-id="34657-p155">如果清单中 `ItemHasKnownEntity`  元素没有匹配 `FilterName` 参数的 `name` 元素值，则该方法返回 `null`。如果 `name` 参数匹配清单中的 `ItemHasKnownEntity` 元素，但在当前匹配的项目中没有实体，则该方法返回一个空数组。</span><span class="sxs-lookup"><span data-stu-id="34657-p155">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="34657-914">类型：Array.<(String|[Contact](/javascript/api/outlook_1_5/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_5/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_5/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_5/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="34657-914">Type: Array.<(String|[Contact](/javascript/api/outlook_1_5/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_5/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_5/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_5/office.tasksuggestion))></span></span>

#### <a name="getregexmatches--object"></a><span data-ttu-id="34657-915">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="34657-915">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="34657-916">返回所选项目中与在清单 XML 文件中定义的正则表达式相匹配的字符串值。</span><span class="sxs-lookup"><span data-stu-id="34657-916">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="34657-917">在 Outlook for iOS 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="34657-917">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="34657-p156">`getRegExMatches` 方法返回匹配在清单 XML 文件中的每个 `ItemHasRegularExpressionMatch` 或 `ItemHasKnownEntity` 规则元素中定义的正则表达式的字符串。对于 `ItemHasRegularExpressionMatch` 规则，匹配字符串必须发生在该规则指定的项目的属性中。`PropertyName` 简单类型定义支持的属性。</span><span class="sxs-lookup"><span data-stu-id="34657-p156">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="34657-921">例如，考虑一个加载项具有以下 `Rule` 元素：</span><span class="sxs-lookup"><span data-stu-id="34657-921">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="34657-922">从 `getRegExMatches` 返回的对象应有两个属性：`fruits` 和 `veggies`。</span><span class="sxs-lookup"><span data-stu-id="34657-922">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="34657-p157">如果在项目的正文属性上指定 `ItemHasRegularExpressionMatch` 规则，则正则表达式应进一步筛选正文，不应尝试返回该项目的整个正文。使用正则表达式（如 `.*`）获取项目的整个正文并不总是返回预期的结果。而该使用 [`Body.getAsync`](/javascript/api/outlook_1_5/office.body#getasync-coerciontype--options--callback-) 方法检索整个正文。</span><span class="sxs-lookup"><span data-stu-id="34657-p157">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook_1_5/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="34657-926">要求</span><span class="sxs-lookup"><span data-stu-id="34657-926">Requirements</span></span>

|<span data-ttu-id="34657-927">要求</span><span class="sxs-lookup"><span data-stu-id="34657-927">Requirement</span></span>| <span data-ttu-id="34657-928">值</span><span class="sxs-lookup"><span data-stu-id="34657-928">Value</span></span>|
|---|---|
|[<span data-ttu-id="34657-929">最低邮箱要求集版本</span><span class="sxs-lookup"><span data-stu-id="34657-929">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="34657-930">1.0</span><span class="sxs-lookup"><span data-stu-id="34657-930">1.0</span></span>|
|[<span data-ttu-id="34657-931">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="34657-931">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="34657-932">ReadItem</span><span class="sxs-lookup"><span data-stu-id="34657-932">ReadItem</span></span>|
|[<span data-ttu-id="34657-933">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="34657-933">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="34657-934">阅读</span><span class="sxs-lookup"><span data-stu-id="34657-934">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="34657-935">返回：</span><span class="sxs-lookup"><span data-stu-id="34657-935">Returns:</span></span>

<span data-ttu-id="34657-p158">一个包含与在清单 XML 文件中定义的正则表达式的字符串数组相匹配的对象。每个数组的名称等于匹配 `ItemHasRegularExpressionMatch` 规则的 `RegExName` 属性或匹配 `ItemHasKnownEntity` 规则的 `FilterName` 属性的相应值。</span><span class="sxs-lookup"><span data-stu-id="34657-p158">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<dl class="param-type"><span data-ttu-id="34657-938">

<dt>类型</dt>

</span><span class="sxs-lookup"><span data-stu-id="34657-938">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="34657-939">Object</span><span class="sxs-lookup"><span data-stu-id="34657-939">Object</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="34657-940">示例</span><span class="sxs-lookup"><span data-stu-id="34657-940">Example</span></span>

<span data-ttu-id="34657-941">以下示例显示了如何访问正则表达式 <rule>元素 `fruits` 和 `veggies` 的匹配项的数组，这些元素在清单中指定。</rule></span><span class="sxs-lookup"><span data-stu-id="34657-941">The following example shows how to access the array of matches for the regular expression <rule>elements `fruits` and `veggies`, which are specified in the manifest.</rule></span></span>

```
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veges = allMatches.veggies;
```

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="34657-942">getRegExMatchesByName(name) → (可空类型) {数组.< 字符串 >}</span><span class="sxs-lookup"><span data-stu-id="34657-942">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="34657-943">返回匹配在清单 XML 文件中定义的命名正则表达式的所选项目中的字符串值。</span><span class="sxs-lookup"><span data-stu-id="34657-943">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="34657-944">在 Outlook for iOS 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="34657-944">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="34657-945">`getRegExMatchesByName` 方法返回匹配在具有指定 `RegExName` 元素值的清单 XML 文件中的 `ItemHasRegularExpressionMatch` 规则元素中定义的正则表达式的字符串。</span><span class="sxs-lookup"><span data-stu-id="34657-945">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="34657-p159">如果在项目正文属性上指定 `ItemHasRegularExpressionMatch` 规则，则正则表达式应进一步筛选正文，不应尝试返回该项目的整个正文。使用正则表达式（如 `.*`）获取项目的整个正文并不总是返回预期的结果。</span><span class="sxs-lookup"><span data-stu-id="34657-p159">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="34657-948">参数：</span><span class="sxs-lookup"><span data-stu-id="34657-948">Parameters:</span></span>

|<span data-ttu-id="34657-949">名称</span><span class="sxs-lookup"><span data-stu-id="34657-949">Name</span></span>| <span data-ttu-id="34657-950">类型</span><span class="sxs-lookup"><span data-stu-id="34657-950">Type</span></span>| <span data-ttu-id="34657-951">说明</span><span class="sxs-lookup"><span data-stu-id="34657-951">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="34657-952">String</span><span class="sxs-lookup"><span data-stu-id="34657-952">String</span></span>|<span data-ttu-id="34657-953">定义筛选器匹配的 `ItemHasRegularExpressionMatch` 规则元素的名称。</span><span class="sxs-lookup"><span data-stu-id="34657-953">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="34657-954">要求</span><span class="sxs-lookup"><span data-stu-id="34657-954">Requirements</span></span>

|<span data-ttu-id="34657-955">要求</span><span class="sxs-lookup"><span data-stu-id="34657-955">Requirement</span></span>| <span data-ttu-id="34657-956">值</span><span class="sxs-lookup"><span data-stu-id="34657-956">Value</span></span>|
|---|---|
|[<span data-ttu-id="34657-957">最低邮箱要求集版本</span><span class="sxs-lookup"><span data-stu-id="34657-957">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="34657-958">1.0</span><span class="sxs-lookup"><span data-stu-id="34657-958">1.0</span></span>|
|[<span data-ttu-id="34657-959">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="34657-959">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="34657-960">ReadItem</span><span class="sxs-lookup"><span data-stu-id="34657-960">ReadItem</span></span>|
|[<span data-ttu-id="34657-961">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="34657-961">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="34657-962">阅读</span><span class="sxs-lookup"><span data-stu-id="34657-962">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="34657-963">返回：</span><span class="sxs-lookup"><span data-stu-id="34657-963">Returns:</span></span>

<span data-ttu-id="34657-964">一个包含与在清单 XML 文件中定义的正则表达式的字符串相匹配的数组。</span><span class="sxs-lookup"><span data-stu-id="34657-964">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<dl class="param-type"><span data-ttu-id="34657-965">

<dt>类型</dt>

</span><span class="sxs-lookup"><span data-stu-id="34657-965">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="34657-966">数组。 < 字符串 ></span><span class="sxs-lookup"><span data-stu-id="34657-966">Array.< String ></span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="34657-967">示例</span><span class="sxs-lookup"><span data-stu-id="34657-967">Example</span></span>

```
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

####  <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="34657-968">getSelectedDataAsync (coercionType，[选项] 回调) → {字符串}</span><span class="sxs-lookup"><span data-stu-id="34657-968">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="34657-969">以异步方式返回邮件的主题或正文中选定的数据。</span><span class="sxs-lookup"><span data-stu-id="34657-969">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="34657-p160">如果没有选定内容，但光标位于正文或主题中，此方法将会为所选数据返回 null。如果选定的是字段，而不是正文或主题，则此方法返回 `InvalidSelection` 错误。</span><span class="sxs-lookup"><span data-stu-id="34657-p160">If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="34657-972">参数：</span><span class="sxs-lookup"><span data-stu-id="34657-972">Parameters:</span></span>

|<span data-ttu-id="34657-973">名称</span><span class="sxs-lookup"><span data-stu-id="34657-973">Name</span></span>| <span data-ttu-id="34657-974">类型</span><span class="sxs-lookup"><span data-stu-id="34657-974">Type</span></span>| <span data-ttu-id="34657-975">属性</span><span class="sxs-lookup"><span data-stu-id="34657-975">Attributes</span></span>| <span data-ttu-id="34657-976">说明</span><span class="sxs-lookup"><span data-stu-id="34657-976">Description</span></span>|
|---|---|---|---|
|`coercionType`| [<span data-ttu-id="34657-977">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="34657-977">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="34657-p161">请求数据的格式。如果为 Text，则此方法返回纯文本作为字符串，删除任何显示的 HTML 标记。如果为 HTML，则此方法返回所选文本，不论是纯文本还是 HTML。</span><span class="sxs-lookup"><span data-stu-id="34657-p161">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`| <span data-ttu-id="34657-981">Object</span><span class="sxs-lookup"><span data-stu-id="34657-981">Object</span></span>| <span data-ttu-id="34657-982">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="34657-982">&lt;optional&gt;</span></span>|<span data-ttu-id="34657-983">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="34657-983">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="34657-984">Object</span><span class="sxs-lookup"><span data-stu-id="34657-984">Object</span></span>| <span data-ttu-id="34657-985">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="34657-985">&lt;optional&gt;</span></span>|<span data-ttu-id="34657-986">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="34657-986">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="34657-987">函数</span><span class="sxs-lookup"><span data-stu-id="34657-987">function</span></span>||<span data-ttu-id="34657-988">方法完成后，使用单个参数 `callback`（一个 [`asyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `AsyncResult` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="34657-988">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="34657-989">若要从回调方法访问所选数据，请调用 `asyncResult.value.data`。</span><span class="sxs-lookup"><span data-stu-id="34657-989">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="34657-990">若要访问选定内容的源属性，请调用 `asyncResult.value.sourceProperty`，这将为 `body` 或 `subject`。</span><span class="sxs-lookup"><span data-stu-id="34657-990">To access the source property that the selection comes from, call , which will be either  or .|</span></span>|

##### <a name="requirements"></a><span data-ttu-id="34657-991">要求</span><span class="sxs-lookup"><span data-stu-id="34657-991">Requirements</span></span>

|<span data-ttu-id="34657-992">要求</span><span class="sxs-lookup"><span data-stu-id="34657-992">Requirement</span></span>| <span data-ttu-id="34657-993">值</span><span class="sxs-lookup"><span data-stu-id="34657-993">Value</span></span>|
|---|---|
|[<span data-ttu-id="34657-994">最低邮箱要求集版本</span><span class="sxs-lookup"><span data-stu-id="34657-994">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="34657-995">1.2</span><span class="sxs-lookup"><span data-stu-id="34657-995">1.2</span></span>|
|[<span data-ttu-id="34657-996">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="34657-996">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="34657-997">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="34657-997">ReadWriteItem</span></span>|
|[<span data-ttu-id="34657-998">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="34657-998">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="34657-999">撰写</span><span class="sxs-lookup"><span data-stu-id="34657-999">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="34657-1000">返回：</span><span class="sxs-lookup"><span data-stu-id="34657-1000">Returns:</span></span>

<span data-ttu-id="34657-1001">作为字符串的所选数据的格式由 `coercionType` 确定。</span><span class="sxs-lookup"><span data-stu-id="34657-1001">The selected data as a string with format determined by `coercionType`.</span></span>

<dl class="param-type"><span data-ttu-id="34657-1002">

<dt>类型</dt>

</span><span class="sxs-lookup"><span data-stu-id="34657-1002">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="34657-1003">String</span><span class="sxs-lookup"><span data-stu-id="34657-1003">String</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="34657-1004">示例</span><span class="sxs-lookup"><span data-stu-id="34657-1004">Example</span></span>

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

####  <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="34657-1005">loadCustomPropertiesAsync （回调、 [userContext]）</span><span class="sxs-lookup"><span data-stu-id="34657-1005">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="34657-1006">为所选项目的加载项异步加载自定义属性。</span><span class="sxs-lookup"><span data-stu-id="34657-1006">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="34657-p163">自定义属性基于每个应用、每个项目存储为键/值对。此方法在回调中返回 `CustomProperties` 对象，该回调提供访问特定于当前项目和当前外接程序的自定义属性的方法。自定义属性未在项目上加密，因此这不应用作安全存储。</span><span class="sxs-lookup"><span data-stu-id="34657-p163">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="34657-1010">参数：</span><span class="sxs-lookup"><span data-stu-id="34657-1010">Parameters:</span></span>

|<span data-ttu-id="34657-1011">名称</span><span class="sxs-lookup"><span data-stu-id="34657-1011">Name</span></span>| <span data-ttu-id="34657-1012">类型</span><span class="sxs-lookup"><span data-stu-id="34657-1012">Type</span></span>| <span data-ttu-id="34657-1013">属性</span><span class="sxs-lookup"><span data-stu-id="34657-1013">Attributes</span></span>| <span data-ttu-id="34657-1014">说明</span><span class="sxs-lookup"><span data-stu-id="34657-1014">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="34657-1015">function</span><span class="sxs-lookup"><span data-stu-id="34657-1015">function</span></span>||<span data-ttu-id="34657-1016">此方法完成时，用单个参数调用 `callback` 参数中传递的函数，`asyncResult`，这是一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象。</span><span class="sxs-lookup"><span data-stu-id="34657-1016">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="34657-1017">自定义属性作为 `asyncResult.value` 属性中的 [`CustomProperties`](/javascript/api/outlook_1_5/office.customproperties) 对象来提供。</span><span class="sxs-lookup"><span data-stu-id="34657-1017">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook_1_5/office.customproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="34657-1018">该对象可用于获取、设置和删除项目中的自定义属性，并将针对自定义属性集的更改保存回服务器。</span><span class="sxs-lookup"><span data-stu-id="34657-1018">The custom properties are provided as a CustomProperties object in the asyncResult.value property. This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`| <span data-ttu-id="34657-1019">Object</span><span class="sxs-lookup"><span data-stu-id="34657-1019">Object</span></span>| <span data-ttu-id="34657-1020">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="34657-1020">&lt;optional&gt;</span></span>|<span data-ttu-id="34657-1021">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="34657-1021">Developers can provide any object they wish to access in the callback method.</span></span> <span data-ttu-id="34657-1022">可以通过回调函数的 `asyncResult.asyncContext` 属性访问该对象。</span><span class="sxs-lookup"><span data-stu-id="34657-1022">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.|</span></span>|

##### <a name="requirements"></a><span data-ttu-id="34657-1023">要求</span><span class="sxs-lookup"><span data-stu-id="34657-1023">Requirements</span></span>

|<span data-ttu-id="34657-1024">要求</span><span class="sxs-lookup"><span data-stu-id="34657-1024">Requirement</span></span>| <span data-ttu-id="34657-1025">值</span><span class="sxs-lookup"><span data-stu-id="34657-1025">Value</span></span>|
|---|---|
|[<span data-ttu-id="34657-1026">最低邮箱要求集版本</span><span class="sxs-lookup"><span data-stu-id="34657-1026">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="34657-1027">1.0</span><span class="sxs-lookup"><span data-stu-id="34657-1027">1.0</span></span>|
|[<span data-ttu-id="34657-1028">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="34657-1028">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="34657-1029">ReadItem</span><span class="sxs-lookup"><span data-stu-id="34657-1029">ReadItem</span></span>|
|[<span data-ttu-id="34657-1030">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="34657-1030">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="34657-1031">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="34657-1031">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="34657-1032">示例</span><span class="sxs-lookup"><span data-stu-id="34657-1032">Example</span></span>

<span data-ttu-id="34657-p166">以下代码示例显示了如何使用 `loadCustomPropertiesAsync` 方法异步加载特定于当前项目的自定义属性。该示例还显示了如何使用 `CustomProperties.saveAsync` 方法将这些属性重新保存到服务器。加载自定义属性后，该代码示例将使用 `CustomProperties.get` 方法读取自定义属性 `myProp`，使用 `CustomProperties.set` 方法写入自定义属性 `otherProp`，最后调用 `saveAsync` 方法保存这些自定义属性。</span><span class="sxs-lookup"><span data-stu-id="34657-p166">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

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

####  <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="34657-1036">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="34657-1036">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="34657-1037">将附件从邮件或约会中删除。</span><span class="sxs-lookup"><span data-stu-id="34657-1037">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="34657-p167">`removeAttachmentAsync` 方法删除项目中带指定标识符的附件。最佳做法是，仅当同一个邮件应用程序在同一会话中添加了一个附件时，你才应使用该附件标识符来删除该附件。在 Outlook Web App 和适用于设备的 OWA 中，附件标识符只在同一个会话中才有效。当用户关闭应用程序，或者如果用户开始在内嵌窗体中撰写，并在随后弹出的内嵌窗体中继续在单独的窗口撰写时，会话即结束。</span><span class="sxs-lookup"><span data-stu-id="34657-p167">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item. As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session. In Outlook Web App and OWA for Devices, the attachment identifier is valid only within the same session. A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="34657-1042">参数：</span><span class="sxs-lookup"><span data-stu-id="34657-1042">Parameters:</span></span>

|<span data-ttu-id="34657-1043">名称</span><span class="sxs-lookup"><span data-stu-id="34657-1043">Name</span></span>| <span data-ttu-id="34657-1044">类型</span><span class="sxs-lookup"><span data-stu-id="34657-1044">Type</span></span>| <span data-ttu-id="34657-1045">属性</span><span class="sxs-lookup"><span data-stu-id="34657-1045">Attributes</span></span>| <span data-ttu-id="34657-1046">说明</span><span class="sxs-lookup"><span data-stu-id="34657-1046">Description</span></span>|
|---|---|---|---|
|`attachmentId`| <span data-ttu-id="34657-1047">String</span><span class="sxs-lookup"><span data-stu-id="34657-1047">String</span></span>||<span data-ttu-id="34657-p168">要删除的附件的标识符。字符串的最大长度为 100 个字符。</span><span class="sxs-lookup"><span data-stu-id="34657-p168">The identifier of the attachment to remove. The maximum length of the string is 100 characters.</span></span>|
|`options`| <span data-ttu-id="34657-1050">Object</span><span class="sxs-lookup"><span data-stu-id="34657-1050">Object</span></span>| <span data-ttu-id="34657-1051">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="34657-1051">&lt;optional&gt;</span></span>|<span data-ttu-id="34657-1052">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="34657-1052">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="34657-1053">Object</span><span class="sxs-lookup"><span data-stu-id="34657-1053">Object</span></span>| <span data-ttu-id="34657-1054">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="34657-1054">&lt;optional&gt;</span></span>|<span data-ttu-id="34657-1055">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="34657-1055">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="34657-1056">function</span><span class="sxs-lookup"><span data-stu-id="34657-1056">function</span></span>| <span data-ttu-id="34657-1057">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="34657-1057">&lt;optional&gt;</span></span>|<span data-ttu-id="34657-1058">此方法完成时，用单个参数调用 `callback` 参数中传递的函数， `asyncResult` ，这是一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象。</span><span class="sxs-lookup"><span data-stu-id="34657-1058">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="34657-1059">如果删除附件失败，`asyncResult.error` 属性将包含一个说明失败原因的错误代码。</span><span class="sxs-lookup"><span data-stu-id="34657-1059">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="34657-1060">错误</span><span class="sxs-lookup"><span data-stu-id="34657-1060">Errors</span></span>

| <span data-ttu-id="34657-1061">错误代码</span><span class="sxs-lookup"><span data-stu-id="34657-1061">Error code</span></span> | <span data-ttu-id="34657-1062">说明</span><span class="sxs-lookup"><span data-stu-id="34657-1062">Description</span></span> |
|------------|-------------|
| `InvalidAttachmentId` | <span data-ttu-id="34657-1063">附件标识符不存在。</span><span class="sxs-lookup"><span data-stu-id="34657-1063">The attachment identifier does not exist.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="34657-1064">要求</span><span class="sxs-lookup"><span data-stu-id="34657-1064">Requirements</span></span>

|<span data-ttu-id="34657-1065">要求</span><span class="sxs-lookup"><span data-stu-id="34657-1065">Requirement</span></span>| <span data-ttu-id="34657-1066">值</span><span class="sxs-lookup"><span data-stu-id="34657-1066">Value</span></span>|
|---|---|
|[<span data-ttu-id="34657-1067">最低的邮箱要求集版本</span><span class="sxs-lookup"><span data-stu-id="34657-1067">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="34657-1068">1.1</span><span class="sxs-lookup"><span data-stu-id="34657-1068">1.1</span></span>|
|[<span data-ttu-id="34657-1069">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="34657-1069">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="34657-1070">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="34657-1070">ReadWriteItem</span></span>|
|[<span data-ttu-id="34657-1071">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="34657-1071">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="34657-1072">撰写</span><span class="sxs-lookup"><span data-stu-id="34657-1072">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="34657-1073">示例</span><span class="sxs-lookup"><span data-stu-id="34657-1073">Example</span></span>

<span data-ttu-id="34657-1074">以下代码删除包含标识符 '0' 的附件。</span><span class="sxs-lookup"><span data-stu-id="34657-1074">The following code removes an attachment with an identifier of '0'.</span></span>

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

####  <a name="saveasyncoptions-callback"></a><span data-ttu-id="34657-1075">saveAsync ([选项] 回调)</span><span class="sxs-lookup"><span data-stu-id="34657-1075">saveAsync([options], callback)</span></span>

<span data-ttu-id="34657-1076">异步保存项目。</span><span class="sxs-lookup"><span data-stu-id="34657-1076">Asynchronously saves an item.</span></span>

<span data-ttu-id="34657-p169">调用时，此方法将当前邮件保存为草稿，并通过回调方法返回项目 ID。在 Outlook Web App 或 Outlook 联机模式下，该项目被保存到服务器中。在 Outlook 缓存模式下，该项目被保存到本地缓存中。</span><span class="sxs-lookup"><span data-stu-id="34657-p169">When invoked, this method saves the current message as a draft and returns the item id via the callback method. In Outlook Web App or Outlook in online mode, the item is saved to the server. In Outlook in cached mode, the item is saved to the local cache.</span></span>

> [!NOTE]
> <span data-ttu-id="34657-1080">如果加载项调用 `saveAsync` 中的项目在撰写模式下才能获取 `itemId` 若要使用 EWS 或 REST API，请注意，缓存模式 Outlook 时，可能需要一些时间才能将项目实际同步到服务器。</span><span class="sxs-lookup"><span data-stu-id="34657-1080">Note: If your add-in calls `saveAsync` on an item in compose mode in order to get an `itemId` to use with EWS or the REST API, be aware that when Outlook is in cached mode, it may take some time before the item is actually synced to the server. Until the item is synced, using the  will return an error.</span></span> <span data-ttu-id="34657-1081">直到该项目同步，使用 `itemId` 将返回错误。</span><span class="sxs-lookup"><span data-stu-id="34657-1081">Until the item is synced, using the `itemId` will return an error.</span></span>

<span data-ttu-id="34657-p171">由于约会没有草稿状态，如果以撰写模式在约会中调用 `saveAsync`，则该项将被保存为用户日历中的正常约会。对于之前未保存过的新约会，则不会发送邀请。保存现有约会将向添加或删除的与会者发送更新。</span><span class="sxs-lookup"><span data-stu-id="34657-p171">Since appointments have no draft state, if `saveAsync` is called on an appointment in compose mode, the item will be saved as a normal appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent. Saving an existing appointment will send an update to added or removed attendees.</span></span>

> [!NOTE]
> <span data-ttu-id="34657-1085">以下客户端在约会上的撰写模式下具有 `saveAsync` 的不同行为：</span><span class="sxs-lookup"><span data-stu-id="34657-1085">The following clients have different behavior for `saveAsync` on appointments in compose mode:</span></span>
>
> - <span data-ttu-id="34657-1086">Mac Outlook 在会议的撰写模式中不支持 `saveAsync` 。</span><span class="sxs-lookup"><span data-stu-id="34657-1086">Note: Mac Outlook does not support `saveAsync` on a meeting in compose mode. Calling  on a meeting in Mac Outlook will return an error.</span></span> <span data-ttu-id="34657-1087">在Mac Outlook 中的会议上调用 `saveAsync` ，则将返回错误。</span><span class="sxs-lookup"><span data-stu-id="34657-1087">Note: Mac Outlook does not support  on a meeting in compose mode. Calling `saveAsync` on a meeting in Mac Outlook will return an error.</span></span>
> - <span data-ttu-id="34657-1088">当 `saveAsync` 在撰写模式调用约会时，Outlook 网页版总会发送一个邀请或更新。</span><span class="sxs-lookup"><span data-stu-id="34657-1088">Outlook on the web always sends an invitation or update when `saveAsync` is called on an appointment in compose mode.</span></span>

##### <a name="parameters"></a><span data-ttu-id="34657-1089">参数：</span><span class="sxs-lookup"><span data-stu-id="34657-1089">Parameters:</span></span>

|<span data-ttu-id="34657-1090">名称</span><span class="sxs-lookup"><span data-stu-id="34657-1090">Name</span></span>| <span data-ttu-id="34657-1091">类型</span><span class="sxs-lookup"><span data-stu-id="34657-1091">Type</span></span>| <span data-ttu-id="34657-1092">属性</span><span class="sxs-lookup"><span data-stu-id="34657-1092">Attributes</span></span>| <span data-ttu-id="34657-1093">说明</span><span class="sxs-lookup"><span data-stu-id="34657-1093">Description</span></span>|
|---|---|---|---|
|`options`| <span data-ttu-id="34657-1094">Object</span><span class="sxs-lookup"><span data-stu-id="34657-1094">Object</span></span>| <span data-ttu-id="34657-1095">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="34657-1095">&lt;optional&gt;</span></span>|<span data-ttu-id="34657-1096">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="34657-1096">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="34657-1097">Object</span><span class="sxs-lookup"><span data-stu-id="34657-1097">Object</span></span>| <span data-ttu-id="34657-1098">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="34657-1098">&lt;optional&gt;</span></span>|<span data-ttu-id="34657-1099">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="34657-1099">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="34657-1100">函数</span><span class="sxs-lookup"><span data-stu-id="34657-1100">function</span></span>||<span data-ttu-id="34657-1101">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="34657-1101">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="34657-1102">如果成功，该项目标识符在 `asyncResult.value` 属性中提供。</span><span class="sxs-lookup"><span data-stu-id="34657-1102">On success, the item identifier is provided in the `asyncResult.value` property.|</span></span>|

##### <a name="requirements"></a><span data-ttu-id="34657-1103">要求</span><span class="sxs-lookup"><span data-stu-id="34657-1103">Requirements</span></span>

|<span data-ttu-id="34657-1104">要求</span><span class="sxs-lookup"><span data-stu-id="34657-1104">Requirement</span></span>| <span data-ttu-id="34657-1105">值</span><span class="sxs-lookup"><span data-stu-id="34657-1105">Value</span></span>|
|---|---|
|[<span data-ttu-id="34657-1106">最低邮箱要求集版本</span><span class="sxs-lookup"><span data-stu-id="34657-1106">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="34657-1107">1.3</span><span class="sxs-lookup"><span data-stu-id="34657-1107">1.3</span></span>|
|[<span data-ttu-id="34657-1108">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="34657-1108">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="34657-1109">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="34657-1109">ReadWriteItem</span></span>|
|[<span data-ttu-id="34657-1110">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="34657-1110">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="34657-1111">撰写</span><span class="sxs-lookup"><span data-stu-id="34657-1111">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="34657-1112">示例</span><span class="sxs-lookup"><span data-stu-id="34657-1112">Examples</span></span>

```
Office.context.mailbox.item.saveAsync(
  function callback(result) {
    // Process the result
  });
```

<span data-ttu-id="34657-p173">下面是传递给回调函数的 `result` 参数示例。`value` 属性包含的该项的项目 ID。</span><span class="sxs-lookup"><span data-stu-id="34657-p173">The following is an example of the `result` parameter passed to the callback function. The `value` property contains the item ID of the item.</span></span>

```
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

####  <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="34657-1115">setSelectedDataAsync (数据，[选项]，回调)</span><span class="sxs-lookup"><span data-stu-id="34657-1115">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="34657-1116">以异步方式将数据插入到邮件的正文或主题中。</span><span class="sxs-lookup"><span data-stu-id="34657-1116">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="34657-p174">`setSelectedDataAsync` 方法将指定字符串插入到项目主题或正文的光标位置，或者，如果在编辑器中已选择文本，则该方法将替换选择的文本。如果光标不在正文或主题字段中，则返回错误。插入之后，光标会位于插入内容的末尾。</span><span class="sxs-lookup"><span data-stu-id="34657-p174">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="34657-1120">参数：</span><span class="sxs-lookup"><span data-stu-id="34657-1120">Parameters:</span></span>

|<span data-ttu-id="34657-1121">名称</span><span class="sxs-lookup"><span data-stu-id="34657-1121">Name</span></span>| <span data-ttu-id="34657-1122">类型</span><span class="sxs-lookup"><span data-stu-id="34657-1122">Type</span></span>| <span data-ttu-id="34657-1123">属性</span><span class="sxs-lookup"><span data-stu-id="34657-1123">Attributes</span></span>| <span data-ttu-id="34657-1124">说明</span><span class="sxs-lookup"><span data-stu-id="34657-1124">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="34657-1125">String</span><span class="sxs-lookup"><span data-stu-id="34657-1125">String</span></span>||<span data-ttu-id="34657-p175">要插入的数据。数据不得超过 1,000,000 个字符。如果传入的数据超过 1,000,000 个字符，则会引发 `ArgumentOutOfRange` 异常。</span><span class="sxs-lookup"><span data-stu-id="34657-p175">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`| <span data-ttu-id="34657-1129">Object</span><span class="sxs-lookup"><span data-stu-id="34657-1129">Object</span></span>| <span data-ttu-id="34657-1130">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="34657-1130">&lt;optional&gt;</span></span>|<span data-ttu-id="34657-1131">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="34657-1131">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="34657-1132">Object</span><span class="sxs-lookup"><span data-stu-id="34657-1132">Object</span></span>| <span data-ttu-id="34657-1133">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="34657-1133">&lt;optional&gt;</span></span>|<span data-ttu-id="34657-1134">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="34657-1134">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`| [<span data-ttu-id="34657-1135">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="34657-1135">Office.CoercionType</span></span>](office.md#coerciontype-string)| <span data-ttu-id="34657-1136">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="34657-1136">&lt;optional&gt;</span></span>|<span data-ttu-id="34657-p176">如果是 `text`，则在 Outlook Web App 和 Outlook 中应用当前样式。如果该字段是 HTML 编辑器，则仅插入文本数据，即使数据为 HTML。</span><span class="sxs-lookup"><span data-stu-id="34657-p176">If `text`, the current style is applied in Outlook Web App and Outlook. If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="34657-p177">如果 `html` 和该字段支持 HTML（主题不支持），则在 Outlook Web App 中应用当前样式，而在 Outlook 中应用默认样式。如果该字段是文本字段，则返回 `InvalidDataFormat` 错误。</span><span class="sxs-lookup"><span data-stu-id="34657-p177">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook Web App and the default style is applied in Outlook. If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="34657-1141">如果未设置 `coercionType`，则结果取决于该字段：如果该字段是 HTML，则使用 HTML；如果该字段是文本，则使用纯文本。</span><span class="sxs-lookup"><span data-stu-id="34657-1141">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`| <span data-ttu-id="34657-1142">函数</span><span class="sxs-lookup"><span data-stu-id="34657-1142">function</span></span>||<span data-ttu-id="34657-1143">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="34657-1143">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="34657-1144">要求</span><span class="sxs-lookup"><span data-stu-id="34657-1144">Requirements</span></span>

|<span data-ttu-id="34657-1145">要求</span><span class="sxs-lookup"><span data-stu-id="34657-1145">Requirement</span></span>| <span data-ttu-id="34657-1146">值</span><span class="sxs-lookup"><span data-stu-id="34657-1146">Value</span></span>|
|---|---|
|[<span data-ttu-id="34657-1147">最低邮箱要求集版本</span><span class="sxs-lookup"><span data-stu-id="34657-1147">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="34657-1148">1.2</span><span class="sxs-lookup"><span data-stu-id="34657-1148">1.2</span></span>|
|[<span data-ttu-id="34657-1149">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="34657-1149">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="34657-1150">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="34657-1150">ReadWriteItem</span></span>|
|[<span data-ttu-id="34657-1151">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="34657-1151">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="34657-1152">撰写</span><span class="sxs-lookup"><span data-stu-id="34657-1152">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="34657-1153">示例</span><span class="sxs-lookup"><span data-stu-id="34657-1153">Example</span></span>

```
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```