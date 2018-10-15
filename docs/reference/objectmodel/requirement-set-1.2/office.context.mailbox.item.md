
# <a name="item"></a><span data-ttu-id="be26b-101">项</span><span class="sxs-lookup"><span data-stu-id="be26b-101">item</span></span>

### <span data-ttu-id="be26b-p101">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md). item</span><span class="sxs-lookup"><span data-stu-id="be26b-p101">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md). item</span></span>

<span data-ttu-id="be26b-p102">`item` 命名空间用于访问当前选定的邮件、会议请求或安排。可以通过使用 [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlook12officemailboxenumsitemtype)  属性确定 `item`  的类型。</span><span class="sxs-lookup"><span data-stu-id="be26b-p102">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlook12officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="be26b-106">要求</span><span class="sxs-lookup"><span data-stu-id="be26b-106">Requirements</span></span>

|<span data-ttu-id="be26b-107">要求</span><span class="sxs-lookup"><span data-stu-id="be26b-107">Requirement</span></span>| <span data-ttu-id="be26b-108">值</span><span class="sxs-lookup"><span data-stu-id="be26b-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="be26b-109">最低邮箱要求集版本</span><span class="sxs-lookup"><span data-stu-id="be26b-109">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="be26b-110">1.0</span><span class="sxs-lookup"><span data-stu-id="be26b-110">1.0</span></span>|
|[<span data-ttu-id="be26b-111">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="be26b-111">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="be26b-112">Restricted</span><span class="sxs-lookup"><span data-stu-id="be26b-112">Restricted</span></span>|
|[<span data-ttu-id="be26b-113">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="be26b-113">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="be26b-114">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="be26b-114">Compose or read</span></span>|

### <a name="example"></a><span data-ttu-id="be26b-115">示例</span><span class="sxs-lookup"><span data-stu-id="be26b-115">Example</span></span>

<span data-ttu-id="be26b-116">以下 JavaScript 代码示例显示了如何访问 Outlook 中当前项目的 `subject` 属性。</span><span class="sxs-lookup"><span data-stu-id="be26b-116">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

```JavaScript
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

### <a name="members"></a><span data-ttu-id="be26b-117">成员</span><span class="sxs-lookup"><span data-stu-id="be26b-117">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlook12officeattachmentdetails"></a><span data-ttu-id="be26b-118">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_2/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="be26b-118">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_2/office.attachmentdetails)></span></span>

<span data-ttu-id="be26b-p103">获取项目的附件数组。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="be26b-p103">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="be26b-121">某些类型的文件因潜在的安全问题被 Outlook 阻止，因此没有返回。</span><span class="sxs-lookup"><span data-stu-id="be26b-121">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="be26b-122">有关详细信息，请参阅 [在 Outlook 中被阻止的附件](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519)。</span><span class="sxs-lookup"><span data-stu-id="be26b-122">For more information see [Payments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="be26b-123">类型：</span><span class="sxs-lookup"><span data-stu-id="be26b-123">Type:</span></span>

*   <span data-ttu-id="be26b-124">数组。 <[AttachmentDetails](/javascript/api/outlook_1_2/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="be26b-124">Array.<[AttachmentDetails](/javascript/api/outlook_1_2/office.attachmentdetails)></span></span>

##### <a name="requirements"></a><span data-ttu-id="be26b-125">要求</span><span class="sxs-lookup"><span data-stu-id="be26b-125">Requirements</span></span>

|<span data-ttu-id="be26b-126">要求</span><span class="sxs-lookup"><span data-stu-id="be26b-126">Requirement</span></span>| <span data-ttu-id="be26b-127">值</span><span class="sxs-lookup"><span data-stu-id="be26b-127">Value</span></span>|
|---|---|
|[<span data-ttu-id="be26b-128">最低邮箱要求集版本</span><span class="sxs-lookup"><span data-stu-id="be26b-128">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="be26b-129">1.0</span><span class="sxs-lookup"><span data-stu-id="be26b-129">1.0</span></span>|
|[<span data-ttu-id="be26b-130">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="be26b-130">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="be26b-131">ReadItem</span><span class="sxs-lookup"><span data-stu-id="be26b-131">ReadItem</span></span>|
|[<span data-ttu-id="be26b-132">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="be26b-132">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="be26b-133">阅读</span><span class="sxs-lookup"><span data-stu-id="be26b-133">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="be26b-134">示例</span><span class="sxs-lookup"><span data-stu-id="be26b-134">Example</span></span>

<span data-ttu-id="be26b-135">以下代码使用当前项目上所有附件的详细信息构成 HTML 字符串。</span><span class="sxs-lookup"><span data-stu-id="be26b-135">The following code builds an HTML string with details of all attachments on the current item.</span></span>

```JavaScript
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

####  <a name="bcc-recipientsjavascriptapioutlook12officerecipients"></a><span data-ttu-id="be26b-136">密件抄送：[收件人](/javascript/api/outlook_1_2/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="be26b-136">bcc :[Recipients](/javascript/api/outlook_1_2/office.recipients)</span></span>

<span data-ttu-id="be26b-137">获取一个对象，提供用于获取或更新邮件的密件抄送 （密件抄送副本） 行上收件人的方法。</span><span class="sxs-lookup"><span data-stu-id="be26b-137">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="be26b-138">仅限撰写模式。</span><span class="sxs-lookup"><span data-stu-id="be26b-138">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="be26b-139">类型：</span><span class="sxs-lookup"><span data-stu-id="be26b-139">Type:</span></span>

*   [<span data-ttu-id="be26b-140">收件人</span><span class="sxs-lookup"><span data-stu-id="be26b-140">Recipients</span></span>](/javascript/api/outlook_1_2/office.recipients)

##### <a name="requirements"></a><span data-ttu-id="be26b-141">要求</span><span class="sxs-lookup"><span data-stu-id="be26b-141">Requirements</span></span>

|<span data-ttu-id="be26b-142">要求</span><span class="sxs-lookup"><span data-stu-id="be26b-142">Requirement</span></span>| <span data-ttu-id="be26b-143">值</span><span class="sxs-lookup"><span data-stu-id="be26b-143">Value</span></span>|
|---|---|
|[<span data-ttu-id="be26b-144">最低邮箱要求集版本</span><span class="sxs-lookup"><span data-stu-id="be26b-144">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="be26b-145">1.1</span><span class="sxs-lookup"><span data-stu-id="be26b-145">1.1</span></span>|
|[<span data-ttu-id="be26b-146">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="be26b-146">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="be26b-147">ReadItem</span><span class="sxs-lookup"><span data-stu-id="be26b-147">ReadItem</span></span>|
|[<span data-ttu-id="be26b-148">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="be26b-148">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="be26b-149">撰写</span><span class="sxs-lookup"><span data-stu-id="be26b-149">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="be26b-150">示例</span><span class="sxs-lookup"><span data-stu-id="be26b-150">Example</span></span>

```JavaScript
Office.context.mailbox.item.bcc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.bcc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.bcc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfBccRecipients = asyncResult.value;
}
```

####  <a name="body-bodyjavascriptapioutlook12officebody"></a><span data-ttu-id="be26b-151">body :[Body](/javascript/api/outlook_1_2/office.body)</span><span class="sxs-lookup"><span data-stu-id="be26b-151">body :[Body](/javascript/api/outlook_1_2/office.body)</span></span>

<span data-ttu-id="be26b-152">获取一个提供用于处理项目正文的方法的对象。</span><span class="sxs-lookup"><span data-stu-id="be26b-152">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="be26b-153">类型：</span><span class="sxs-lookup"><span data-stu-id="be26b-153">Type:</span></span>

*   [<span data-ttu-id="be26b-154">Body</span><span class="sxs-lookup"><span data-stu-id="be26b-154">Body</span></span>](/javascript/api/outlook_1_2/office.body)

##### <a name="requirements"></a><span data-ttu-id="be26b-155">要求</span><span class="sxs-lookup"><span data-stu-id="be26b-155">Requirements</span></span>

|<span data-ttu-id="be26b-156">要求</span><span class="sxs-lookup"><span data-stu-id="be26b-156">Requirement</span></span>| <span data-ttu-id="be26b-157">值</span><span class="sxs-lookup"><span data-stu-id="be26b-157">Value</span></span>|
|---|---|
|[<span data-ttu-id="be26b-158">最低的邮箱要求集版本</span><span class="sxs-lookup"><span data-stu-id="be26b-158">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="be26b-159">1.1</span><span class="sxs-lookup"><span data-stu-id="be26b-159">1.1</span></span>|
|[<span data-ttu-id="be26b-160">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="be26b-160">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="be26b-161">ReadItem</span><span class="sxs-lookup"><span data-stu-id="be26b-161">ReadItem</span></span>|
|[<span data-ttu-id="be26b-162">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="be26b-162">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="be26b-163">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="be26b-163">Compose or read</span></span>|

####  <a name="cc-arrayemailaddressdetailsjavascriptapioutlook12officeemailaddressdetailsrecipientsjavascriptapioutlook12officerecipients"></a><span data-ttu-id="be26b-164">抄送： 数组。 <[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)>|[收件人](/javascript/api/outlook_1_2/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="be26b-164">cc :Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_2/office.recipients)</span></span>

<span data-ttu-id="be26b-165">提供对邮件抄送 (cc) 收件人的访问。</span><span class="sxs-lookup"><span data-stu-id="be26b-165">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="be26b-166">对象的类型和访问级别取决于当前项的模式。</span><span class="sxs-lookup"><span data-stu-id="be26b-166">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="be26b-167">阅读模式</span><span class="sxs-lookup"><span data-stu-id="be26b-167">Read mode</span></span>

<span data-ttu-id="be26b-p107">`cc`属性返回包含邮件的**抄送**行上所列每个收件人的 `EmailAddressDetails` 对象数组。集合上限为 100 个成员。</span><span class="sxs-lookup"><span data-stu-id="be26b-p107">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="be26b-170">撰写模式</span><span class="sxs-lookup"><span data-stu-id="be26b-170">Compose mode</span></span>

<span data-ttu-id="be26b-171">`cc` 属性返回 `Recipients` 对象，该对象提供用于处理邮件**抄送**行上收件人的方法。</span><span class="sxs-lookup"><span data-stu-id="be26b-171">The `cc` property returns a `Recipients` object that provides methods for manipulating the recipients on the **Cc** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="be26b-172">类型：</span><span class="sxs-lookup"><span data-stu-id="be26b-172">Type:</span></span>

*   <span data-ttu-id="be26b-173">数组。 <[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)> |[收件人](/javascript/api/outlook_1_2/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="be26b-173">Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_2/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="be26b-174">要求</span><span class="sxs-lookup"><span data-stu-id="be26b-174">Requirements</span></span>

|<span data-ttu-id="be26b-175">要求</span><span class="sxs-lookup"><span data-stu-id="be26b-175">Requirement</span></span>| <span data-ttu-id="be26b-176">值</span><span class="sxs-lookup"><span data-stu-id="be26b-176">Value</span></span>|
|---|---|
|[<span data-ttu-id="be26b-177">最低邮箱要求集版本</span><span class="sxs-lookup"><span data-stu-id="be26b-177">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="be26b-178">1.0</span><span class="sxs-lookup"><span data-stu-id="be26b-178">1.0</span></span>|
|[<span data-ttu-id="be26b-179">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="be26b-179">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="be26b-180">ReadItem</span><span class="sxs-lookup"><span data-stu-id="be26b-180">ReadItem</span></span>|
|[<span data-ttu-id="be26b-181">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="be26b-181">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="be26b-182">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="be26b-182">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="be26b-183">示例</span><span class="sxs-lookup"><span data-stu-id="be26b-183">Example</span></span>

```JavaScript
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

####  <a name="nullable-conversationid-string"></a><span data-ttu-id="be26b-184">(nullable) conversationId :String</span><span class="sxs-lookup"><span data-stu-id="be26b-184">(nullable) conversationId :String</span></span>

<span data-ttu-id="be26b-185">获取包含特定消息的电子邮件会话的标识符。</span><span class="sxs-lookup"><span data-stu-id="be26b-185">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="be26b-p108">如果在阅读窗体或撰写窗体的回复中激活邮件应用程序，则此属性可以获得一个整数值。如果用户随后更改了回复邮件的主题（若发送回复），则该邮件的对话 ID 将改变且之前获取的值将不适用。</span><span class="sxs-lookup"><span data-stu-id="be26b-p108">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="be26b-p109">对于撰写窗体的新项目，此属性获得一个 null 值。如果用户设置一个主题并保存该项目，`conversationId` 属性将返回一个值。</span><span class="sxs-lookup"><span data-stu-id="be26b-p109">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="be26b-190">类型：</span><span class="sxs-lookup"><span data-stu-id="be26b-190">Type:</span></span>

*   <span data-ttu-id="be26b-191">字符串</span><span class="sxs-lookup"><span data-stu-id="be26b-191">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="be26b-192">要求</span><span class="sxs-lookup"><span data-stu-id="be26b-192">Requirements</span></span>

|<span data-ttu-id="be26b-193">要求</span><span class="sxs-lookup"><span data-stu-id="be26b-193">Requirement</span></span>| <span data-ttu-id="be26b-194">值</span><span class="sxs-lookup"><span data-stu-id="be26b-194">Value</span></span>|
|---|---|
|[<span data-ttu-id="be26b-195">最低邮箱要求集版本</span><span class="sxs-lookup"><span data-stu-id="be26b-195">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="be26b-196">1.0</span><span class="sxs-lookup"><span data-stu-id="be26b-196">1.0</span></span>|
|[<span data-ttu-id="be26b-197">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="be26b-197">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="be26b-198">ReadItem</span><span class="sxs-lookup"><span data-stu-id="be26b-198">ReadItem</span></span>|
|[<span data-ttu-id="be26b-199">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="be26b-199">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="be26b-200">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="be26b-200">Compose or read</span></span>|

#### <a name="datetimecreated-date"></a><span data-ttu-id="be26b-201">dateTimeCreated：日期</span><span class="sxs-lookup"><span data-stu-id="be26b-201">dateTimeCreated :Date</span></span>

<span data-ttu-id="be26b-p110">获取项目创建的日期和时间。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="be26b-p110">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="be26b-204">类型：</span><span class="sxs-lookup"><span data-stu-id="be26b-204">Type:</span></span>

*   <span data-ttu-id="be26b-205">Date</span><span class="sxs-lookup"><span data-stu-id="be26b-205">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="be26b-206">要求</span><span class="sxs-lookup"><span data-stu-id="be26b-206">Requirements</span></span>

|<span data-ttu-id="be26b-207">要求</span><span class="sxs-lookup"><span data-stu-id="be26b-207">Requirement</span></span>| <span data-ttu-id="be26b-208">值</span><span class="sxs-lookup"><span data-stu-id="be26b-208">Value</span></span>|
|---|---|
|[<span data-ttu-id="be26b-209">最低邮箱要求集版本</span><span class="sxs-lookup"><span data-stu-id="be26b-209">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="be26b-210">1.0</span><span class="sxs-lookup"><span data-stu-id="be26b-210">1.0</span></span>|
|[<span data-ttu-id="be26b-211">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="be26b-211">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="be26b-212">ReadItem</span><span class="sxs-lookup"><span data-stu-id="be26b-212">ReadItem</span></span>|
|[<span data-ttu-id="be26b-213">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="be26b-213">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="be26b-214">阅读</span><span class="sxs-lookup"><span data-stu-id="be26b-214">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="be26b-215">示例</span><span class="sxs-lookup"><span data-stu-id="be26b-215">Example</span></span>

```JavaScript
var created = Office.context.mailbox.item.dateTimeCreated;
```

#### <a name="datetimemodified-date"></a><span data-ttu-id="be26b-216">dateTimeModified： 日期</span><span class="sxs-lookup"><span data-stu-id="be26b-216">dateTimeModified :Date</span></span>

<span data-ttu-id="be26b-p111">获取项目最近一次修改的日期和时间。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="be26b-p111">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="be26b-219">注意：在 iOS 版 Outlook 或  Android 版 Outlook 中不支持此成员。</span><span class="sxs-lookup"><span data-stu-id="be26b-219">Note: This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="type"></a><span data-ttu-id="be26b-220">类型：</span><span class="sxs-lookup"><span data-stu-id="be26b-220">Type:</span></span>

*   <span data-ttu-id="be26b-221">Date</span><span class="sxs-lookup"><span data-stu-id="be26b-221">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="be26b-222">要求</span><span class="sxs-lookup"><span data-stu-id="be26b-222">Requirements</span></span>

|<span data-ttu-id="be26b-223">要求</span><span class="sxs-lookup"><span data-stu-id="be26b-223">Requirement</span></span>| <span data-ttu-id="be26b-224">值</span><span class="sxs-lookup"><span data-stu-id="be26b-224">Value</span></span>|
|---|---|
|[<span data-ttu-id="be26b-225">最低邮箱要求集版本</span><span class="sxs-lookup"><span data-stu-id="be26b-225">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="be26b-226">1.0</span><span class="sxs-lookup"><span data-stu-id="be26b-226">1.0</span></span>|
|[<span data-ttu-id="be26b-227">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="be26b-227">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="be26b-228">ReadItem</span><span class="sxs-lookup"><span data-stu-id="be26b-228">ReadItem</span></span>|
|[<span data-ttu-id="be26b-229">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="be26b-229">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="be26b-230">阅读</span><span class="sxs-lookup"><span data-stu-id="be26b-230">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="be26b-231">示例</span><span class="sxs-lookup"><span data-stu-id="be26b-231">Example</span></span>

```JavaScript
var modified = Office.context.mailbox.item.dateTimeModified;
```

####  <a name="end-datetimejavascriptapioutlook12officetime"></a><span data-ttu-id="be26b-232">end :Date|[Time](/javascript/api/outlook_1_2/office.time)</span><span class="sxs-lookup"><span data-stu-id="be26b-232">end :Date|[Time](/javascript/api/outlook_1_2/office.time)</span></span>

<span data-ttu-id="be26b-233">获取或设置约会结束的日期和时间。</span><span class="sxs-lookup"><span data-stu-id="be26b-233">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="be26b-p112">将 `end` 属性表示为协调世界时 (UTC) 的日期和时间值。可使用 [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook12officelocalclienttime) 方法将 end 属性值转换为客户端的本地日期和时间。</span><span class="sxs-lookup"><span data-stu-id="be26b-p112">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook12officelocalclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="be26b-236">阅读模式</span><span class="sxs-lookup"><span data-stu-id="be26b-236">Read mode</span></span>

<span data-ttu-id="be26b-237">`end` 属性返回 `Date` 对象。</span><span class="sxs-lookup"><span data-stu-id="be26b-237">The `end` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="be26b-238">撰写模式</span><span class="sxs-lookup"><span data-stu-id="be26b-238">Compose mode</span></span>

<span data-ttu-id="be26b-239">`end` 属性返回 `Time` 对象。</span><span class="sxs-lookup"><span data-stu-id="be26b-239">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="be26b-240">使用  方法设置结束时间时，应使用  方法将客户端的本地时间转换为服务器的 UTC。 [`Time.setAsync`](/javascript/api/outlook_1_2/office.time#setasync-datetime--options--callback-) [ `convertToUtcClientTime`  ](office.context.mailbox.md#converttoutcclienttimeinput--date)</span><span class="sxs-lookup"><span data-stu-id="be26b-240">When you use the [`Time.setAsync`](/javascript/api/outlook_1_2/office.time#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="be26b-241">类型：</span><span class="sxs-lookup"><span data-stu-id="be26b-241">Type:</span></span>

*   <span data-ttu-id="be26b-242">日期 | [时间](/javascript/api/outlook_1_2/office.time)</span><span class="sxs-lookup"><span data-stu-id="be26b-242">Date | [Time](/javascript/api/outlook_1_2/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="be26b-243">要求</span><span class="sxs-lookup"><span data-stu-id="be26b-243">Requirements</span></span>

|<span data-ttu-id="be26b-244">要求</span><span class="sxs-lookup"><span data-stu-id="be26b-244">Requirement</span></span>| <span data-ttu-id="be26b-245">值</span><span class="sxs-lookup"><span data-stu-id="be26b-245">Value</span></span>|
|---|---|
|[<span data-ttu-id="be26b-246">最低邮箱要求集版本</span><span class="sxs-lookup"><span data-stu-id="be26b-246">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="be26b-247">1.0</span><span class="sxs-lookup"><span data-stu-id="be26b-247">1.0</span></span>|
|[<span data-ttu-id="be26b-248">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="be26b-248">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="be26b-249">ReadItem</span><span class="sxs-lookup"><span data-stu-id="be26b-249">ReadItem</span></span>|
|[<span data-ttu-id="be26b-250">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="be26b-250">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="be26b-251">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="be26b-251">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="be26b-252">示例</span><span class="sxs-lookup"><span data-stu-id="be26b-252">Example</span></span>

<span data-ttu-id="be26b-253">以下示例通过使用 `Time` 对象的 [`setAsync`](/javascript/api/outlook_1_2/office.time#setasync-datetime--options--callback-) 方法，设置撰写模式下约会的结束时间。</span><span class="sxs-lookup"><span data-stu-id="be26b-253">The following example sets the end time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_2/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

```JavaScript
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

#### <a name="from-emailaddressdetailsjavascriptapioutlook12officeemailaddressdetails"></a><span data-ttu-id="be26b-254">从：[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="be26b-254">from :[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)</span></span>

<span data-ttu-id="be26b-p113">获取邮件发件人的电子邮件地址。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="be26b-p113">Gets the email address of the sender of a message. Read mode only.</span></span>

<span data-ttu-id="be26b-p114">`from` 和 [`sender`](#sender-emailaddressdetailsjavascriptapioutlook12officeemailaddressdetails) 属性表示同一个人，邮件由代理人发送的除外。在此情况下，`from` 属性表示代理程序，而 sender 属性表示代理人。</span><span class="sxs-lookup"><span data-stu-id="be26b-p114">The `from` and [`sender`](#sender-emailaddressdetailsjavascriptapioutlook12officeemailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="be26b-259">`EmailAddressDetails` 对象的 `recipientType` 属性 在 `from` 属性是 `undefined`。</span><span class="sxs-lookup"><span data-stu-id="be26b-259">Note: The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="be26b-260">类型：</span><span class="sxs-lookup"><span data-stu-id="be26b-260">Type:</span></span>

*   [<span data-ttu-id="be26b-261">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="be26b-261">EmailAddressDetails</span></span>](/javascript/api/outlook_1_2/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="be26b-262">要求</span><span class="sxs-lookup"><span data-stu-id="be26b-262">Requirements</span></span>

|<span data-ttu-id="be26b-263">要求</span><span class="sxs-lookup"><span data-stu-id="be26b-263">Requirement</span></span>| <span data-ttu-id="be26b-264">值</span><span class="sxs-lookup"><span data-stu-id="be26b-264">Value</span></span>|
|---|---|
|[<span data-ttu-id="be26b-265">最低邮箱要求集版本</span><span class="sxs-lookup"><span data-stu-id="be26b-265">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="be26b-266">1.0</span><span class="sxs-lookup"><span data-stu-id="be26b-266">1.0</span></span>|
|[<span data-ttu-id="be26b-267">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="be26b-267">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="be26b-268">ReadItem</span><span class="sxs-lookup"><span data-stu-id="be26b-268">ReadItem</span></span>|
|[<span data-ttu-id="be26b-269">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="be26b-269">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="be26b-270">阅读</span><span class="sxs-lookup"><span data-stu-id="be26b-270">Read</span></span>|

#### <a name="internetmessageid-string"></a><span data-ttu-id="be26b-271">internetMessageId :String</span><span class="sxs-lookup"><span data-stu-id="be26b-271">internetMessageId :String</span></span>

<span data-ttu-id="be26b-p115">获取电子邮件的 Internet 消息标识符。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="be26b-p115">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="be26b-274">类型：</span><span class="sxs-lookup"><span data-stu-id="be26b-274">Type:</span></span>

*   <span data-ttu-id="be26b-275">字符串</span><span class="sxs-lookup"><span data-stu-id="be26b-275">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="be26b-276">要求</span><span class="sxs-lookup"><span data-stu-id="be26b-276">Requirements</span></span>

|<span data-ttu-id="be26b-277">要求</span><span class="sxs-lookup"><span data-stu-id="be26b-277">Requirement</span></span>| <span data-ttu-id="be26b-278">值</span><span class="sxs-lookup"><span data-stu-id="be26b-278">Value</span></span>|
|---|---|
|[<span data-ttu-id="be26b-279">最低邮箱要求集版本</span><span class="sxs-lookup"><span data-stu-id="be26b-279">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="be26b-280">1.0</span><span class="sxs-lookup"><span data-stu-id="be26b-280">1.0</span></span>|
|[<span data-ttu-id="be26b-281">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="be26b-281">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="be26b-282">ReadItem</span><span class="sxs-lookup"><span data-stu-id="be26b-282">ReadItem</span></span>|
|[<span data-ttu-id="be26b-283">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="be26b-283">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="be26b-284">阅读</span><span class="sxs-lookup"><span data-stu-id="be26b-284">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="be26b-285">示例</span><span class="sxs-lookup"><span data-stu-id="be26b-285">Example</span></span>

```JavaScript
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

#### <a name="itemclass-string"></a><span data-ttu-id="be26b-286">itemClass :String</span><span class="sxs-lookup"><span data-stu-id="be26b-286">itemClass :String</span></span>

<span data-ttu-id="be26b-p116">获取选定项目的 Exchange Web 服务项目类。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="be26b-p116">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="be26b-p117">`itemClass` 属性指定所选项目的邮件类别。以下是邮件或约会项目的默认邮件类别。</span><span class="sxs-lookup"><span data-stu-id="be26b-p117">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

| <span data-ttu-id="be26b-291">类型</span><span class="sxs-lookup"><span data-stu-id="be26b-291">Type</span></span> | <span data-ttu-id="be26b-292">说明</span><span class="sxs-lookup"><span data-stu-id="be26b-292">Description</span></span> | <span data-ttu-id="be26b-293">项目类</span><span class="sxs-lookup"><span data-stu-id="be26b-293">item class</span></span> |
| --- | --- | --- |
| <span data-ttu-id="be26b-294">约会项目</span><span class="sxs-lookup"><span data-stu-id="be26b-294">Appointment items</span></span> | <span data-ttu-id="be26b-295">这些是项目类 `IPM.Appointment` 或 `IPM.Appointment.Occurence` 的日历项目。</span><span class="sxs-lookup"><span data-stu-id="be26b-295">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurence`.</span></span> | `IPM.Appointment`<br />`IPM.Appointment.Occurence` |
| <span data-ttu-id="be26b-296">邮件项目</span><span class="sxs-lookup"><span data-stu-id="be26b-296">Message items</span></span> | <span data-ttu-id="be26b-297">这些项目包括具有默认邮件类别 `IPM.Note` 的电子邮件，以及使用 `IPM.Schedule.Meeting` 作为基础邮件类别的会议请求、响应和取消。</span><span class="sxs-lookup"><span data-stu-id="be26b-297">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span> | `IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled` |

<span data-ttu-id="be26b-298">你可以创建用于扩展默认邮件类别的自定义邮件类别，例如，自定义约会邮件类别 `IPM.Appointment.Contoso`。</span><span class="sxs-lookup"><span data-stu-id="be26b-298">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="be26b-299">类型：</span><span class="sxs-lookup"><span data-stu-id="be26b-299">Type:</span></span>

*   <span data-ttu-id="be26b-300">String</span><span class="sxs-lookup"><span data-stu-id="be26b-300">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="be26b-301">要求</span><span class="sxs-lookup"><span data-stu-id="be26b-301">Requirements</span></span>

|<span data-ttu-id="be26b-302">要求</span><span class="sxs-lookup"><span data-stu-id="be26b-302">Requirement</span></span>| <span data-ttu-id="be26b-303">值</span><span class="sxs-lookup"><span data-stu-id="be26b-303">Value</span></span>|
|---|---|
|[<span data-ttu-id="be26b-304">最低邮箱要求集版本</span><span class="sxs-lookup"><span data-stu-id="be26b-304">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="be26b-305">1.0</span><span class="sxs-lookup"><span data-stu-id="be26b-305">1.0</span></span>|
|[<span data-ttu-id="be26b-306">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="be26b-306">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="be26b-307">ReadItem</span><span class="sxs-lookup"><span data-stu-id="be26b-307">ReadItem</span></span>|
|[<span data-ttu-id="be26b-308">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="be26b-308">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="be26b-309">阅读</span><span class="sxs-lookup"><span data-stu-id="be26b-309">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="be26b-310">示例</span><span class="sxs-lookup"><span data-stu-id="be26b-310">Example</span></span>

```JavaScript
var itemClass = Office.context.mailbox.item.itemClass;
```

#### <a name="nullable-itemid-string"></a><span data-ttu-id="be26b-311">（可空类型）itemId ：字符串</span><span class="sxs-lookup"><span data-stu-id="be26b-311">(nullable) itemId :String</span></span>

<span data-ttu-id="be26b-p118">获取当前项目的 Exchange Web 服务项目标识符。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="be26b-p118">Gets the Exchange Web Services item identifier for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="be26b-314">`itemId` 属性返回的标识符与 Exchange Web 服务项目标识符相同。</span><span class="sxs-lookup"><span data-stu-id="be26b-314">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="be26b-315">`itemId` 属性与 Outlook 条目 ID 或使用 Outlook REST API 的 ID不同。</span><span class="sxs-lookup"><span data-stu-id="be26b-315">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="be26b-316">使用此值进行 REST API 调用之前，应使用 `Office.context.mailbox.convertToRestId` 对其转换，该功能自要求集 1.3 起提供。</span><span class="sxs-lookup"><span data-stu-id="be26b-316">Before making REST API calls using this value, it should be converted using `Office.context.mailbox.convertToRestId`, which is available starting in requirement set 1.3.</span></span> <span data-ttu-id="be26b-317">更多详情请参阅[从 Outlook 加载项中使用 Outlook REST API](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id)。</span><span class="sxs-lookup"><span data-stu-id="be26b-317">For more details, see [Use the Outlook REST APIs from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

##### <a name="type"></a><span data-ttu-id="be26b-318">类型：</span><span class="sxs-lookup"><span data-stu-id="be26b-318">Type:</span></span>

*   <span data-ttu-id="be26b-319">字符串</span><span class="sxs-lookup"><span data-stu-id="be26b-319">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="be26b-320">要求</span><span class="sxs-lookup"><span data-stu-id="be26b-320">Requirements</span></span>

|<span data-ttu-id="be26b-321">要求</span><span class="sxs-lookup"><span data-stu-id="be26b-321">Requirement</span></span>| <span data-ttu-id="be26b-322">值</span><span class="sxs-lookup"><span data-stu-id="be26b-322">Value</span></span>|
|---|---|
|[<span data-ttu-id="be26b-323">最低邮箱要求集版本</span><span class="sxs-lookup"><span data-stu-id="be26b-323">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="be26b-324">1.0</span><span class="sxs-lookup"><span data-stu-id="be26b-324">1.0</span></span>|
|[<span data-ttu-id="be26b-325">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="be26b-325">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="be26b-326">ReadItem</span><span class="sxs-lookup"><span data-stu-id="be26b-326">ReadItem</span></span>|
|[<span data-ttu-id="be26b-327">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="be26b-327">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="be26b-328">阅读</span><span class="sxs-lookup"><span data-stu-id="be26b-328">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="be26b-329">示例</span><span class="sxs-lookup"><span data-stu-id="be26b-329">Example</span></span>

<span data-ttu-id="be26b-p120">以下代码检查项目标识符是否存在。如果 `itemId` 属性返回 `null` 或 `undefined`，则将项目保存到存储，并从异步结果中获取项目标识符。</span><span class="sxs-lookup"><span data-stu-id="be26b-p120">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

```JavaScript
var itemId = Office.context.mailbox.item.itemId;
if (itemId === null || itemId == undefined) {
  Office.context.mailbox.item.saveAsync(function(result){
    itemId = result.value;
  });
}
```

####  <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlook12officemailboxenumsitemtype"></a><span data-ttu-id="be26b-332">itemType:[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_2/office.mailboxenums.itemtype)</span><span class="sxs-lookup"><span data-stu-id="be26b-332">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_2/office.mailboxenums.itemtype)</span></span>

<span data-ttu-id="be26b-333">获取实例代表项的类型。</span><span class="sxs-lookup"><span data-stu-id="be26b-333">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="be26b-334">`itemType` 属性返回其中一个 `ItemType` 枚举值，指示 `item` 对象实例是邮件还是约会。</span><span class="sxs-lookup"><span data-stu-id="be26b-334">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="be26b-335">类型：</span><span class="sxs-lookup"><span data-stu-id="be26b-335">Type:</span></span>

*   [<span data-ttu-id="be26b-336">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="be26b-336">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook_1_2/office.mailboxenums.itemtype)

##### <a name="requirements"></a><span data-ttu-id="be26b-337">要求</span><span class="sxs-lookup"><span data-stu-id="be26b-337">Requirements</span></span>

|<span data-ttu-id="be26b-338">要求</span><span class="sxs-lookup"><span data-stu-id="be26b-338">Requirement</span></span>| <span data-ttu-id="be26b-339">值</span><span class="sxs-lookup"><span data-stu-id="be26b-339">Value</span></span>|
|---|---|
|[<span data-ttu-id="be26b-340">最低邮箱要求集版本</span><span class="sxs-lookup"><span data-stu-id="be26b-340">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="be26b-341">1.0</span><span class="sxs-lookup"><span data-stu-id="be26b-341">1.0</span></span>|
|[<span data-ttu-id="be26b-342">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="be26b-342">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="be26b-343">ReadItem</span><span class="sxs-lookup"><span data-stu-id="be26b-343">ReadItem</span></span>|
|[<span data-ttu-id="be26b-344">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="be26b-344">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="be26b-345">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="be26b-345">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="be26b-346">示例</span><span class="sxs-lookup"><span data-stu-id="be26b-346">Example</span></span>

```JavaScript
if (Office.context.mailbox.item.itemType == Office.MailboxEnums.ItemType.Message)
  // do something
else
  // do something else
```

####  <a name="location-stringlocationjavascriptapioutlook12officelocation"></a><span data-ttu-id="be26b-347">location :String|[Location](/javascript/api/outlook_1_2/office.location)</span><span class="sxs-lookup"><span data-stu-id="be26b-347">location :String|[Location](/javascript/api/outlook_1_2/office.location)</span></span>

<span data-ttu-id="be26b-348">获取或设置约会的位置。</span><span class="sxs-lookup"><span data-stu-id="be26b-348">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="be26b-349">阅读模式</span><span class="sxs-lookup"><span data-stu-id="be26b-349">Read mode</span></span>

<span data-ttu-id="be26b-350">`location` 属性返回一个包含约会位置的字符串。</span><span class="sxs-lookup"><span data-stu-id="be26b-350">The `location` property returns a string that contains the location of the appointment.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="be26b-351">撰写模式</span><span class="sxs-lookup"><span data-stu-id="be26b-351">Compose mode</span></span>

<span data-ttu-id="be26b-352">`location` 属性返回一个 `Location` 对象，该对象提供用于获取和设置约会位置的方法。</span><span class="sxs-lookup"><span data-stu-id="be26b-352">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="be26b-353">类型：</span><span class="sxs-lookup"><span data-stu-id="be26b-353">Type:</span></span>

*   <span data-ttu-id="be26b-354">字符串 | [位置](/javascript/api/outlook_1_2/office.location)</span><span class="sxs-lookup"><span data-stu-id="be26b-354">String | [Location](/javascript/api/outlook_1_2/office.location)</span></span>

##### <a name="requirements"></a><span data-ttu-id="be26b-355">要求</span><span class="sxs-lookup"><span data-stu-id="be26b-355">Requirements</span></span>

|<span data-ttu-id="be26b-356">要求</span><span class="sxs-lookup"><span data-stu-id="be26b-356">Requirement</span></span>| <span data-ttu-id="be26b-357">值</span><span class="sxs-lookup"><span data-stu-id="be26b-357">Value</span></span>|
|---|---|
|[<span data-ttu-id="be26b-358">最低邮箱要求集版本</span><span class="sxs-lookup"><span data-stu-id="be26b-358">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="be26b-359">1.0</span><span class="sxs-lookup"><span data-stu-id="be26b-359">1.0</span></span>|
|[<span data-ttu-id="be26b-360">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="be26b-360">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="be26b-361">ReadItem</span><span class="sxs-lookup"><span data-stu-id="be26b-361">ReadItem</span></span>|
|[<span data-ttu-id="be26b-362">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="be26b-362">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="be26b-363">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="be26b-363">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="be26b-364">示例</span><span class="sxs-lookup"><span data-stu-id="be26b-364">Example</span></span>

```JavaScript
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

#### <a name="normalizedsubject-string"></a><span data-ttu-id="be26b-365">normalizedSubject :String</span><span class="sxs-lookup"><span data-stu-id="be26b-365">normalizedSubject :String</span></span>

<span data-ttu-id="be26b-p121">获取删除了所有前缀（包括 `RE:` 和 `FWD:`）的项目主题。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="be26b-p121">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="be26b-p122">normalizedSubject 属性获取包含由电子邮件程序添加的任何标准前缀（如 `RE:` 和 `FW:`）的项目主题。若要获取包含完整前缀的项目主题，请使用 [`subject`](#subject-stringsubjectjavascriptapioutlook12officesubject) 属性。</span><span class="sxs-lookup"><span data-stu-id="be26b-p122">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubjectjavascriptapioutlook12officesubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="be26b-370">类型：</span><span class="sxs-lookup"><span data-stu-id="be26b-370">Type:</span></span>

*   <span data-ttu-id="be26b-371">String</span><span class="sxs-lookup"><span data-stu-id="be26b-371">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="be26b-372">要求</span><span class="sxs-lookup"><span data-stu-id="be26b-372">Requirements</span></span>

|<span data-ttu-id="be26b-373">要求</span><span class="sxs-lookup"><span data-stu-id="be26b-373">Requirement</span></span>| <span data-ttu-id="be26b-374">值</span><span class="sxs-lookup"><span data-stu-id="be26b-374">Value</span></span>|
|---|---|
|[<span data-ttu-id="be26b-375">最低邮箱要求集版本</span><span class="sxs-lookup"><span data-stu-id="be26b-375">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="be26b-376">1.0</span><span class="sxs-lookup"><span data-stu-id="be26b-376">1.0</span></span>|
|[<span data-ttu-id="be26b-377">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="be26b-377">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="be26b-378">ReadItem</span><span class="sxs-lookup"><span data-stu-id="be26b-378">ReadItem</span></span>|
|[<span data-ttu-id="be26b-379">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="be26b-379">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="be26b-380">阅读</span><span class="sxs-lookup"><span data-stu-id="be26b-380">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="be26b-381">示例</span><span class="sxs-lookup"><span data-stu-id="be26b-381">Example</span></span>

```JavaScript
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
```

####  <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlook12officeemailaddressdetailsrecipientsjavascriptapioutlook12officerecipients"></a><span data-ttu-id="be26b-382">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_2/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="be26b-382">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_2/office.recipients)</span></span>

<span data-ttu-id="be26b-383">提供对事件可选与会者的访问。</span><span class="sxs-lookup"><span data-stu-id="be26b-383">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="be26b-384">对象的类型和访问级别取决于当前项的模式。</span><span class="sxs-lookup"><span data-stu-id="be26b-384">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="be26b-385">阅读模式</span><span class="sxs-lookup"><span data-stu-id="be26b-385">Read mode</span></span>

<span data-ttu-id="be26b-386">`optionalAttendees` 属性返回一个数组，其中包含每个可选与会者的 `EmailAddressDetails` 对象。</span><span class="sxs-lookup"><span data-stu-id="be26b-386">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="be26b-387">撰写模式</span><span class="sxs-lookup"><span data-stu-id="be26b-387">Compose mode</span></span>

<span data-ttu-id="be26b-388">`optionalAttendees` 属性返回一个 `Recipients` 对象，该对象提供用于获取和设置可选与会者的方法。</span><span class="sxs-lookup"><span data-stu-id="be26b-388">The `optionalAttendees` property returns a `Recipients` object that provides methods to get and set the optional attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="be26b-389">类型：</span><span class="sxs-lookup"><span data-stu-id="be26b-389">Type:</span></span>

*   <span data-ttu-id="be26b-390">数组。 <[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)> |[收件人](/javascript/api/outlook_1_2/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="be26b-390">Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_2/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="be26b-391">要求</span><span class="sxs-lookup"><span data-stu-id="be26b-391">Requirements</span></span>

|<span data-ttu-id="be26b-392">要求</span><span class="sxs-lookup"><span data-stu-id="be26b-392">Requirement</span></span>| <span data-ttu-id="be26b-393">值</span><span class="sxs-lookup"><span data-stu-id="be26b-393">Value</span></span>|
|---|---|
|[<span data-ttu-id="be26b-394">最低邮箱要求集版本</span><span class="sxs-lookup"><span data-stu-id="be26b-394">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="be26b-395">1.0</span><span class="sxs-lookup"><span data-stu-id="be26b-395">1.0</span></span>|
|[<span data-ttu-id="be26b-396">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="be26b-396">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="be26b-397">ReadItem</span><span class="sxs-lookup"><span data-stu-id="be26b-397">ReadItem</span></span>|
|[<span data-ttu-id="be26b-398">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="be26b-398">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="be26b-399">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="be26b-399">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="be26b-400">示例</span><span class="sxs-lookup"><span data-stu-id="be26b-400">Example</span></span>

```JavaScript
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

#### <a name="organizer-emailaddressdetailsjavascriptapioutlook12officeemailaddressdetails"></a><span data-ttu-id="be26b-401">组织者：[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="be26b-401">organizer :[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)</span></span>

<span data-ttu-id="be26b-p124">获取指定会议的会议组织者的电子邮件地址。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="be26b-p124">Gets the email address of the meeting organizer for a specified meeting. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="be26b-404">类型：</span><span class="sxs-lookup"><span data-stu-id="be26b-404">Type:</span></span>

*   [<span data-ttu-id="be26b-405">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="be26b-405">EmailAddressDetails</span></span>](/javascript/api/outlook_1_2/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="be26b-406">要求</span><span class="sxs-lookup"><span data-stu-id="be26b-406">Requirements</span></span>

|<span data-ttu-id="be26b-407">要求</span><span class="sxs-lookup"><span data-stu-id="be26b-407">Requirement</span></span>| <span data-ttu-id="be26b-408">值</span><span class="sxs-lookup"><span data-stu-id="be26b-408">Value</span></span>|
|---|---|
|[<span data-ttu-id="be26b-409">最低邮箱要求集版本</span><span class="sxs-lookup"><span data-stu-id="be26b-409">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="be26b-410">1.0</span><span class="sxs-lookup"><span data-stu-id="be26b-410">1.0</span></span>|
|[<span data-ttu-id="be26b-411">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="be26b-411">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="be26b-412">ReadItem</span><span class="sxs-lookup"><span data-stu-id="be26b-412">ReadItem</span></span>|
|[<span data-ttu-id="be26b-413">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="be26b-413">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="be26b-414">阅读</span><span class="sxs-lookup"><span data-stu-id="be26b-414">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="be26b-415">示例</span><span class="sxs-lookup"><span data-stu-id="be26b-415">Example</span></span>

```JavaScript
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
```

####  <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlook12officeemailaddressdetailsrecipientsjavascriptapioutlook12officerecipients"></a><span data-ttu-id="be26b-416">requiredAttendees： 数组。 <[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)>|[收件人](/javascript/api/outlook_1_2/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="be26b-416">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_2/office.recipients)</span></span>

<span data-ttu-id="be26b-417">提供对事件必需与会者的访问。</span><span class="sxs-lookup"><span data-stu-id="be26b-417">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="be26b-418">对象的类型和访问级别取决于当前项的模式。</span><span class="sxs-lookup"><span data-stu-id="be26b-418">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="be26b-419">阅读模式</span><span class="sxs-lookup"><span data-stu-id="be26b-419">Read mode</span></span>

<span data-ttu-id="be26b-420">`requiredAttendees` 属性返回一个数组，其中包含每个必需与会者的 `EmailAddressDetails` 对象。</span><span class="sxs-lookup"><span data-stu-id="be26b-420">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="be26b-421">撰写模式</span><span class="sxs-lookup"><span data-stu-id="be26b-421">Compose mode</span></span>

<span data-ttu-id="be26b-422">`requiredAttendees` 属性返回一个 `Recipients` 对象，该对象提供用于获取和设置可选与会者的方法。</span><span class="sxs-lookup"><span data-stu-id="be26b-422">The `requiredAttendees` property returns a `Recipients` object that provides methods to get and set the required attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="be26b-423">类型：</span><span class="sxs-lookup"><span data-stu-id="be26b-423">Type:</span></span>

*   <span data-ttu-id="be26b-424">数组。 <[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)> |[收件人](/javascript/api/outlook_1_2/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="be26b-424">Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_2/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="be26b-425">要求</span><span class="sxs-lookup"><span data-stu-id="be26b-425">Requirements</span></span>

|<span data-ttu-id="be26b-426">要求</span><span class="sxs-lookup"><span data-stu-id="be26b-426">Requirement</span></span>| <span data-ttu-id="be26b-427">值</span><span class="sxs-lookup"><span data-stu-id="be26b-427">Value</span></span>|
|---|---|
|[<span data-ttu-id="be26b-428">最低邮箱要求集版本</span><span class="sxs-lookup"><span data-stu-id="be26b-428">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="be26b-429">1.0</span><span class="sxs-lookup"><span data-stu-id="be26b-429">1.0</span></span>|
|[<span data-ttu-id="be26b-430">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="be26b-430">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="be26b-431">ReadItem</span><span class="sxs-lookup"><span data-stu-id="be26b-431">ReadItem</span></span>|
|[<span data-ttu-id="be26b-432">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="be26b-432">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="be26b-433">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="be26b-433">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="be26b-434">示例</span><span class="sxs-lookup"><span data-stu-id="be26b-434">Example</span></span>

```JavaScript
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
}
```

#### <a name="sender-emailaddressdetailsjavascriptapioutlook12officeemailaddressdetails"></a><span data-ttu-id="be26b-435">发件人：[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="be26b-435">sender :[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)</span></span>

<span data-ttu-id="be26b-p126">获取电子邮件发件人的电子邮件地址。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="be26b-p126">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="be26b-p127">[`from`](#from-emailaddressdetailsjavascriptapioutlook12officeemailaddressdetails) 和 `sender` 属性表示同一个人，邮件由代理人发送的除外。在此情况下，`from` 属性表示代理程序，而 sender 属性表示代理人。</span><span class="sxs-lookup"><span data-stu-id="be26b-p127">The [`from`](#from-emailaddressdetailsjavascriptapioutlook12officeemailaddressdetails) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="be26b-440">`EmailAddressDetails` 对象的 `recipientType` 属性在 `sender` 属性是 `undefined`。</span><span class="sxs-lookup"><span data-stu-id="be26b-440">Note: The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="be26b-441">类型：</span><span class="sxs-lookup"><span data-stu-id="be26b-441">Type:</span></span>

*   [<span data-ttu-id="be26b-442">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="be26b-442">EmailAddressDetails</span></span>](/javascript/api/outlook_1_2/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="be26b-443">要求</span><span class="sxs-lookup"><span data-stu-id="be26b-443">Requirements</span></span>

|<span data-ttu-id="be26b-444">要求</span><span class="sxs-lookup"><span data-stu-id="be26b-444">Requirement</span></span>| <span data-ttu-id="be26b-445">值</span><span class="sxs-lookup"><span data-stu-id="be26b-445">Value</span></span>|
|---|---|
|[<span data-ttu-id="be26b-446">最低邮箱要求集版本</span><span class="sxs-lookup"><span data-stu-id="be26b-446">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="be26b-447">1.0</span><span class="sxs-lookup"><span data-stu-id="be26b-447">1.0</span></span>|
|[<span data-ttu-id="be26b-448">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="be26b-448">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="be26b-449">ReadItem</span><span class="sxs-lookup"><span data-stu-id="be26b-449">ReadItem</span></span>|
|[<span data-ttu-id="be26b-450">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="be26b-450">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="be26b-451">阅读</span><span class="sxs-lookup"><span data-stu-id="be26b-451">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="be26b-452">示例</span><span class="sxs-lookup"><span data-stu-id="be26b-452">Example</span></span>

```JavaScript
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
```

####  <a name="start-datetimejavascriptapioutlook12officetime"></a><span data-ttu-id="be26b-453">start :Date|[Time](/javascript/api/outlook_1_2/office.time)</span><span class="sxs-lookup"><span data-stu-id="be26b-453">start :Date|[Time](/javascript/api/outlook_1_2/office.time)</span></span>

<span data-ttu-id="be26b-454">获取或设置约会开始的日期和时间。</span><span class="sxs-lookup"><span data-stu-id="be26b-454">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="be26b-p128">将 `start` 属性表示为协调世界时 (UTC) 的日期和时间值。可使用 [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook12officelocalclienttime) 方法将该值转换为客户端的本地日期和时间。</span><span class="sxs-lookup"><span data-stu-id="be26b-p128">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook12officelocalclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="be26b-457">阅读模式</span><span class="sxs-lookup"><span data-stu-id="be26b-457">Read mode</span></span>

<span data-ttu-id="be26b-458">`start` 属性返回 `Date` 对象。</span><span class="sxs-lookup"><span data-stu-id="be26b-458">The `start` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="be26b-459">撰写模式</span><span class="sxs-lookup"><span data-stu-id="be26b-459">Compose mode</span></span>

<span data-ttu-id="be26b-460">`start` 属性返回 `Time` 对象。</span><span class="sxs-lookup"><span data-stu-id="be26b-460">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="be26b-461">使用 [`Time.setAsync`](/javascript/api/outlook_1_2/office.time#setasync-datetime--options--callback-) 方法设置开始时间时，应使用 [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) 方法将客户端的本地时间转换为服务器的 UTC。</span><span class="sxs-lookup"><span data-stu-id="be26b-461">When you use the [`Time.setAsync`](/javascript/api/outlook_1_2/office.time#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="be26b-462">类型：</span><span class="sxs-lookup"><span data-stu-id="be26b-462">Type:</span></span>

*   <span data-ttu-id="be26b-463">日期 | [时间](/javascript/api/outlook_1_2/office.time)</span><span class="sxs-lookup"><span data-stu-id="be26b-463">Date | [Time](/javascript/api/outlook_1_2/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="be26b-464">要求</span><span class="sxs-lookup"><span data-stu-id="be26b-464">Requirements</span></span>

|<span data-ttu-id="be26b-465">要求</span><span class="sxs-lookup"><span data-stu-id="be26b-465">Requirement</span></span>| <span data-ttu-id="be26b-466">值</span><span class="sxs-lookup"><span data-stu-id="be26b-466">Value</span></span>|
|---|---|
|[<span data-ttu-id="be26b-467">最低邮箱要求集版本</span><span class="sxs-lookup"><span data-stu-id="be26b-467">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="be26b-468">1.0</span><span class="sxs-lookup"><span data-stu-id="be26b-468">1.0</span></span>|
|[<span data-ttu-id="be26b-469">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="be26b-469">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="be26b-470">ReadItem</span><span class="sxs-lookup"><span data-stu-id="be26b-470">ReadItem</span></span>|
|[<span data-ttu-id="be26b-471">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="be26b-471">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="be26b-472">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="be26b-472">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="be26b-473">示例</span><span class="sxs-lookup"><span data-stu-id="be26b-473">Example</span></span>

<span data-ttu-id="be26b-474">以下示例通过使用 `Time` 对象的 [`setAsync`](/javascript/api/outlook_1_2/office.time#setasync-datetime--options--callback-) 方法，设置撰写模式下约会的开始时间。</span><span class="sxs-lookup"><span data-stu-id="be26b-474">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_2/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

```JavaScript
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

####  <a name="subject-stringsubjectjavascriptapioutlook12officesubject"></a><span data-ttu-id="be26b-475">主题： 字符串 |[主题](/javascript/api/outlook_1_2/office.subject)</span><span class="sxs-lookup"><span data-stu-id="be26b-475">subject :String|[Subject](/javascript/api/outlook_1_2/office.subject)</span></span>

<span data-ttu-id="be26b-476">获取或设置显示在项目的主题字段中的说明。</span><span class="sxs-lookup"><span data-stu-id="be26b-476">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="be26b-477">`subject` 属性获取或设置由电子邮件服务器发送项目时的整个主题。</span><span class="sxs-lookup"><span data-stu-id="be26b-477">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="be26b-478">阅读模式</span><span class="sxs-lookup"><span data-stu-id="be26b-478">Read mode</span></span>

<span data-ttu-id="be26b-p129">`subject` 属性返回一个字符串。使用 [`normalizedSubject`](#normalizedsubject-string) 属性获取不带任何前导前缀（如 `RE:` 和 `FW:`）的主题。</span><span class="sxs-lookup"><span data-stu-id="be26b-p129">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

```
var subject = Office.context.mailbox.item.subject;
```

##### <a name="compose-mode"></a><span data-ttu-id="be26b-481">撰写模式</span><span class="sxs-lookup"><span data-stu-id="be26b-481">Compose mode</span></span>

<span data-ttu-id="be26b-482">`subject` 属性返回一个 `Subject` 对象，该对象提供用于获取和设置主题的方法。</span><span class="sxs-lookup"><span data-stu-id="be26b-482">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```JavaScript
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="be26b-483">类型：</span><span class="sxs-lookup"><span data-stu-id="be26b-483">Type:</span></span>

*   <span data-ttu-id="be26b-484">字符串 | [主题](/javascript/api/outlook_1_2/office.subject)</span><span class="sxs-lookup"><span data-stu-id="be26b-484">String | [Subject](/javascript/api/outlook_1_2/office.subject)</span></span>

##### <a name="requirements"></a><span data-ttu-id="be26b-485">要求</span><span class="sxs-lookup"><span data-stu-id="be26b-485">Requirements</span></span>

|<span data-ttu-id="be26b-486">要求</span><span class="sxs-lookup"><span data-stu-id="be26b-486">Requirement</span></span>| <span data-ttu-id="be26b-487">值</span><span class="sxs-lookup"><span data-stu-id="be26b-487">Value</span></span>|
|---|---|
|[<span data-ttu-id="be26b-488">最低邮箱要求集版本</span><span class="sxs-lookup"><span data-stu-id="be26b-488">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="be26b-489">1.0</span><span class="sxs-lookup"><span data-stu-id="be26b-489">1.0</span></span>|
|[<span data-ttu-id="be26b-490">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="be26b-490">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="be26b-491">ReadItem</span><span class="sxs-lookup"><span data-stu-id="be26b-491">ReadItem</span></span>|
|[<span data-ttu-id="be26b-492">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="be26b-492">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="be26b-493">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="be26b-493">Compose or read</span></span>|

####  <a name="to-arrayemailaddressdetailsjavascriptapioutlook12officeemailaddressdetailsrecipientsjavascriptapioutlook12officerecipients"></a><span data-ttu-id="be26b-494">发送： 数组。 <[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)>|[收件人](/javascript/api/outlook_1_2/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="be26b-494">to :Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_2/office.recipients)</span></span>

<span data-ttu-id="be26b-495">提供对邮件的 **发送** 行上收件人的访问。</span><span class="sxs-lookup"><span data-stu-id="be26b-495">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="be26b-496">对象的类型和访问级别取决于当前项的模式。</span><span class="sxs-lookup"><span data-stu-id="be26b-496">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="be26b-497">阅读模式</span><span class="sxs-lookup"><span data-stu-id="be26b-497">Read mode</span></span>

<span data-ttu-id="be26b-p131">`to` 属性返回包含邮件的**收件人**行上所列的每个收件人的 `EmailAddressDetails` 对象的数组。集合上限为 100 个成员。</span><span class="sxs-lookup"><span data-stu-id="be26b-p131">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="be26b-500">撰写模式</span><span class="sxs-lookup"><span data-stu-id="be26b-500">Compose mode</span></span>

<span data-ttu-id="be26b-501">`to` 属性返回 `Recipients` 对象，该对象提供用于处理邮件**收件人**行上收件人的方法。</span><span class="sxs-lookup"><span data-stu-id="be26b-501">The to`to` property returns a Recipients`Recipients` object that provides methods for manipulating the recipients on the **To** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="be26b-502">类型：</span><span class="sxs-lookup"><span data-stu-id="be26b-502">Type:</span></span>

*   <span data-ttu-id="be26b-503">数组。 <[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)> |[收件人](/javascript/api/outlook_1_2/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="be26b-503">Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_2/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="be26b-504">要求</span><span class="sxs-lookup"><span data-stu-id="be26b-504">Requirements</span></span>

|<span data-ttu-id="be26b-505">要求</span><span class="sxs-lookup"><span data-stu-id="be26b-505">Requirement</span></span>| <span data-ttu-id="be26b-506">值</span><span class="sxs-lookup"><span data-stu-id="be26b-506">Value</span></span>|
|---|---|
|[<span data-ttu-id="be26b-507">最低邮箱要求集版本</span><span class="sxs-lookup"><span data-stu-id="be26b-507">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="be26b-508">1.0</span><span class="sxs-lookup"><span data-stu-id="be26b-508">1.0</span></span>|
|[<span data-ttu-id="be26b-509">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="be26b-509">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="be26b-510">ReadItem</span><span class="sxs-lookup"><span data-stu-id="be26b-510">ReadItem</span></span>|
|[<span data-ttu-id="be26b-511">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="be26b-511">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="be26b-512">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="be26b-512">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="be26b-513">示例</span><span class="sxs-lookup"><span data-stu-id="be26b-513">Example</span></span>

```JavaScript
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

### <a name="methods"></a><span data-ttu-id="be26b-514">方法</span><span class="sxs-lookup"><span data-stu-id="be26b-514">Methods</span></span>

####  <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="be26b-515">addFileAttachmentAsync (uri，attachmentName，[选项] [回调])</span><span class="sxs-lookup"><span data-stu-id="be26b-515">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="be26b-516">将文件作为附件添加到邮件或约会。</span><span class="sxs-lookup"><span data-stu-id="be26b-516">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="be26b-517">`addFileAttachmentAsync` 方法在指定的 URI 上载文件并将其附加到撰写窗体中的项目。</span><span class="sxs-lookup"><span data-stu-id="be26b-517">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="be26b-518">你随后可以将该标识符与 [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) 方法一同使用，以删除同一个会话中的附件。</span><span class="sxs-lookup"><span data-stu-id="be26b-518">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="be26b-519">参数：</span><span class="sxs-lookup"><span data-stu-id="be26b-519">Parameters:</span></span>

|<span data-ttu-id="be26b-520">名称</span><span class="sxs-lookup"><span data-stu-id="be26b-520">Name</span></span>| <span data-ttu-id="be26b-521">类型</span><span class="sxs-lookup"><span data-stu-id="be26b-521">Type</span></span>| <span data-ttu-id="be26b-522">属性</span><span class="sxs-lookup"><span data-stu-id="be26b-522">Attributes</span></span>| <span data-ttu-id="be26b-523">说明</span><span class="sxs-lookup"><span data-stu-id="be26b-523">Description</span></span>|
|---|---|---|---|
|`uri`| <span data-ttu-id="be26b-524">字符串</span><span class="sxs-lookup"><span data-stu-id="be26b-524">String</span></span>||<span data-ttu-id="be26b-p132">提供附加到邮件或约会的文件的位置的 URI。最大长度为 2048 个字符。</span><span class="sxs-lookup"><span data-stu-id="be26b-p132">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="be26b-527">字符串</span><span class="sxs-lookup"><span data-stu-id="be26b-527">String</span></span>||<span data-ttu-id="be26b-p133">在附件上载过程中显示的附件名称。最大长度为 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="be26b-p133">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="be26b-530">Object</span><span class="sxs-lookup"><span data-stu-id="be26b-530">Object</span></span>| <span data-ttu-id="be26b-531">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="be26b-531">&lt;optional&gt;</span></span>|<span data-ttu-id="be26b-532">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="be26b-532">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="be26b-533">Object</span><span class="sxs-lookup"><span data-stu-id="be26b-533">Object</span></span>| <span data-ttu-id="be26b-534">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="be26b-534">&lt;optional&gt;</span></span>|<span data-ttu-id="be26b-535">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="be26b-535">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="be26b-536">function</span><span class="sxs-lookup"><span data-stu-id="be26b-536">function</span></span>| <span data-ttu-id="be26b-537">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="be26b-537">&lt;optional&gt;</span></span>|<span data-ttu-id="be26b-538">方法完成后，使用单个参数 `asyncResult` （一个 [`AsyncResult`](/javascript/api/office/office.asyncresult)   对象）调用在 `callback`  参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="be26b-538">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="be26b-539">如果成功，附件标识符将在 `asyncResult.value` 属性中提供。</span><span class="sxs-lookup"><span data-stu-id="be26b-539">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="be26b-540">如果上传附件失败，`asyncResult` 对象将包含一个提供错误说明的 `Error` 对象。</span><span class="sxs-lookup"><span data-stu-id="be26b-540">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="be26b-541">错误</span><span class="sxs-lookup"><span data-stu-id="be26b-541">Errors</span></span>

| <span data-ttu-id="be26b-542">错误代码</span><span class="sxs-lookup"><span data-stu-id="be26b-542">Error code</span></span> | <span data-ttu-id="be26b-543">说明</span><span class="sxs-lookup"><span data-stu-id="be26b-543">Description</span></span> |
|------------|-------------|
| `AttachmentSizeExceeded` | <span data-ttu-id="be26b-544">附件大小超过了允许的大小。</span><span class="sxs-lookup"><span data-stu-id="be26b-544">The attachment is larger than allowed.</span></span> |
| `FileTypeNotSupported` | <span data-ttu-id="be26b-545">该附件的扩展名不是允许的扩展名。</span><span class="sxs-lookup"><span data-stu-id="be26b-545">The attachment has an extension that is not allowed.</span></span> |
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="be26b-546">邮件或约会具有的附件过多。</span><span class="sxs-lookup"><span data-stu-id="be26b-546">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="be26b-547">要求</span><span class="sxs-lookup"><span data-stu-id="be26b-547">Requirements</span></span>

|<span data-ttu-id="be26b-548">要求</span><span class="sxs-lookup"><span data-stu-id="be26b-548">Requirement</span></span>| <span data-ttu-id="be26b-549">值</span><span class="sxs-lookup"><span data-stu-id="be26b-549">Value</span></span>|
|---|---|
|[<span data-ttu-id="be26b-550">最低的邮箱要求集版本</span><span class="sxs-lookup"><span data-stu-id="be26b-550">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="be26b-551">1.1</span><span class="sxs-lookup"><span data-stu-id="be26b-551">1.1</span></span>|
|[<span data-ttu-id="be26b-552">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="be26b-552">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="be26b-553">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="be26b-553">ReadWriteItem</span></span>|
|[<span data-ttu-id="be26b-554">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="be26b-554">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="be26b-555">撰写</span><span class="sxs-lookup"><span data-stu-id="be26b-555">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="be26b-556">示例</span><span class="sxs-lookup"><span data-stu-id="be26b-556">Example</span></span>

```JavaScript
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

####  <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="be26b-557">addItemAttachmentAsync (itemId，attachmentName，[选项] [回调])</span><span class="sxs-lookup"><span data-stu-id="be26b-557">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="be26b-558">将 Exchange 项目（如邮件）作为附件添加到邮件或约会。</span><span class="sxs-lookup"><span data-stu-id="be26b-558">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="be26b-p134">`addItemAttachmentAsync` 方法将包含指定 Exchange 标识符的项目附加到撰写窗体中的项目。如果指定一个回调方法，此方法使用 `asyncResult` 参数调用，该参数包含一个附件标识符或代码，指示附加项目过程中出现的任何错误。可以使用 `options` 参数将状态信息传递给回调方法（如果需要）。</span><span class="sxs-lookup"><span data-stu-id="be26b-p134">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="be26b-562">你随后可以将该标识符与 [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) 方法一同使用，以删除同一个会话中的附件。</span><span class="sxs-lookup"><span data-stu-id="be26b-562">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="be26b-563">如果 Office 外接程序在 Outlook Web App 中运行，则 `addItemAttachmentAsync` 方法可以将项目附加到项目（正在编辑的项目除外）中；然而，不支持也不建议这样做。</span><span class="sxs-lookup"><span data-stu-id="be26b-563">If your Office add-in is running in Outlook Web App, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="be26b-564">参数：</span><span class="sxs-lookup"><span data-stu-id="be26b-564">Parameters:</span></span>

|<span data-ttu-id="be26b-565">名称</span><span class="sxs-lookup"><span data-stu-id="be26b-565">Name</span></span>| <span data-ttu-id="be26b-566">类型</span><span class="sxs-lookup"><span data-stu-id="be26b-566">Type</span></span>| <span data-ttu-id="be26b-567">属性</span><span class="sxs-lookup"><span data-stu-id="be26b-567">Attributes</span></span>| <span data-ttu-id="be26b-568">说明</span><span class="sxs-lookup"><span data-stu-id="be26b-568">Description</span></span>|
|---|---|---|---|
|`itemId`| <span data-ttu-id="be26b-569">字符串</span><span class="sxs-lookup"><span data-stu-id="be26b-569">String</span></span>||<span data-ttu-id="be26b-p135">要附加项目的 Exchange 标识符。最大长度为 100 个字符。</span><span class="sxs-lookup"><span data-stu-id="be26b-p135">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="be26b-572">String</span><span class="sxs-lookup"><span data-stu-id="be26b-572">String</span></span>||<span data-ttu-id="be26b-p136">要附加项目的主题。最大长度为 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="be26b-p136">The sujbect of the item to be attached. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="be26b-575">Object</span><span class="sxs-lookup"><span data-stu-id="be26b-575">Object</span></span>| <span data-ttu-id="be26b-576">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="be26b-576">&lt;optional&gt;</span></span>|<span data-ttu-id="be26b-577">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="be26b-577">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="be26b-578">Object</span><span class="sxs-lookup"><span data-stu-id="be26b-578">Object</span></span>| <span data-ttu-id="be26b-579">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="be26b-579">&lt;optional&gt;</span></span>|<span data-ttu-id="be26b-580">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="be26b-580">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="be26b-581">函数</span><span class="sxs-lookup"><span data-stu-id="be26b-581">function</span></span>| <span data-ttu-id="be26b-582">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="be26b-582">&lt;optional&gt;</span></span>|<span data-ttu-id="be26b-583">方法完成后，使用单个参数 `asyncResult` （一个 [`AsyncResult`](/javascript/api/office/office.asyncresult)   对象）调用在 `callback`  参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="be26b-583">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="be26b-584">如果成功，附件标识符将在 `asyncResult.value` 属性中提供。</span><span class="sxs-lookup"><span data-stu-id="be26b-584">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="be26b-585">如果添加附件失败，`asyncResult` 对象将包含一个提供错误说明的 `Error` 对象。</span><span class="sxs-lookup"><span data-stu-id="be26b-585">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="be26b-586">错误</span><span class="sxs-lookup"><span data-stu-id="be26b-586">Errors</span></span>

| <span data-ttu-id="be26b-587">错误代码</span><span class="sxs-lookup"><span data-stu-id="be26b-587">Error code</span></span> | <span data-ttu-id="be26b-588">说明</span><span class="sxs-lookup"><span data-stu-id="be26b-588">Description</span></span> |
|------------|-------------|
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="be26b-589">邮件或约会具有的附件过多。</span><span class="sxs-lookup"><span data-stu-id="be26b-589">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="be26b-590">要求</span><span class="sxs-lookup"><span data-stu-id="be26b-590">Requirements</span></span>

|<span data-ttu-id="be26b-591">要求</span><span class="sxs-lookup"><span data-stu-id="be26b-591">Requirement</span></span>| <span data-ttu-id="be26b-592">值</span><span class="sxs-lookup"><span data-stu-id="be26b-592">Value</span></span>|
|---|---|
|[<span data-ttu-id="be26b-593">最低邮箱要求集版本</span><span class="sxs-lookup"><span data-stu-id="be26b-593">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="be26b-594">1.1</span><span class="sxs-lookup"><span data-stu-id="be26b-594">1.1</span></span>|
|[<span data-ttu-id="be26b-595">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="be26b-595">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="be26b-596">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="be26b-596">ReadWriteItem</span></span>|
|[<span data-ttu-id="be26b-597">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="be26b-597">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="be26b-598">撰写</span><span class="sxs-lookup"><span data-stu-id="be26b-598">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="be26b-599">示例</span><span class="sxs-lookup"><span data-stu-id="be26b-599">Example</span></span>

<span data-ttu-id="be26b-600">以下示例将现有 Outlook 项目添加为名为 `My Attachment` 的附件。</span><span class="sxs-lookup"><span data-stu-id="be26b-600">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

```JavaScript
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

#### <a name="displayreplyallformformdata"></a><span data-ttu-id="be26b-601">displayReplyAllForm(formData)</span><span class="sxs-lookup"><span data-stu-id="be26b-601">displayReplyAllForm(formData)</span></span>

<span data-ttu-id="be26b-602">显示答复窗体，其中包括所选邮件的发件人和所有收件人或所选约会的组织者和所有与会者。</span><span class="sxs-lookup"><span data-stu-id="be26b-602">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="be26b-603">在 Outlook for iOS 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="be26b-603">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="be26b-604">在 Outlook Web App 中，答复窗体显示为包含 3 列视图的弹出式窗体以及包含 2 列或 1 列视图的弹出式窗体。</span><span class="sxs-lookup"><span data-stu-id="be26b-604">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="be26b-605">如果任意字符串参数超出其限制， `displayReplyAllForm` 将引发异常。</span><span class="sxs-lookup"><span data-stu-id="be26b-605">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="be26b-p137">当在 `formData.attachments` 参数中指定附件时，Outlook 和 Outlook Web App 尝试下载所有附件并将其附加到答复窗体。如果无法添加任何附件，则在窗体 UI 中显示错误。如果这不可能，则不引发错误消息。</span><span class="sxs-lookup"><span data-stu-id="be26b-p137">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="be26b-609">参数：</span><span class="sxs-lookup"><span data-stu-id="be26b-609">Parameters:</span></span>

|<span data-ttu-id="be26b-610">名称</span><span class="sxs-lookup"><span data-stu-id="be26b-610">Name</span></span>| <span data-ttu-id="be26b-611">类型</span><span class="sxs-lookup"><span data-stu-id="be26b-611">Type</span></span>| <span data-ttu-id="be26b-612">说明</span><span class="sxs-lookup"><span data-stu-id="be26b-612">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="be26b-613">字符串 | 对象</span><span class="sxs-lookup"><span data-stu-id="be26b-613">String &#124; Object</span></span>| |<span data-ttu-id="be26b-p138">一个包含文本和 HTML 且代表答复窗体正文的字符串。字符串限制为 32 KB。</span><span class="sxs-lookup"><span data-stu-id="be26b-p138">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="be26b-616">**OR**</span><span class="sxs-lookup"><span data-stu-id="be26b-616">**OR**</span></span><br/><span data-ttu-id="be26b-p139">包含正文或附件数据和回调函数的对象。对象定义如下。</span><span class="sxs-lookup"><span data-stu-id="be26b-p139">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="be26b-619">字符串</span><span class="sxs-lookup"><span data-stu-id="be26b-619">String</span></span> | <span data-ttu-id="be26b-620">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="be26b-620">&lt;optional&gt;</span></span> | <span data-ttu-id="be26b-p140">一个包含文本和 HTML 且表示答复窗体正文的字符串。字符串限制为 32 KB。</span><span class="sxs-lookup"><span data-stu-id="be26b-p140">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="be26b-623">数组。&lt;对象&gt;</span><span class="sxs-lookup"><span data-stu-id="be26b-623">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="be26b-624">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="be26b-624">&lt;optional&gt;</span></span> | <span data-ttu-id="be26b-625">JSON 对象（是文件或项目附件）的数组。</span><span class="sxs-lookup"><span data-stu-id="be26b-625">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="be26b-626">字符串</span><span class="sxs-lookup"><span data-stu-id="be26b-626">String</span></span> | | <span data-ttu-id="be26b-p141">指示附件的类型。必须是文件附件的 `file` 或项目附件的 `item` 。</span><span class="sxs-lookup"><span data-stu-id="be26b-p141">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="be26b-629">字符串</span><span class="sxs-lookup"><span data-stu-id="be26b-629">String</span></span> | | <span data-ttu-id="be26b-630">一个包含附件的名称的字符串，最多包含 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="be26b-630">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="be26b-631">字符串</span><span class="sxs-lookup"><span data-stu-id="be26b-631">String</span></span> | | <span data-ttu-id="be26b-p142">仅在将 `type` 设置为 `file` 时使用。文件位置的 URI。</span><span class="sxs-lookup"><span data-stu-id="be26b-p142">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="be26b-634">字符串</span><span class="sxs-lookup"><span data-stu-id="be26b-634">String</span></span> | | <span data-ttu-id="be26b-p143">仅在将 `type` 设置为 `item` 时使用。附件的 EWS 项目 ID。字符串最长为 100 个字符。</span><span class="sxs-lookup"><span data-stu-id="be26b-p143">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="be26b-638">function</span><span class="sxs-lookup"><span data-stu-id="be26b-638">function</span></span> | <span data-ttu-id="be26b-639">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="be26b-639">&lt;optional&gt;</span></span> | <span data-ttu-id="be26b-640">方法完成后，使用单个参数 `asyncResult`（一个 [AsyncResult](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="be26b-640">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="be26b-641">要求</span><span class="sxs-lookup"><span data-stu-id="be26b-641">Requirements</span></span>

|<span data-ttu-id="be26b-642">要求</span><span class="sxs-lookup"><span data-stu-id="be26b-642">Requirement</span></span>| <span data-ttu-id="be26b-643">值</span><span class="sxs-lookup"><span data-stu-id="be26b-643">Value</span></span>|
|---|---|
|[<span data-ttu-id="be26b-644">最低邮箱要求集版本</span><span class="sxs-lookup"><span data-stu-id="be26b-644">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="be26b-645">1.0</span><span class="sxs-lookup"><span data-stu-id="be26b-645">1.0</span></span>|
|[<span data-ttu-id="be26b-646">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="be26b-646">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="be26b-647">ReadItem</span><span class="sxs-lookup"><span data-stu-id="be26b-647">ReadItem</span></span>|
|[<span data-ttu-id="be26b-648">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="be26b-648">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="be26b-649">阅读</span><span class="sxs-lookup"><span data-stu-id="be26b-649">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="be26b-650">示例</span><span class="sxs-lookup"><span data-stu-id="be26b-650">Examples</span></span>

<span data-ttu-id="be26b-651">以下代码将一个字符串传递到 `displayReplyAllForm` 函数。</span><span class="sxs-lookup"><span data-stu-id="be26b-651">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="be26b-652">使用空白正文答复。</span><span class="sxs-lookup"><span data-stu-id="be26b-652">Reply with an empty body.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="be26b-653">仅使用正文答复。</span><span class="sxs-lookup"><span data-stu-id="be26b-653">Reply with just a body.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="be26b-654">使用正文和文件附件答复。</span><span class="sxs-lookup"><span data-stu-id="be26b-654">Reply with a body and a file attachment.</span></span>

```JavaScript
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

<span data-ttu-id="be26b-655">使用正文和项目附件答复。</span><span class="sxs-lookup"><span data-stu-id="be26b-655">Reply with a body and an item attachment.</span></span>

```JavaScript
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

<span data-ttu-id="be26b-656">使用正文、文件附件、项目附件和回调答复。</span><span class="sxs-lookup"><span data-stu-id="be26b-656">Reply with a body, file attachment, item attachment, and a callback.</span></span>

```JavaScript
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

#### <a name="displayreplyformformdata"></a><span data-ttu-id="be26b-657">displayReplyForm(formData)</span><span class="sxs-lookup"><span data-stu-id="be26b-657">displayReplyForm(formData)</span></span>

<span data-ttu-id="be26b-658">显示答复窗体，其中仅包括所选邮件的发件人或所选约会的组织者。</span><span class="sxs-lookup"><span data-stu-id="be26b-658">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="be26b-659">在 Outlook for iOS 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="be26b-659">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="be26b-660">在 Outlook Web App 中，答复窗体显示为包含 3 列视图的弹出式窗体以及包含 2 列或 1 列视图的弹出式窗体。</span><span class="sxs-lookup"><span data-stu-id="be26b-660">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="be26b-661">如果任意字符串参数超出其限制， `displayReplyForm` 将引发异常。</span><span class="sxs-lookup"><span data-stu-id="be26b-661">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="be26b-p144">当在 `formData.attachments` 参数中指定附件时，Outlook 和 Outlook Web App 尝试下载所有附件并将其附加到答复窗体。如果无法添加任何附件，则在窗体 UI 中显示错误。如果这不可能，则不引发错误消息。</span><span class="sxs-lookup"><span data-stu-id="be26b-p144">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="be26b-665">参数：</span><span class="sxs-lookup"><span data-stu-id="be26b-665">Parameters:</span></span>

|<span data-ttu-id="be26b-666">名称</span><span class="sxs-lookup"><span data-stu-id="be26b-666">Name</span></span>| <span data-ttu-id="be26b-667">类型</span><span class="sxs-lookup"><span data-stu-id="be26b-667">Type</span></span>| <span data-ttu-id="be26b-668">说明</span><span class="sxs-lookup"><span data-stu-id="be26b-668">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="be26b-669">字符串 | 对象</span><span class="sxs-lookup"><span data-stu-id="be26b-669">String &#124; Object</span></span>| | <span data-ttu-id="be26b-p145">一个包含文本和 HTML 且代表答复窗体正文的字符串。字符串限制为 32 KB。</span><span class="sxs-lookup"><span data-stu-id="be26b-p145">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="be26b-672">**OR**</span><span class="sxs-lookup"><span data-stu-id="be26b-672">**OR**</span></span><br/><span data-ttu-id="be26b-p146">包含正文或附件数据和回调函数的对象。对象定义如下。</span><span class="sxs-lookup"><span data-stu-id="be26b-p146">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="be26b-675">字符串</span><span class="sxs-lookup"><span data-stu-id="be26b-675">String</span></span> | <span data-ttu-id="be26b-676">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="be26b-676">&lt;optional&gt;</span></span> | <span data-ttu-id="be26b-p147">一个包含文本和 HTML 且表示答复窗体正文的字符串。字符串限制为 32 KB。</span><span class="sxs-lookup"><span data-stu-id="be26b-p147">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="be26b-679">数组。&lt;对象&gt;</span><span class="sxs-lookup"><span data-stu-id="be26b-679">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="be26b-680">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="be26b-680">&lt;optional&gt;</span></span> | <span data-ttu-id="be26b-681">JSON 对象（是文件或项目附件）的数组。</span><span class="sxs-lookup"><span data-stu-id="be26b-681">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="be26b-682">字符串</span><span class="sxs-lookup"><span data-stu-id="be26b-682">String</span></span> | | <span data-ttu-id="be26b-p148">指示附件的类型。必须是文件附件的 `file` 或项目附件的 `item` 。</span><span class="sxs-lookup"><span data-stu-id="be26b-p148">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="be26b-685">字符串</span><span class="sxs-lookup"><span data-stu-id="be26b-685">String</span></span> | | <span data-ttu-id="be26b-686">一个包含附件的名称的字符串，最多包含 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="be26b-686">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="be26b-687">字符串</span><span class="sxs-lookup"><span data-stu-id="be26b-687">String</span></span> | | <span data-ttu-id="be26b-p149">仅在将 `type` 设置为 `file` 时使用。文件位置的 URI。</span><span class="sxs-lookup"><span data-stu-id="be26b-p149">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="be26b-690">字符串</span><span class="sxs-lookup"><span data-stu-id="be26b-690">String</span></span> | | <span data-ttu-id="be26b-p150">仅在将 `type` 设置为 `item` 时使用。附件的 EWS 项目 ID。字符串最长为 100 个字符。</span><span class="sxs-lookup"><span data-stu-id="be26b-p150">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="be26b-694">function</span><span class="sxs-lookup"><span data-stu-id="be26b-694">function</span></span> | <span data-ttu-id="be26b-695">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="be26b-695">&lt;optional&gt;</span></span> | <span data-ttu-id="be26b-696">方法完成后，使用单个参数 `asyncResult`（一个 [AsyncResult](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="be26b-696">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="be26b-697">要求</span><span class="sxs-lookup"><span data-stu-id="be26b-697">Requirements</span></span>

|<span data-ttu-id="be26b-698">要求</span><span class="sxs-lookup"><span data-stu-id="be26b-698">Requirement</span></span>| <span data-ttu-id="be26b-699">值</span><span class="sxs-lookup"><span data-stu-id="be26b-699">Value</span></span>|
|---|---|
|[<span data-ttu-id="be26b-700">最低邮箱要求集版本</span><span class="sxs-lookup"><span data-stu-id="be26b-700">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="be26b-701">1.0</span><span class="sxs-lookup"><span data-stu-id="be26b-701">1.0</span></span>|
|[<span data-ttu-id="be26b-702">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="be26b-702">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="be26b-703">ReadItem</span><span class="sxs-lookup"><span data-stu-id="be26b-703">ReadItem</span></span>|
|[<span data-ttu-id="be26b-704">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="be26b-704">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="be26b-705">阅读</span><span class="sxs-lookup"><span data-stu-id="be26b-705">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="be26b-706">示例</span><span class="sxs-lookup"><span data-stu-id="be26b-706">Examples</span></span>

<span data-ttu-id="be26b-707">以下代码将一个字符串传递到 `displayReplyForm` 函数。</span><span class="sxs-lookup"><span data-stu-id="be26b-707">The following code passes a string to the `displayReplyForm` function.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="be26b-708">使用空白正文答复。</span><span class="sxs-lookup"><span data-stu-id="be26b-708">Reply with an empty body.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="be26b-709">仅使用正文答复。</span><span class="sxs-lookup"><span data-stu-id="be26b-709">Reply with just a body.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="be26b-710">使用正文和文件附件答复。</span><span class="sxs-lookup"><span data-stu-id="be26b-710">Reply with a body and a file attachment.</span></span>

```JavaScript
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

<span data-ttu-id="be26b-711">使用正文和项目附件答复。</span><span class="sxs-lookup"><span data-stu-id="be26b-711">Reply with a body and an item attachment.</span></span>

```JavaScript
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

<span data-ttu-id="be26b-712">使用正文、文件附件、项目附件和回调答复。</span><span class="sxs-lookup"><span data-stu-id="be26b-712">Reply with a body, file attachment, item attachment, and a callback.</span></span>

```JavaScript
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

#### <a name="getentities--entitiesjavascriptapioutlook12officeentities"></a><span data-ttu-id="be26b-713">getEntities() → {[实体](/javascript/api/outlook_1_2/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="be26b-713">getEntities() → {[Entities](/javascript/api/outlook_1_2/office.entities)}</span></span>

<span data-ttu-id="be26b-714">获取在所选项正文中找到的实体。</span><span class="sxs-lookup"><span data-stu-id="be26b-714">Gets the entities found in the selected item.</span></span>

> [!NOTE]
> <span data-ttu-id="be26b-715">在 Outlook for iOS 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="be26b-715">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="be26b-716">要求</span><span class="sxs-lookup"><span data-stu-id="be26b-716">Requirements</span></span>

|<span data-ttu-id="be26b-717">要求</span><span class="sxs-lookup"><span data-stu-id="be26b-717">Requirement</span></span>| <span data-ttu-id="be26b-718">值</span><span class="sxs-lookup"><span data-stu-id="be26b-718">Value</span></span>|
|---|---|
|[<span data-ttu-id="be26b-719">最低邮箱要求集版本</span><span class="sxs-lookup"><span data-stu-id="be26b-719">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="be26b-720">1.0</span><span class="sxs-lookup"><span data-stu-id="be26b-720">1.0</span></span>|
|[<span data-ttu-id="be26b-721">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="be26b-721">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="be26b-722">ReadItem</span><span class="sxs-lookup"><span data-stu-id="be26b-722">ReadItem</span></span>|
|[<span data-ttu-id="be26b-723">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="be26b-723">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="be26b-724">阅读</span><span class="sxs-lookup"><span data-stu-id="be26b-724">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="be26b-725">返回：</span><span class="sxs-lookup"><span data-stu-id="be26b-725">Returns:</span></span>

<span data-ttu-id="be26b-726">类型：[Entities](/javascript/api/outlook_1_2/office.entities)</span><span class="sxs-lookup"><span data-stu-id="be26b-726">Type: [Entities](/javascript/api/outlook_1_2/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="be26b-727">示例</span><span class="sxs-lookup"><span data-stu-id="be26b-727">Example</span></span>

<span data-ttu-id="be26b-728">以下示例访问当前项正文中的联系人实体。</span><span class="sxs-lookup"><span data-stu-id="be26b-728">The following example accesses the contacts entities on the current item.</span></span>

```
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlook12officecontactmeetingsuggestionjavascriptapioutlook12officemeetingsuggestionphonenumberjavascriptapioutlook12officephonenumbertasksuggestionjavascriptapioutlook12officetasksuggestion"></a><span data-ttu-id="be26b-729">getEntitiesByType(entityType) → (nullable)  {数组。 <(String|[联系人](/javascript/api/outlook_1_2/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_2/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_2/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_2/office.tasksuggestion)) >}</span><span class="sxs-lookup"><span data-stu-id="be26b-729">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_2/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_2/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_2/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_2/office.tasksuggestion))>}</span></span>

<span data-ttu-id="be26b-730">获取所选项目正文中找到的指定实体类型的所有实体的数组。</span><span class="sxs-lookup"><span data-stu-id="be26b-730">Gets an array of all the entities of the specified entity type found in the selected item.</span></span>

> [!NOTE]
> <span data-ttu-id="be26b-731">在 Outlook for iOS 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="be26b-731">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="be26b-732">参数：</span><span class="sxs-lookup"><span data-stu-id="be26b-732">Parameters:</span></span>

|<span data-ttu-id="be26b-733">名称</span><span class="sxs-lookup"><span data-stu-id="be26b-733">Name</span></span>| <span data-ttu-id="be26b-734">类型</span><span class="sxs-lookup"><span data-stu-id="be26b-734">Type</span></span>| <span data-ttu-id="be26b-735">说明</span><span class="sxs-lookup"><span data-stu-id="be26b-735">Description</span></span>|
|---|---|---|
|`entityType`| [<span data-ttu-id="be26b-736">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="be26b-736">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook_1_2/office.mailboxenums.entitytype)|<span data-ttu-id="be26b-737">EntityType 枚举值之一。</span><span class="sxs-lookup"><span data-stu-id="be26b-737">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="be26b-738">要求</span><span class="sxs-lookup"><span data-stu-id="be26b-738">Requirements</span></span>

|<span data-ttu-id="be26b-739">要求</span><span class="sxs-lookup"><span data-stu-id="be26b-739">Requirement</span></span>| <span data-ttu-id="be26b-740">值</span><span class="sxs-lookup"><span data-stu-id="be26b-740">Value</span></span>|
|---|---|
|[<span data-ttu-id="be26b-741">最低邮箱要求集版本</span><span class="sxs-lookup"><span data-stu-id="be26b-741">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="be26b-742">1.0</span><span class="sxs-lookup"><span data-stu-id="be26b-742">1.0</span></span>|
|[<span data-ttu-id="be26b-743">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="be26b-743">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="be26b-744">Restricted</span><span class="sxs-lookup"><span data-stu-id="be26b-744">Restricted</span></span>|
|[<span data-ttu-id="be26b-745">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="be26b-745">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="be26b-746">阅读</span><span class="sxs-lookup"><span data-stu-id="be26b-746">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="be26b-747">返回：</span><span class="sxs-lookup"><span data-stu-id="be26b-747">Returns:</span></span>

<span data-ttu-id="be26b-748">如果在 `entityType` 中传递的值不是 `EntityType` 枚举的有效成员，该方法返回 null。</span><span class="sxs-lookup"><span data-stu-id="be26b-748">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="be26b-749">如果指定类型的任何实体都不存在于该项目上，该方法将返回空数组。</span><span class="sxs-lookup"><span data-stu-id="be26b-749">If no entities of the specified type are present on the item, the method returns an empty array.</span></span> <span data-ttu-id="be26b-750">否则，返回数组中的对象类型取决于 `entityType` 参数中请求实体的类型。</span><span class="sxs-lookup"><span data-stu-id="be26b-750">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="be26b-751">当使用此方法的最低权限级别**受限**时，某些实体类型需要 **ReadItem** 才能进行访问，如下表中所指定。</span><span class="sxs-lookup"><span data-stu-id="be26b-751">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

| <span data-ttu-id="be26b-752">值对应于 `entityType`</span><span class="sxs-lookup"><span data-stu-id="be26b-752">Value of `entityType`</span></span> | <span data-ttu-id="be26b-753">返回的数组中对象的类型</span><span class="sxs-lookup"><span data-stu-id="be26b-753">Type of objects in returned array</span></span> | <span data-ttu-id="be26b-754">所需权限级别</span><span class="sxs-lookup"><span data-stu-id="be26b-754">Required Permission Level</span></span> |
| --- | --- | --- |
| `Address` | <span data-ttu-id="be26b-755">字符串</span><span class="sxs-lookup"><span data-stu-id="be26b-755">String</span></span> | <span data-ttu-id="be26b-756">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="be26b-756">**Restricted**</span></span> |
| `Contact` | <span data-ttu-id="be26b-757">联系人</span><span class="sxs-lookup"><span data-stu-id="be26b-757">Contact</span></span> | <span data-ttu-id="be26b-758">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="be26b-758">**ReadItem**</span></span> |
| `EmailAddress` | <span data-ttu-id="be26b-759">字符串</span><span class="sxs-lookup"><span data-stu-id="be26b-759">String</span></span> | <span data-ttu-id="be26b-760">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="be26b-760">**ReadItem**</span></span> |
| `MeetingSuggestion` | <span data-ttu-id="be26b-761">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="be26b-761">MeetingSuggestion</span></span> | <span data-ttu-id="be26b-762">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="be26b-762">**ReadItem**</span></span> |
| `PhoneNumber` | <span data-ttu-id="be26b-763">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="be26b-763">PhoneNumber</span></span> | <span data-ttu-id="be26b-764">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="be26b-764">**Restricted**</span></span> |
| `TaskSuggestion` | <span data-ttu-id="be26b-765">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="be26b-765">TaskSuggestion</span></span> | <span data-ttu-id="be26b-766">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="be26b-766">**ReadItem**</span></span> |
| `URL` | <span data-ttu-id="be26b-767">字符串</span><span class="sxs-lookup"><span data-stu-id="be26b-767">String</span></span> | <span data-ttu-id="be26b-768">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="be26b-768">**Restricted**</span></span> |

<span data-ttu-id="be26b-769">类型：数组.<(字符串|[Contact](/javascript/api/outlook_1_2/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_2/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_2/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_2/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="be26b-769">Type: Array.<(String|[Contact](/javascript/api/outlook_1_2/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_2/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_2/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_2/office.tasksuggestion))></span></span>

##### <a name="example"></a><span data-ttu-id="be26b-770">示例</span><span class="sxs-lookup"><span data-stu-id="be26b-770">Example</span></span>

<span data-ttu-id="be26b-771">以下示例显示了如何访问代表当前项正文中邮政地址的字符串数组。</span><span class="sxs-lookup"><span data-stu-id="be26b-771">The following example shows how to access an array of strings that represent postal addresses in the subject or body of the current item.</span></span>

```JavaScript
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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlook12officecontactmeetingsuggestionjavascriptapioutlook12officemeetingsuggestionphonenumberjavascriptapioutlook12officephonenumbertasksuggestionjavascriptapioutlook12officetasksuggestion"></a><span data-ttu-id="be26b-772">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_2/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_2/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_2/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_2/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="be26b-772">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_2/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_2/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_2/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_2/office.tasksuggestion))>}</span></span>

<span data-ttu-id="be26b-773">返回传递清单 XML 文件中定义的命名筛选器所选项中的已知实体。</span><span class="sxs-lookup"><span data-stu-id="be26b-773">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="be26b-774">在 Outlook for iOS 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="be26b-774">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="be26b-775">`getFilteredEntitiesByName` 方法返回与具有指定 `FilterName` 元素值的清单 XML 文件中的 [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) 规则元素中定义的规则表达式相匹配的实体。</span><span class="sxs-lookup"><span data-stu-id="be26b-775">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="be26b-776">参数：</span><span class="sxs-lookup"><span data-stu-id="be26b-776">Parameters:</span></span>

|<span data-ttu-id="be26b-777">名称</span><span class="sxs-lookup"><span data-stu-id="be26b-777">Name</span></span>| <span data-ttu-id="be26b-778">类型</span><span class="sxs-lookup"><span data-stu-id="be26b-778">Type</span></span>| <span data-ttu-id="be26b-779">说明</span><span class="sxs-lookup"><span data-stu-id="be26b-779">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="be26b-780">字符串</span><span class="sxs-lookup"><span data-stu-id="be26b-780">String</span></span>|<span data-ttu-id="be26b-781">定义筛选器匹配的 `ItemHasKnownEntity` 规则元素的名称。</span><span class="sxs-lookup"><span data-stu-id="be26b-781">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="be26b-782">要求</span><span class="sxs-lookup"><span data-stu-id="be26b-782">Requirements</span></span>

|<span data-ttu-id="be26b-783">要求</span><span class="sxs-lookup"><span data-stu-id="be26b-783">Requirement</span></span>| <span data-ttu-id="be26b-784">值</span><span class="sxs-lookup"><span data-stu-id="be26b-784">Value</span></span>|
|---|---|
|[<span data-ttu-id="be26b-785">最低邮箱要求集版本</span><span class="sxs-lookup"><span data-stu-id="be26b-785">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="be26b-786">1.0</span><span class="sxs-lookup"><span data-stu-id="be26b-786">1.0</span></span>|
|[<span data-ttu-id="be26b-787">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="be26b-787">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="be26b-788">ReadItem</span><span class="sxs-lookup"><span data-stu-id="be26b-788">ReadItem</span></span>|
|[<span data-ttu-id="be26b-789">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="be26b-789">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="be26b-790">阅读</span><span class="sxs-lookup"><span data-stu-id="be26b-790">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="be26b-791">返回：</span><span class="sxs-lookup"><span data-stu-id="be26b-791">Returns:</span></span>

<span data-ttu-id="be26b-p152">如果清单中 `ItemHasKnownEntity`  元素没有匹配 `FilterName` 参数的 `name` 元素值，则该方法返回 `null`。如果 `name` 参数匹配清单中的 `ItemHasKnownEntity` 元素，但在当前匹配的项目中没有实体，则该方法返回一个空数组。</span><span class="sxs-lookup"><span data-stu-id="be26b-p152">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="be26b-794">类型：Array.<(String|[Contact](/javascript/api/outlook_1_2/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_2/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_2/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_2/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="be26b-794">Type: Array.<(String|[Contact](/javascript/api/outlook_1_2/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_2/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_2/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_2/office.tasksuggestion))></span></span>

#### <a name="getregexmatches--object"></a><span data-ttu-id="be26b-795">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="be26b-795">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="be26b-796">返回所选项目中与在清单 XML 文件中定义的正则表达式相匹配的字符串值。</span><span class="sxs-lookup"><span data-stu-id="be26b-796">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="be26b-797">在 Outlook for iOS 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="be26b-797">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="be26b-p153">`getRegExMatches` 方法返回匹配在清单 XML 文件中的每个 `ItemHasRegularExpressionMatch` 或 `ItemHasKnownEntity` 规则元素中定义的正则表达式的字符串。对于 `ItemHasRegularExpressionMatch` 规则，匹配字符串必须发生在该规则指定的项目的属性中。`PropertyName` 简单类型定义支持的属性。</span><span class="sxs-lookup"><span data-stu-id="be26b-p153">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="be26b-801">例如，考虑一个加载项具有以下 `Rule` 元素：</span><span class="sxs-lookup"><span data-stu-id="be26b-801">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```JavaScript
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="be26b-802">从 `getRegExMatches` 返回的对象应有两个属性：`fruits` 和 `veggies`。</span><span class="sxs-lookup"><span data-stu-id="be26b-802">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```JavaScript
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

> [!NOTE]
> <span data-ttu-id="be26b-p154">如果在项目正文属性上指定 `ItemHasRegularExpressionMatch` 规则，则正则表达式应进一步筛选正文，不应尝试返回该项目的整个正文。使用正则表达式（如 `.*`）获取项目的整个正文并不总是返回预期的结果。</span><span class="sxs-lookup"><span data-stu-id="be26b-p154">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="requirements"></a><span data-ttu-id="be26b-805">要求</span><span class="sxs-lookup"><span data-stu-id="be26b-805">Requirements</span></span>

|<span data-ttu-id="be26b-806">要求</span><span class="sxs-lookup"><span data-stu-id="be26b-806">Requirement</span></span>| <span data-ttu-id="be26b-807">值</span><span class="sxs-lookup"><span data-stu-id="be26b-807">Value</span></span>|
|---|---|
|[<span data-ttu-id="be26b-808">最低邮箱要求集版本</span><span class="sxs-lookup"><span data-stu-id="be26b-808">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="be26b-809">1.0</span><span class="sxs-lookup"><span data-stu-id="be26b-809">1.0</span></span>|
|[<span data-ttu-id="be26b-810">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="be26b-810">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="be26b-811">ReadItem</span><span class="sxs-lookup"><span data-stu-id="be26b-811">ReadItem</span></span>|
|[<span data-ttu-id="be26b-812">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="be26b-812">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="be26b-813">阅读</span><span class="sxs-lookup"><span data-stu-id="be26b-813">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="be26b-814">返回：</span><span class="sxs-lookup"><span data-stu-id="be26b-814">Returns:</span></span>

<span data-ttu-id="be26b-p155">一个包含与在清单 XML 文件中定义的正则表达式的字符串数组相匹配的对象。每个数组的名称等于匹配 `ItemHasRegularExpressionMatch` 规则的 `RegExName` 属性或匹配 `ItemHasKnownEntity` 规则的 `FilterName` 属性的相应值。</span><span class="sxs-lookup"><span data-stu-id="be26b-p155">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<dl class="param-type"><span data-ttu-id="be26b-817">

<dt>类型</dt>

</span><span class="sxs-lookup"><span data-stu-id="be26b-817">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="be26b-818">Object</span><span class="sxs-lookup"><span data-stu-id="be26b-818">Object</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="be26b-819">示例</span><span class="sxs-lookup"><span data-stu-id="be26b-819">Example</span></span>

<span data-ttu-id="be26b-820">以下示例显示了如何访问正则表达式 <rule>元素 `fruits` 和 `veggies` 的匹配项的数组，这些元素在清单中指定。</rule></span><span class="sxs-lookup"><span data-stu-id="be26b-820">The following example shows how to access the array of matches for the regular expression <rule>elements `fruits` and `veggies`, which are specified in the manifest.</rule></span></span>

```JavaScript
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veges = allMatches.veggies;
```

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="be26b-821">getRegExMatchesByName(name) → (可空类型) {数组.< 字符串 >}</span><span class="sxs-lookup"><span data-stu-id="be26b-821">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="be26b-822">返回匹配在清单 XML 文件中定义的命名正则表达式的所选项目中的字符串值。</span><span class="sxs-lookup"><span data-stu-id="be26b-822">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="be26b-823">在 Outlook for iOS 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="be26b-823">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="be26b-824">`getRegExMatchesByName` 方法返回匹配在具有指定 `RegExName` 元素值的清单 XML 文件中的 `ItemHasRegularExpressionMatch` 规则元素中定义的正则表达式的字符串。</span><span class="sxs-lookup"><span data-stu-id="be26b-824">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="be26b-p156">如果在项目正文属性上指定 `ItemHasRegularExpressionMatch` 规则，则正则表达式应进一步筛选正文，不应尝试返回该项目的整个正文。使用正则表达式（如 `.*`）获取项目的整个正文并不总是返回预期的结果。</span><span class="sxs-lookup"><span data-stu-id="be26b-p156">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="be26b-827">参数：</span><span class="sxs-lookup"><span data-stu-id="be26b-827">Parameters:</span></span>

|<span data-ttu-id="be26b-828">名称</span><span class="sxs-lookup"><span data-stu-id="be26b-828">Name</span></span>| <span data-ttu-id="be26b-829">类型</span><span class="sxs-lookup"><span data-stu-id="be26b-829">Type</span></span>| <span data-ttu-id="be26b-830">说明</span><span class="sxs-lookup"><span data-stu-id="be26b-830">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="be26b-831">字符串</span><span class="sxs-lookup"><span data-stu-id="be26b-831">String</span></span>|<span data-ttu-id="be26b-832">定义筛选器匹配的 `ItemHasRegularExpressionMatch` 规则元素的名称。</span><span class="sxs-lookup"><span data-stu-id="be26b-832">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="be26b-833">要求</span><span class="sxs-lookup"><span data-stu-id="be26b-833">Requirements</span></span>

|<span data-ttu-id="be26b-834">要求</span><span class="sxs-lookup"><span data-stu-id="be26b-834">Requirement</span></span>| <span data-ttu-id="be26b-835">值</span><span class="sxs-lookup"><span data-stu-id="be26b-835">Value</span></span>|
|---|---|
|[<span data-ttu-id="be26b-836">最低邮箱要求集版本</span><span class="sxs-lookup"><span data-stu-id="be26b-836">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="be26b-837">1.0</span><span class="sxs-lookup"><span data-stu-id="be26b-837">1.0</span></span>|
|[<span data-ttu-id="be26b-838">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="be26b-838">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="be26b-839">ReadItem</span><span class="sxs-lookup"><span data-stu-id="be26b-839">ReadItem</span></span>|
|[<span data-ttu-id="be26b-840">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="be26b-840">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="be26b-841">阅读</span><span class="sxs-lookup"><span data-stu-id="be26b-841">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="be26b-842">返回：</span><span class="sxs-lookup"><span data-stu-id="be26b-842">Returns:</span></span>

<span data-ttu-id="be26b-843">一个包含与在清单 XML 文件中定义的正则表达式的字符串相匹配的数组。</span><span class="sxs-lookup"><span data-stu-id="be26b-843">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<dl class="param-type"><span data-ttu-id="be26b-844">

<dt>类型</dt>

</span><span class="sxs-lookup"><span data-stu-id="be26b-844">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="be26b-845">数组。 < 字符串 ></span><span class="sxs-lookup"><span data-stu-id="be26b-845">Array.< String ></span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="be26b-846">示例</span><span class="sxs-lookup"><span data-stu-id="be26b-846">Example</span></span>

```JavaScript
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

####  <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="be26b-847">getSelectedDataAsync (coercionType，[选项] 回调) → {字符串}</span><span class="sxs-lookup"><span data-stu-id="be26b-847">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="be26b-848">以异步方式返回邮件的主题或正文中选定的数据。</span><span class="sxs-lookup"><span data-stu-id="be26b-848">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="be26b-p157">如果没有选定内容，但光标位于正文或主题中，此方法将会为所选数据返回 null。如果选定的是字段，而不是正文或主题，则此方法返回 `InvalidSelection` 错误。</span><span class="sxs-lookup"><span data-stu-id="be26b-p157">If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="be26b-851">参数：</span><span class="sxs-lookup"><span data-stu-id="be26b-851">Parameters:</span></span>

|<span data-ttu-id="be26b-852">名称</span><span class="sxs-lookup"><span data-stu-id="be26b-852">Name</span></span>| <span data-ttu-id="be26b-853">类型</span><span class="sxs-lookup"><span data-stu-id="be26b-853">Type</span></span>| <span data-ttu-id="be26b-854">属性</span><span class="sxs-lookup"><span data-stu-id="be26b-854">Attributes</span></span>| <span data-ttu-id="be26b-855">说明</span><span class="sxs-lookup"><span data-stu-id="be26b-855">Description</span></span>|
|---|---|---|---|
|`coercionType`| [<span data-ttu-id="be26b-856">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="be26b-856">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="be26b-p158">请求数据的格式。如果为 Text，则此方法返回纯文本作为字符串，删除任何显示的 HTML 标记。如果为 HTML，则此方法返回所选文本，不论是纯文本还是 HTML。</span><span class="sxs-lookup"><span data-stu-id="be26b-p158">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`| <span data-ttu-id="be26b-860">Object</span><span class="sxs-lookup"><span data-stu-id="be26b-860">Object</span></span>| <span data-ttu-id="be26b-861">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="be26b-861">&lt;optional&gt;</span></span>|<span data-ttu-id="be26b-862">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="be26b-862">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="be26b-863">Object</span><span class="sxs-lookup"><span data-stu-id="be26b-863">Object</span></span>| <span data-ttu-id="be26b-864">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="be26b-864">&lt;optional&gt;</span></span>|<span data-ttu-id="be26b-865">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="be26b-865">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="be26b-866">函数</span><span class="sxs-lookup"><span data-stu-id="be26b-866">function</span></span>||<span data-ttu-id="be26b-867">方法完成后，使用单个参数 `callback`（一个 [`asyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `AsyncResult` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="be26b-867">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="be26b-868">若要从回调方法访问所选数据，请调用 `asyncResult.value.data`。</span><span class="sxs-lookup"><span data-stu-id="be26b-868">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="be26b-869">若要访问选定内容的源属性，请调用 `asyncResult.value.sourceProperty`，这将为 `body` 或 `subject`。</span><span class="sxs-lookup"><span data-stu-id="be26b-869">To access the source property that the selection comes from, call , which will be either  or .|</span></span>|

##### <a name="requirements"></a><span data-ttu-id="be26b-870">要求</span><span class="sxs-lookup"><span data-stu-id="be26b-870">Requirements</span></span>

|<span data-ttu-id="be26b-871">要求</span><span class="sxs-lookup"><span data-stu-id="be26b-871">Requirement</span></span>| <span data-ttu-id="be26b-872">值</span><span class="sxs-lookup"><span data-stu-id="be26b-872">Value</span></span>|
|---|---|
|[<span data-ttu-id="be26b-873">最低邮箱要求集版本</span><span class="sxs-lookup"><span data-stu-id="be26b-873">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="be26b-874">1.2</span><span class="sxs-lookup"><span data-stu-id="be26b-874">1.2</span></span>|
|[<span data-ttu-id="be26b-875">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="be26b-875">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="be26b-876">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="be26b-876">ReadWriteItem</span></span>|
|[<span data-ttu-id="be26b-877">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="be26b-877">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="be26b-878">撰写</span><span class="sxs-lookup"><span data-stu-id="be26b-878">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="be26b-879">返回：</span><span class="sxs-lookup"><span data-stu-id="be26b-879">Returns:</span></span>

<span data-ttu-id="be26b-880">作为字符串的所选数据的格式由 `coercionType` 确定。</span><span class="sxs-lookup"><span data-stu-id="be26b-880">The selected data as a string with format determined by `coercionType`.</span></span>

<dl class="param-type"><span data-ttu-id="be26b-881">

<dt>类型</dt>

</span><span class="sxs-lookup"><span data-stu-id="be26b-881">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="be26b-882">字符串</span><span class="sxs-lookup"><span data-stu-id="be26b-882">String</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="be26b-883">示例</span><span class="sxs-lookup"><span data-stu-id="be26b-883">Example</span></span>

```JavaScript
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

####  <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="be26b-884">loadCustomPropertiesAsync （回调、 [userContext]）</span><span class="sxs-lookup"><span data-stu-id="be26b-884">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="be26b-885">为所选项目的加载项异步加载自定义属性。</span><span class="sxs-lookup"><span data-stu-id="be26b-885">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="be26b-p160">自定义属性基于每个应用、每个项目存储为键/值对。此方法在回调中返回 `CustomProperties` 对象，该回调提供访问特定于当前项目和当前外接程序的自定义属性的方法。自定义属性未在项目上加密，因此这不应用作安全存储。</span><span class="sxs-lookup"><span data-stu-id="be26b-p160">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="be26b-889">参数：</span><span class="sxs-lookup"><span data-stu-id="be26b-889">Parameters:</span></span>

|<span data-ttu-id="be26b-890">名称</span><span class="sxs-lookup"><span data-stu-id="be26b-890">Name</span></span>| <span data-ttu-id="be26b-891">类型</span><span class="sxs-lookup"><span data-stu-id="be26b-891">Type</span></span>| <span data-ttu-id="be26b-892">属性</span><span class="sxs-lookup"><span data-stu-id="be26b-892">Attributes</span></span>| <span data-ttu-id="be26b-893">说明</span><span class="sxs-lookup"><span data-stu-id="be26b-893">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="be26b-894">函数</span><span class="sxs-lookup"><span data-stu-id="be26b-894">function</span></span>||<span data-ttu-id="be26b-895">此方法完成时，用单个参数调用 `callback` 参数中传递的函数，`asyncResult`，这是一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象。</span><span class="sxs-lookup"><span data-stu-id="be26b-895">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="be26b-896">自定义属性作为 `asyncResult.value` 属性中的 [`CustomProperties`](/javascript/api/outlook_1_2/office.customproperties) 对象来提供。</span><span class="sxs-lookup"><span data-stu-id="be26b-896">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook_1_2/office.customproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="be26b-897">该对象可用于获取、设置和删除项目中的自定义属性，并将针对自定义属性集的更改保存回服务器。</span><span class="sxs-lookup"><span data-stu-id="be26b-897">The custom properties are provided as a CustomProperties object in the asyncResult.value property. This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`| <span data-ttu-id="be26b-898">Object</span><span class="sxs-lookup"><span data-stu-id="be26b-898">Object</span></span>| <span data-ttu-id="be26b-899">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="be26b-899">&lt;optional&gt;</span></span>|<span data-ttu-id="be26b-900">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="be26b-900">Developers can provide any object they wish to access in the callback method.</span></span> <span data-ttu-id="be26b-901">可以通过回调函数的 `asyncResult.asyncContext` 属性访问该对象。</span><span class="sxs-lookup"><span data-stu-id="be26b-901">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.|</span></span>|

##### <a name="requirements"></a><span data-ttu-id="be26b-902">要求</span><span class="sxs-lookup"><span data-stu-id="be26b-902">Requirements</span></span>

|<span data-ttu-id="be26b-903">要求</span><span class="sxs-lookup"><span data-stu-id="be26b-903">Requirement</span></span>| <span data-ttu-id="be26b-904">值</span><span class="sxs-lookup"><span data-stu-id="be26b-904">Value</span></span>|
|---|---|
|[<span data-ttu-id="be26b-905">最低邮箱要求集版本</span><span class="sxs-lookup"><span data-stu-id="be26b-905">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="be26b-906">1.0</span><span class="sxs-lookup"><span data-stu-id="be26b-906">1.0</span></span>|
|[<span data-ttu-id="be26b-907">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="be26b-907">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="be26b-908">ReadItem</span><span class="sxs-lookup"><span data-stu-id="be26b-908">ReadItem</span></span>|
|[<span data-ttu-id="be26b-909">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="be26b-909">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="be26b-910">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="be26b-910">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="be26b-911">示例</span><span class="sxs-lookup"><span data-stu-id="be26b-911">Example</span></span>

<span data-ttu-id="be26b-p163">以下代码示例显示了如何使用 `loadCustomPropertiesAsync` 方法异步加载特定于当前项目的自定义属性。该示例还显示了如何使用 `CustomProperties.saveAsync` 方法将这些属性重新保存到服务器。加载自定义属性后，该代码示例将使用 `CustomProperties.get` 方法读取自定义属性 `myProp`，使用 `CustomProperties.set` 方法写入自定义属性 `otherProp`，最后调用 `saveAsync` 方法保存这些自定义属性。</span><span class="sxs-lookup"><span data-stu-id="be26b-p163">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

```JavaScript
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

####  <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="be26b-915">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="be26b-915">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="be26b-916">将附件从邮件或约会中删除。</span><span class="sxs-lookup"><span data-stu-id="be26b-916">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="be26b-p164">`removeAttachmentAsync` 方法删除项目中带指定标识符的附件。最佳做法是，仅当同一个邮件应用程序在同一会话中添加了一个附件时，你才应使用该附件标识符来删除该附件。在 Outlook Web App 和适用于设备的 OWA 中，附件标识符只在同一个会话中才有效。当用户关闭应用程序，或者如果用户开始在内嵌窗体中撰写，并在随后弹出的内嵌窗体中继续在单独的窗口撰写时，会话即结束。</span><span class="sxs-lookup"><span data-stu-id="be26b-p164">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item. As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session. In Outlook Web App and OWA for Devices, the attachment identifier is valid only within the same session. A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="be26b-921">参数：</span><span class="sxs-lookup"><span data-stu-id="be26b-921">Parameters:</span></span>

|<span data-ttu-id="be26b-922">名称</span><span class="sxs-lookup"><span data-stu-id="be26b-922">Name</span></span>| <span data-ttu-id="be26b-923">类型</span><span class="sxs-lookup"><span data-stu-id="be26b-923">Type</span></span>| <span data-ttu-id="be26b-924">属性</span><span class="sxs-lookup"><span data-stu-id="be26b-924">Attributes</span></span>| <span data-ttu-id="be26b-925">说明</span><span class="sxs-lookup"><span data-stu-id="be26b-925">Description</span></span>|
|---|---|---|---|
|`attachmentId`| <span data-ttu-id="be26b-926">字符串</span><span class="sxs-lookup"><span data-stu-id="be26b-926">String</span></span>||<span data-ttu-id="be26b-p165">要删除的附件的标识符。字符串的最大长度为 100 个字符。</span><span class="sxs-lookup"><span data-stu-id="be26b-p165">The identifier of the attachment to remove. The maximum length of the string is 100 characters.</span></span>|
|`options`| <span data-ttu-id="be26b-929">Object</span><span class="sxs-lookup"><span data-stu-id="be26b-929">Object</span></span>| <span data-ttu-id="be26b-930">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="be26b-930">&lt;optional&gt;</span></span>|<span data-ttu-id="be26b-931">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="be26b-931">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="be26b-932">Object</span><span class="sxs-lookup"><span data-stu-id="be26b-932">Object</span></span>| <span data-ttu-id="be26b-933">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="be26b-933">&lt;optional&gt;</span></span>|<span data-ttu-id="be26b-934">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="be26b-934">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="be26b-935">function</span><span class="sxs-lookup"><span data-stu-id="be26b-935">function</span></span>| <span data-ttu-id="be26b-936">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="be26b-936">&lt;optional&gt;</span></span>|<span data-ttu-id="be26b-937">此方法完成时，用单个参数调用 `callback` 参数中传递的函数， `asyncResult` ，这是一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象。</span><span class="sxs-lookup"><span data-stu-id="be26b-937">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="be26b-938">如果删除附件失败，`asyncResult.error` 属性将包含一个说明失败原因的错误代码。</span><span class="sxs-lookup"><span data-stu-id="be26b-938">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="be26b-939">错误</span><span class="sxs-lookup"><span data-stu-id="be26b-939">Errors</span></span>

| <span data-ttu-id="be26b-940">错误代码</span><span class="sxs-lookup"><span data-stu-id="be26b-940">Error code</span></span> | <span data-ttu-id="be26b-941">说明</span><span class="sxs-lookup"><span data-stu-id="be26b-941">Description</span></span> |
|------------|-------------|
| `InvalidAttachmentId` | <span data-ttu-id="be26b-942">附件标识符不存在。</span><span class="sxs-lookup"><span data-stu-id="be26b-942">The attachment identifier does not exist.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="be26b-943">要求</span><span class="sxs-lookup"><span data-stu-id="be26b-943">Requirements</span></span>

|<span data-ttu-id="be26b-944">要求</span><span class="sxs-lookup"><span data-stu-id="be26b-944">Requirement</span></span>| <span data-ttu-id="be26b-945">值</span><span class="sxs-lookup"><span data-stu-id="be26b-945">Value</span></span>|
|---|---|
|[<span data-ttu-id="be26b-946">最低邮箱要求集版本</span><span class="sxs-lookup"><span data-stu-id="be26b-946">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="be26b-947">1.1</span><span class="sxs-lookup"><span data-stu-id="be26b-947">1.1</span></span>|
|[<span data-ttu-id="be26b-948">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="be26b-948">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="be26b-949">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="be26b-949">ReadWriteItem</span></span>|
|[<span data-ttu-id="be26b-950">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="be26b-950">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="be26b-951">撰写</span><span class="sxs-lookup"><span data-stu-id="be26b-951">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="be26b-952">示例</span><span class="sxs-lookup"><span data-stu-id="be26b-952">Example</span></span>

<span data-ttu-id="be26b-953">以下代码删除一个带  '0' 标识符的附件。</span><span class="sxs-lookup"><span data-stu-id="be26b-953">The following code removes an attachment with an identifier of '0'.</span></span>

```JavaScript
Office.context.mailbox.item.removeAttachmentAsync(
  '0',
  { asyncContext : null },
  function (asyncResult)
  {
    console.log(asyncResult.status);
  }
);
```

####  <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="be26b-954">setSelectedDataAsync(数据, [选项], 回调)</span><span class="sxs-lookup"><span data-stu-id="be26b-954">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="be26b-955">以异步方式将数据插入到邮件的正文或主题中。</span><span class="sxs-lookup"><span data-stu-id="be26b-955">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="be26b-p166">`setSelectedDataAsync` 方法将指定字符串插入到项目主题或正文的光标位置，或者，如果在编辑器中已选择文本，则该方法将替换选择的文本。如果光标不在正文或主题字段中，则返回错误。插入之后，光标会位于插入内容的末尾。</span><span class="sxs-lookup"><span data-stu-id="be26b-p166">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="be26b-959">参数：</span><span class="sxs-lookup"><span data-stu-id="be26b-959">Parameters:</span></span>

|<span data-ttu-id="be26b-960">名称</span><span class="sxs-lookup"><span data-stu-id="be26b-960">Name</span></span>| <span data-ttu-id="be26b-961">类型</span><span class="sxs-lookup"><span data-stu-id="be26b-961">Type</span></span>| <span data-ttu-id="be26b-962">属性</span><span class="sxs-lookup"><span data-stu-id="be26b-962">Attributes</span></span>| <span data-ttu-id="be26b-963">说明</span><span class="sxs-lookup"><span data-stu-id="be26b-963">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="be26b-964">字符串</span><span class="sxs-lookup"><span data-stu-id="be26b-964">String</span></span>||<span data-ttu-id="be26b-p167">要插入的数据。数据不得超过 1,000,000 个字符。如果传入的数据超过 1,000,000 个字符，则会引发 `ArgumentOutOfRange` 异常。</span><span class="sxs-lookup"><span data-stu-id="be26b-p167">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`| <span data-ttu-id="be26b-968">Object</span><span class="sxs-lookup"><span data-stu-id="be26b-968">Object</span></span>| <span data-ttu-id="be26b-969">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="be26b-969">&lt;optional&gt;</span></span>|<span data-ttu-id="be26b-970">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="be26b-970">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="be26b-971">Object</span><span class="sxs-lookup"><span data-stu-id="be26b-971">Object</span></span>| <span data-ttu-id="be26b-972">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="be26b-972">&lt;optional&gt;</span></span>|<span data-ttu-id="be26b-973">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="be26b-973">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`| [<span data-ttu-id="be26b-974">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="be26b-974">Office.CoercionType</span></span>](office.md#coerciontype-string)| <span data-ttu-id="be26b-975">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="be26b-975">&lt;optional&gt;</span></span>|<span data-ttu-id="be26b-p168">如果是 `text`，则在 Outlook Web App 和 Outlook 中应用当前样式。如果该字段是 HTML 编辑器，则仅插入文本数据，即使数据为 HTML。</span><span class="sxs-lookup"><span data-stu-id="be26b-p168">If `text`, the current style is applied in Outlook Web App and Outlook. If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="be26b-p169">如果 `html` 和该字段支持 HTML（主题不支持），则在 Outlook Web App 中应用当前样式，而在 Outlook 中应用默认样式。如果该字段是文本字段，则返回 `InvalidDataFormat` 错误。</span><span class="sxs-lookup"><span data-stu-id="be26b-p169">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook Web App and the default style is applied in Outlook. If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="be26b-980">如果未设置 `coercionType`，则结果取决于该字段：如果该字段是 HTML，则使用 HTML；如果该字段是文本，则使用纯文本。</span><span class="sxs-lookup"><span data-stu-id="be26b-980">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`| <span data-ttu-id="be26b-981">function</span><span class="sxs-lookup"><span data-stu-id="be26b-981">function</span></span>||<span data-ttu-id="be26b-982">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="be26b-982">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="be26b-983">要求</span><span class="sxs-lookup"><span data-stu-id="be26b-983">Requirements</span></span>

|<span data-ttu-id="be26b-984">要求</span><span class="sxs-lookup"><span data-stu-id="be26b-984">Requirement</span></span>| <span data-ttu-id="be26b-985">值</span><span class="sxs-lookup"><span data-stu-id="be26b-985">Value</span></span>|
|---|---|
|[<span data-ttu-id="be26b-986">最低的邮箱要求集版本</span><span class="sxs-lookup"><span data-stu-id="be26b-986">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="be26b-987">1.2</span><span class="sxs-lookup"><span data-stu-id="be26b-987">1.2</span></span>|
|[<span data-ttu-id="be26b-988">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="be26b-988">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="be26b-989">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="be26b-989">ReadWriteItem</span></span>|
|[<span data-ttu-id="be26b-990">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="be26b-990">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="be26b-991">撰写</span><span class="sxs-lookup"><span data-stu-id="be26b-991">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="be26b-992">示例</span><span class="sxs-lookup"><span data-stu-id="be26b-992">Example</span></span>

```JavaScript
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```