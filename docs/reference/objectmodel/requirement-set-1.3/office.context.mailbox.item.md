
# <a name="item"></a><span data-ttu-id="13674-101">item</span><span class="sxs-lookup"><span data-stu-id="13674-101">item</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmditem"></a><span data-ttu-id="13674-102">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span><span class="sxs-lookup"><span data-stu-id="13674-102">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span></span>

<span data-ttu-id="13674-p101">`item` 命名空间用于访问当前选定的邮件、会议请求或约会。可以通过使用 [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlook13officemailboxenumsitemtype) 属性确定 `item` 的类型。</span><span class="sxs-lookup"><span data-stu-id="13674-p101">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlook13officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="13674-105">Requirements</span><span class="sxs-lookup"><span data-stu-id="13674-105">Requirements</span></span>

|<span data-ttu-id="13674-106">要求</span><span class="sxs-lookup"><span data-stu-id="13674-106">Requirement</span></span>| <span data-ttu-id="13674-107">值</span><span class="sxs-lookup"><span data-stu-id="13674-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="13674-108">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="13674-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="13674-109">1.0</span><span class="sxs-lookup"><span data-stu-id="13674-109">1.0</span></span>|
|[<span data-ttu-id="13674-110">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="13674-110">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="13674-111">受限</span><span class="sxs-lookup"><span data-stu-id="13674-111">Restricted</span></span>|
|[<span data-ttu-id="13674-112">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="13674-112">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="13674-113">Compose 或 Read</span><span class="sxs-lookup"><span data-stu-id="13674-113">Compose or read</span></span>|

### <a name="example"></a><span data-ttu-id="13674-114">示例</span><span class="sxs-lookup"><span data-stu-id="13674-114">Example</span></span>

<span data-ttu-id="13674-115">以下 JavaScript 代码示例显示了如何访问 Outlook 中当前项目的 `subject` 属性。</span><span class="sxs-lookup"><span data-stu-id="13674-115">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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

### <a name="members"></a><span data-ttu-id="13674-116">成员</span><span class="sxs-lookup"><span data-stu-id="13674-116">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlook13officeattachmentdetails"></a><span data-ttu-id="13674-117">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_3/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="13674-117">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_3/office.attachmentdetails)></span></span>

<span data-ttu-id="13674-p102">获取项目的附件数组。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="13674-p102">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="13674-120">由于潜在的安全问题，某些类型的文件会受到 Outlook 阻止，并且不会返回。</span><span class="sxs-lookup"><span data-stu-id="13674-120">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="13674-121">如需了解更多信息，请参阅 [Outlook 中阻止的附件](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519)。</span><span class="sxs-lookup"><span data-stu-id="13674-121">For more information see [Payments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="13674-122">类型：</span><span class="sxs-lookup"><span data-stu-id="13674-122">Type:</span></span>

*   <span data-ttu-id="13674-123">Array.<[AttachmentDetails](/javascript/api/outlook_1_3/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="13674-123">Array.<[AttachmentDetails](/javascript/api/outlook_1_3/office.attachmentdetails)></span></span>

##### <a name="requirements"></a><span data-ttu-id="13674-124">要求</span><span class="sxs-lookup"><span data-stu-id="13674-124">Requirements</span></span>

|<span data-ttu-id="13674-125">要求</span><span class="sxs-lookup"><span data-stu-id="13674-125">Requirement</span></span>| <span data-ttu-id="13674-126">值</span><span class="sxs-lookup"><span data-stu-id="13674-126">Value</span></span>|
|---|---|
|[<span data-ttu-id="13674-127">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="13674-127">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="13674-128">1.0</span><span class="sxs-lookup"><span data-stu-id="13674-128">1.0</span></span>|
|[<span data-ttu-id="13674-129">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="13674-129">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="13674-130">ReadItem</span><span class="sxs-lookup"><span data-stu-id="13674-130">ReadItem</span></span>|
|[<span data-ttu-id="13674-131">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="13674-131">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="13674-132">阅读</span><span class="sxs-lookup"><span data-stu-id="13674-132">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="13674-133">示例</span><span class="sxs-lookup"><span data-stu-id="13674-133">Example</span></span>

<span data-ttu-id="13674-134">以下代码使用当前项目上所有附件的详细信息构成 HTML 字符串。</span><span class="sxs-lookup"><span data-stu-id="13674-134">The following code builds an HTML string with details of all attachments on the current item.</span></span>

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

####  <a name="bcc-recipientsjavascriptapioutlook13officerecipients"></a><span data-ttu-id="13674-135">bcc :[Recipients](/javascript/api/outlook_1_3/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="13674-135">bcc :[Recipients](/javascript/api/outlook_1_3/office.recipients)</span></span>

<span data-ttu-id="13674-136">获取对象，该对象提供用于获取或更新邮件的密件抄送 (Bcc) 行上的收件人的方法。</span><span class="sxs-lookup"><span data-stu-id="13674-136">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="13674-137">仅限撰写模式。</span><span class="sxs-lookup"><span data-stu-id="13674-137">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="13674-138">类型：</span><span class="sxs-lookup"><span data-stu-id="13674-138">Type:</span></span>

*   [<span data-ttu-id="13674-139">收件人</span><span class="sxs-lookup"><span data-stu-id="13674-139">Recipients</span></span>](/javascript/api/outlook_1_3/office.recipients)

##### <a name="requirements"></a><span data-ttu-id="13674-140">要求</span><span class="sxs-lookup"><span data-stu-id="13674-140">Requirements</span></span>

|<span data-ttu-id="13674-141">要求</span><span class="sxs-lookup"><span data-stu-id="13674-141">Requirement</span></span>| <span data-ttu-id="13674-142">值</span><span class="sxs-lookup"><span data-stu-id="13674-142">Value</span></span>|
|---|---|
|[<span data-ttu-id="13674-143">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="13674-143">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="13674-144">1.1</span><span class="sxs-lookup"><span data-stu-id="13674-144">1.1</span></span>|
|[<span data-ttu-id="13674-145">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="13674-145">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="13674-146">ReadItem</span><span class="sxs-lookup"><span data-stu-id="13674-146">ReadItem</span></span>|
|[<span data-ttu-id="13674-147">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="13674-147">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="13674-148">撰写</span><span class="sxs-lookup"><span data-stu-id="13674-148">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="13674-149">示例</span><span class="sxs-lookup"><span data-stu-id="13674-149">Example</span></span>

```js
Office.context.mailbox.item.bcc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.bcc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.bcc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfBccRecipients = asyncResult.value;
}
```

####  <a name="body-bodyjavascriptapioutlook13officebody"></a><span data-ttu-id="13674-150">body :[Body](/javascript/api/outlook_1_3/office.body)</span><span class="sxs-lookup"><span data-stu-id="13674-150">body :[Body](/javascript/api/outlook_1_3/office.body)</span></span>

<span data-ttu-id="13674-151">获取一个提供用于处理项目正文的方法的对象。</span><span class="sxs-lookup"><span data-stu-id="13674-151">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="13674-152">类型：</span><span class="sxs-lookup"><span data-stu-id="13674-152">Type:</span></span>

*   [<span data-ttu-id="13674-153">Body</span><span class="sxs-lookup"><span data-stu-id="13674-153">Body</span></span>](/javascript/api/outlook_1_3/office.body)

##### <a name="requirements"></a><span data-ttu-id="13674-154">要求</span><span class="sxs-lookup"><span data-stu-id="13674-154">Requirements</span></span>

|<span data-ttu-id="13674-155">要求</span><span class="sxs-lookup"><span data-stu-id="13674-155">Requirement</span></span>| <span data-ttu-id="13674-156">值</span><span class="sxs-lookup"><span data-stu-id="13674-156">Value</span></span>|
|---|---|
|[<span data-ttu-id="13674-157">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="13674-157">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="13674-158">1.1</span><span class="sxs-lookup"><span data-stu-id="13674-158">1.1</span></span>|
|[<span data-ttu-id="13674-159">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="13674-159">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="13674-160">ReadItem</span><span class="sxs-lookup"><span data-stu-id="13674-160">ReadItem</span></span>|
|[<span data-ttu-id="13674-161">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="13674-161">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="13674-162">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="13674-162">Compose or read</span></span>|

####  <a name="cc-arrayemailaddressdetailsjavascriptapioutlook13officeemailaddressdetailsrecipientsjavascriptapioutlook13officerecipients"></a><span data-ttu-id="13674-163">cc :Array.<[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_3/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="13674-163">cc :Array.<[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_3/office.recipients)</span></span>

<span data-ttu-id="13674-164">提供对邮件的抄送 (Cc) 收件人的访问权限。</span><span class="sxs-lookup"><span data-stu-id="13674-164">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="13674-165">对象的类型和访问级别取决于当前项目的模式。</span><span class="sxs-lookup"><span data-stu-id="13674-165">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="13674-166">读取模式</span><span class="sxs-lookup"><span data-stu-id="13674-166">Read mode</span></span>

<span data-ttu-id="13674-p106">`cc` 属性返回包含邮件的**抄送**行上所列的每个收件人的 `EmailAddressDetails` 对象的数组。集合上限为 100 个成员。</span><span class="sxs-lookup"><span data-stu-id="13674-p106">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="13674-169">撰写模式</span><span class="sxs-lookup"><span data-stu-id="13674-169">Compose mode</span></span>

<span data-ttu-id="13674-170">`cc` 属性返回一个 `Recipients` 对象，该对象提供用于获取或更新邮件的**抄送**行上收件人的方法。</span><span class="sxs-lookup"><span data-stu-id="13674-170">The `cc` property returns a `Recipients` object that provides methods for manipulating the recipients on the **Cc** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="13674-171">类型：</span><span class="sxs-lookup"><span data-stu-id="13674-171">Type:</span></span>

*   <span data-ttu-id="13674-172">Array.<[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_3/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="13674-172">Array.<[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_3/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="13674-173">要求</span><span class="sxs-lookup"><span data-stu-id="13674-173">Requirements</span></span>

|<span data-ttu-id="13674-174">要求</span><span class="sxs-lookup"><span data-stu-id="13674-174">Requirement</span></span>| <span data-ttu-id="13674-175">值</span><span class="sxs-lookup"><span data-stu-id="13674-175">Value</span></span>|
|---|---|
|[<span data-ttu-id="13674-176">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="13674-176">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="13674-177">1.0</span><span class="sxs-lookup"><span data-stu-id="13674-177">1.0</span></span>|
|[<span data-ttu-id="13674-178">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="13674-178">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="13674-179">ReadItem</span><span class="sxs-lookup"><span data-stu-id="13674-179">ReadItem</span></span>|
|[<span data-ttu-id="13674-180">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="13674-180">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="13674-181">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="13674-181">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="13674-182">示例</span><span class="sxs-lookup"><span data-stu-id="13674-182">Example</span></span>

```js
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

####  <a name="nullable-conversationid-string"></a><span data-ttu-id="13674-183">(nullable) conversationId :String</span><span class="sxs-lookup"><span data-stu-id="13674-183">(nullable) conversationId :String</span></span>

<span data-ttu-id="13674-184">获取包含特定消息的电子邮件会话的标识符。</span><span class="sxs-lookup"><span data-stu-id="13674-184">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="13674-p107">如果在阅读窗体或撰写窗体的回复中激活邮件应用程序，则此属性可以获得一个整数值。如果用户随后更改了回复邮件的主题（若发送回复），则该邮件的对话 ID 将改变且之前获取的值将不适用。</span><span class="sxs-lookup"><span data-stu-id="13674-p107">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="13674-p108">对于撰写窗体的新项目，此属性获得一个 null 值。如果用户设置一个主题并保存该项目，`conversationId` 属性将返回一个值。</span><span class="sxs-lookup"><span data-stu-id="13674-p108">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="13674-189">类型:</span><span class="sxs-lookup"><span data-stu-id="13674-189">Type:</span></span>

*   <span data-ttu-id="13674-190">String</span><span class="sxs-lookup"><span data-stu-id="13674-190">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="13674-191">要求</span><span class="sxs-lookup"><span data-stu-id="13674-191">Requirements</span></span>

|<span data-ttu-id="13674-192">要求</span><span class="sxs-lookup"><span data-stu-id="13674-192">Requirement</span></span>| <span data-ttu-id="13674-193">值</span><span class="sxs-lookup"><span data-stu-id="13674-193">Value</span></span>|
|---|---|
|[<span data-ttu-id="13674-194">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="13674-194">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="13674-195">1.0</span><span class="sxs-lookup"><span data-stu-id="13674-195">1.0</span></span>|
|[<span data-ttu-id="13674-196">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="13674-196">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="13674-197">ReadItem</span><span class="sxs-lookup"><span data-stu-id="13674-197">ReadItem</span></span>|
|[<span data-ttu-id="13674-198">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="13674-198">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="13674-199">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="13674-199">Compose or read</span></span>|

#### <a name="datetimecreated-date"></a><span data-ttu-id="13674-200">dateTimeCreated :Date</span><span class="sxs-lookup"><span data-stu-id="13674-200">dateTimeCreated :Date</span></span>

<span data-ttu-id="13674-p109">获取项目创建的日期和时间。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="13674-p109">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="13674-203">类型：</span><span class="sxs-lookup"><span data-stu-id="13674-203">Type:</span></span>

*   <span data-ttu-id="13674-204">日期</span><span class="sxs-lookup"><span data-stu-id="13674-204">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="13674-205">要求</span><span class="sxs-lookup"><span data-stu-id="13674-205">Requirements</span></span>

|<span data-ttu-id="13674-206">要求</span><span class="sxs-lookup"><span data-stu-id="13674-206">Requirement</span></span>| <span data-ttu-id="13674-207">值</span><span class="sxs-lookup"><span data-stu-id="13674-207">Value</span></span>|
|---|---|
|[<span data-ttu-id="13674-208">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="13674-208">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="13674-209">1.0</span><span class="sxs-lookup"><span data-stu-id="13674-209">1.0</span></span>|
|[<span data-ttu-id="13674-210">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="13674-210">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="13674-211">ReadItem</span><span class="sxs-lookup"><span data-stu-id="13674-211">ReadItem</span></span>|
|[<span data-ttu-id="13674-212">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="13674-212">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="13674-213">阅读</span><span class="sxs-lookup"><span data-stu-id="13674-213">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="13674-214">示例</span><span class="sxs-lookup"><span data-stu-id="13674-214">Example</span></span>

```js
var created = Office.context.mailbox.item.dateTimeCreated;
```

#### <a name="datetimemodified-date"></a><span data-ttu-id="13674-215">dateTimeModified :Date</span><span class="sxs-lookup"><span data-stu-id="13674-215">dateTimeModified :Date</span></span>

<span data-ttu-id="13674-p110">获取项目最近一次修改的日期和时间。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="13674-p110">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="13674-218">在 Outlook for iOS 或 Outlook for Android 中不支持此成员。</span><span class="sxs-lookup"><span data-stu-id="13674-218">Note: This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="type"></a><span data-ttu-id="13674-219">类型:</span><span class="sxs-lookup"><span data-stu-id="13674-219">Type:</span></span>

*   <span data-ttu-id="13674-220">日期</span><span class="sxs-lookup"><span data-stu-id="13674-220">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="13674-221">要求</span><span class="sxs-lookup"><span data-stu-id="13674-221">Requirements</span></span>

|<span data-ttu-id="13674-222">要求</span><span class="sxs-lookup"><span data-stu-id="13674-222">Requirement</span></span>| <span data-ttu-id="13674-223">值</span><span class="sxs-lookup"><span data-stu-id="13674-223">Value</span></span>|
|---|---|
|[<span data-ttu-id="13674-224">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="13674-224">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="13674-225">1.0</span><span class="sxs-lookup"><span data-stu-id="13674-225">1.0</span></span>|
|[<span data-ttu-id="13674-226">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="13674-226">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="13674-227">ReadItem</span><span class="sxs-lookup"><span data-stu-id="13674-227">ReadItem</span></span>|
|[<span data-ttu-id="13674-228">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="13674-228">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="13674-229">阅读</span><span class="sxs-lookup"><span data-stu-id="13674-229">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="13674-230">示例</span><span class="sxs-lookup"><span data-stu-id="13674-230">Example</span></span>

```js
var modified = Office.context.mailbox.item.dateTimeModified;
```

####  <a name="end-datetimejavascriptapioutlook13officetime"></a><span data-ttu-id="13674-231">end :Date|[Time](/javascript/api/outlook_1_3/office.time)</span><span class="sxs-lookup"><span data-stu-id="13674-231">end :Date|[Time](/javascript/api/outlook_1_3/office.time)</span></span>

<span data-ttu-id="13674-232">获取或设置约会结束的日期和时间。</span><span class="sxs-lookup"><span data-stu-id="13674-232">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="13674-p111">将 `end` 属性表示为协调世界时 (UTC) 的日期和时间值。可使用 [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook13officelocalclienttime) 方法将 end 属性值转换为客户端的本地日期和时间。</span><span class="sxs-lookup"><span data-stu-id="13674-p111">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook13officelocalclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="13674-235">阅读模式</span><span class="sxs-lookup"><span data-stu-id="13674-235">Read mode</span></span>

<span data-ttu-id="13674-236">`end` 属性返回 `Date` 对象。</span><span class="sxs-lookup"><span data-stu-id="13674-236">The `end` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="13674-237">撰写模式</span><span class="sxs-lookup"><span data-stu-id="13674-237">Compose mode</span></span>

<span data-ttu-id="13674-238">`end` 属性返回 `Time` 对象。</span><span class="sxs-lookup"><span data-stu-id="13674-238">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="13674-239">使用 [`Time.setAsync`](/javascript/api/outlook_1_3/office.time#setasync-datetime--options--callback-) 方法设置结束时间时，应使用 [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) 方法将客户端的本地时间转换为服务器的 UTC。</span><span class="sxs-lookup"><span data-stu-id="13674-239">When you use the [`Time.setAsync`](/javascript/api/outlook_1_3/office.time#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="13674-240">类型：</span><span class="sxs-lookup"><span data-stu-id="13674-240">Type:</span></span>

*   <span data-ttu-id="13674-241">Date | [Time](/javascript/api/outlook_1_3/office.time)</span><span class="sxs-lookup"><span data-stu-id="13674-241">Date | [Time](/javascript/api/outlook_1_3/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="13674-242">要求</span><span class="sxs-lookup"><span data-stu-id="13674-242">Requirements</span></span>

|<span data-ttu-id="13674-243">要求</span><span class="sxs-lookup"><span data-stu-id="13674-243">Requirement</span></span>| <span data-ttu-id="13674-244">值</span><span class="sxs-lookup"><span data-stu-id="13674-244">Value</span></span>|
|---|---|
|[<span data-ttu-id="13674-245">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="13674-245">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="13674-246">1.0</span><span class="sxs-lookup"><span data-stu-id="13674-246">1.0</span></span>|
|[<span data-ttu-id="13674-247">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="13674-247">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="13674-248">ReadItem</span><span class="sxs-lookup"><span data-stu-id="13674-248">ReadItem</span></span>|
|[<span data-ttu-id="13674-249">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="13674-249">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="13674-250">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="13674-250">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="13674-251">示例</span><span class="sxs-lookup"><span data-stu-id="13674-251">Example</span></span>

<span data-ttu-id="13674-252">以下示例通过使用 `Time` 对象的 [`setAsync`](/javascript/api/outlook_1_3/office.time#setasync-datetime--options--callback-) 方法，设置撰写模式下约会的结束时间。</span><span class="sxs-lookup"><span data-stu-id="13674-252">The following example sets the end time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_3/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

#### <a name="from-emailaddressdetailsjavascriptapioutlook13officeemailaddressdetails"></a><span data-ttu-id="13674-253">from :[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="13674-253">from :[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)</span></span>

<span data-ttu-id="13674-p112">获取邮件发件人的电子邮件地址。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="13674-p112">Gets the email address of the sender of a message. Read mode only.</span></span>

<span data-ttu-id="13674-p113">`from` 和 [`sender`](#sender-emailaddressdetailsjavascriptapioutlook13officeemailaddressdetails) 属性表示同一个人，邮件由代理人发送的除外。在此情况下，`from` 属性表示代理程序，而 sender 属性表示代理人。</span><span class="sxs-lookup"><span data-stu-id="13674-p113">The `from` and [`sender`](#sender-emailaddressdetailsjavascriptapioutlook13officeemailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="13674-258">`from` 属性中 `EmailAddressDetails` 对象的 `recipientType` 属性为 `undefined`。</span><span class="sxs-lookup"><span data-stu-id="13674-258">Note: The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="13674-259">类型：</span><span class="sxs-lookup"><span data-stu-id="13674-259">Type:</span></span>

*   [<span data-ttu-id="13674-260">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="13674-260">EmailAddressDetails</span></span>](/javascript/api/outlook_1_3/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="13674-261">要求</span><span class="sxs-lookup"><span data-stu-id="13674-261">Requirements</span></span>

|<span data-ttu-id="13674-262">要求</span><span class="sxs-lookup"><span data-stu-id="13674-262">Requirement</span></span>| <span data-ttu-id="13674-263">值</span><span class="sxs-lookup"><span data-stu-id="13674-263">Value</span></span>|
|---|---|
|[<span data-ttu-id="13674-264">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="13674-264">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="13674-265">1.0</span><span class="sxs-lookup"><span data-stu-id="13674-265">1.0</span></span>|
|[<span data-ttu-id="13674-266">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="13674-266">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="13674-267">ReadItem</span><span class="sxs-lookup"><span data-stu-id="13674-267">ReadItem</span></span>|
|[<span data-ttu-id="13674-268">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="13674-268">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="13674-269">阅读</span><span class="sxs-lookup"><span data-stu-id="13674-269">Read</span></span>|

#### <a name="internetmessageid-string"></a><span data-ttu-id="13674-270">internetMessageId :String</span><span class="sxs-lookup"><span data-stu-id="13674-270">internetMessageId :String</span></span>

<span data-ttu-id="13674-p114">获取电子邮件的 Internet 消息标识符。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="13674-p114">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="13674-273">类型：</span><span class="sxs-lookup"><span data-stu-id="13674-273">Type:</span></span>

*   <span data-ttu-id="13674-274">String</span><span class="sxs-lookup"><span data-stu-id="13674-274">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="13674-275">要求</span><span class="sxs-lookup"><span data-stu-id="13674-275">Requirements</span></span>

|<span data-ttu-id="13674-276">要求</span><span class="sxs-lookup"><span data-stu-id="13674-276">Requirement</span></span>| <span data-ttu-id="13674-277">值</span><span class="sxs-lookup"><span data-stu-id="13674-277">Value</span></span>|
|---|---|
|[<span data-ttu-id="13674-278">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="13674-278">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="13674-279">1.0</span><span class="sxs-lookup"><span data-stu-id="13674-279">1.0</span></span>|
|[<span data-ttu-id="13674-280">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="13674-280">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="13674-281">ReadItem</span><span class="sxs-lookup"><span data-stu-id="13674-281">ReadItem</span></span>|
|[<span data-ttu-id="13674-282">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="13674-282">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="13674-283">阅读</span><span class="sxs-lookup"><span data-stu-id="13674-283">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="13674-284">示例</span><span class="sxs-lookup"><span data-stu-id="13674-284">Example</span></span>

```js
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

#### <a name="itemclass-string"></a><span data-ttu-id="13674-285">itemClass :String</span><span class="sxs-lookup"><span data-stu-id="13674-285">itemClass :String</span></span>

<span data-ttu-id="13674-p115">获取选定项目的 Exchange Web 服务项目类。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="13674-p115">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="13674-p116">`itemClass` 属性指定所选项目的邮件类别。以下是邮件或约会项目的默认邮件类别。</span><span class="sxs-lookup"><span data-stu-id="13674-p116">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

| <span data-ttu-id="13674-290">类型</span><span class="sxs-lookup"><span data-stu-id="13674-290">Type</span></span> | <span data-ttu-id="13674-291">说明</span><span class="sxs-lookup"><span data-stu-id="13674-291">Description</span></span> | <span data-ttu-id="13674-292">项目类</span><span class="sxs-lookup"><span data-stu-id="13674-292">item class</span></span> |
| --- | --- | --- |
| <span data-ttu-id="13674-293">约会项目</span><span class="sxs-lookup"><span data-stu-id="13674-293">Appointment items</span></span> | <span data-ttu-id="13674-294">这些是项目类 `IPM.Appointment` 或 `IPM.Appointment.Occurence` 的日历项目。</span><span class="sxs-lookup"><span data-stu-id="13674-294">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurence`.</span></span> | `IPM.Appointment`<br />`IPM.Appointment.Occurence` |
| <span data-ttu-id="13674-295">邮件项目</span><span class="sxs-lookup"><span data-stu-id="13674-295">Message items</span></span> | <span data-ttu-id="13674-296">这些项目包括具有默认邮件类别 `IPM.Note` 的电子邮件，以及使用 `IPM.Schedule.Meeting` 作为基础邮件类别的会议请求、响应和取消。</span><span class="sxs-lookup"><span data-stu-id="13674-296">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span> | `IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled` |

<span data-ttu-id="13674-297">你可以创建用于扩展默认邮件类别的自定义邮件类别，例如，自定义约会邮件类别 `IPM.Appointment.Contoso`。</span><span class="sxs-lookup"><span data-stu-id="13674-297">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="13674-298">类型：</span><span class="sxs-lookup"><span data-stu-id="13674-298">Type:</span></span>

*   <span data-ttu-id="13674-299">String</span><span class="sxs-lookup"><span data-stu-id="13674-299">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="13674-300">要求</span><span class="sxs-lookup"><span data-stu-id="13674-300">Requirements</span></span>

|<span data-ttu-id="13674-301">要求</span><span class="sxs-lookup"><span data-stu-id="13674-301">Requirement</span></span>| <span data-ttu-id="13674-302">值</span><span class="sxs-lookup"><span data-stu-id="13674-302">Value</span></span>|
|---|---|
|[<span data-ttu-id="13674-303">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="13674-303">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="13674-304">1.0</span><span class="sxs-lookup"><span data-stu-id="13674-304">1.0</span></span>|
|[<span data-ttu-id="13674-305">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="13674-305">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="13674-306">ReadItem</span><span class="sxs-lookup"><span data-stu-id="13674-306">ReadItem</span></span>|
|[<span data-ttu-id="13674-307">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="13674-307">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="13674-308">阅读</span><span class="sxs-lookup"><span data-stu-id="13674-308">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="13674-309">示例</span><span class="sxs-lookup"><span data-stu-id="13674-309">Example</span></span>

```js
var itemClass = Office.context.mailbox.item.itemClass;
```

#### <a name="nullable-itemid-string"></a><span data-ttu-id="13674-310">(nullable) itemId :String</span><span class="sxs-lookup"><span data-stu-id="13674-310">(nullable) itemId :String</span></span>

<span data-ttu-id="13674-p117">获取当前项目的 Exchange Web 服务项目标识符。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="13674-p117">Gets the Exchange Web Services item identifier for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="13674-313">`itemId` 属性返回的标识符与 Exchange Web 服务项目标识符相同。</span><span class="sxs-lookup"><span data-stu-id="13674-313">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="13674-314">`itemId` 属性与 Outlook 条目 ID 或 Outlook REST API 所使用的 ID 不同。</span><span class="sxs-lookup"><span data-stu-id="13674-314">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="13674-315">使用此值进行 REST API 调用前，应使用 [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string) 对它进行转换。</span><span class="sxs-lookup"><span data-stu-id="13674-315">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="13674-316">有关详细信息，请参阅[从 Outlook 外接程序使用 Outlook REST API](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id)。</span><span class="sxs-lookup"><span data-stu-id="13674-316">For more details, see [Use the Outlook REST APIs from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

<span data-ttu-id="13674-p119">`itemId` 属性在撰写模式下不可用。如果需要项目标识符，[`saveAsync`](#saveasyncoptions-callback) 方法可用于将项目保存到存储，这将在回调函数的 [`AsyncResult.value`](/javascript/api/office/office.asyncresult) 参数中返回项目标识符。</span><span class="sxs-lookup"><span data-stu-id="13674-p119">The `itemId` property is not available in compose mode. If an item identifier is required, the [`saveAsync`](#saveasyncoptions-callback) method can be used to save the item to the store, which will return the item identifier in the [`AsyncResult.value`](/javascript/api/office/office.asyncresult) parameter in the callback function.</span></span>

##### <a name="type"></a><span data-ttu-id="13674-319">类型:</span><span class="sxs-lookup"><span data-stu-id="13674-319">Type:</span></span>

*   <span data-ttu-id="13674-320">String</span><span class="sxs-lookup"><span data-stu-id="13674-320">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="13674-321">要求</span><span class="sxs-lookup"><span data-stu-id="13674-321">Requirements</span></span>

|<span data-ttu-id="13674-322">要求</span><span class="sxs-lookup"><span data-stu-id="13674-322">Requirement</span></span>| <span data-ttu-id="13674-323">值</span><span class="sxs-lookup"><span data-stu-id="13674-323">Value</span></span>|
|---|---|
|[<span data-ttu-id="13674-324">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="13674-324">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="13674-325">1.0</span><span class="sxs-lookup"><span data-stu-id="13674-325">1.0</span></span>|
|[<span data-ttu-id="13674-326">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="13674-326">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="13674-327">ReadItem</span><span class="sxs-lookup"><span data-stu-id="13674-327">ReadItem</span></span>|
|[<span data-ttu-id="13674-328">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="13674-328">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="13674-329">阅读</span><span class="sxs-lookup"><span data-stu-id="13674-329">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="13674-330">示例</span><span class="sxs-lookup"><span data-stu-id="13674-330">Example</span></span>

<span data-ttu-id="13674-p120">以下代码检查项目标识符是否存在。如果 `itemId` 属性返回 `null` 或 `undefined`，则将项目保存到存储，并从异步结果中获取项目标识符。</span><span class="sxs-lookup"><span data-stu-id="13674-p120">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

```js
var itemId = Office.context.mailbox.item.itemId;
if (itemId === null || itemId == undefined) {
  Office.context.mailbox.item.saveAsync(function(result){
    itemId = result.value;
  });
}
```

####  <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlook13officemailboxenumsitemtype"></a><span data-ttu-id="13674-333">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_3/office.mailboxenums.itemtype)</span><span class="sxs-lookup"><span data-stu-id="13674-333">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_3/office.mailboxenums.itemtype)</span></span>

<span data-ttu-id="13674-334">获取实例表示的项的类型。</span><span class="sxs-lookup"><span data-stu-id="13674-334">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="13674-335">`itemType` 属性返回其中一个 `ItemType` 枚举值，指示 `item` 对象实例是邮件还是约会。</span><span class="sxs-lookup"><span data-stu-id="13674-335">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="13674-336">类型：</span><span class="sxs-lookup"><span data-stu-id="13674-336">Type:</span></span>

*   [<span data-ttu-id="13674-337">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="13674-337">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook_1_3/office.mailboxenums.itemtype)

##### <a name="requirements"></a><span data-ttu-id="13674-338">要求</span><span class="sxs-lookup"><span data-stu-id="13674-338">Requirements</span></span>

|<span data-ttu-id="13674-339">要求</span><span class="sxs-lookup"><span data-stu-id="13674-339">Requirement</span></span>| <span data-ttu-id="13674-340">值</span><span class="sxs-lookup"><span data-stu-id="13674-340">Value</span></span>|
|---|---|
|[<span data-ttu-id="13674-341">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="13674-341">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="13674-342">1.0</span><span class="sxs-lookup"><span data-stu-id="13674-342">1.0</span></span>|
|[<span data-ttu-id="13674-343">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="13674-343">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="13674-344">ReadItem</span><span class="sxs-lookup"><span data-stu-id="13674-344">ReadItem</span></span>|
|[<span data-ttu-id="13674-345">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="13674-345">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="13674-346">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="13674-346">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="13674-347">示例</span><span class="sxs-lookup"><span data-stu-id="13674-347">Example</span></span>

```js
if (Office.context.mailbox.item.itemType == Office.MailboxEnums.ItemType.Message)
  // do something
else
  // do something else
```

####  <a name="location-stringlocationjavascriptapioutlook13officelocation"></a><span data-ttu-id="13674-348">location :String|[Location](/javascript/api/outlook_1_3/office.location)</span><span class="sxs-lookup"><span data-stu-id="13674-348">location :String|[Location](/javascript/api/outlook_1_3/office.location)</span></span>

<span data-ttu-id="13674-349">获取或设置约会的位置。</span><span class="sxs-lookup"><span data-stu-id="13674-349">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="13674-350">阅读模式</span><span class="sxs-lookup"><span data-stu-id="13674-350">Read mode</span></span>

<span data-ttu-id="13674-351">`location` 属性返回一个包含约会位置的字符串。</span><span class="sxs-lookup"><span data-stu-id="13674-351">The `location` property returns a string that contains the location of the appointment.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="13674-352">撰写模式</span><span class="sxs-lookup"><span data-stu-id="13674-352">Compose mode</span></span>

<span data-ttu-id="13674-353">`location` 属性返回一个 `Location` 对象，该对象提供用于获取和设置约会位置的方法。</span><span class="sxs-lookup"><span data-stu-id="13674-353">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="13674-354">类型：</span><span class="sxs-lookup"><span data-stu-id="13674-354">Type:</span></span>

*   <span data-ttu-id="13674-355">String | [Location](/javascript/api/outlook_1_3/office.location)</span><span class="sxs-lookup"><span data-stu-id="13674-355">String | [Location](/javascript/api/outlook_1_3/office.location)</span></span>

##### <a name="requirements"></a><span data-ttu-id="13674-356">要求</span><span class="sxs-lookup"><span data-stu-id="13674-356">Requirements</span></span>

|<span data-ttu-id="13674-357">要求</span><span class="sxs-lookup"><span data-stu-id="13674-357">Requirement</span></span>| <span data-ttu-id="13674-358">值</span><span class="sxs-lookup"><span data-stu-id="13674-358">Value</span></span>|
|---|---|
|[<span data-ttu-id="13674-359">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="13674-359">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="13674-360">1.0</span><span class="sxs-lookup"><span data-stu-id="13674-360">1.0</span></span>|
|[<span data-ttu-id="13674-361">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="13674-361">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="13674-362">ReadItem</span><span class="sxs-lookup"><span data-stu-id="13674-362">ReadItem</span></span>|
|[<span data-ttu-id="13674-363">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="13674-363">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="13674-364">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="13674-364">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="13674-365">示例</span><span class="sxs-lookup"><span data-stu-id="13674-365">Example</span></span>

```js
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

#### <a name="normalizedsubject-string"></a><span data-ttu-id="13674-366">normalizedSubject :String</span><span class="sxs-lookup"><span data-stu-id="13674-366">normalizedSubject :String</span></span>

<span data-ttu-id="13674-p121">获取删除了所有前缀（包括 `RE:` 和 `FWD:`）的项目主题。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="13674-p121">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="13674-p122">normalizedSubject 属性获取包含由电子邮件程序添加的任何标准前缀（如 `RE:` 和 `FW:`）的项目主题。若要获取包含完整前缀的项目主题，请使用 [`subject`](#subject-stringsubjectjavascriptapioutlook13officesubject) 属性。</span><span class="sxs-lookup"><span data-stu-id="13674-p122">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubjectjavascriptapioutlook13officesubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="13674-371">类型:</span><span class="sxs-lookup"><span data-stu-id="13674-371">Type:</span></span>

*   <span data-ttu-id="13674-372">String</span><span class="sxs-lookup"><span data-stu-id="13674-372">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="13674-373">要求</span><span class="sxs-lookup"><span data-stu-id="13674-373">Requirements</span></span>

|<span data-ttu-id="13674-374">要求</span><span class="sxs-lookup"><span data-stu-id="13674-374">Requirement</span></span>| <span data-ttu-id="13674-375">值</span><span class="sxs-lookup"><span data-stu-id="13674-375">Value</span></span>|
|---|---|
|[<span data-ttu-id="13674-376">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="13674-376">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="13674-377">1.0</span><span class="sxs-lookup"><span data-stu-id="13674-377">1.0</span></span>|
|[<span data-ttu-id="13674-378">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="13674-378">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="13674-379">ReadItem</span><span class="sxs-lookup"><span data-stu-id="13674-379">ReadItem</span></span>|
|[<span data-ttu-id="13674-380">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="13674-380">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="13674-381">阅读</span><span class="sxs-lookup"><span data-stu-id="13674-381">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="13674-382">示例</span><span class="sxs-lookup"><span data-stu-id="13674-382">Example</span></span>

```js
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
```

####  <a name="notificationmessages-notificationmessagesjavascriptapioutlook13officenotificationmessages"></a><span data-ttu-id="13674-383">notificationMessages :[NotificationMessages](/javascript/api/outlook_1_3/office.notificationmessages)</span><span class="sxs-lookup"><span data-stu-id="13674-383">notificationMessages :[NotificationMessages](/javascript/api/outlook_1_3/office.notificationmessages)</span></span>

<span data-ttu-id="13674-384">获取项目的通知邮件。</span><span class="sxs-lookup"><span data-stu-id="13674-384">Gets the notification messages for an item.</span></span>

##### <a name="type"></a><span data-ttu-id="13674-385">类型：</span><span class="sxs-lookup"><span data-stu-id="13674-385">Type:</span></span>

*   [<span data-ttu-id="13674-386">NotificationMessages</span><span class="sxs-lookup"><span data-stu-id="13674-386">NotificationMessages</span></span>](/javascript/api/outlook_1_3/office.notificationmessages)

##### <a name="requirements"></a><span data-ttu-id="13674-387">要求</span><span class="sxs-lookup"><span data-stu-id="13674-387">Requirements</span></span>

|<span data-ttu-id="13674-388">要求</span><span class="sxs-lookup"><span data-stu-id="13674-388">Requirement</span></span>| <span data-ttu-id="13674-389">值</span><span class="sxs-lookup"><span data-stu-id="13674-389">Value</span></span>|
|---|---|
|[<span data-ttu-id="13674-390">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="13674-390">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="13674-391">1.3</span><span class="sxs-lookup"><span data-stu-id="13674-391">1.3</span></span>|
|[<span data-ttu-id="13674-392">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="13674-392">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="13674-393">ReadItem</span><span class="sxs-lookup"><span data-stu-id="13674-393">ReadItem</span></span>|
|[<span data-ttu-id="13674-394">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="13674-394">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="13674-395">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="13674-395">Compose or read</span></span>|

####  <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlook13officeemailaddressdetailsrecipientsjavascriptapioutlook13officerecipients"></a><span data-ttu-id="13674-396">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_3/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="13674-396">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_3/office.recipients)</span></span>

<span data-ttu-id="13674-397">提供对事件的可选与会者的访问权限。</span><span class="sxs-lookup"><span data-stu-id="13674-397">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="13674-398">对象的类型和访问级别取决于当前项目的模式。</span><span class="sxs-lookup"><span data-stu-id="13674-398">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="13674-399">读取模式</span><span class="sxs-lookup"><span data-stu-id="13674-399">Read mode</span></span>

<span data-ttu-id="13674-400">`optionalAttendees` 属性返回一个数组，其中包含每个可选与会者的 `EmailAddressDetails` 对象。</span><span class="sxs-lookup"><span data-stu-id="13674-400">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="13674-401">撰写模式</span><span class="sxs-lookup"><span data-stu-id="13674-401">Compose mode</span></span>

<span data-ttu-id="13674-402">`optionalAttendees` 属性返回一个 `Recipients` 对象，该对象提供用于获取或更新可选与会者的方法。</span><span class="sxs-lookup"><span data-stu-id="13674-402">The `optionalAttendees` property returns a `Recipients` object that provides methods to get and set the optional attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="13674-403">类型：</span><span class="sxs-lookup"><span data-stu-id="13674-403">Type:</span></span>

*   <span data-ttu-id="13674-404">Array.<[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_3/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="13674-404">Array.<[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_3/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="13674-405">要求</span><span class="sxs-lookup"><span data-stu-id="13674-405">Requirements</span></span>

|<span data-ttu-id="13674-406">要求</span><span class="sxs-lookup"><span data-stu-id="13674-406">Requirement</span></span>| <span data-ttu-id="13674-407">值</span><span class="sxs-lookup"><span data-stu-id="13674-407">Value</span></span>|
|---|---|
|[<span data-ttu-id="13674-408">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="13674-408">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="13674-409">1.0</span><span class="sxs-lookup"><span data-stu-id="13674-409">1.0</span></span>|
|[<span data-ttu-id="13674-410">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="13674-410">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="13674-411">ReadItem</span><span class="sxs-lookup"><span data-stu-id="13674-411">ReadItem</span></span>|
|[<span data-ttu-id="13674-412">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="13674-412">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="13674-413">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="13674-413">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="13674-414">示例</span><span class="sxs-lookup"><span data-stu-id="13674-414">Example</span></span>

```js
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

#### <a name="organizer-emailaddressdetailsjavascriptapioutlook13officeemailaddressdetails"></a><span data-ttu-id="13674-415">organizer :[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="13674-415">organizer :[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)</span></span>

<span data-ttu-id="13674-p124">获取指定会议的会议组织者的电子邮件地址。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="13674-p124">Gets the email address of the meeting organizer for a specified meeting. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="13674-418">类型:</span><span class="sxs-lookup"><span data-stu-id="13674-418">Type:</span></span>

*   [<span data-ttu-id="13674-419">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="13674-419">EmailAddressDetails</span></span>](/javascript/api/outlook_1_3/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="13674-420">要求</span><span class="sxs-lookup"><span data-stu-id="13674-420">Requirements</span></span>

|<span data-ttu-id="13674-421">要求</span><span class="sxs-lookup"><span data-stu-id="13674-421">Requirement</span></span>| <span data-ttu-id="13674-422">值</span><span class="sxs-lookup"><span data-stu-id="13674-422">Value</span></span>|
|---|---|
|[<span data-ttu-id="13674-423">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="13674-423">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="13674-424">1.0</span><span class="sxs-lookup"><span data-stu-id="13674-424">1.0</span></span>|
|[<span data-ttu-id="13674-425">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="13674-425">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="13674-426">ReadItem</span><span class="sxs-lookup"><span data-stu-id="13674-426">ReadItem</span></span>|
|[<span data-ttu-id="13674-427">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="13674-427">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="13674-428">阅读</span><span class="sxs-lookup"><span data-stu-id="13674-428">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="13674-429">示例</span><span class="sxs-lookup"><span data-stu-id="13674-429">Example</span></span>

```js
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
```

####  <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlook13officeemailaddressdetailsrecipientsjavascriptapioutlook13officerecipients"></a><span data-ttu-id="13674-430">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_3/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="13674-430">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_3/office.recipients)</span></span>

<span data-ttu-id="13674-431">提供对事件的必需与会者的访问权限。</span><span class="sxs-lookup"><span data-stu-id="13674-431">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="13674-432">对象的类型和访问级别取决于当前项目的模式。</span><span class="sxs-lookup"><span data-stu-id="13674-432">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="13674-433">读取模式</span><span class="sxs-lookup"><span data-stu-id="13674-433">Read mode</span></span>

<span data-ttu-id="13674-434">`requiredAttendees` 属性返回一个数组，其中包含每个必需与会者的 `EmailAddressDetails` 对象。</span><span class="sxs-lookup"><span data-stu-id="13674-434">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="13674-435">撰写模式</span><span class="sxs-lookup"><span data-stu-id="13674-435">Compose mode</span></span>

<span data-ttu-id="13674-436">`requiredAttendees` 属性返回一个 `Recipients` 对象，该对象提供用于获取或更新必需与会者的方法。</span><span class="sxs-lookup"><span data-stu-id="13674-436">The `requiredAttendees` property returns a `Recipients` object that provides methods to get and set the required attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="13674-437">类型：</span><span class="sxs-lookup"><span data-stu-id="13674-437">Type:</span></span>

*   <span data-ttu-id="13674-438">Array.<[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_3/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="13674-438">Array.<[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_3/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="13674-439">要求</span><span class="sxs-lookup"><span data-stu-id="13674-439">Requirements</span></span>

|<span data-ttu-id="13674-440">要求</span><span class="sxs-lookup"><span data-stu-id="13674-440">Requirement</span></span>| <span data-ttu-id="13674-441">值</span><span class="sxs-lookup"><span data-stu-id="13674-441">Value</span></span>|
|---|---|
|[<span data-ttu-id="13674-442">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="13674-442">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="13674-443">1.0</span><span class="sxs-lookup"><span data-stu-id="13674-443">1.0</span></span>|
|[<span data-ttu-id="13674-444">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="13674-444">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="13674-445">ReadItem</span><span class="sxs-lookup"><span data-stu-id="13674-445">ReadItem</span></span>|
|[<span data-ttu-id="13674-446">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="13674-446">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="13674-447">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="13674-447">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="13674-448">示例</span><span class="sxs-lookup"><span data-stu-id="13674-448">Example</span></span>

```js
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
}
```

#### <a name="sender-emailaddressdetailsjavascriptapioutlook13officeemailaddressdetails"></a><span data-ttu-id="13674-449">sender :[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="13674-449">sender :[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)</span></span>

<span data-ttu-id="13674-p126">获取电子邮件发件人的电子邮件地址。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="13674-p126">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="13674-p127">[`from`](#from-emailaddressdetailsjavascriptapioutlook13officeemailaddressdetails) 和 `sender` 属性表示同一个人，邮件由代理人发送的除外。在此情况下，`from` 属性表示代理程序，而 sender 属性表示代理人。</span><span class="sxs-lookup"><span data-stu-id="13674-p127">The [`from`](#from-emailaddressdetailsjavascriptapioutlook13officeemailaddressdetails) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="13674-454">`sender` 属性中 `EmailAddressDetails` 对象的 `recipientType` 属性为 `undefined`。</span><span class="sxs-lookup"><span data-stu-id="13674-454">Note: The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="13674-455">类型：</span><span class="sxs-lookup"><span data-stu-id="13674-455">Type:</span></span>

*   [<span data-ttu-id="13674-456">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="13674-456">EmailAddressDetails</span></span>](/javascript/api/outlook_1_3/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="13674-457">要求</span><span class="sxs-lookup"><span data-stu-id="13674-457">Requirements</span></span>

|<span data-ttu-id="13674-458">要求</span><span class="sxs-lookup"><span data-stu-id="13674-458">Requirement</span></span>| <span data-ttu-id="13674-459">值</span><span class="sxs-lookup"><span data-stu-id="13674-459">Value</span></span>|
|---|---|
|[<span data-ttu-id="13674-460">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="13674-460">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="13674-461">1.0</span><span class="sxs-lookup"><span data-stu-id="13674-461">1.0</span></span>|
|[<span data-ttu-id="13674-462">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="13674-462">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="13674-463">ReadItem</span><span class="sxs-lookup"><span data-stu-id="13674-463">ReadItem</span></span>|
|[<span data-ttu-id="13674-464">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="13674-464">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="13674-465">阅读</span><span class="sxs-lookup"><span data-stu-id="13674-465">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="13674-466">示例</span><span class="sxs-lookup"><span data-stu-id="13674-466">Example</span></span>

```js
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
```

####  <a name="start-datetimejavascriptapioutlook13officetime"></a><span data-ttu-id="13674-467">start :Date|[Time](/javascript/api/outlook_1_3/office.time)</span><span class="sxs-lookup"><span data-stu-id="13674-467">start :Date|[Time](/javascript/api/outlook_1_3/office.time)</span></span>

<span data-ttu-id="13674-468">获取或设置约会开始的日期和时间。</span><span class="sxs-lookup"><span data-stu-id="13674-468">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="13674-p128">将 `start` 属性表示为协调世界时 (UTC) 的日期和时间值。可使用 [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook13officelocalclienttime) 方法将该值转换为客户端的本地日期和时间。</span><span class="sxs-lookup"><span data-stu-id="13674-p128">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook13officelocalclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="13674-471">阅读模式</span><span class="sxs-lookup"><span data-stu-id="13674-471">Read mode</span></span>

<span data-ttu-id="13674-472">`start` 属性返回 `Date` 对象。</span><span class="sxs-lookup"><span data-stu-id="13674-472">The `start` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="13674-473">撰写模式</span><span class="sxs-lookup"><span data-stu-id="13674-473">Compose mode</span></span>

<span data-ttu-id="13674-474">`start` 属性返回 `Time` 对象。</span><span class="sxs-lookup"><span data-stu-id="13674-474">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="13674-475">使用 [`Time.setAsync`](/javascript/api/outlook_1_3/office.time#setasync-datetime--options--callback-) 方法设置开始时间时，应使用 [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) 方法将客户端的本地时间转换为服务器的 UTC。</span><span class="sxs-lookup"><span data-stu-id="13674-475">When you use the [`Time.setAsync`](/javascript/api/outlook_1_3/office.time#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="13674-476">类型：</span><span class="sxs-lookup"><span data-stu-id="13674-476">Type:</span></span>

*   <span data-ttu-id="13674-477">Date | [Time](/javascript/api/outlook_1_3/office.time)</span><span class="sxs-lookup"><span data-stu-id="13674-477">Date | [Time](/javascript/api/outlook_1_3/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="13674-478">要求</span><span class="sxs-lookup"><span data-stu-id="13674-478">Requirements</span></span>

|<span data-ttu-id="13674-479">要求</span><span class="sxs-lookup"><span data-stu-id="13674-479">Requirement</span></span>| <span data-ttu-id="13674-480">值</span><span class="sxs-lookup"><span data-stu-id="13674-480">Value</span></span>|
|---|---|
|[<span data-ttu-id="13674-481">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="13674-481">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="13674-482">1.0</span><span class="sxs-lookup"><span data-stu-id="13674-482">1.0</span></span>|
|[<span data-ttu-id="13674-483">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="13674-483">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="13674-484">ReadItem</span><span class="sxs-lookup"><span data-stu-id="13674-484">ReadItem</span></span>|
|[<span data-ttu-id="13674-485">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="13674-485">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="13674-486">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="13674-486">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="13674-487">示例</span><span class="sxs-lookup"><span data-stu-id="13674-487">Example</span></span>

<span data-ttu-id="13674-488">以下示例通过使用 `Time` 对象的 [`setAsync`](/javascript/api/outlook_1_3/office.time#setasync-datetime--options--callback-) 方法，设置撰写模式下约会的开始时间。</span><span class="sxs-lookup"><span data-stu-id="13674-488">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_3/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

####  <a name="subject-stringsubjectjavascriptapioutlook13officesubject"></a><span data-ttu-id="13674-489">subject :String|[Subject](/javascript/api/outlook_1_3/office.subject)</span><span class="sxs-lookup"><span data-stu-id="13674-489">subject :String|[Subject](/javascript/api/outlook_1_3/office.subject)</span></span>

<span data-ttu-id="13674-490">获取或设置显示在项目的主题字段中的说明。</span><span class="sxs-lookup"><span data-stu-id="13674-490">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="13674-491">`subject` 属性获取或设置由电子邮件服务器发送项目时的整个主题。</span><span class="sxs-lookup"><span data-stu-id="13674-491">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="13674-492">阅读模式</span><span class="sxs-lookup"><span data-stu-id="13674-492">Read mode</span></span>

<span data-ttu-id="13674-p129">`subject` 属性返回一个字符串。使用 [`normalizedSubject`](#normalizedsubject-string) 属性获取不带任何前导前缀（如 `RE:` 和 `FW:`）的主题。</span><span class="sxs-lookup"><span data-stu-id="13674-p129">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

```js
var subject = Office.context.mailbox.item.subject;
```

##### <a name="compose-mode"></a><span data-ttu-id="13674-495">撰写模式</span><span class="sxs-lookup"><span data-stu-id="13674-495">Compose mode</span></span>

<span data-ttu-id="13674-496">`subject` 属性返回一个 `Subject` 对象，该对象提供用于获取和设置主题的方法。</span><span class="sxs-lookup"><span data-stu-id="13674-496">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```js
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="13674-497">类型：</span><span class="sxs-lookup"><span data-stu-id="13674-497">Type:</span></span>

*   <span data-ttu-id="13674-498">String | [Subject](/javascript/api/outlook_1_3/office.subject)</span><span class="sxs-lookup"><span data-stu-id="13674-498">String | [Subject](/javascript/api/outlook_1_3/office.subject)</span></span>

##### <a name="requirements"></a><span data-ttu-id="13674-499">要求</span><span class="sxs-lookup"><span data-stu-id="13674-499">Requirements</span></span>

|<span data-ttu-id="13674-500">要求</span><span class="sxs-lookup"><span data-stu-id="13674-500">Requirement</span></span>| <span data-ttu-id="13674-501">值</span><span class="sxs-lookup"><span data-stu-id="13674-501">Value</span></span>|
|---|---|
|[<span data-ttu-id="13674-502">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="13674-502">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="13674-503">1.0</span><span class="sxs-lookup"><span data-stu-id="13674-503">1.0</span></span>|
|[<span data-ttu-id="13674-504">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="13674-504">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="13674-505">ReadItem</span><span class="sxs-lookup"><span data-stu-id="13674-505">ReadItem</span></span>|
|[<span data-ttu-id="13674-506">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="13674-506">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="13674-507">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="13674-507">Compose or read</span></span>|

####  <a name="to-arrayemailaddressdetailsjavascriptapioutlook13officeemailaddressdetailsrecipientsjavascriptapioutlook13officerecipients"></a><span data-ttu-id="13674-508">to :Array.<[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_3/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="13674-508">to :Array.<[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_3/office.recipients)</span></span>

<span data-ttu-id="13674-509">提供对邮件的**发送至**行上的收件人的访问权限。</span><span class="sxs-lookup"><span data-stu-id="13674-509">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="13674-510">对象的类型和访问级别取决于当前项目的模式。</span><span class="sxs-lookup"><span data-stu-id="13674-510">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="13674-511">读取模式</span><span class="sxs-lookup"><span data-stu-id="13674-511">Read mode</span></span>

<span data-ttu-id="13674-p131">`to` 属性返回包含邮件的**收件人**行上所列的每个收件人的 `EmailAddressDetails` 对象的数组。集合上限为 100 个成员。</span><span class="sxs-lookup"><span data-stu-id="13674-p131">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="13674-514">撰写模式</span><span class="sxs-lookup"><span data-stu-id="13674-514">Compose mode</span></span>

<span data-ttu-id="13674-515">`to` 属性返回一个 `Recipients` 对象，该对象提供用于获取或更新邮件的**发送至**行上收件人的方法。</span><span class="sxs-lookup"><span data-stu-id="13674-515">The to`to` property returns a Recipients`Recipients` object that provides methods for manipulating the recipients on the **To** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="13674-516">类型：</span><span class="sxs-lookup"><span data-stu-id="13674-516">Type:</span></span>

*   <span data-ttu-id="13674-517">Array.<[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_3/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="13674-517">Array.<[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_3/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="13674-518">要求</span><span class="sxs-lookup"><span data-stu-id="13674-518">Requirements</span></span>

|<span data-ttu-id="13674-519">要求</span><span class="sxs-lookup"><span data-stu-id="13674-519">Requirement</span></span>| <span data-ttu-id="13674-520">值</span><span class="sxs-lookup"><span data-stu-id="13674-520">Value</span></span>|
|---|---|
|[<span data-ttu-id="13674-521">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="13674-521">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="13674-522">1.0</span><span class="sxs-lookup"><span data-stu-id="13674-522">1.0</span></span>|
|[<span data-ttu-id="13674-523">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="13674-523">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="13674-524">ReadItem</span><span class="sxs-lookup"><span data-stu-id="13674-524">ReadItem</span></span>|
|[<span data-ttu-id="13674-525">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="13674-525">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="13674-526">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="13674-526">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="13674-527">示例</span><span class="sxs-lookup"><span data-stu-id="13674-527">Example</span></span>

```js
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

### <a name="methods"></a><span data-ttu-id="13674-528">方法</span><span class="sxs-lookup"><span data-stu-id="13674-528">Methods</span></span>

####  <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="13674-529">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="13674-529">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="13674-530">将文件作为附件添加到邮件或约会。</span><span class="sxs-lookup"><span data-stu-id="13674-530">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="13674-531">`addFileAttachmentAsync` 方法在指定的 URI 上载文件并将其附加到撰写窗体中的项目。</span><span class="sxs-lookup"><span data-stu-id="13674-531">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="13674-532">你随后可以将该标识符与 [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) 方法一同使用，以删除同一个会话中的附件。</span><span class="sxs-lookup"><span data-stu-id="13674-532">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="13674-533">参数：</span><span class="sxs-lookup"><span data-stu-id="13674-533">Parameters:</span></span>

|<span data-ttu-id="13674-534">名称</span><span class="sxs-lookup"><span data-stu-id="13674-534">Name</span></span>| <span data-ttu-id="13674-535">类型</span><span class="sxs-lookup"><span data-stu-id="13674-535">Type</span></span>| <span data-ttu-id="13674-536">属性</span><span class="sxs-lookup"><span data-stu-id="13674-536">Attributes</span></span>| <span data-ttu-id="13674-537">说明</span><span class="sxs-lookup"><span data-stu-id="13674-537">Description</span></span>|
|---|---|---|---|
|`uri`| <span data-ttu-id="13674-538">字符串</span><span class="sxs-lookup"><span data-stu-id="13674-538">String</span></span>||<span data-ttu-id="13674-p132">提供附加到邮件或约会的文件的位置的 URI。最大长度为 2048 个字符。</span><span class="sxs-lookup"><span data-stu-id="13674-p132">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="13674-541">字符串</span><span class="sxs-lookup"><span data-stu-id="13674-541">String</span></span>||<span data-ttu-id="13674-p133">在附件上载过程中显示的附件名称。最大长度为 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="13674-p133">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="13674-544">Object</span><span class="sxs-lookup"><span data-stu-id="13674-544">Object</span></span>| <span data-ttu-id="13674-545">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="13674-545">&lt;optional&gt;</span></span>|<span data-ttu-id="13674-546">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="13674-546">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="13674-547">对象</span><span class="sxs-lookup"><span data-stu-id="13674-547">Object</span></span>| <span data-ttu-id="13674-548">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="13674-548">&lt;optional&gt;</span></span>|<span data-ttu-id="13674-549">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="13674-549">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="13674-550">函数</span><span class="sxs-lookup"><span data-stu-id="13674-550">function</span></span>| <span data-ttu-id="13674-551">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="13674-551">&lt;optional&gt;</span></span>|<span data-ttu-id="13674-552">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="13674-552">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="13674-553">如果成功，附件标识符将在 `asyncResult.value` 属性中提供。</span><span class="sxs-lookup"><span data-stu-id="13674-553">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="13674-554">如果上传附件失败，`asyncResult` 对象将包含一个提供错误说明的 `Error` 对象。</span><span class="sxs-lookup"><span data-stu-id="13674-554">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="13674-555">错误</span><span class="sxs-lookup"><span data-stu-id="13674-555">Errors</span></span>

| <span data-ttu-id="13674-556">错误代码</span><span class="sxs-lookup"><span data-stu-id="13674-556">Error code</span></span> | <span data-ttu-id="13674-557">说明</span><span class="sxs-lookup"><span data-stu-id="13674-557">Description</span></span> |
|------------|-------------|
| `AttachmentSizeExceeded` | <span data-ttu-id="13674-558">附件大小超过了允许的大小。</span><span class="sxs-lookup"><span data-stu-id="13674-558">The attachment is larger than allowed.</span></span> |
| `FileTypeNotSupported` | <span data-ttu-id="13674-559">该附件的扩展名不是允许的扩展名。</span><span class="sxs-lookup"><span data-stu-id="13674-559">The attachment has an extension that is not allowed.</span></span> |
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="13674-560">邮件或约会具有的附件过多。</span><span class="sxs-lookup"><span data-stu-id="13674-560">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="13674-561">要求</span><span class="sxs-lookup"><span data-stu-id="13674-561">Requirements</span></span>

|<span data-ttu-id="13674-562">要求</span><span class="sxs-lookup"><span data-stu-id="13674-562">Requirement</span></span>| <span data-ttu-id="13674-563">值</span><span class="sxs-lookup"><span data-stu-id="13674-563">Value</span></span>|
|---|---|
|[<span data-ttu-id="13674-564">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="13674-564">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="13674-565">1.1</span><span class="sxs-lookup"><span data-stu-id="13674-565">1.1</span></span>|
|[<span data-ttu-id="13674-566">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="13674-566">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="13674-567">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="13674-567">ReadWriteItem</span></span>|
|[<span data-ttu-id="13674-568">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="13674-568">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="13674-569">撰写</span><span class="sxs-lookup"><span data-stu-id="13674-569">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="13674-570">示例</span><span class="sxs-lookup"><span data-stu-id="13674-570">Example</span></span>

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

####  <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="13674-571">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="13674-571">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="13674-572">将 Exchange 项目（如邮件）作为附件添加到邮件或约会。</span><span class="sxs-lookup"><span data-stu-id="13674-572">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="13674-p134">`addItemAttachmentAsync` 方法将包含指定 Exchange 标识符的项目附加到撰写窗体中的项目。如果指定一个回调方法，此方法使用 `asyncResult` 参数调用，该参数包含一个附件标识符或代码，指示附加项目过程中出现的任何错误。可以使用 `options` 参数将状态信息传递给回调方法（如果需要）。</span><span class="sxs-lookup"><span data-stu-id="13674-p134">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="13674-576">随后可以将该标识符与 [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) 方法一同使用，以删除同一个会话中的附件。</span><span class="sxs-lookup"><span data-stu-id="13674-576">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="13674-577">如果 Office 外接程序在 Outlook Web App 中运行，则 `addItemAttachmentAsync` 方法可以将项目附加到项目（正在编辑的项目除外）中；然而，不支持也不建议这样做。</span><span class="sxs-lookup"><span data-stu-id="13674-577">If your Office add-in is running in Outlook Web App, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="13674-578">参数：</span><span class="sxs-lookup"><span data-stu-id="13674-578">Parameters:</span></span>

|<span data-ttu-id="13674-579">名称</span><span class="sxs-lookup"><span data-stu-id="13674-579">Name</span></span>| <span data-ttu-id="13674-580">类型</span><span class="sxs-lookup"><span data-stu-id="13674-580">Type</span></span>| <span data-ttu-id="13674-581">属性</span><span class="sxs-lookup"><span data-stu-id="13674-581">Attributes</span></span>| <span data-ttu-id="13674-582">说明</span><span class="sxs-lookup"><span data-stu-id="13674-582">Description</span></span>|
|---|---|---|---|
|`itemId`| <span data-ttu-id="13674-583">字符串</span><span class="sxs-lookup"><span data-stu-id="13674-583">String</span></span>||<span data-ttu-id="13674-p135">要附加的项目的 Exchange 标识符。最大长度为 100 个字符。</span><span class="sxs-lookup"><span data-stu-id="13674-p135">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="13674-586">字符串</span><span class="sxs-lookup"><span data-stu-id="13674-586">String</span></span>||<span data-ttu-id="13674-p136">要附加的项目的主题。最大长度为 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="13674-p136">The sujbect of the item to be attached. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="13674-589">Object</span><span class="sxs-lookup"><span data-stu-id="13674-589">Object</span></span>| <span data-ttu-id="13674-590">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="13674-590">&lt;optional&gt;</span></span>|<span data-ttu-id="13674-591">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="13674-591">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="13674-592">对象</span><span class="sxs-lookup"><span data-stu-id="13674-592">Object</span></span>| <span data-ttu-id="13674-593">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="13674-593">&lt;optional&gt;</span></span>|<span data-ttu-id="13674-594">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="13674-594">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="13674-595">函数</span><span class="sxs-lookup"><span data-stu-id="13674-595">function</span></span>| <span data-ttu-id="13674-596">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="13674-596">&lt;optional&gt;</span></span>|<span data-ttu-id="13674-597">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="13674-597">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="13674-598">如果成功，附件标识符将在 `asyncResult.value` 属性中提供。</span><span class="sxs-lookup"><span data-stu-id="13674-598">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="13674-599">如果添加附件失败，`asyncResult` 对象将包含一个提供错误说明的 `Error` 对象。</span><span class="sxs-lookup"><span data-stu-id="13674-599">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="13674-600">错误</span><span class="sxs-lookup"><span data-stu-id="13674-600">Errors</span></span>

| <span data-ttu-id="13674-601">错误代码</span><span class="sxs-lookup"><span data-stu-id="13674-601">Error code</span></span> | <span data-ttu-id="13674-602">说明</span><span class="sxs-lookup"><span data-stu-id="13674-602">Description</span></span> |
|------------|-------------|
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="13674-603">邮件或约会具有的附件过多。</span><span class="sxs-lookup"><span data-stu-id="13674-603">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="13674-604">要求</span><span class="sxs-lookup"><span data-stu-id="13674-604">Requirements</span></span>

|<span data-ttu-id="13674-605">要求</span><span class="sxs-lookup"><span data-stu-id="13674-605">Requirement</span></span>| <span data-ttu-id="13674-606">值</span><span class="sxs-lookup"><span data-stu-id="13674-606">Value</span></span>|
|---|---|
|[<span data-ttu-id="13674-607">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="13674-607">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="13674-608">1.1</span><span class="sxs-lookup"><span data-stu-id="13674-608">1.1</span></span>|
|[<span data-ttu-id="13674-609">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="13674-609">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="13674-610">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="13674-610">ReadWriteItem</span></span>|
|[<span data-ttu-id="13674-611">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="13674-611">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="13674-612">撰写</span><span class="sxs-lookup"><span data-stu-id="13674-612">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="13674-613">示例</span><span class="sxs-lookup"><span data-stu-id="13674-613">Example</span></span>

<span data-ttu-id="13674-614">以下示例将现有的 Outlook 项目添加为名为 `My Attachment` 的附件。</span><span class="sxs-lookup"><span data-stu-id="13674-614">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

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

####  <a name="close"></a><span data-ttu-id="13674-615">close()</span><span class="sxs-lookup"><span data-stu-id="13674-615">close()</span></span>

<span data-ttu-id="13674-616">关闭当前正在撰写的项目。</span><span class="sxs-lookup"><span data-stu-id="13674-616">Closes the current item that is being composed.</span></span>

<span data-ttu-id="13674-p137">
            \`close\` 方法的行为取决于要撰写的项目的当前状态。如果项目具有未保存的更改，客户端将提示用户保存、放弃或取消关闭操作。</span><span class="sxs-lookup"><span data-stu-id="13674-p137">The behavior of the `close` method depends on the current state of the item being composed. If the item has unsaved changes, the client prompts the user to save, discard, or cancel the close action.</span></span>

> [!NOTE]
> <span data-ttu-id="13674-619">在 Outlook 网页版中，如果该项目是约会并且之前已使用 `saveAsync` 保存，则即使自上次保存项目后未发生任何更改，也会提示用户保存、放弃或取消。</span><span class="sxs-lookup"><span data-stu-id="13674-619">In Outlook on the web, if the item is an appointment and it has previously been saved using `saveAsync`, the user is prompted to save, discard, or cancel even if no changes have occurred since the item was last saved.</span></span>

<span data-ttu-id="13674-620">在 Outlook 桌面版客户端中，如果邮件是内联答复，`close` 方法不起作用。</span><span class="sxs-lookup"><span data-stu-id="13674-620">In the Outlook desktop client, if the message is an inline reply, the `close` method has no effect.</span></span>

##### <a name="requirements"></a><span data-ttu-id="13674-621">要求</span><span class="sxs-lookup"><span data-stu-id="13674-621">Requirements</span></span>

|<span data-ttu-id="13674-622">要求</span><span class="sxs-lookup"><span data-stu-id="13674-622">Requirement</span></span>| <span data-ttu-id="13674-623">值</span><span class="sxs-lookup"><span data-stu-id="13674-623">Value</span></span>|
|---|---|
|[<span data-ttu-id="13674-624">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="13674-624">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="13674-625">1.3</span><span class="sxs-lookup"><span data-stu-id="13674-625">1.3</span></span>|
|[<span data-ttu-id="13674-626">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="13674-626">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="13674-627">受限</span><span class="sxs-lookup"><span data-stu-id="13674-627">Restricted</span></span>|
|[<span data-ttu-id="13674-628">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="13674-628">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="13674-629">撰写</span><span class="sxs-lookup"><span data-stu-id="13674-629">Compose</span></span>|

#### <a name="displayreplyallformformdata"></a><span data-ttu-id="13674-630">displayReplyAllForm(formData)</span><span class="sxs-lookup"><span data-stu-id="13674-630">displayReplyAllForm(formData)</span></span>

<span data-ttu-id="13674-631">显示答复窗体，其中包括所选邮件的发件人和所有收件人或所选约会的组织者和所有与会者。</span><span class="sxs-lookup"><span data-stu-id="13674-631">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="13674-632">在 Outlook for iOS 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="13674-632">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="13674-633">在 Outlook Web App 中，答复窗体显示为包含 3 列视图的弹出式窗体以及包含 2 列或 1 列视图的弹出式窗体。</span><span class="sxs-lookup"><span data-stu-id="13674-633">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="13674-634">如果任意字符串参数超出其限制，`displayReplyAllForm` 将引发异常。</span><span class="sxs-lookup"><span data-stu-id="13674-634">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="13674-p138">当在 `formData.attachments` 参数中指定附件时，Outlook 和 Outlook Web App 尝试下载所有附件并将其附加到答复窗体。如果无法添加任何附件，则在窗体 UI 中显示错误。如果这不可能，则不引发错误消息。</span><span class="sxs-lookup"><span data-stu-id="13674-p138">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="13674-638">参数：</span><span class="sxs-lookup"><span data-stu-id="13674-638">Parameters:</span></span>

|<span data-ttu-id="13674-639">名称</span><span class="sxs-lookup"><span data-stu-id="13674-639">Name</span></span>| <span data-ttu-id="13674-640">类型</span><span class="sxs-lookup"><span data-stu-id="13674-640">Type</span></span>| <span data-ttu-id="13674-641">说明</span><span class="sxs-lookup"><span data-stu-id="13674-641">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="13674-642">字符串 &#124; 对象</span><span class="sxs-lookup"><span data-stu-id="13674-642">String &#124; Object</span></span>| |<span data-ttu-id="13674-p139">一个包含文本和 HTML 且表示答复窗体的正文的字符串。字符串限制为 32 KB。</span><span class="sxs-lookup"><span data-stu-id="13674-p139">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="13674-645">**或**</span><span class="sxs-lookup"><span data-stu-id="13674-645">**OR**</span></span><br/><span data-ttu-id="13674-p140">包含正文或附件数据和回调函数的对象。对象定义如下。</span><span class="sxs-lookup"><span data-stu-id="13674-p140">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="13674-648">字符串</span><span class="sxs-lookup"><span data-stu-id="13674-648">String</span></span> | <span data-ttu-id="13674-649">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="13674-649">&lt;optional&gt;</span></span> | <span data-ttu-id="13674-p141">一个包含文本和 HTML 且表示答复窗体的正文的字符串。字符串限制为 32 KB。</span><span class="sxs-lookup"><span data-stu-id="13674-p141">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="13674-652">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="13674-652">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="13674-653">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="13674-653">&lt;optional&gt;</span></span> | <span data-ttu-id="13674-654">JSON 对象（是文件或项目附件）的数组。</span><span class="sxs-lookup"><span data-stu-id="13674-654">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="13674-655">字符串</span><span class="sxs-lookup"><span data-stu-id="13674-655">String</span></span> | | <span data-ttu-id="13674-p142">指示附件的类型。必须是文件附件的 `file` 或项目附件的 `item`。</span><span class="sxs-lookup"><span data-stu-id="13674-p142">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="13674-658">字符串</span><span class="sxs-lookup"><span data-stu-id="13674-658">String</span></span> | | <span data-ttu-id="13674-659">一个包含附件的名称的字符串，最多包含 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="13674-659">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="13674-660">String</span><span class="sxs-lookup"><span data-stu-id="13674-660">String</span></span> | | <span data-ttu-id="13674-p143">仅在将 `type` 设置为 `file` 时使用。文件的位置的 URI。</span><span class="sxs-lookup"><span data-stu-id="13674-p143">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="13674-663">字符串</span><span class="sxs-lookup"><span data-stu-id="13674-663">String</span></span> | | <span data-ttu-id="13674-p144">仅在将 `type` 设置为 `item` 时使用。附件的 EWS 项目 ID。字符串最长为 100 个字符。</span><span class="sxs-lookup"><span data-stu-id="13674-p144">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="13674-667">函数</span><span class="sxs-lookup"><span data-stu-id="13674-667">function</span></span> | <span data-ttu-id="13674-668">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="13674-668">&lt;optional&gt;</span></span> | <span data-ttu-id="13674-669">方法完成后，使用单个参数 `asyncResult`（一个 [AsyncResult](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="13674-669">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="13674-670">要求</span><span class="sxs-lookup"><span data-stu-id="13674-670">Requirements</span></span>

|<span data-ttu-id="13674-671">要求</span><span class="sxs-lookup"><span data-stu-id="13674-671">Requirement</span></span>| <span data-ttu-id="13674-672">值</span><span class="sxs-lookup"><span data-stu-id="13674-672">Value</span></span>|
|---|---|
|[<span data-ttu-id="13674-673">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="13674-673">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="13674-674">1.0</span><span class="sxs-lookup"><span data-stu-id="13674-674">1.0</span></span>|
|[<span data-ttu-id="13674-675">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="13674-675">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="13674-676">ReadItem</span><span class="sxs-lookup"><span data-stu-id="13674-676">ReadItem</span></span>|
|[<span data-ttu-id="13674-677">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="13674-677">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="13674-678">阅读</span><span class="sxs-lookup"><span data-stu-id="13674-678">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="13674-679">示例</span><span class="sxs-lookup"><span data-stu-id="13674-679">Examples</span></span>

<span data-ttu-id="13674-680">以下代码将一个字符串传递到 `displayReplyAllForm` 函数。</span><span class="sxs-lookup"><span data-stu-id="13674-680">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="13674-681">使用空白正文答复。</span><span class="sxs-lookup"><span data-stu-id="13674-681">Reply with an empty body.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="13674-682">仅使用正文答复。</span><span class="sxs-lookup"><span data-stu-id="13674-682">Reply with just a body.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="13674-683">使用正文和文件附件答复。</span><span class="sxs-lookup"><span data-stu-id="13674-683">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="13674-684">使用正文和项目附件答复。</span><span class="sxs-lookup"><span data-stu-id="13674-684">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="13674-685">使用正文、文件附件、项目附件和回调答复。</span><span class="sxs-lookup"><span data-stu-id="13674-685">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="displayreplyformformdata"></a><span data-ttu-id="13674-686">displayReplyForm(formData)</span><span class="sxs-lookup"><span data-stu-id="13674-686">displayReplyForm(formData)</span></span>

<span data-ttu-id="13674-687">显示答复窗体，其中仅包括所选邮件的发件人或所选约会的组织者。</span><span class="sxs-lookup"><span data-stu-id="13674-687">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="13674-688">在 Outlook for iOS 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="13674-688">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="13674-689">在 Outlook Web App 中，答复窗体显示为包含 3 列视图的弹出式窗体以及包含 2 列或 1 列视图的弹出式窗体。</span><span class="sxs-lookup"><span data-stu-id="13674-689">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="13674-690">如果任意字符串参数超出其限制，`displayReplyForm` 将引发异常。</span><span class="sxs-lookup"><span data-stu-id="13674-690">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="13674-p145">当在 `formData.attachments` 参数中指定附件时，Outlook 和 Outlook Web App 尝试下载所有附件并将其附加到答复窗体。如果无法添加任何附件，则在窗体 UI 中显示错误。如果这不可能，则不引发错误消息。</span><span class="sxs-lookup"><span data-stu-id="13674-p145">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="13674-694">参数：</span><span class="sxs-lookup"><span data-stu-id="13674-694">Parameters:</span></span>

|<span data-ttu-id="13674-695">名称</span><span class="sxs-lookup"><span data-stu-id="13674-695">Name</span></span>| <span data-ttu-id="13674-696">类型</span><span class="sxs-lookup"><span data-stu-id="13674-696">Type</span></span>| <span data-ttu-id="13674-697">说明</span><span class="sxs-lookup"><span data-stu-id="13674-697">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="13674-698">字符串 &#124; 对象</span><span class="sxs-lookup"><span data-stu-id="13674-698">String &#124; Object</span></span>| | <span data-ttu-id="13674-p146">一个包含文本和 HTML 且表示答复窗体的正文的字符串。字符串限制为 32 KB。</span><span class="sxs-lookup"><span data-stu-id="13674-p146">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="13674-701">**或**</span><span class="sxs-lookup"><span data-stu-id="13674-701">**OR**</span></span><br/><span data-ttu-id="13674-p147">包含正文或附件数据和回调函数的对象。对象定义如下。</span><span class="sxs-lookup"><span data-stu-id="13674-p147">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="13674-704">字符串</span><span class="sxs-lookup"><span data-stu-id="13674-704">String</span></span> | <span data-ttu-id="13674-705">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="13674-705">&lt;optional&gt;</span></span> | <span data-ttu-id="13674-p148">一个包含文本和 HTML 且表示答复窗体的正文的字符串。字符串限制为 32 KB。</span><span class="sxs-lookup"><span data-stu-id="13674-p148">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="13674-708">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="13674-708">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="13674-709">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="13674-709">&lt;optional&gt;</span></span> | <span data-ttu-id="13674-710">JSON 对象（是文件或项目附件）的数组。</span><span class="sxs-lookup"><span data-stu-id="13674-710">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="13674-711">字符串</span><span class="sxs-lookup"><span data-stu-id="13674-711">String</span></span> | | <span data-ttu-id="13674-p149">指示附件的类型。必须是文件附件的 `file` 或项目附件的 `item`。</span><span class="sxs-lookup"><span data-stu-id="13674-p149">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="13674-714">字符串</span><span class="sxs-lookup"><span data-stu-id="13674-714">String</span></span> | | <span data-ttu-id="13674-715">一个包含附件的名称的字符串，最多包含 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="13674-715">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="13674-716">String</span><span class="sxs-lookup"><span data-stu-id="13674-716">String</span></span> | | <span data-ttu-id="13674-p150">仅在将 `type` 设置为 `file` 时使用。文件的位置的 URI。</span><span class="sxs-lookup"><span data-stu-id="13674-p150">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="13674-719">字符串</span><span class="sxs-lookup"><span data-stu-id="13674-719">String</span></span> | | <span data-ttu-id="13674-p151">仅在将 `type` 设置为 `item` 时使用。附件的 EWS 项目 ID。字符串最长为 100 个字符。</span><span class="sxs-lookup"><span data-stu-id="13674-p151">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="13674-723">函数</span><span class="sxs-lookup"><span data-stu-id="13674-723">function</span></span> | <span data-ttu-id="13674-724">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="13674-724">&lt;optional&gt;</span></span> | <span data-ttu-id="13674-725">方法完成后，使用单个参数 `asyncResult`（一个 [AsyncResult](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="13674-725">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="13674-726">要求</span><span class="sxs-lookup"><span data-stu-id="13674-726">Requirements</span></span>

|<span data-ttu-id="13674-727">要求</span><span class="sxs-lookup"><span data-stu-id="13674-727">Requirement</span></span>| <span data-ttu-id="13674-728">值</span><span class="sxs-lookup"><span data-stu-id="13674-728">Value</span></span>|
|---|---|
|[<span data-ttu-id="13674-729">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="13674-729">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="13674-730">1.0</span><span class="sxs-lookup"><span data-stu-id="13674-730">1.0</span></span>|
|[<span data-ttu-id="13674-731">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="13674-731">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="13674-732">ReadItem</span><span class="sxs-lookup"><span data-stu-id="13674-732">ReadItem</span></span>|
|[<span data-ttu-id="13674-733">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="13674-733">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="13674-734">阅读</span><span class="sxs-lookup"><span data-stu-id="13674-734">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="13674-735">示例</span><span class="sxs-lookup"><span data-stu-id="13674-735">Examples</span></span>

<span data-ttu-id="13674-736">以下代码将一个字符串传递到 `displayReplyForm` 函数。</span><span class="sxs-lookup"><span data-stu-id="13674-736">The following code passes a string to the `displayReplyForm` function.</span></span>

```js
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="13674-737">使用空白正文答复。</span><span class="sxs-lookup"><span data-stu-id="13674-737">Reply with an empty body.</span></span>

```js
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="13674-738">仅使用正文答复。</span><span class="sxs-lookup"><span data-stu-id="13674-738">Reply with just a body.</span></span>

```js
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="13674-739">使用正文和文件附件答复。</span><span class="sxs-lookup"><span data-stu-id="13674-739">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="13674-740">使用正文和项目附件答复。</span><span class="sxs-lookup"><span data-stu-id="13674-740">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="13674-741">使用正文、文件附件、项目附件和回调答复。</span><span class="sxs-lookup"><span data-stu-id="13674-741">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="getentities--entitiesjavascriptapioutlook13officeentities"></a><span data-ttu-id="13674-742">getEntities() → {[Entities](/javascript/api/outlook_1_3/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="13674-742">getEntities() → {[Entities](/javascript/api/outlook_1_3/office.entities)}</span></span>

<span data-ttu-id="13674-743">获取在所选项目的正文中找到的实体。</span><span class="sxs-lookup"><span data-stu-id="13674-743">Gets the entities found in the selected item.</span></span>

> [!NOTE]
> <span data-ttu-id="13674-744">在 Outlook for iOS 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="13674-744">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="13674-745">要求</span><span class="sxs-lookup"><span data-stu-id="13674-745">Requirements</span></span>

|<span data-ttu-id="13674-746">要求</span><span class="sxs-lookup"><span data-stu-id="13674-746">Requirement</span></span>| <span data-ttu-id="13674-747">值</span><span class="sxs-lookup"><span data-stu-id="13674-747">Value</span></span>|
|---|---|
|[<span data-ttu-id="13674-748">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="13674-748">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="13674-749">1.0</span><span class="sxs-lookup"><span data-stu-id="13674-749">1.0</span></span>|
|[<span data-ttu-id="13674-750">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="13674-750">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="13674-751">ReadItem</span><span class="sxs-lookup"><span data-stu-id="13674-751">ReadItem</span></span>|
|[<span data-ttu-id="13674-752">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="13674-752">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="13674-753">阅读</span><span class="sxs-lookup"><span data-stu-id="13674-753">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="13674-754">返回：</span><span class="sxs-lookup"><span data-stu-id="13674-754">Returns:</span></span>

<span data-ttu-id="13674-755">类型：[Entities](/javascript/api/outlook_1_3/office.entities)</span><span class="sxs-lookup"><span data-stu-id="13674-755">Type: [Entities](/javascript/api/outlook_1_3/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="13674-756">示例</span><span class="sxs-lookup"><span data-stu-id="13674-756">Example</span></span>

<span data-ttu-id="13674-757">以下示例访问当前项目的正文中的联系人实体。</span><span class="sxs-lookup"><span data-stu-id="13674-757">The following example accesses the contacts entities on the current item.</span></span>

```js
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlook13officecontactmeetingsuggestionjavascriptapioutlook13officemeetingsuggestionphonenumberjavascriptapioutlook13officephonenumbertasksuggestionjavascriptapioutlook13officetasksuggestion"></a><span data-ttu-id="13674-758">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_3/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_3/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_3/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_3/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="13674-758">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_3/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_3/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_3/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_3/office.tasksuggestion))>}</span></span>

<span data-ttu-id="13674-759">获取所选项目的正文中找到的指定实体类型的所有实体的数组。</span><span class="sxs-lookup"><span data-stu-id="13674-759">Gets an array of all the entities of the specified entity type found in the selected item.</span></span>

> [!NOTE]
> <span data-ttu-id="13674-760">在 Outlook for iOS 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="13674-760">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="13674-761">参数：</span><span class="sxs-lookup"><span data-stu-id="13674-761">Parameters:</span></span>

|<span data-ttu-id="13674-762">名称</span><span class="sxs-lookup"><span data-stu-id="13674-762">Name</span></span>| <span data-ttu-id="13674-763">类型</span><span class="sxs-lookup"><span data-stu-id="13674-763">Type</span></span>| <span data-ttu-id="13674-764">说明</span><span class="sxs-lookup"><span data-stu-id="13674-764">Description</span></span>|
|---|---|---|
|`entityType`| [<span data-ttu-id="13674-765">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="13674-765">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook_1_3/office.mailboxenums.entitytype)|<span data-ttu-id="13674-766">EntityType 枚举值之一。</span><span class="sxs-lookup"><span data-stu-id="13674-766">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="13674-767">要求</span><span class="sxs-lookup"><span data-stu-id="13674-767">Requirements</span></span>

|<span data-ttu-id="13674-768">要求</span><span class="sxs-lookup"><span data-stu-id="13674-768">Requirement</span></span>| <span data-ttu-id="13674-769">值</span><span class="sxs-lookup"><span data-stu-id="13674-769">Value</span></span>|
|---|---|
|[<span data-ttu-id="13674-770">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="13674-770">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="13674-771">1.0</span><span class="sxs-lookup"><span data-stu-id="13674-771">1.0</span></span>|
|[<span data-ttu-id="13674-772">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="13674-772">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="13674-773">受限</span><span class="sxs-lookup"><span data-stu-id="13674-773">Restricted</span></span>|
|[<span data-ttu-id="13674-774">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="13674-774">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="13674-775">阅读</span><span class="sxs-lookup"><span data-stu-id="13674-775">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="13674-776">返回：</span><span class="sxs-lookup"><span data-stu-id="13674-776">Returns:</span></span>

<span data-ttu-id="13674-777">如果在 `entityType` 中传递的值不是 `EntityType` 枚举的有效成员，该方法返回 null。</span><span class="sxs-lookup"><span data-stu-id="13674-777">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="13674-778">如果指定类型的任何实体都不存在于该项目的正文中，该方法将返回空数组。</span><span class="sxs-lookup"><span data-stu-id="13674-778">If no entities of the specified type are present on the item, the method returns an empty array.</span></span> <span data-ttu-id="13674-779">否则，返回的数组中对象的类型取决于 `entityType` 参数中请求实体的类型。</span><span class="sxs-lookup"><span data-stu-id="13674-779">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="13674-780">当使用此方法的最低权限级别**受限**时，某些实体类型需要 **ReadItem** 才能进行访问，如下表中所指定。</span><span class="sxs-lookup"><span data-stu-id="13674-780">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

| <span data-ttu-id="13674-781">`entityType` 的值</span><span class="sxs-lookup"><span data-stu-id="13674-781">Value of `entityType`</span></span> | <span data-ttu-id="13674-782">返回的数组中对象的类型</span><span class="sxs-lookup"><span data-stu-id="13674-782">Type of objects in returned array</span></span> | <span data-ttu-id="13674-783">所需权限级别</span><span class="sxs-lookup"><span data-stu-id="13674-783">Required Permission Level</span></span> |
| --- | --- | --- |
| `Address` | <span data-ttu-id="13674-784">String</span><span class="sxs-lookup"><span data-stu-id="13674-784">String</span></span> | <span data-ttu-id="13674-785">**受限**</span><span class="sxs-lookup"><span data-stu-id="13674-785">**Restricted**</span></span> |
| `Contact` | <span data-ttu-id="13674-786">Contact</span><span class="sxs-lookup"><span data-stu-id="13674-786">Contact</span></span> | <span data-ttu-id="13674-787">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="13674-787">**ReadItem**</span></span> |
| `EmailAddress` | <span data-ttu-id="13674-788">String</span><span class="sxs-lookup"><span data-stu-id="13674-788">String</span></span> | <span data-ttu-id="13674-789">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="13674-789">**ReadItem**</span></span> |
| `MeetingSuggestion` | <span data-ttu-id="13674-790">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="13674-790">MeetingSuggestion</span></span> | <span data-ttu-id="13674-791">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="13674-791">**ReadItem**</span></span> |
| `PhoneNumber` | <span data-ttu-id="13674-792">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="13674-792">PhoneNumber</span></span> | <span data-ttu-id="13674-793">**受限**</span><span class="sxs-lookup"><span data-stu-id="13674-793">**Restricted**</span></span> |
| `TaskSuggestion` | <span data-ttu-id="13674-794">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="13674-794">TaskSuggestion</span></span> | <span data-ttu-id="13674-795">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="13674-795">**ReadItem**</span></span> |
| `URL` | <span data-ttu-id="13674-796">String</span><span class="sxs-lookup"><span data-stu-id="13674-796">String</span></span> | <span data-ttu-id="13674-797">**受限**</span><span class="sxs-lookup"><span data-stu-id="13674-797">**Restricted**</span></span> |

<span data-ttu-id="13674-798">类型：Array.<(String|[Contact](/javascript/api/outlook_1_3/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_3/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_3/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_3/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="13674-798">Type: Array.<(String|[Contact](/javascript/api/outlook_1_3/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_3/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_3/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_3/office.tasksuggestion))></span></span>

##### <a name="example"></a><span data-ttu-id="13674-799">示例</span><span class="sxs-lookup"><span data-stu-id="13674-799">Example</span></span>

<span data-ttu-id="13674-800">以下示例显示了如何访问表示当前项目的正文中的邮政地址的字符串数组。</span><span class="sxs-lookup"><span data-stu-id="13674-800">The following example shows how to access an array of strings that represent postal addresses in the subject or body of the current item.</span></span>

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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlook13officecontactmeetingsuggestionjavascriptapioutlook13officemeetingsuggestionphonenumberjavascriptapioutlook13officephonenumbertasksuggestionjavascriptapioutlook13officetasksuggestion"></a><span data-ttu-id="13674-801">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_3/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_3/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_3/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_3/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="13674-801">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_3/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_3/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_3/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_3/office.tasksuggestion))>}</span></span>

<span data-ttu-id="13674-802">返回传递清单 XML 文件中定义的命名筛选器的所选项目中的已知实体。</span><span class="sxs-lookup"><span data-stu-id="13674-802">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="13674-803">在 Outlook for iOS 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="13674-803">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="13674-804">`getFilteredEntitiesByName` 方法返回匹配在具有指定 `FilterName` 元素值的清单 XML 文件中的 [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) 规则元素中定义的正则表达式的实体。</span><span class="sxs-lookup"><span data-stu-id="13674-804">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="13674-805">参数：</span><span class="sxs-lookup"><span data-stu-id="13674-805">Parameters:</span></span>

|<span data-ttu-id="13674-806">名称</span><span class="sxs-lookup"><span data-stu-id="13674-806">Name</span></span>| <span data-ttu-id="13674-807">类型</span><span class="sxs-lookup"><span data-stu-id="13674-807">Type</span></span>| <span data-ttu-id="13674-808">描述</span><span class="sxs-lookup"><span data-stu-id="13674-808">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="13674-809">字符串</span><span class="sxs-lookup"><span data-stu-id="13674-809">String</span></span>|<span data-ttu-id="13674-810">定义筛选器匹配的 `ItemHasKnownEntity` 规则元素的名称。</span><span class="sxs-lookup"><span data-stu-id="13674-810">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="13674-811">要求</span><span class="sxs-lookup"><span data-stu-id="13674-811">Requirements</span></span>

|<span data-ttu-id="13674-812">要求</span><span class="sxs-lookup"><span data-stu-id="13674-812">Requirement</span></span>| <span data-ttu-id="13674-813">值</span><span class="sxs-lookup"><span data-stu-id="13674-813">Value</span></span>|
|---|---|
|[<span data-ttu-id="13674-814">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="13674-814">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="13674-815">1.0</span><span class="sxs-lookup"><span data-stu-id="13674-815">1.0</span></span>|
|[<span data-ttu-id="13674-816">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="13674-816">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="13674-817">ReadItem</span><span class="sxs-lookup"><span data-stu-id="13674-817">ReadItem</span></span>|
|[<span data-ttu-id="13674-818">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="13674-818">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="13674-819">阅读</span><span class="sxs-lookup"><span data-stu-id="13674-819">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="13674-820">返回：</span><span class="sxs-lookup"><span data-stu-id="13674-820">Returns:</span></span>

<span data-ttu-id="13674-p153">如果具有匹配 `name` 参数的 `FilterName` 元素值的清单中没有任何 `ItemHasKnownEntity` 元素，则该方法返回 `null`。如果 `name` 参数匹配清单中的 `ItemHasKnownEntity` 元素，但在匹配的当前项目中没有实体，则该方法返回一个空数组。</span><span class="sxs-lookup"><span data-stu-id="13674-p153">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="13674-823">类型：Array.<(String|[Contact](/javascript/api/outlook_1_3/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_3/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_3/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_3/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="13674-823">Type: Array.<(String|[Contact](/javascript/api/outlook_1_3/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_3/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_3/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_3/office.tasksuggestion))></span></span>

#### <a name="getregexmatches--object"></a><span data-ttu-id="13674-824">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="13674-824">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="13674-825">返回所选项目中匹配在清单 XML 文件中定义的正则表达式的字符串值。</span><span class="sxs-lookup"><span data-stu-id="13674-825">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="13674-826">在 Outlook for iOS 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="13674-826">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="13674-p154">`getRegExMatches` 方法返回匹配在清单 XML 文件中的每个 `ItemHasRegularExpressionMatch` 或 `ItemHasKnownEntity` 规则元素中定义的正则表达式的字符串。对于 `ItemHasRegularExpressionMatch` 规则，匹配字符串必须发生在该规则指定的项目的属性中。`PropertyName` 简单类型定义支持的属性。</span><span class="sxs-lookup"><span data-stu-id="13674-p154">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="13674-830">例如，考虑一个外接程序清单具有以下 `Rule` 元素：</span><span class="sxs-lookup"><span data-stu-id="13674-830">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="13674-831">从 `getRegExMatches` 返回的对象应有两个属性：`fruits` 和 `veggies`。</span><span class="sxs-lookup"><span data-stu-id="13674-831">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="13674-p155">如果在项目的正文属性上指定 `ItemHasRegularExpressionMatch` 规则，则正则表达式应进一步筛选正文，不应尝试返回该项目的整个正文。使用正则表达式（如 `.*`）获取项目的整个正文并不总是返回预期的结果。而是使用 [`Body.getAsync`](/javascript/api/outlook_1_3/office.body#getasync-coerciontype--options--callback-) 方法检索整个正文。</span><span class="sxs-lookup"><span data-stu-id="13674-p155">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook_1_3/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="13674-835">Requirements</span><span class="sxs-lookup"><span data-stu-id="13674-835">Requirements</span></span>

|<span data-ttu-id="13674-836">要求</span><span class="sxs-lookup"><span data-stu-id="13674-836">Requirement</span></span>| <span data-ttu-id="13674-837">值</span><span class="sxs-lookup"><span data-stu-id="13674-837">Value</span></span>|
|---|---|
|[<span data-ttu-id="13674-838">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="13674-838">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="13674-839">1.0</span><span class="sxs-lookup"><span data-stu-id="13674-839">1.0</span></span>|
|[<span data-ttu-id="13674-840">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="13674-840">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="13674-841">ReadItem</span><span class="sxs-lookup"><span data-stu-id="13674-841">ReadItem</span></span>|
|[<span data-ttu-id="13674-842">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="13674-842">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="13674-843">阅读</span><span class="sxs-lookup"><span data-stu-id="13674-843">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="13674-844">返回：</span><span class="sxs-lookup"><span data-stu-id="13674-844">Returns:</span></span>

<span data-ttu-id="13674-p156">一个包含匹配在清单 XML 文件中定义的正则表达式的字符串数组的对象。每个数组的名称等于匹配 `ItemHasRegularExpressionMatch` 规则的 `RegExName` 属性或匹配 `ItemHasKnownEntity` 规则的 `FilterName` 属性的相应值。</span><span class="sxs-lookup"><span data-stu-id="13674-p156">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<dl class="param-type"><span data-ttu-id="13674-847">

<dt>
类型</dt>


</span><span class="sxs-lookup"><span data-stu-id="13674-847">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="13674-848">对象</span><span class="sxs-lookup"><span data-stu-id="13674-848">Object</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="13674-849">示例</span><span class="sxs-lookup"><span data-stu-id="13674-849">Example</span></span>

<span data-ttu-id="13674-850">以下示例显示了如何访问正则表达式 <rule> 元素 `fruits` 和 `veggies` 的匹配项的数组，这些元素在清单中指定。</rule></span><span class="sxs-lookup"><span data-stu-id="13674-850">The following example shows how to access the array of matches for the regular expression <rule>elements `fruits` and `veggies`, which are specified in the manifest.</rule></span></span>

```js
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veges = allMatches.veggies;
```

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="13674-851">getRegExMatchesByName(name) → (nullable) {Array.<String>}</span><span class="sxs-lookup"><span data-stu-id="13674-851">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="13674-852">返回匹配在清单 XML 文件中定义的命名正则表达式的所选项目中的字符串值。</span><span class="sxs-lookup"><span data-stu-id="13674-852">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="13674-853">在 Outlook for iOS 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="13674-853">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="13674-854">`getRegExMatchesByName` 方法返回匹配在具有指定 `RegExName` 元素值的清单 XML 文件中的 `ItemHasRegularExpressionMatch` 规则元素中定义的正则表达式的字符串。</span><span class="sxs-lookup"><span data-stu-id="13674-854">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="13674-p157">如果在项目的正文属性上指定 `ItemHasRegularExpressionMatch` 规则，则正则表达式应进一步筛选正文，不应尝试返回该项目的整个正文。使用正则表达式（如 `.*`）获取项目的整个正文并不总是返回预期的结果。</span><span class="sxs-lookup"><span data-stu-id="13674-p157">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="13674-857">参数：</span><span class="sxs-lookup"><span data-stu-id="13674-857">Parameters:</span></span>

|<span data-ttu-id="13674-858">名称</span><span class="sxs-lookup"><span data-stu-id="13674-858">Name</span></span>| <span data-ttu-id="13674-859">类型</span><span class="sxs-lookup"><span data-stu-id="13674-859">Type</span></span>| <span data-ttu-id="13674-860">描述</span><span class="sxs-lookup"><span data-stu-id="13674-860">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="13674-861">字符串</span><span class="sxs-lookup"><span data-stu-id="13674-861">String</span></span>|<span data-ttu-id="13674-862">定义筛选器匹配的 `ItemHasRegularExpressionMatch` 规则元素的名称。</span><span class="sxs-lookup"><span data-stu-id="13674-862">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="13674-863">要求</span><span class="sxs-lookup"><span data-stu-id="13674-863">Requirements</span></span>

|<span data-ttu-id="13674-864">要求</span><span class="sxs-lookup"><span data-stu-id="13674-864">Requirement</span></span>| <span data-ttu-id="13674-865">值</span><span class="sxs-lookup"><span data-stu-id="13674-865">Value</span></span>|
|---|---|
|[<span data-ttu-id="13674-866">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="13674-866">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="13674-867">1.0</span><span class="sxs-lookup"><span data-stu-id="13674-867">1.0</span></span>|
|[<span data-ttu-id="13674-868">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="13674-868">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="13674-869">ReadItem</span><span class="sxs-lookup"><span data-stu-id="13674-869">ReadItem</span></span>|
|[<span data-ttu-id="13674-870">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="13674-870">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="13674-871">阅读</span><span class="sxs-lookup"><span data-stu-id="13674-871">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="13674-872">返回：</span><span class="sxs-lookup"><span data-stu-id="13674-872">Returns:</span></span>

<span data-ttu-id="13674-873">一个包含匹配在清单 XML 文件中定义的正则表达式的字符串的数组。</span><span class="sxs-lookup"><span data-stu-id="13674-873">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<dl class="param-type"><span data-ttu-id="13674-874">

<dt>类型</dt>

</span><span class="sxs-lookup"><span data-stu-id="13674-874">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="13674-875">Array.< String ></span><span class="sxs-lookup"><span data-stu-id="13674-875">Array.< String ></span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="13674-876">示例</span><span class="sxs-lookup"><span data-stu-id="13674-876">Example</span></span>

```js
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

####  <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="13674-877">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="13674-877">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="13674-878">以异步方式返回邮件的主题或正文中选定的数据。</span><span class="sxs-lookup"><span data-stu-id="13674-878">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="13674-p158">如果没有选定内容，但光标位于正文或主题中，此方法将会为所选数据返回 null。如果选定的是字段，而不是正文或主题，则此方法返回 `InvalidSelection` 错误。</span><span class="sxs-lookup"><span data-stu-id="13674-p158">If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="13674-881">参数：</span><span class="sxs-lookup"><span data-stu-id="13674-881">Parameters:</span></span>

|<span data-ttu-id="13674-882">名称</span><span class="sxs-lookup"><span data-stu-id="13674-882">Name</span></span>| <span data-ttu-id="13674-883">类型</span><span class="sxs-lookup"><span data-stu-id="13674-883">Type</span></span>| <span data-ttu-id="13674-884">属性</span><span class="sxs-lookup"><span data-stu-id="13674-884">Attributes</span></span>| <span data-ttu-id="13674-885">说明</span><span class="sxs-lookup"><span data-stu-id="13674-885">Description</span></span>|
|---|---|---|---|
|`coercionType`| [<span data-ttu-id="13674-886">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="13674-886">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="13674-p159">请求数据的格式。如果为文本，则此方法返回纯文本作为字符串，删除任何显示的 HTML 标记。如果为 HTML，则此方法返回所选文本，不论是纯文本还是 HTML。</span><span class="sxs-lookup"><span data-stu-id="13674-p159">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`| <span data-ttu-id="13674-890">对象</span><span class="sxs-lookup"><span data-stu-id="13674-890">Object</span></span>| <span data-ttu-id="13674-891">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="13674-891">&lt;optional&gt;</span></span>|<span data-ttu-id="13674-892">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="13674-892">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="13674-893">对象</span><span class="sxs-lookup"><span data-stu-id="13674-893">Object</span></span>| <span data-ttu-id="13674-894">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="13674-894">&lt;optional&gt;</span></span>|<span data-ttu-id="13674-895">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="13674-895">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="13674-896">函数</span><span class="sxs-lookup"><span data-stu-id="13674-896">function</span></span>||<span data-ttu-id="13674-897">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="13674-897">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="13674-898">若要从回调方法访问所选数据，请调用 `asyncResult.value.data`。</span><span class="sxs-lookup"><span data-stu-id="13674-898">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="13674-899">若要访问选定内容的源属性，请调用 `asyncResult.value.sourceProperty`，这将为 `body` 或 `subject`。</span><span class="sxs-lookup"><span data-stu-id="13674-899">To access the source property that the selection comes from, call , which will be either  or .|</span></span>|

##### <a name="requirements"></a><span data-ttu-id="13674-900">要求</span><span class="sxs-lookup"><span data-stu-id="13674-900">Requirements</span></span>

|<span data-ttu-id="13674-901">要求</span><span class="sxs-lookup"><span data-stu-id="13674-901">Requirement</span></span>| <span data-ttu-id="13674-902">值</span><span class="sxs-lookup"><span data-stu-id="13674-902">Value</span></span>|
|---|---|
|[<span data-ttu-id="13674-903">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="13674-903">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="13674-904">1.2</span><span class="sxs-lookup"><span data-stu-id="13674-904">1.2</span></span>|
|[<span data-ttu-id="13674-905">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="13674-905">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="13674-906">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="13674-906">ReadWriteItem</span></span>|
|[<span data-ttu-id="13674-907">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="13674-907">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="13674-908">撰写</span><span class="sxs-lookup"><span data-stu-id="13674-908">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="13674-909">返回：</span><span class="sxs-lookup"><span data-stu-id="13674-909">Returns:</span></span>

<span data-ttu-id="13674-910">作为字符串的所选数据的格式由 `coercionType` 确定。</span><span class="sxs-lookup"><span data-stu-id="13674-910">The selected data as a string with format determined by `coercionType`.</span></span>

<dl class="param-type"><span data-ttu-id="13674-911">

<dt>
类型</dt>


</span><span class="sxs-lookup"><span data-stu-id="13674-911">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="13674-912">字符串</span><span class="sxs-lookup"><span data-stu-id="13674-912">String</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="13674-913">示例</span><span class="sxs-lookup"><span data-stu-id="13674-913">Example</span></span>

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

####  <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="13674-914">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="13674-914">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="13674-915">异步加载所选项目上此外接程序的自定义属性。</span><span class="sxs-lookup"><span data-stu-id="13674-915">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="13674-p161">自定义属性基于每个应用、每个项目存储为键/值对。此方法在回调中返回 `CustomProperties` 对象，该回调提供访问特定于当前项目和当前外接程序的自定义属性的方法。自定义属性未在项目上加密，因此这不应用作安全存储。</span><span class="sxs-lookup"><span data-stu-id="13674-p161">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="13674-919">参数：</span><span class="sxs-lookup"><span data-stu-id="13674-919">Parameters:</span></span>

|<span data-ttu-id="13674-920">名称</span><span class="sxs-lookup"><span data-stu-id="13674-920">Name</span></span>| <span data-ttu-id="13674-921">类型</span><span class="sxs-lookup"><span data-stu-id="13674-921">Type</span></span>| <span data-ttu-id="13674-922">属性</span><span class="sxs-lookup"><span data-stu-id="13674-922">Attributes</span></span>| <span data-ttu-id="13674-923">说明</span><span class="sxs-lookup"><span data-stu-id="13674-923">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="13674-924">函数</span><span class="sxs-lookup"><span data-stu-id="13674-924">function</span></span>||<span data-ttu-id="13674-925">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="13674-925">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="13674-926">自定义属性作为 `asyncResult.value` 属性中的 [`CustomProperties`](/javascript/api/outlook_1_3/office.customproperties) 对象提供。</span><span class="sxs-lookup"><span data-stu-id="13674-926">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook_1_3/office.customproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="13674-927">此对象可用于获取、设置以及从项目中删除自定义属性，并将自定义属性集的更改重新保存到服务器。</span><span class="sxs-lookup"><span data-stu-id="13674-927">The custom properties are provided as a CustomProperties object in the asyncResult.value property. This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`| <span data-ttu-id="13674-928">对象</span><span class="sxs-lookup"><span data-stu-id="13674-928">Object</span></span>| <span data-ttu-id="13674-929">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="13674-929">&lt;optional&gt;</span></span>|<span data-ttu-id="13674-930">开发人员可以提供他们想要在回调函数中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="13674-930">Developers can provide any object they wish to access in the callback method.</span></span> <span data-ttu-id="13674-931">此对象可以通过回调函数中的 `asyncResult.asyncContext` 属性进行访问。</span><span class="sxs-lookup"><span data-stu-id="13674-931">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.|</span></span>|

##### <a name="requirements"></a><span data-ttu-id="13674-932">要求</span><span class="sxs-lookup"><span data-stu-id="13674-932">Requirements</span></span>

|<span data-ttu-id="13674-933">要求</span><span class="sxs-lookup"><span data-stu-id="13674-933">Requirement</span></span>| <span data-ttu-id="13674-934">值</span><span class="sxs-lookup"><span data-stu-id="13674-934">Value</span></span>|
|---|---|
|[<span data-ttu-id="13674-935">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="13674-935">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="13674-936">1.0</span><span class="sxs-lookup"><span data-stu-id="13674-936">1.0</span></span>|
|[<span data-ttu-id="13674-937">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="13674-937">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="13674-938">ReadItem</span><span class="sxs-lookup"><span data-stu-id="13674-938">ReadItem</span></span>|
|[<span data-ttu-id="13674-939">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="13674-939">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="13674-940">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="13674-940">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="13674-941">示例</span><span class="sxs-lookup"><span data-stu-id="13674-941">Example</span></span>

<span data-ttu-id="13674-p164">以下代码示例显示了如何使用 `loadCustomPropertiesAsync` 方法异步加载特定于当前项目的自定义属性。该示例还显示了如何使用 `CustomProperties.saveAsync` 方法将这些属性重新保存到服务器。加载自定义属性后，该代码示例将使用 `CustomProperties.get` 方法读取自定义属性 `myProp`，使用 `CustomProperties.set` 方法写入自定义属性 `otherProp`，最后调用 `saveAsync` 方法保存这些自定义属性。</span><span class="sxs-lookup"><span data-stu-id="13674-p164">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

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

####  <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="13674-945">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="13674-945">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="13674-946">将附件从邮件或约会中删除。</span><span class="sxs-lookup"><span data-stu-id="13674-946">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="13674-p165">`removeAttachmentAsync` 方法删除项目中带指定标识符的附件。最佳做法是，仅当同一个邮件应用程序在同一会话中添加了一个附件时，你才应使用该附件标识符来删除该附件。在 Outlook Web App 和适用于设备的 OWA 中，附件标识符只在同一个会话中才有效。当用户关闭应用程序，或者如果用户开始在内嵌窗体中撰写，并在随后弹出的内嵌窗体中继续在单独的窗口撰写时，会话即结束。</span><span class="sxs-lookup"><span data-stu-id="13674-p165">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item. As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session. In Outlook Web App and OWA for Devices, the attachment identifier is valid only within the same session. A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="13674-951">参数：</span><span class="sxs-lookup"><span data-stu-id="13674-951">Parameters:</span></span>

|<span data-ttu-id="13674-952">名称</span><span class="sxs-lookup"><span data-stu-id="13674-952">Name</span></span>| <span data-ttu-id="13674-953">类型</span><span class="sxs-lookup"><span data-stu-id="13674-953">Type</span></span>| <span data-ttu-id="13674-954">属性</span><span class="sxs-lookup"><span data-stu-id="13674-954">Attributes</span></span>| <span data-ttu-id="13674-955">说明</span><span class="sxs-lookup"><span data-stu-id="13674-955">Description</span></span>|
|---|---|---|---|
|`attachmentId`| <span data-ttu-id="13674-956">字符串</span><span class="sxs-lookup"><span data-stu-id="13674-956">String</span></span>||<span data-ttu-id="13674-p166">要删除的附件的标识符。字符串的最大长度为 100 个字符。</span><span class="sxs-lookup"><span data-stu-id="13674-p166">The identifier of the attachment to remove. The maximum length of the string is 100 characters.</span></span>|
|`options`| <span data-ttu-id="13674-959">Object</span><span class="sxs-lookup"><span data-stu-id="13674-959">Object</span></span>| <span data-ttu-id="13674-960">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="13674-960">&lt;optional&gt;</span></span>|<span data-ttu-id="13674-961">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="13674-961">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="13674-962">对象</span><span class="sxs-lookup"><span data-stu-id="13674-962">Object</span></span>| <span data-ttu-id="13674-963">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="13674-963">&lt;optional&gt;</span></span>|<span data-ttu-id="13674-964">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="13674-964">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="13674-965">函数</span><span class="sxs-lookup"><span data-stu-id="13674-965">function</span></span>| <span data-ttu-id="13674-966">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="13674-966">&lt;optional&gt;</span></span>|<span data-ttu-id="13674-967">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="13674-967">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="13674-968">如果删除附件失败，`asyncResult.error` 属性将包含一个说明失败原因的错误代码。</span><span class="sxs-lookup"><span data-stu-id="13674-968">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="13674-969">错误</span><span class="sxs-lookup"><span data-stu-id="13674-969">Errors</span></span>

| <span data-ttu-id="13674-970">错误代码</span><span class="sxs-lookup"><span data-stu-id="13674-970">Error code</span></span> | <span data-ttu-id="13674-971">说明</span><span class="sxs-lookup"><span data-stu-id="13674-971">Description</span></span> |
|------------|-------------|
| `InvalidAttachmentId` | <span data-ttu-id="13674-972">附件标识符不存在。</span><span class="sxs-lookup"><span data-stu-id="13674-972">The attachment identifier does not exist.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="13674-973">要求</span><span class="sxs-lookup"><span data-stu-id="13674-973">Requirements</span></span>

|<span data-ttu-id="13674-974">要求</span><span class="sxs-lookup"><span data-stu-id="13674-974">Requirement</span></span>| <span data-ttu-id="13674-975">值</span><span class="sxs-lookup"><span data-stu-id="13674-975">Value</span></span>|
|---|---|
|[<span data-ttu-id="13674-976">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="13674-976">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="13674-977">1.1</span><span class="sxs-lookup"><span data-stu-id="13674-977">1.1</span></span>|
|[<span data-ttu-id="13674-978">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="13674-978">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="13674-979">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="13674-979">ReadWriteItem</span></span>|
|[<span data-ttu-id="13674-980">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="13674-980">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="13674-981">撰写</span><span class="sxs-lookup"><span data-stu-id="13674-981">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="13674-982">示例</span><span class="sxs-lookup"><span data-stu-id="13674-982">Example</span></span>

<span data-ttu-id="13674-983">以下代码删除包含标识符 '0' 的附件。</span><span class="sxs-lookup"><span data-stu-id="13674-983">The following code removes an attachment with an identifier of '0'.</span></span>

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

####  <a name="saveasyncoptions-callback"></a><span data-ttu-id="13674-984">saveAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="13674-984">saveAsync([options], callback)</span></span>

<span data-ttu-id="13674-985">异步保存项目。</span><span class="sxs-lookup"><span data-stu-id="13674-985">Asynchronously saves an item.</span></span>

<span data-ttu-id="13674-p167">调用时，此方法将当前邮件保存为草稿，并通过回调方法返回项目 ID。在 Outlook Web App 或 Outlook 联机模式下，该项目被保存到服务器中。在 Outlook 缓存模式下，该项目被保存到本地缓存中。</span><span class="sxs-lookup"><span data-stu-id="13674-p167">When invoked, this method saves the current message as a draft and returns the item id via the callback method. In Outlook Web App or Outlook in online mode, the item is saved to the server. In Outlook in cached mode, the item is saved to the local cache.</span></span>

> [!NOTE]
> <span data-ttu-id="13674-989">如果外接程序在撰写模式下对某个项目调用 `saveAsync` 来获得 `itemId`，以便与 EWS 或 REST API 一同使用，请注意，当 Outlook 处于高速缓存模式时，可能需要一段时间项目才能真正同步到服务器。</span><span class="sxs-lookup"><span data-stu-id="13674-989">Note: If your add-in calls `saveAsync` on an item in compose mode in order to get an `itemId` to use with EWS or the REST API, be aware that when Outlook is in cached mode, it may take some time before the item is actually synced to the server. Until the item is synced, using the  will return an error.</span></span> <span data-ttu-id="13674-990">在项目同步前，使用 `itemId` 将返回一个错误。</span><span class="sxs-lookup"><span data-stu-id="13674-990">Until the item is synced, using the `itemId` will return an error.</span></span>

<span data-ttu-id="13674-p169">由于约会没有草稿状态，如果以撰写模式在约会中调用 `saveAsync`，则该项将被保存为用户日历中的正常约会。对于之前未保存过的新约会，则不会发送邀请。保存现有约会将向添加或删除的与会者发送更新。</span><span class="sxs-lookup"><span data-stu-id="13674-p169">Since appointments have no draft state, if `saveAsync` is called on an appointment in compose mode, the item will be saved as a normal appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent. Saving an existing appointment will send an update to added or removed attendees.</span></span>

> [!NOTE]
> <span data-ttu-id="13674-994">以下客户端在撰写模式下对约会上的 `saveAsync` 具有不同的行为：</span><span class="sxs-lookup"><span data-stu-id="13674-994">The following clients have different behavior for `saveAsync` on appointments in compose mode:</span></span>
>
> - <span data-ttu-id="13674-995">Mac Outlook 不支持在撰写模式下对会议执行 `saveAsync` 操作。</span><span class="sxs-lookup"><span data-stu-id="13674-995">Note: Mac Outlook does not support `saveAsync` on a meeting in compose mode. Calling  on a meeting in Mac Outlook will return an error.</span></span> <span data-ttu-id="13674-996">对 Mac Outlook 中的会议调用 `saveAsync` 将会返回一个错误。</span><span class="sxs-lookup"><span data-stu-id="13674-996">Note: Mac Outlook does not support  on a meeting in compose mode. Calling `saveAsync` on a meeting in Mac Outlook will return an error.</span></span>
> - <span data-ttu-id="13674-997">在撰写模式下的约会上调用 `saveAsync` 时，Outlook 网页版始终发送邀请或更新。</span><span class="sxs-lookup"><span data-stu-id="13674-997">Outlook on the web always sends an invitation or update when `saveAsync` is called on an appointment in compose mode.</span></span>

##### <a name="parameters"></a><span data-ttu-id="13674-998">参数：</span><span class="sxs-lookup"><span data-stu-id="13674-998">Parameters:</span></span>

|<span data-ttu-id="13674-999">名称</span><span class="sxs-lookup"><span data-stu-id="13674-999">Name</span></span>| <span data-ttu-id="13674-1000">类型</span><span class="sxs-lookup"><span data-stu-id="13674-1000">Type</span></span>| <span data-ttu-id="13674-1001">属性</span><span class="sxs-lookup"><span data-stu-id="13674-1001">Attributes</span></span>| <span data-ttu-id="13674-1002">说明</span><span class="sxs-lookup"><span data-stu-id="13674-1002">Description</span></span>|
|---|---|---|---|
|`options`| <span data-ttu-id="13674-1003">Object</span><span class="sxs-lookup"><span data-stu-id="13674-1003">Object</span></span>| <span data-ttu-id="13674-1004">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="13674-1004">&lt;optional&gt;</span></span>|<span data-ttu-id="13674-1005">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="13674-1005">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="13674-1006">对象</span><span class="sxs-lookup"><span data-stu-id="13674-1006">Object</span></span>| <span data-ttu-id="13674-1007">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="13674-1007">&lt;optional&gt;</span></span>|<span data-ttu-id="13674-1008">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="13674-1008">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="13674-1009">函数</span><span class="sxs-lookup"><span data-stu-id="13674-1009">function</span></span>||<span data-ttu-id="13674-1010">方法完成后，使用一个参数 `asyncResult`（[`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="13674-1010">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="13674-1011">如果成功，该项目标识符将在 `asyncResult.value` 属性中提供。</span><span class="sxs-lookup"><span data-stu-id="13674-1011">On success, the item identifier is provided in the `asyncResult.value` property.|</span></span>|

##### <a name="requirements"></a><span data-ttu-id="13674-1012">要求</span><span class="sxs-lookup"><span data-stu-id="13674-1012">Requirements</span></span>

|<span data-ttu-id="13674-1013">要求</span><span class="sxs-lookup"><span data-stu-id="13674-1013">Requirement</span></span>| <span data-ttu-id="13674-1014">值</span><span class="sxs-lookup"><span data-stu-id="13674-1014">Value</span></span>|
|---|---|
|[<span data-ttu-id="13674-1015">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="13674-1015">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="13674-1016">1.3</span><span class="sxs-lookup"><span data-stu-id="13674-1016">1.3</span></span>|
|[<span data-ttu-id="13674-1017">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="13674-1017">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="13674-1018">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="13674-1018">ReadWriteItem</span></span>|
|[<span data-ttu-id="13674-1019">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="13674-1019">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="13674-1020">撰写</span><span class="sxs-lookup"><span data-stu-id="13674-1020">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="13674-1021">示例</span><span class="sxs-lookup"><span data-stu-id="13674-1021">Examples</span></span>

```js
Office.context.mailbox.item.saveAsync(
  function callback(result) {
    // Process the result
  });
```

<span data-ttu-id="13674-p171">下面是传递给回调函数的 `result` 参数的示例。`value` 属性包含的项目的项目 ID。</span><span class="sxs-lookup"><span data-stu-id="13674-p171">The following is an example of the `result` parameter passed to the callback function. The `value` property contains the item ID of the item.</span></span>

```json
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

####  <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="13674-1024">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="13674-1024">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="13674-1025">以异步方式将数据插入到邮件的正文或主题中。</span><span class="sxs-lookup"><span data-stu-id="13674-1025">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="13674-p172">`setSelectedDataAsync` 方法将指定字符串插入到项目主题或正文的光标位置，或者，如果在编辑器中已选择文本，则该方法将替换选择的文本。如果光标不在正文或主题字段中，则返回错误。插入之后，光标会位于插入内容的末尾。</span><span class="sxs-lookup"><span data-stu-id="13674-p172">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="13674-1029">参数：</span><span class="sxs-lookup"><span data-stu-id="13674-1029">Parameters:</span></span>

|<span data-ttu-id="13674-1030">名称</span><span class="sxs-lookup"><span data-stu-id="13674-1030">Name</span></span>| <span data-ttu-id="13674-1031">类型</span><span class="sxs-lookup"><span data-stu-id="13674-1031">Type</span></span>| <span data-ttu-id="13674-1032">属性</span><span class="sxs-lookup"><span data-stu-id="13674-1032">Attributes</span></span>| <span data-ttu-id="13674-1033">说明</span><span class="sxs-lookup"><span data-stu-id="13674-1033">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="13674-1034">字符串</span><span class="sxs-lookup"><span data-stu-id="13674-1034">String</span></span>||<span data-ttu-id="13674-p173">要插入的数据。数据不得超过 1,000,000 个字符。如果传入的数据超过 1,000,000 个字符，则会引发 `ArgumentOutOfRange` 异常。</span><span class="sxs-lookup"><span data-stu-id="13674-p173">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`| <span data-ttu-id="13674-1038">Object</span><span class="sxs-lookup"><span data-stu-id="13674-1038">Object</span></span>| <span data-ttu-id="13674-1039">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="13674-1039">&lt;optional&gt;</span></span>|<span data-ttu-id="13674-1040">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="13674-1040">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="13674-1041">对象</span><span class="sxs-lookup"><span data-stu-id="13674-1041">Object</span></span>| <span data-ttu-id="13674-1042">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="13674-1042">&lt;optional&gt;</span></span>|<span data-ttu-id="13674-1043">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="13674-1043">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`| [<span data-ttu-id="13674-1044">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="13674-1044">Office.CoercionType</span></span>](office.md#coerciontype-string)| <span data-ttu-id="13674-1045">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="13674-1045">&lt;optional&gt;</span></span>|<span data-ttu-id="13674-p174">如果为 `text`，则在 Outlook Web App 和 Outlook 中应用当前样式。如果该字段是 HTML 编辑器，则仅插入文本数据，即使数据为 HTML。</span><span class="sxs-lookup"><span data-stu-id="13674-p174">If `text`, the current style is applied in Outlook Web App and Outlook. If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="13674-p175">如果 `html` 和该字段支持 HTML（主题不支持），则在 Outlook Web App 中应用当前样式，而在 Outlook 中应用默认样式。如果该字段是文本字段，则返回 `InvalidDataFormat` 错误。</span><span class="sxs-lookup"><span data-stu-id="13674-p175">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook Web App and the default style is applied in Outlook. If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="13674-1050">如果未设置 `coercionType`，则结果取决于该字段：如果该字段是 HTML，则使用 HTML；如果该字段是文本，则使用纯文本。</span><span class="sxs-lookup"><span data-stu-id="13674-1050">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`| <span data-ttu-id="13674-1051">function</span><span class="sxs-lookup"><span data-stu-id="13674-1051">function</span></span>||<span data-ttu-id="13674-1052">方法完成后，使用单个参数 `callback`（一个 [`asyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `AsyncResult` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="13674-1052">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="13674-1053">Requirements</span><span class="sxs-lookup"><span data-stu-id="13674-1053">Requirements</span></span>

|<span data-ttu-id="13674-1054">要求</span><span class="sxs-lookup"><span data-stu-id="13674-1054">Requirement</span></span>| <span data-ttu-id="13674-1055">值</span><span class="sxs-lookup"><span data-stu-id="13674-1055">Value</span></span>|
|---|---|
|[<span data-ttu-id="13674-1056">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="13674-1056">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="13674-1057">1.2</span><span class="sxs-lookup"><span data-stu-id="13674-1057">1.2</span></span>|
|[<span data-ttu-id="13674-1058">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="13674-1058">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="13674-1059">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="13674-1059">ReadWriteItem</span></span>|
|[<span data-ttu-id="13674-1060">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="13674-1060">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="13674-1061">撰写</span><span class="sxs-lookup"><span data-stu-id="13674-1061">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="13674-1062">示例</span><span class="sxs-lookup"><span data-stu-id="13674-1062">Example</span></span>

```js
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```