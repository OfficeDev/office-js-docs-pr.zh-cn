
# <a name="item"></a><span data-ttu-id="f40a0-101">item</span><span class="sxs-lookup"><span data-stu-id="f40a0-101">item</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmditem"></a><span data-ttu-id="f40a0-102">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span><span class="sxs-lookup"><span data-stu-id="f40a0-102">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span></span>

<span data-ttu-id="f40a0-p101">`item`命名空间用于访问当前选定的邮件、会议请求或安排。可以通过使用[itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlook14officemailboxenumsitemtype) 属性确定`item`的类型。</span><span class="sxs-lookup"><span data-stu-id="f40a0-p101">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlook14officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="f40a0-105">要求</span><span class="sxs-lookup"><span data-stu-id="f40a0-105">Requirements</span></span>

|<span data-ttu-id="f40a0-106">要求</span><span class="sxs-lookup"><span data-stu-id="f40a0-106">Requirement</span></span>| <span data-ttu-id="f40a0-107">值</span><span class="sxs-lookup"><span data-stu-id="f40a0-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="f40a0-108">最低邮箱要求集版本</span><span class="sxs-lookup"><span data-stu-id="f40a0-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f40a0-109">1.0</span><span class="sxs-lookup"><span data-stu-id="f40a0-109">1.0</span></span>|
|[<span data-ttu-id="f40a0-110">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="f40a0-110">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f40a0-111">Restricted</span><span class="sxs-lookup"><span data-stu-id="f40a0-111">Restricted</span></span>|
|[<span data-ttu-id="f40a0-112">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="f40a0-112">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f40a0-113">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="f40a0-113">Compose or read</span></span>|

### <a name="example"></a><span data-ttu-id="f40a0-114">示例</span><span class="sxs-lookup"><span data-stu-id="f40a0-114">Example</span></span>

<span data-ttu-id="f40a0-115">以下 JavaScript 代码示例显示了如何访问 Outlook 中当前项目的 `subject` 属性。</span><span class="sxs-lookup"><span data-stu-id="f40a0-115">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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

### <a name="members"></a><span data-ttu-id="f40a0-116">成员</span><span class="sxs-lookup"><span data-stu-id="f40a0-116">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlook14officeattachmentdetails"></a><span data-ttu-id="f40a0-117">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_4/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="f40a0-117">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_4/office.attachmentdetails)></span></span>

<span data-ttu-id="f40a0-p102">获取项目的附件数组。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="f40a0-p102">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="f40a0-120">某些类型的文件因潜在的安全问题被 Outlook 阻止，因此没有返回。</span><span class="sxs-lookup"><span data-stu-id="f40a0-120">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="f40a0-121">有关详细信息，请参阅 [在 Outlook 中被阻止的附件](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519)。</span><span class="sxs-lookup"><span data-stu-id="f40a0-121">For more information see [Payments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="f40a0-122">类型：</span><span class="sxs-lookup"><span data-stu-id="f40a0-122">Type:</span></span>

*   <span data-ttu-id="f40a0-123">数组。 <[AttachmentDetails](/javascript/api/outlook_1_4/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="f40a0-123">Array.<[AttachmentDetails](/javascript/api/outlook_1_4/office.attachmentdetails)></span></span>

##### <a name="requirements"></a><span data-ttu-id="f40a0-124">要求</span><span class="sxs-lookup"><span data-stu-id="f40a0-124">Requirements</span></span>

|<span data-ttu-id="f40a0-125">要求</span><span class="sxs-lookup"><span data-stu-id="f40a0-125">Requirement</span></span>| <span data-ttu-id="f40a0-126">值</span><span class="sxs-lookup"><span data-stu-id="f40a0-126">Value</span></span>|
|---|---|
|[<span data-ttu-id="f40a0-127">最低邮箱要求集版本</span><span class="sxs-lookup"><span data-stu-id="f40a0-127">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f40a0-128">1.0</span><span class="sxs-lookup"><span data-stu-id="f40a0-128">1.0</span></span>|
|[<span data-ttu-id="f40a0-129">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="f40a0-129">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f40a0-130">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f40a0-130">ReadItem</span></span>|
|[<span data-ttu-id="f40a0-131">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="f40a0-131">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f40a0-132">阅读</span><span class="sxs-lookup"><span data-stu-id="f40a0-132">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="f40a0-133">示例</span><span class="sxs-lookup"><span data-stu-id="f40a0-133">Example</span></span>

<span data-ttu-id="f40a0-134">以下代码使用当前项目上所有附件的详细信息构成 HTML 字符串。</span><span class="sxs-lookup"><span data-stu-id="f40a0-134">The following code builds an HTML string with details of all attachments on the current item.</span></span>

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

####  <a name="bcc-recipientsjavascriptapioutlook14officerecipients"></a><span data-ttu-id="f40a0-135">密件抄送：[收件人](/javascript/api/outlook_1_4/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="f40a0-135">bcc :[Recipients](/javascript/api/outlook_1_4/office.recipients)</span></span>

<span data-ttu-id="f40a0-136">获取一个对象，提供用于获取或更新邮件的密件抄送 （密件抄送副本） 行的方法。</span><span class="sxs-lookup"><span data-stu-id="f40a0-136">Gets an object that provides methods to get or update the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="f40a0-137">仅限撰写模式。</span><span class="sxs-lookup"><span data-stu-id="f40a0-137">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="f40a0-138">类型：</span><span class="sxs-lookup"><span data-stu-id="f40a0-138">Type:</span></span>

*   [<span data-ttu-id="f40a0-139">收件人</span><span class="sxs-lookup"><span data-stu-id="f40a0-139">Recipients</span></span>](/javascript/api/outlook_1_4/office.recipients)

##### <a name="requirements"></a><span data-ttu-id="f40a0-140">要求</span><span class="sxs-lookup"><span data-stu-id="f40a0-140">Requirements</span></span>

|<span data-ttu-id="f40a0-141">要求</span><span class="sxs-lookup"><span data-stu-id="f40a0-141">Requirement</span></span>| <span data-ttu-id="f40a0-142">值</span><span class="sxs-lookup"><span data-stu-id="f40a0-142">Value</span></span>|
|---|---|
|[<span data-ttu-id="f40a0-143">最低的邮箱要求集版本</span><span class="sxs-lookup"><span data-stu-id="f40a0-143">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f40a0-144">1.1</span><span class="sxs-lookup"><span data-stu-id="f40a0-144">1.1</span></span>|
|[<span data-ttu-id="f40a0-145">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="f40a0-145">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f40a0-146">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f40a0-146">ReadItem</span></span>|
|[<span data-ttu-id="f40a0-147">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="f40a0-147">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f40a0-148">撰写</span><span class="sxs-lookup"><span data-stu-id="f40a0-148">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="f40a0-149">示例</span><span class="sxs-lookup"><span data-stu-id="f40a0-149">Example</span></span>

```
Office.context.mailbox.item.bcc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.bcc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.bcc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfBccRecipients = asyncResult.value;
}
```

####  <a name="body-bodyjavascriptapioutlook14officebody"></a><span data-ttu-id="f40a0-150">body :[Body](/javascript/api/outlook_1_4/office.body)</span><span class="sxs-lookup"><span data-stu-id="f40a0-150">body :[Body](/javascript/api/outlook_1_4/office.body)</span></span>

<span data-ttu-id="f40a0-151">获取一个提供用于处理项目正文的方法的对象。</span><span class="sxs-lookup"><span data-stu-id="f40a0-151">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="f40a0-152">类型：</span><span class="sxs-lookup"><span data-stu-id="f40a0-152">Type:</span></span>

*   [<span data-ttu-id="f40a0-153">Body</span><span class="sxs-lookup"><span data-stu-id="f40a0-153">Body</span></span>](/javascript/api/outlook_1_4/office.body)

##### <a name="requirements"></a><span data-ttu-id="f40a0-154">要求</span><span class="sxs-lookup"><span data-stu-id="f40a0-154">Requirements</span></span>

|<span data-ttu-id="f40a0-155">要求</span><span class="sxs-lookup"><span data-stu-id="f40a0-155">Requirement</span></span>| <span data-ttu-id="f40a0-156">值</span><span class="sxs-lookup"><span data-stu-id="f40a0-156">Value</span></span>|
|---|---|
|[<span data-ttu-id="f40a0-157">最低邮箱要求集版本</span><span class="sxs-lookup"><span data-stu-id="f40a0-157">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f40a0-158">1.1</span><span class="sxs-lookup"><span data-stu-id="f40a0-158">1.1</span></span>|
|[<span data-ttu-id="f40a0-159">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="f40a0-159">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f40a0-160">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f40a0-160">ReadItem</span></span>|
|[<span data-ttu-id="f40a0-161">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="f40a0-161">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f40a0-162">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="f40a0-162">Compose or read</span></span>|

####  <a name="cc-arrayemailaddressdetailsjavascriptapioutlook14officeemailaddressdetailsrecipientsjavascriptapioutlook14officerecipients"></a><span data-ttu-id="f40a0-163">抄送： 数组。 <[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)>|[收件人](/javascript/api/outlook_1_4/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="f40a0-163">cc :Array.<[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_4/office.recipients)</span></span>

<span data-ttu-id="f40a0-164">提供对邮件抄送 (cc) 收件人的访问。</span><span class="sxs-lookup"><span data-stu-id="f40a0-164">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="f40a0-165">对象的类型和访问级别取决于当前项的模式。</span><span class="sxs-lookup"><span data-stu-id="f40a0-165">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="f40a0-166">阅读模式</span><span class="sxs-lookup"><span data-stu-id="f40a0-166">Read mode</span></span>

<span data-ttu-id="f40a0-p106">`cc`属性返回包含邮件的**抄送**行上所列每个收件人的 `EmailAddressDetails` 对象数组。集合上限为 100 个成员。</span><span class="sxs-lookup"><span data-stu-id="f40a0-p106">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="f40a0-169">撰写模式</span><span class="sxs-lookup"><span data-stu-id="f40a0-169">Compose mode</span></span>

<span data-ttu-id="f40a0-170">`cc` 属性返回 `Recipients` 对象，该对象提供用于处理邮件**抄送**行上收件人的方法。</span><span class="sxs-lookup"><span data-stu-id="f40a0-170">The `cc` property returns a `Recipients` object that provides methods for manipulating the recipients on the **Cc** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="f40a0-171">类型：</span><span class="sxs-lookup"><span data-stu-id="f40a0-171">Type:</span></span>

*   <span data-ttu-id="f40a0-172">数组。 <[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)> |[收件人](/javascript/api/outlook_1_4/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="f40a0-172">Array.<[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_4/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="f40a0-173">要求</span><span class="sxs-lookup"><span data-stu-id="f40a0-173">Requirements</span></span>

|<span data-ttu-id="f40a0-174">要求</span><span class="sxs-lookup"><span data-stu-id="f40a0-174">Requirement</span></span>| <span data-ttu-id="f40a0-175">值</span><span class="sxs-lookup"><span data-stu-id="f40a0-175">Value</span></span>|
|---|---|
|[<span data-ttu-id="f40a0-176">最低邮箱要求集版本</span><span class="sxs-lookup"><span data-stu-id="f40a0-176">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f40a0-177">1.0</span><span class="sxs-lookup"><span data-stu-id="f40a0-177">1.0</span></span>|
|[<span data-ttu-id="f40a0-178">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="f40a0-178">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f40a0-179">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f40a0-179">ReadItem</span></span>|
|[<span data-ttu-id="f40a0-180">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="f40a0-180">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f40a0-181">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="f40a0-181">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="f40a0-182">示例</span><span class="sxs-lookup"><span data-stu-id="f40a0-182">Example</span></span>

```
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

####  <a name="nullable-conversationid-string"></a><span data-ttu-id="f40a0-183">(nullable) conversationId :String</span><span class="sxs-lookup"><span data-stu-id="f40a0-183">(nullable) conversationId :String</span></span>

<span data-ttu-id="f40a0-184">获取包含特定消息的电子邮件会话的标识符。</span><span class="sxs-lookup"><span data-stu-id="f40a0-184">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="f40a0-p107">如果在阅读窗体或撰写窗体的回复中激活邮件应用程序，则此属性可以获得一个整数值。如果用户随后更改了回复邮件的主题（若发送回复），则该邮件的对话 ID 将改变且之前获取的值将不适用。</span><span class="sxs-lookup"><span data-stu-id="f40a0-p107">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="f40a0-p108">对于撰写窗体的新项目，此属性获得一个 null 值。如果用户设置一个主题并保存该项目，`conversationId` 属性将返回一个值。</span><span class="sxs-lookup"><span data-stu-id="f40a0-p108">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="f40a0-189">类型：</span><span class="sxs-lookup"><span data-stu-id="f40a0-189">Type:</span></span>

*   <span data-ttu-id="f40a0-190">String</span><span class="sxs-lookup"><span data-stu-id="f40a0-190">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="f40a0-191">要求</span><span class="sxs-lookup"><span data-stu-id="f40a0-191">Requirements</span></span>

|<span data-ttu-id="f40a0-192">要求</span><span class="sxs-lookup"><span data-stu-id="f40a0-192">Requirement</span></span>| <span data-ttu-id="f40a0-193">值</span><span class="sxs-lookup"><span data-stu-id="f40a0-193">Value</span></span>|
|---|---|
|[<span data-ttu-id="f40a0-194">最低邮箱要求集版本</span><span class="sxs-lookup"><span data-stu-id="f40a0-194">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f40a0-195">1.0</span><span class="sxs-lookup"><span data-stu-id="f40a0-195">1.0</span></span>|
|[<span data-ttu-id="f40a0-196">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="f40a0-196">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f40a0-197">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f40a0-197">ReadItem</span></span>|
|[<span data-ttu-id="f40a0-198">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="f40a0-198">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f40a0-199">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="f40a0-199">Compose or read</span></span>|

#### <a name="datetimecreated-date"></a><span data-ttu-id="f40a0-200">dateTimeCreated：日期</span><span class="sxs-lookup"><span data-stu-id="f40a0-200">dateTimeCreated :Date</span></span>

<span data-ttu-id="f40a0-p109">获取项目创建的日期和时间。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="f40a0-p109">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="f40a0-203">类型：</span><span class="sxs-lookup"><span data-stu-id="f40a0-203">Type:</span></span>

*   <span data-ttu-id="f40a0-204">Date</span><span class="sxs-lookup"><span data-stu-id="f40a0-204">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="f40a0-205">要求</span><span class="sxs-lookup"><span data-stu-id="f40a0-205">Requirements</span></span>

|<span data-ttu-id="f40a0-206">要求</span><span class="sxs-lookup"><span data-stu-id="f40a0-206">Requirement</span></span>| <span data-ttu-id="f40a0-207">值</span><span class="sxs-lookup"><span data-stu-id="f40a0-207">Value</span></span>|
|---|---|
|[<span data-ttu-id="f40a0-208">最低邮箱要求集版本</span><span class="sxs-lookup"><span data-stu-id="f40a0-208">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f40a0-209">1.0</span><span class="sxs-lookup"><span data-stu-id="f40a0-209">1.0</span></span>|
|[<span data-ttu-id="f40a0-210">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="f40a0-210">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f40a0-211">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f40a0-211">ReadItem</span></span>|
|[<span data-ttu-id="f40a0-212">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="f40a0-212">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f40a0-213">阅读</span><span class="sxs-lookup"><span data-stu-id="f40a0-213">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="f40a0-214">示例</span><span class="sxs-lookup"><span data-stu-id="f40a0-214">Example</span></span>

```
var created = Office.context.mailbox.item.dateTimeCreated;
```

#### <a name="datetimemodified-date"></a><span data-ttu-id="f40a0-215">dateTimeModified： 日期</span><span class="sxs-lookup"><span data-stu-id="f40a0-215">dateTimeModified :Date</span></span>

<span data-ttu-id="f40a0-p110">获取项目最近一次修改的日期和时间。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="f40a0-p110">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="f40a0-218">注意：在 iOS 版 Outlook 或  Android 版 Outlook 中不支持此成员。</span><span class="sxs-lookup"><span data-stu-id="f40a0-218">Note: This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="type"></a><span data-ttu-id="f40a0-219">类型：</span><span class="sxs-lookup"><span data-stu-id="f40a0-219">Type:</span></span>

*   <span data-ttu-id="f40a0-220">Date</span><span class="sxs-lookup"><span data-stu-id="f40a0-220">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="f40a0-221">要求</span><span class="sxs-lookup"><span data-stu-id="f40a0-221">Requirements</span></span>

|<span data-ttu-id="f40a0-222">要求</span><span class="sxs-lookup"><span data-stu-id="f40a0-222">Requirement</span></span>| <span data-ttu-id="f40a0-223">值</span><span class="sxs-lookup"><span data-stu-id="f40a0-223">Value</span></span>|
|---|---|
|[<span data-ttu-id="f40a0-224">最低邮箱要求集版本</span><span class="sxs-lookup"><span data-stu-id="f40a0-224">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f40a0-225">1.0</span><span class="sxs-lookup"><span data-stu-id="f40a0-225">1.0</span></span>|
|[<span data-ttu-id="f40a0-226">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="f40a0-226">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f40a0-227">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f40a0-227">ReadItem</span></span>|
|[<span data-ttu-id="f40a0-228">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="f40a0-228">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f40a0-229">阅读</span><span class="sxs-lookup"><span data-stu-id="f40a0-229">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="f40a0-230">示例</span><span class="sxs-lookup"><span data-stu-id="f40a0-230">Example</span></span>

```
var modified = Office.context.mailbox.item.dateTimeModified;
```

####  <a name="end-datetimejavascriptapioutlook14officetime"></a><span data-ttu-id="f40a0-231">end :Date|[Time](/javascript/api/outlook_1_4/office.time)</span><span class="sxs-lookup"><span data-stu-id="f40a0-231">end :Date|[Time](/javascript/api/outlook_1_4/office.time)</span></span>

<span data-ttu-id="f40a0-232">获取或设置约会结束的日期和时间。</span><span class="sxs-lookup"><span data-stu-id="f40a0-232">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="f40a0-p111">将 `end` 属性表示为协调世界时 (UTC) 的日期和时间值。可使用 [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook14officelocalclienttime) 方法将 end 属性值转换为客户端的本地日期和时间。</span><span class="sxs-lookup"><span data-stu-id="f40a0-p111">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook14officelocalclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="f40a0-235">阅读模式</span><span class="sxs-lookup"><span data-stu-id="f40a0-235">Read mode</span></span>

<span data-ttu-id="f40a0-236">`end` 属性返回 `Date` 对象。</span><span class="sxs-lookup"><span data-stu-id="f40a0-236">The `end` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="f40a0-237">撰写模式</span><span class="sxs-lookup"><span data-stu-id="f40a0-237">Compose mode</span></span>

<span data-ttu-id="f40a0-238">`end` 属性返回 `Time` 对象。</span><span class="sxs-lookup"><span data-stu-id="f40a0-238">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="f40a0-239">使用  方法设置结束时间时，应使用  方法将客户端的本地时间转换为服务器的 UTC。 [`Time.setAsync`](/javascript/api/outlook_1_4/office.time#setasync-datetime--options--callback-) [ `convertToUtcClientTime`  ](office.context.mailbox.md#converttoutcclienttimeinput--date)</span><span class="sxs-lookup"><span data-stu-id="f40a0-239">When you use the [`Time.setAsync`](/javascript/api/outlook_1_4/office.time#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="f40a0-240">类型：</span><span class="sxs-lookup"><span data-stu-id="f40a0-240">Type:</span></span>

*   <span data-ttu-id="f40a0-241">日期 | [时间](/javascript/api/outlook_1_4/office.time)</span><span class="sxs-lookup"><span data-stu-id="f40a0-241">Date | [Time](/javascript/api/outlook_1_4/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="f40a0-242">要求</span><span class="sxs-lookup"><span data-stu-id="f40a0-242">Requirements</span></span>

|<span data-ttu-id="f40a0-243">要求</span><span class="sxs-lookup"><span data-stu-id="f40a0-243">Requirement</span></span>| <span data-ttu-id="f40a0-244">值</span><span class="sxs-lookup"><span data-stu-id="f40a0-244">Value</span></span>|
|---|---|
|[<span data-ttu-id="f40a0-245">最低邮箱要求集版本</span><span class="sxs-lookup"><span data-stu-id="f40a0-245">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f40a0-246">1.0</span><span class="sxs-lookup"><span data-stu-id="f40a0-246">1.0</span></span>|
|[<span data-ttu-id="f40a0-247">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="f40a0-247">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f40a0-248">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f40a0-248">ReadItem</span></span>|
|[<span data-ttu-id="f40a0-249">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="f40a0-249">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f40a0-250">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="f40a0-250">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="f40a0-251">示例</span><span class="sxs-lookup"><span data-stu-id="f40a0-251">Example</span></span>

<span data-ttu-id="f40a0-252">以下示例通过使用 `Time` 对象的 [`setAsync`](/javascript/api/outlook_1_4/office.time#setasync-datetime--options--callback-) 方法，设置撰写模式下约会的结束时间。</span><span class="sxs-lookup"><span data-stu-id="f40a0-252">The following example sets the end time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_4/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

#### <a name="from-emailaddressdetailsjavascriptapioutlook14officeemailaddressdetails"></a><span data-ttu-id="f40a0-253">从：[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="f40a0-253">from :[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)</span></span>

<span data-ttu-id="f40a0-p112">获取邮件发件人的电子邮件地址。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="f40a0-p112">Gets the email address of the sender of a message. Read mode only.</span></span>

<span data-ttu-id="f40a0-p113">`from` 和 [`sender`](#sender-emailaddressdetailsjavascriptapioutlook14officeemailaddressdetails) 属性表示同一个人，邮件由代理人发送的除外。在此情况下，`from` 属性表示代理程序，而 sender 属性表示代理人。</span><span class="sxs-lookup"><span data-stu-id="f40a0-p113">The `from` and [`sender`](#sender-emailaddressdetailsjavascriptapioutlook14officeemailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="f40a0-258">`EmailAddressDetails` 对象的 `recipientType` 属性 在 `from` 属性是 `undefined`。</span><span class="sxs-lookup"><span data-stu-id="f40a0-258">Note: The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="f40a0-259">类型：</span><span class="sxs-lookup"><span data-stu-id="f40a0-259">Type:</span></span>

*   [<span data-ttu-id="f40a0-260">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="f40a0-260">EmailAddressDetails</span></span>](/javascript/api/outlook_1_4/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="f40a0-261">要求</span><span class="sxs-lookup"><span data-stu-id="f40a0-261">Requirements</span></span>

|<span data-ttu-id="f40a0-262">要求</span><span class="sxs-lookup"><span data-stu-id="f40a0-262">Requirement</span></span>| <span data-ttu-id="f40a0-263">值</span><span class="sxs-lookup"><span data-stu-id="f40a0-263">Value</span></span>|
|---|---|
|[<span data-ttu-id="f40a0-264">最低邮箱要求集版本</span><span class="sxs-lookup"><span data-stu-id="f40a0-264">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f40a0-265">1.0</span><span class="sxs-lookup"><span data-stu-id="f40a0-265">1.0</span></span>|
|[<span data-ttu-id="f40a0-266">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="f40a0-266">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f40a0-267">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f40a0-267">ReadItem</span></span>|
|[<span data-ttu-id="f40a0-268">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="f40a0-268">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f40a0-269">阅读</span><span class="sxs-lookup"><span data-stu-id="f40a0-269">Read</span></span>|

#### <a name="internetmessageid-string"></a><span data-ttu-id="f40a0-270">internetMessageId :String</span><span class="sxs-lookup"><span data-stu-id="f40a0-270">internetMessageId :String</span></span>

<span data-ttu-id="f40a0-p114">获取电子邮件的 Internet 消息标识符。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="f40a0-p114">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="f40a0-273">类型：</span><span class="sxs-lookup"><span data-stu-id="f40a0-273">Type:</span></span>

*   <span data-ttu-id="f40a0-274">String</span><span class="sxs-lookup"><span data-stu-id="f40a0-274">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="f40a0-275">要求</span><span class="sxs-lookup"><span data-stu-id="f40a0-275">Requirements</span></span>

|<span data-ttu-id="f40a0-276">要求</span><span class="sxs-lookup"><span data-stu-id="f40a0-276">Requirement</span></span>| <span data-ttu-id="f40a0-277">值</span><span class="sxs-lookup"><span data-stu-id="f40a0-277">Value</span></span>|
|---|---|
|[<span data-ttu-id="f40a0-278">最低邮箱要求集版本</span><span class="sxs-lookup"><span data-stu-id="f40a0-278">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f40a0-279">1.0</span><span class="sxs-lookup"><span data-stu-id="f40a0-279">1.0</span></span>|
|[<span data-ttu-id="f40a0-280">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="f40a0-280">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f40a0-281">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f40a0-281">ReadItem</span></span>|
|[<span data-ttu-id="f40a0-282">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="f40a0-282">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f40a0-283">阅读</span><span class="sxs-lookup"><span data-stu-id="f40a0-283">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="f40a0-284">示例</span><span class="sxs-lookup"><span data-stu-id="f40a0-284">Example</span></span>

```
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

#### <a name="itemclass-string"></a><span data-ttu-id="f40a0-285">itemClass :String</span><span class="sxs-lookup"><span data-stu-id="f40a0-285">itemClass :String</span></span>

<span data-ttu-id="f40a0-p115">获取选定项目的 Exchange Web 服务项目类。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="f40a0-p115">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="f40a0-p116">`itemClass` 属性指定所选项目的邮件类别。以下是邮件或约会项目的默认邮件类别。</span><span class="sxs-lookup"><span data-stu-id="f40a0-p116">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

| <span data-ttu-id="f40a0-290">类型</span><span class="sxs-lookup"><span data-stu-id="f40a0-290">Type</span></span> | <span data-ttu-id="f40a0-291">说明</span><span class="sxs-lookup"><span data-stu-id="f40a0-291">Description</span></span> | <span data-ttu-id="f40a0-292">项目类</span><span class="sxs-lookup"><span data-stu-id="f40a0-292">item class</span></span> |
| --- | --- | --- |
| <span data-ttu-id="f40a0-293">约会项目</span><span class="sxs-lookup"><span data-stu-id="f40a0-293">Appointment items</span></span> | <span data-ttu-id="f40a0-294">这些是项目类 `IPM.Appointment` 或 `IPM.Appointment.Occurence` 的日历项目。</span><span class="sxs-lookup"><span data-stu-id="f40a0-294">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurence`.</span></span> | `IPM.Appointment`<br />`IPM.Appointment.Occurence` |
| <span data-ttu-id="f40a0-295">邮件项目</span><span class="sxs-lookup"><span data-stu-id="f40a0-295">Message items</span></span> | <span data-ttu-id="f40a0-296">这些项目包括具有默认邮件类别 `IPM.Note` 的电子邮件，以及使用 `IPM.Schedule.Meeting` 作为基础邮件类别的会议请求、响应和取消。</span><span class="sxs-lookup"><span data-stu-id="f40a0-296">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span> | `IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled` |

<span data-ttu-id="f40a0-297">你可以创建用于扩展默认邮件类别的自定义邮件类别，例如，自定义约会邮件类别 `IPM.Appointment.Contoso`。</span><span class="sxs-lookup"><span data-stu-id="f40a0-297">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="f40a0-298">类型：</span><span class="sxs-lookup"><span data-stu-id="f40a0-298">Type:</span></span>

*   <span data-ttu-id="f40a0-299">字符串</span><span class="sxs-lookup"><span data-stu-id="f40a0-299">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="f40a0-300">要求</span><span class="sxs-lookup"><span data-stu-id="f40a0-300">Requirements</span></span>

|<span data-ttu-id="f40a0-301">要求</span><span class="sxs-lookup"><span data-stu-id="f40a0-301">Requirement</span></span>| <span data-ttu-id="f40a0-302">值</span><span class="sxs-lookup"><span data-stu-id="f40a0-302">Value</span></span>|
|---|---|
|[<span data-ttu-id="f40a0-303">最低邮箱要求集版本</span><span class="sxs-lookup"><span data-stu-id="f40a0-303">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f40a0-304">1.0</span><span class="sxs-lookup"><span data-stu-id="f40a0-304">1.0</span></span>|
|[<span data-ttu-id="f40a0-305">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="f40a0-305">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f40a0-306">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f40a0-306">ReadItem</span></span>|
|[<span data-ttu-id="f40a0-307">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="f40a0-307">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f40a0-308">阅读</span><span class="sxs-lookup"><span data-stu-id="f40a0-308">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="f40a0-309">示例</span><span class="sxs-lookup"><span data-stu-id="f40a0-309">Example</span></span>

```
var itemClass = Office.context.mailbox.item.itemClass;
```

#### <a name="nullable-itemid-string"></a><span data-ttu-id="f40a0-310">（可空类型）itemId ：字符串</span><span class="sxs-lookup"><span data-stu-id="f40a0-310">(nullable) itemId :String</span></span>

<span data-ttu-id="f40a0-p117">获取当前项目的 Exchange Web 服务项目标识符。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="f40a0-p117">Gets the Exchange Web Services item identifier for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="f40a0-313">`itemId` 属性返回的标识符与 Exchange Web 服务项目标识符相同。</span><span class="sxs-lookup"><span data-stu-id="f40a0-313">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="f40a0-314">`itemId` 属性与 Outlook 条目 ID 或使用 Outlook REST API 的 ID不同。</span><span class="sxs-lookup"><span data-stu-id="f40a0-314">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="f40a0-315">使用此值的 REST API 调用之前，应使用 [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string)将其转换。</span><span class="sxs-lookup"><span data-stu-id="f40a0-315">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="f40a0-316">有关详细信息，请参阅 [从 Outlook 外接程序使用 Outlook REST API](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id)。</span><span class="sxs-lookup"><span data-stu-id="f40a0-316">For more details, see [Use the Outlook REST APIs from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

<span data-ttu-id="f40a0-p119">`itemId` 属性在撰写模式下不可用。如果需要项目标识符，[`saveAsync`](#saveasyncoptions-callback) 方法可用于将项目保存到存储，这将在回调函数的 [`AsyncResult.value`](/javascript/api/office/office.asyncresult) 参数中返回项目标识符。</span><span class="sxs-lookup"><span data-stu-id="f40a0-p119">The `itemId` property is not available in compose mode. If an item identifier is required, the [`saveAsync`](#saveasyncoptions-callback) method can be used to save the item to the store, which will return the item identifier in the [`AsyncResult.value`](/javascript/api/office/office.asyncresult) parameter in the callback function.</span></span>

##### <a name="type"></a><span data-ttu-id="f40a0-319">类型：</span><span class="sxs-lookup"><span data-stu-id="f40a0-319">Type:</span></span>

*   <span data-ttu-id="f40a0-320">String</span><span class="sxs-lookup"><span data-stu-id="f40a0-320">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="f40a0-321">要求</span><span class="sxs-lookup"><span data-stu-id="f40a0-321">Requirements</span></span>

|<span data-ttu-id="f40a0-322">要求</span><span class="sxs-lookup"><span data-stu-id="f40a0-322">Requirement</span></span>| <span data-ttu-id="f40a0-323">值</span><span class="sxs-lookup"><span data-stu-id="f40a0-323">Value</span></span>|
|---|---|
|[<span data-ttu-id="f40a0-324">最低邮箱要求集版本</span><span class="sxs-lookup"><span data-stu-id="f40a0-324">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f40a0-325">1.0</span><span class="sxs-lookup"><span data-stu-id="f40a0-325">1.0</span></span>|
|[<span data-ttu-id="f40a0-326">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="f40a0-326">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f40a0-327">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f40a0-327">ReadItem</span></span>|
|[<span data-ttu-id="f40a0-328">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="f40a0-328">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f40a0-329">阅读</span><span class="sxs-lookup"><span data-stu-id="f40a0-329">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="f40a0-330">示例</span><span class="sxs-lookup"><span data-stu-id="f40a0-330">Example</span></span>

<span data-ttu-id="f40a0-p120">以下代码检查项目标识符是否存在。如果 `itemId` 属性返回 `null` 或 `undefined`，则将项目保存到存储，并从异步结果中获取项目标识符。</span><span class="sxs-lookup"><span data-stu-id="f40a0-p120">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

```
var itemId = Office.context.mailbox.item.itemId;
if (itemId === null || itemId == undefined) {
  Office.context.mailbox.item.saveAsync(function(result){
    itemId = result.value;
  });
}
```

####  <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlook14officemailboxenumsitemtype"></a><span data-ttu-id="f40a0-333">itemType:[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_4/office.mailboxenums.itemtype)</span><span class="sxs-lookup"><span data-stu-id="f40a0-333">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_4/office.mailboxenums.itemtype)</span></span>

<span data-ttu-id="f40a0-334">获取实例代表项的类型。</span><span class="sxs-lookup"><span data-stu-id="f40a0-334">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="f40a0-335">`itemType` 属性返回其中一个 `ItemType` 枚举值，指示 `item` 对象实例是邮件还是约会。</span><span class="sxs-lookup"><span data-stu-id="f40a0-335">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="f40a0-336">类型：</span><span class="sxs-lookup"><span data-stu-id="f40a0-336">Type:</span></span>

*   [<span data-ttu-id="f40a0-337">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="f40a0-337">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook_1_4/office.mailboxenums.itemtype)

##### <a name="requirements"></a><span data-ttu-id="f40a0-338">要求</span><span class="sxs-lookup"><span data-stu-id="f40a0-338">Requirements</span></span>

|<span data-ttu-id="f40a0-339">要求</span><span class="sxs-lookup"><span data-stu-id="f40a0-339">Requirement</span></span>| <span data-ttu-id="f40a0-340">值</span><span class="sxs-lookup"><span data-stu-id="f40a0-340">Value</span></span>|
|---|---|
|[<span data-ttu-id="f40a0-341">最低邮箱要求集版本</span><span class="sxs-lookup"><span data-stu-id="f40a0-341">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f40a0-342">1.0</span><span class="sxs-lookup"><span data-stu-id="f40a0-342">1.0</span></span>|
|[<span data-ttu-id="f40a0-343">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="f40a0-343">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f40a0-344">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f40a0-344">ReadItem</span></span>|
|[<span data-ttu-id="f40a0-345">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="f40a0-345">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f40a0-346">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="f40a0-346">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="f40a0-347">示例</span><span class="sxs-lookup"><span data-stu-id="f40a0-347">Example</span></span>

```
if (Office.context.mailbox.item.itemType == Office.MailboxEnums.ItemType.Message)
  // do something
else
  // do something else
```

####  <a name="location-stringlocationjavascriptapioutlook14officelocation"></a><span data-ttu-id="f40a0-348">location :String|[Location](/javascript/api/outlook_1_4/office.location)</span><span class="sxs-lookup"><span data-stu-id="f40a0-348">location :String|[Location](/javascript/api/outlook_1_4/office.location)</span></span>

<span data-ttu-id="f40a0-349">获取或设置约会的位置。</span><span class="sxs-lookup"><span data-stu-id="f40a0-349">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="f40a0-350">阅读模式</span><span class="sxs-lookup"><span data-stu-id="f40a0-350">Read mode</span></span>

<span data-ttu-id="f40a0-351">`location` 属性返回一个包含约会位置的字符串。</span><span class="sxs-lookup"><span data-stu-id="f40a0-351">The `location` property returns a string that contains the location of the appointment.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="f40a0-352">撰写模式</span><span class="sxs-lookup"><span data-stu-id="f40a0-352">Compose mode</span></span>

<span data-ttu-id="f40a0-353">`location` 属性返回一个 `Location` 对象，该对象提供用于获取和设置约会位置的方法。</span><span class="sxs-lookup"><span data-stu-id="f40a0-353">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="f40a0-354">类型：</span><span class="sxs-lookup"><span data-stu-id="f40a0-354">Type:</span></span>

*   <span data-ttu-id="f40a0-355">字符串 | [位置](/javascript/api/outlook_1_4/office.location)</span><span class="sxs-lookup"><span data-stu-id="f40a0-355">String | [Location](/javascript/api/outlook_1_4/office.location)</span></span>

##### <a name="requirements"></a><span data-ttu-id="f40a0-356">要求</span><span class="sxs-lookup"><span data-stu-id="f40a0-356">Requirements</span></span>

|<span data-ttu-id="f40a0-357">要求</span><span class="sxs-lookup"><span data-stu-id="f40a0-357">Requirement</span></span>| <span data-ttu-id="f40a0-358">值</span><span class="sxs-lookup"><span data-stu-id="f40a0-358">Value</span></span>|
|---|---|
|[<span data-ttu-id="f40a0-359">最低邮箱要求集版本</span><span class="sxs-lookup"><span data-stu-id="f40a0-359">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f40a0-360">1.0</span><span class="sxs-lookup"><span data-stu-id="f40a0-360">1.0</span></span>|
|[<span data-ttu-id="f40a0-361">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="f40a0-361">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f40a0-362">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f40a0-362">ReadItem</span></span>|
|[<span data-ttu-id="f40a0-363">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="f40a0-363">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f40a0-364">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="f40a0-364">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="f40a0-365">示例</span><span class="sxs-lookup"><span data-stu-id="f40a0-365">Example</span></span>

```
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

#### <a name="normalizedsubject-string"></a><span data-ttu-id="f40a0-366">normalizedSubject :String</span><span class="sxs-lookup"><span data-stu-id="f40a0-366">normalizedSubject :String</span></span>

<span data-ttu-id="f40a0-p121">获取删除了所有前缀（包括 `RE:` 和 `FWD:`）的项目主题。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="f40a0-p121">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="f40a0-p122">normalizedSubject 属性获取包含由电子邮件程序添加的任何标准前缀（如 `RE:` 和 `FW:`）的项目主题。若要获取包含完整前缀的项目主题，请使用 [`subject`](#subject-stringsubjectjavascriptapioutlook14officesubject) 属性。</span><span class="sxs-lookup"><span data-stu-id="f40a0-p122">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubjectjavascriptapioutlook14officesubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="f40a0-371">类型：</span><span class="sxs-lookup"><span data-stu-id="f40a0-371">Type:</span></span>

*   <span data-ttu-id="f40a0-372">String</span><span class="sxs-lookup"><span data-stu-id="f40a0-372">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="f40a0-373">要求</span><span class="sxs-lookup"><span data-stu-id="f40a0-373">Requirements</span></span>

|<span data-ttu-id="f40a0-374">要求</span><span class="sxs-lookup"><span data-stu-id="f40a0-374">Requirement</span></span>| <span data-ttu-id="f40a0-375">值</span><span class="sxs-lookup"><span data-stu-id="f40a0-375">Value</span></span>|
|---|---|
|[<span data-ttu-id="f40a0-376">最低邮箱要求集版本</span><span class="sxs-lookup"><span data-stu-id="f40a0-376">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f40a0-377">1.0</span><span class="sxs-lookup"><span data-stu-id="f40a0-377">1.0</span></span>|
|[<span data-ttu-id="f40a0-378">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="f40a0-378">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f40a0-379">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f40a0-379">ReadItem</span></span>|
|[<span data-ttu-id="f40a0-380">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="f40a0-380">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f40a0-381">阅读</span><span class="sxs-lookup"><span data-stu-id="f40a0-381">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="f40a0-382">示例</span><span class="sxs-lookup"><span data-stu-id="f40a0-382">Example</span></span>

```
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
```

####  <a name="notificationmessages-notificationmessagesjavascriptapioutlook14officenotificationmessages"></a><span data-ttu-id="f40a0-383">notificationMessages:[NotificationMessages](/javascript/api/outlook_1_4/office.notificationmessages)</span><span class="sxs-lookup"><span data-stu-id="f40a0-383">notificationMessages :[NotificationMessages](/javascript/api/outlook_1_4/office.notificationmessages)</span></span>

<span data-ttu-id="f40a0-384">获取一个项目的通知邮件。</span><span class="sxs-lookup"><span data-stu-id="f40a0-384">Gets the notification messages for an item.</span></span>

##### <a name="type"></a><span data-ttu-id="f40a0-385">类型：</span><span class="sxs-lookup"><span data-stu-id="f40a0-385">Type:</span></span>

*   [<span data-ttu-id="f40a0-386">NotificationMessages</span><span class="sxs-lookup"><span data-stu-id="f40a0-386">NotificationMessages</span></span>](/javascript/api/outlook_1_4/office.notificationmessages)

##### <a name="requirements"></a><span data-ttu-id="f40a0-387">要求</span><span class="sxs-lookup"><span data-stu-id="f40a0-387">Requirements</span></span>

|<span data-ttu-id="f40a0-388">要求</span><span class="sxs-lookup"><span data-stu-id="f40a0-388">Requirement</span></span>| <span data-ttu-id="f40a0-389">值</span><span class="sxs-lookup"><span data-stu-id="f40a0-389">Value</span></span>|
|---|---|
|[<span data-ttu-id="f40a0-390">最低邮箱要求集版本</span><span class="sxs-lookup"><span data-stu-id="f40a0-390">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f40a0-391">1.3</span><span class="sxs-lookup"><span data-stu-id="f40a0-391">1.3</span></span>|
|[<span data-ttu-id="f40a0-392">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="f40a0-392">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f40a0-393">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f40a0-393">ReadItem</span></span>|
|[<span data-ttu-id="f40a0-394">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="f40a0-394">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f40a0-395">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="f40a0-395">Compose or read</span></span>|

####  <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlook14officeemailaddressdetailsrecipientsjavascriptapioutlook14officerecipients"></a><span data-ttu-id="f40a0-396">optionalAttendees :数组.<[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)>|[收件人](/javascript/api/outlook_1_4/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="f40a0-396">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_4/office.recipients)</span></span>

<span data-ttu-id="f40a0-397">提供对事件可选与会者的访问。</span><span class="sxs-lookup"><span data-stu-id="f40a0-397">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="f40a0-398">对象的类型和访问级别取决于当前项的模式。</span><span class="sxs-lookup"><span data-stu-id="f40a0-398">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="f40a0-399">阅读模式</span><span class="sxs-lookup"><span data-stu-id="f40a0-399">Read mode</span></span>

<span data-ttu-id="f40a0-400">`optionalAttendees` 属性返回一个数组，其中包含每个可选与会者的 `EmailAddressDetails` 对象。</span><span class="sxs-lookup"><span data-stu-id="f40a0-400">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="f40a0-401">撰写模式</span><span class="sxs-lookup"><span data-stu-id="f40a0-401">Compose mode</span></span>

<span data-ttu-id="f40a0-402">`optionalAttendees` 属性返回一个 `Recipients` 对象，该对象提供用于获取和设置可选与会者的方法。</span><span class="sxs-lookup"><span data-stu-id="f40a0-402">The `optionalAttendees` property returns a `Recipients` object that provides methods to get and set the optional attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="f40a0-403">类型：</span><span class="sxs-lookup"><span data-stu-id="f40a0-403">Type:</span></span>

*   <span data-ttu-id="f40a0-404">数组。 <[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)> |[收件人](/javascript/api/outlook_1_4/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="f40a0-404">Array.<[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_4/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="f40a0-405">要求</span><span class="sxs-lookup"><span data-stu-id="f40a0-405">Requirements</span></span>

|<span data-ttu-id="f40a0-406">要求</span><span class="sxs-lookup"><span data-stu-id="f40a0-406">Requirement</span></span>| <span data-ttu-id="f40a0-407">值</span><span class="sxs-lookup"><span data-stu-id="f40a0-407">Value</span></span>|
|---|---|
|[<span data-ttu-id="f40a0-408">最低邮箱要求集版本</span><span class="sxs-lookup"><span data-stu-id="f40a0-408">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f40a0-409">1.0</span><span class="sxs-lookup"><span data-stu-id="f40a0-409">1.0</span></span>|
|[<span data-ttu-id="f40a0-410">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="f40a0-410">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f40a0-411">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f40a0-411">ReadItem</span></span>|
|[<span data-ttu-id="f40a0-412">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="f40a0-412">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f40a0-413">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="f40a0-413">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="f40a0-414">示例</span><span class="sxs-lookup"><span data-stu-id="f40a0-414">Example</span></span>

```
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

#### <a name="organizer-emailaddressdetailsjavascriptapioutlook14officeemailaddressdetails"></a><span data-ttu-id="f40a0-415">组织者：[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="f40a0-415">organizer :[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)</span></span>

<span data-ttu-id="f40a0-p124">获取指定会议的会议组织者的电子邮件地址。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="f40a0-p124">Gets the email address of the meeting organizer for a specified meeting. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="f40a0-418">类型：</span><span class="sxs-lookup"><span data-stu-id="f40a0-418">Type:</span></span>

*   [<span data-ttu-id="f40a0-419">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="f40a0-419">EmailAddressDetails</span></span>](/javascript/api/outlook_1_4/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="f40a0-420">要求</span><span class="sxs-lookup"><span data-stu-id="f40a0-420">Requirements</span></span>

|<span data-ttu-id="f40a0-421">要求</span><span class="sxs-lookup"><span data-stu-id="f40a0-421">Requirement</span></span>| <span data-ttu-id="f40a0-422">值</span><span class="sxs-lookup"><span data-stu-id="f40a0-422">Value</span></span>|
|---|---|
|[<span data-ttu-id="f40a0-423">最低邮箱要求集版本</span><span class="sxs-lookup"><span data-stu-id="f40a0-423">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f40a0-424">1.0</span><span class="sxs-lookup"><span data-stu-id="f40a0-424">1.0</span></span>|
|[<span data-ttu-id="f40a0-425">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="f40a0-425">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f40a0-426">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f40a0-426">ReadItem</span></span>|
|[<span data-ttu-id="f40a0-427">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="f40a0-427">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f40a0-428">阅读</span><span class="sxs-lookup"><span data-stu-id="f40a0-428">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="f40a0-429">示例</span><span class="sxs-lookup"><span data-stu-id="f40a0-429">Example</span></span>

```
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
```

####  <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlook14officeemailaddressdetailsrecipientsjavascriptapioutlook14officerecipients"></a><span data-ttu-id="f40a0-430">requiredAttendees： 数组。 <[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)>|[收件人](/javascript/api/outlook_1_4/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="f40a0-430">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_4/office.recipients)</span></span>

<span data-ttu-id="f40a0-431">提供对事件必需与会者的访问。</span><span class="sxs-lookup"><span data-stu-id="f40a0-431">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="f40a0-432">对象的类型和访问级别取决于当前项的模式。</span><span class="sxs-lookup"><span data-stu-id="f40a0-432">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="f40a0-433">阅读模式</span><span class="sxs-lookup"><span data-stu-id="f40a0-433">Read mode</span></span>

<span data-ttu-id="f40a0-434">`requiredAttendees` 属性返回一个数组，其中包含每个必需与会者的 `EmailAddressDetails` 对象。</span><span class="sxs-lookup"><span data-stu-id="f40a0-434">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="f40a0-435">撰写模式</span><span class="sxs-lookup"><span data-stu-id="f40a0-435">Compose mode</span></span>

<span data-ttu-id="f40a0-436">`requiredAttendees` 属性返回一个 `Recipients` 对象，该对象提供用于获取和设置可选与会者的方法。</span><span class="sxs-lookup"><span data-stu-id="f40a0-436">The `requiredAttendees` property returns a `Recipients` object that provides methods to get and set the required attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="f40a0-437">类型：</span><span class="sxs-lookup"><span data-stu-id="f40a0-437">Type:</span></span>

*   <span data-ttu-id="f40a0-438">数组。 <[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)> |[收件人](/javascript/api/outlook_1_4/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="f40a0-438">Array.<[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_4/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="f40a0-439">要求</span><span class="sxs-lookup"><span data-stu-id="f40a0-439">Requirements</span></span>

|<span data-ttu-id="f40a0-440">要求</span><span class="sxs-lookup"><span data-stu-id="f40a0-440">Requirement</span></span>| <span data-ttu-id="f40a0-441">值</span><span class="sxs-lookup"><span data-stu-id="f40a0-441">Value</span></span>|
|---|---|
|[<span data-ttu-id="f40a0-442">最低邮箱要求集版本</span><span class="sxs-lookup"><span data-stu-id="f40a0-442">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f40a0-443">1.0</span><span class="sxs-lookup"><span data-stu-id="f40a0-443">1.0</span></span>|
|[<span data-ttu-id="f40a0-444">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="f40a0-444">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f40a0-445">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f40a0-445">ReadItem</span></span>|
|[<span data-ttu-id="f40a0-446">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="f40a0-446">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f40a0-447">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="f40a0-447">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="f40a0-448">示例</span><span class="sxs-lookup"><span data-stu-id="f40a0-448">Example</span></span>

```
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
}
```

#### <a name="sender-emailaddressdetailsjavascriptapioutlook14officeemailaddressdetails"></a><span data-ttu-id="f40a0-449">发件人：[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="f40a0-449">sender :[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)</span></span>

<span data-ttu-id="f40a0-p126">获取电子邮件发件人的电子邮件地址。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="f40a0-p126">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="f40a0-p127">[`from`](#from-emailaddressdetailsjavascriptapioutlook14officeemailaddressdetails) 和 `sender` 属性表示同一个人，邮件由代理人发送的除外。在此情况下，`from` 属性表示代理程序，而 sender 属性表示代理人。</span><span class="sxs-lookup"><span data-stu-id="f40a0-p127">The [`from`](#from-emailaddressdetailsjavascriptapioutlook14officeemailaddressdetails) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="f40a0-454">`EmailAddressDetails` 对象的 `recipientType` 属性在 `sender` 属性是 `undefined`。</span><span class="sxs-lookup"><span data-stu-id="f40a0-454">Note: The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="f40a0-455">类型：</span><span class="sxs-lookup"><span data-stu-id="f40a0-455">Type:</span></span>

*   [<span data-ttu-id="f40a0-456">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="f40a0-456">EmailAddressDetails</span></span>](/javascript/api/outlook_1_4/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="f40a0-457">要求</span><span class="sxs-lookup"><span data-stu-id="f40a0-457">Requirements</span></span>

|<span data-ttu-id="f40a0-458">要求</span><span class="sxs-lookup"><span data-stu-id="f40a0-458">Requirement</span></span>| <span data-ttu-id="f40a0-459">值</span><span class="sxs-lookup"><span data-stu-id="f40a0-459">Value</span></span>|
|---|---|
|[<span data-ttu-id="f40a0-460">最低邮箱要求集版本</span><span class="sxs-lookup"><span data-stu-id="f40a0-460">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f40a0-461">1.0</span><span class="sxs-lookup"><span data-stu-id="f40a0-461">1.0</span></span>|
|[<span data-ttu-id="f40a0-462">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="f40a0-462">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f40a0-463">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f40a0-463">ReadItem</span></span>|
|[<span data-ttu-id="f40a0-464">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="f40a0-464">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f40a0-465">阅读</span><span class="sxs-lookup"><span data-stu-id="f40a0-465">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="f40a0-466">示例</span><span class="sxs-lookup"><span data-stu-id="f40a0-466">Example</span></span>

```
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
```

####  <a name="start-datetimejavascriptapioutlook14officetime"></a><span data-ttu-id="f40a0-467">start :Date|[Time](/javascript/api/outlook_1_4/office.time)</span><span class="sxs-lookup"><span data-stu-id="f40a0-467">start :Date|[Time](/javascript/api/outlook_1_4/office.time)</span></span>

<span data-ttu-id="f40a0-468">获取或设置约会开始的日期和时间。</span><span class="sxs-lookup"><span data-stu-id="f40a0-468">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="f40a0-p128">将 `start` 属性表示为协调世界时 (UTC) 的日期和时间值。可使用 [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook14officelocalclienttime) 方法将该值转换为客户端的本地日期和时间。</span><span class="sxs-lookup"><span data-stu-id="f40a0-p128">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook14officelocalclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="f40a0-471">阅读模式</span><span class="sxs-lookup"><span data-stu-id="f40a0-471">Read mode</span></span>

<span data-ttu-id="f40a0-472">`start` 属性返回 `Date` 对象。</span><span class="sxs-lookup"><span data-stu-id="f40a0-472">The `start` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="f40a0-473">撰写模式</span><span class="sxs-lookup"><span data-stu-id="f40a0-473">Compose mode</span></span>

<span data-ttu-id="f40a0-474">`start` 属性返回 `Time` 对象。</span><span class="sxs-lookup"><span data-stu-id="f40a0-474">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="f40a0-475">使用 [`Time.setAsync`](/javascript/api/outlook_1_4/office.time#setasync-datetime--options--callback-) 方法设置开始时间时，应使用 [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) 方法将客户端的本地时间转换为服务器的 UTC。</span><span class="sxs-lookup"><span data-stu-id="f40a0-475">When you use the [`Time.setAsync`](/javascript/api/outlook_1_4/office.time#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="f40a0-476">类型：</span><span class="sxs-lookup"><span data-stu-id="f40a0-476">Type:</span></span>

*   <span data-ttu-id="f40a0-477">日期 | [时间](/javascript/api/outlook_1_4/office.time)</span><span class="sxs-lookup"><span data-stu-id="f40a0-477">Date | [Time](/javascript/api/outlook_1_4/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="f40a0-478">要求</span><span class="sxs-lookup"><span data-stu-id="f40a0-478">Requirements</span></span>

|<span data-ttu-id="f40a0-479">要求</span><span class="sxs-lookup"><span data-stu-id="f40a0-479">Requirement</span></span>| <span data-ttu-id="f40a0-480">值</span><span class="sxs-lookup"><span data-stu-id="f40a0-480">Value</span></span>|
|---|---|
|[<span data-ttu-id="f40a0-481">最低邮箱要求集版本</span><span class="sxs-lookup"><span data-stu-id="f40a0-481">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f40a0-482">1.0</span><span class="sxs-lookup"><span data-stu-id="f40a0-482">1.0</span></span>|
|[<span data-ttu-id="f40a0-483">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="f40a0-483">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f40a0-484">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f40a0-484">ReadItem</span></span>|
|[<span data-ttu-id="f40a0-485">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="f40a0-485">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f40a0-486">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="f40a0-486">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="f40a0-487">示例</span><span class="sxs-lookup"><span data-stu-id="f40a0-487">Example</span></span>

<span data-ttu-id="f40a0-488">以下示例通过使用 `Time` 对象的 [`setAsync`](/javascript/api/outlook_1_4/office.time#setasync-datetime--options--callback-) 方法，设置撰写模式下约会的开始时间。</span><span class="sxs-lookup"><span data-stu-id="f40a0-488">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_4/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

####  <a name="subject-stringsubjectjavascriptapioutlook14officesubject"></a><span data-ttu-id="f40a0-489">主题： 字符串 |[主题](/javascript/api/outlook_1_4/office.subject)</span><span class="sxs-lookup"><span data-stu-id="f40a0-489">subject :String|[Subject](/javascript/api/outlook_1_4/office.subject)</span></span>

<span data-ttu-id="f40a0-490">获取或设置显示在项目的主题字段中的说明。</span><span class="sxs-lookup"><span data-stu-id="f40a0-490">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="f40a0-491">`subject` 属性获取或设置由电子邮件服务器发送项目时的整个主题。</span><span class="sxs-lookup"><span data-stu-id="f40a0-491">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="f40a0-492">阅读模式</span><span class="sxs-lookup"><span data-stu-id="f40a0-492">Read mode</span></span>

<span data-ttu-id="f40a0-p129">`subject` 属性返回一个字符串。使用 [`normalizedSubject`](#normalizedsubject-string) 属性获取不带任何前导前缀（如 `RE:` 和 `FW:`）的主题。</span><span class="sxs-lookup"><span data-stu-id="f40a0-p129">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

```
var subject = Office.context.mailbox.item.subject;
```

##### <a name="compose-mode"></a><span data-ttu-id="f40a0-495">撰写模式</span><span class="sxs-lookup"><span data-stu-id="f40a0-495">Compose mode</span></span>

<span data-ttu-id="f40a0-496">`subject` 属性返回一个 `Subject` 对象，该对象提供用于获取和设置主题的方法。</span><span class="sxs-lookup"><span data-stu-id="f40a0-496">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="f40a0-497">类型：</span><span class="sxs-lookup"><span data-stu-id="f40a0-497">Type:</span></span>

*   <span data-ttu-id="f40a0-498">字符串 | [主题](/javascript/api/outlook_1_4/office.subject)</span><span class="sxs-lookup"><span data-stu-id="f40a0-498">String | [Subject](/javascript/api/outlook_1_4/office.subject)</span></span>

##### <a name="requirements"></a><span data-ttu-id="f40a0-499">要求</span><span class="sxs-lookup"><span data-stu-id="f40a0-499">Requirements</span></span>

|<span data-ttu-id="f40a0-500">要求</span><span class="sxs-lookup"><span data-stu-id="f40a0-500">Requirement</span></span>| <span data-ttu-id="f40a0-501">值</span><span class="sxs-lookup"><span data-stu-id="f40a0-501">Value</span></span>|
|---|---|
|[<span data-ttu-id="f40a0-502">最低邮箱要求集版本</span><span class="sxs-lookup"><span data-stu-id="f40a0-502">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f40a0-503">1.0</span><span class="sxs-lookup"><span data-stu-id="f40a0-503">1.0</span></span>|
|[<span data-ttu-id="f40a0-504">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="f40a0-504">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f40a0-505">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f40a0-505">ReadItem</span></span>|
|[<span data-ttu-id="f40a0-506">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="f40a0-506">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f40a0-507">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="f40a0-507">Compose or read</span></span>|

####  <a name="to-arrayemailaddressdetailsjavascriptapioutlook14officeemailaddressdetailsrecipientsjavascriptapioutlook14officerecipients"></a><span data-ttu-id="f40a0-508">发送： 数组。 <[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)>|[收件人](/javascript/api/outlook_1_4/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="f40a0-508">to :Array.<[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_4/office.recipients)</span></span>

<span data-ttu-id="f40a0-509">提供对邮件的 **发送** 行上收件人的访问。</span><span class="sxs-lookup"><span data-stu-id="f40a0-509">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="f40a0-510">对象的类型和访问级别取决于当前项的模式。</span><span class="sxs-lookup"><span data-stu-id="f40a0-510">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="f40a0-511">阅读模式</span><span class="sxs-lookup"><span data-stu-id="f40a0-511">Read mode</span></span>

<span data-ttu-id="f40a0-p131">`to` 属性返回包含邮件的**收件人**行上所列的每个收件人的 `EmailAddressDetails` 对象的数组。集合上限为 100 个成员。</span><span class="sxs-lookup"><span data-stu-id="f40a0-p131">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="f40a0-514">撰写模式</span><span class="sxs-lookup"><span data-stu-id="f40a0-514">Compose mode</span></span>

<span data-ttu-id="f40a0-515">`to` 属性返回 `Recipients` 对象，该对象提供用于处理邮件**收件人**行上收件人的方法。</span><span class="sxs-lookup"><span data-stu-id="f40a0-515">The to`to` property returns a Recipients`Recipients` object that provides methods for manipulating the recipients on the **To** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="f40a0-516">类型：</span><span class="sxs-lookup"><span data-stu-id="f40a0-516">Type:</span></span>

*   <span data-ttu-id="f40a0-517">数组。 <[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)> |[收件人](/javascript/api/outlook_1_4/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="f40a0-517">Array.<[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_4/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="f40a0-518">要求</span><span class="sxs-lookup"><span data-stu-id="f40a0-518">Requirements</span></span>

|<span data-ttu-id="f40a0-519">要求</span><span class="sxs-lookup"><span data-stu-id="f40a0-519">Requirement</span></span>| <span data-ttu-id="f40a0-520">值</span><span class="sxs-lookup"><span data-stu-id="f40a0-520">Value</span></span>|
|---|---|
|[<span data-ttu-id="f40a0-521">最低邮箱要求集版本</span><span class="sxs-lookup"><span data-stu-id="f40a0-521">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f40a0-522">1.0</span><span class="sxs-lookup"><span data-stu-id="f40a0-522">1.0</span></span>|
|[<span data-ttu-id="f40a0-523">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="f40a0-523">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f40a0-524">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f40a0-524">ReadItem</span></span>|
|[<span data-ttu-id="f40a0-525">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="f40a0-525">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f40a0-526">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="f40a0-526">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="f40a0-527">示例</span><span class="sxs-lookup"><span data-stu-id="f40a0-527">Example</span></span>

```
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

### <a name="methods"></a><span data-ttu-id="f40a0-528">方法</span><span class="sxs-lookup"><span data-stu-id="f40a0-528">Methods</span></span>

####  <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="f40a0-529">addFileAttachmentAsync (uri，attachmentName，[选项] [回调])</span><span class="sxs-lookup"><span data-stu-id="f40a0-529">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="f40a0-530">将文件作为附件添加到邮件或约会。</span><span class="sxs-lookup"><span data-stu-id="f40a0-530">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="f40a0-531">`addFileAttachmentAsync` 方法在指定的 URI 上载文件并将其附加到撰写窗体中的项目。</span><span class="sxs-lookup"><span data-stu-id="f40a0-531">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="f40a0-532">你随后可以将该标识符与 [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) 方法一同使用，以删除同一个会话中的附件。</span><span class="sxs-lookup"><span data-stu-id="f40a0-532">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="f40a0-533">参数：</span><span class="sxs-lookup"><span data-stu-id="f40a0-533">Parameters:</span></span>

|<span data-ttu-id="f40a0-534">名称</span><span class="sxs-lookup"><span data-stu-id="f40a0-534">Name</span></span>| <span data-ttu-id="f40a0-535">类型</span><span class="sxs-lookup"><span data-stu-id="f40a0-535">Type</span></span>| <span data-ttu-id="f40a0-536">属性</span><span class="sxs-lookup"><span data-stu-id="f40a0-536">Attributes</span></span>| <span data-ttu-id="f40a0-537">说明</span><span class="sxs-lookup"><span data-stu-id="f40a0-537">Description</span></span>|
|---|---|---|---|
|`uri`| <span data-ttu-id="f40a0-538">String</span><span class="sxs-lookup"><span data-stu-id="f40a0-538">String</span></span>||<span data-ttu-id="f40a0-p132">提供附加到邮件或约会的文件的位置的 URI。最大长度为 2048 个字符。</span><span class="sxs-lookup"><span data-stu-id="f40a0-p132">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="f40a0-541">String</span><span class="sxs-lookup"><span data-stu-id="f40a0-541">String</span></span>||<span data-ttu-id="f40a0-p133">在附件上载过程中显示的附件名称。最大长度为 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="f40a0-p133">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="f40a0-544">Object</span><span class="sxs-lookup"><span data-stu-id="f40a0-544">Object</span></span>| <span data-ttu-id="f40a0-545">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="f40a0-545">&lt;optional&gt;</span></span>|<span data-ttu-id="f40a0-546">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="f40a0-546">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="f40a0-547">Object</span><span class="sxs-lookup"><span data-stu-id="f40a0-547">Object</span></span>| <span data-ttu-id="f40a0-548">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="f40a0-548">&lt;optional&gt;</span></span>|<span data-ttu-id="f40a0-549">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="f40a0-549">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="f40a0-550">function</span><span class="sxs-lookup"><span data-stu-id="f40a0-550">function</span></span>| <span data-ttu-id="f40a0-551">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="f40a0-551">&lt;optional&gt;</span></span>|<span data-ttu-id="f40a0-552">方法完成后，使用单个参数 `asyncResult` （一个 [`AsyncResult`](/javascript/api/office/office.asyncresult)   对象）调用在 `callback`  参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="f40a0-552">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="f40a0-553">如果成功，附件标识符将在 `asyncResult.value` 属性中提供。</span><span class="sxs-lookup"><span data-stu-id="f40a0-553">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="f40a0-554">如果上传附件失败，`asyncResult` 对象将包含一个提供错误说明的 `Error` 对象。</span><span class="sxs-lookup"><span data-stu-id="f40a0-554">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="f40a0-555">错误</span><span class="sxs-lookup"><span data-stu-id="f40a0-555">Errors</span></span>

| <span data-ttu-id="f40a0-556">错误代码</span><span class="sxs-lookup"><span data-stu-id="f40a0-556">Error code</span></span> | <span data-ttu-id="f40a0-557">说明</span><span class="sxs-lookup"><span data-stu-id="f40a0-557">Description</span></span> |
|------------|-------------|
| `AttachmentSizeExceeded` | <span data-ttu-id="f40a0-558">附件大小超过了允许的大小。</span><span class="sxs-lookup"><span data-stu-id="f40a0-558">The attachment is larger than allowed.</span></span> |
| `FileTypeNotSupported` | <span data-ttu-id="f40a0-559">该附件的扩展名不是允许的扩展名。</span><span class="sxs-lookup"><span data-stu-id="f40a0-559">The attachment has an extension that is not allowed.</span></span> |
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="f40a0-560">邮件或约会具有的附件过多。</span><span class="sxs-lookup"><span data-stu-id="f40a0-560">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="f40a0-561">要求</span><span class="sxs-lookup"><span data-stu-id="f40a0-561">Requirements</span></span>

|<span data-ttu-id="f40a0-562">要求</span><span class="sxs-lookup"><span data-stu-id="f40a0-562">Requirement</span></span>| <span data-ttu-id="f40a0-563">值</span><span class="sxs-lookup"><span data-stu-id="f40a0-563">Value</span></span>|
|---|---|
|[<span data-ttu-id="f40a0-564">最低邮箱要求集版本</span><span class="sxs-lookup"><span data-stu-id="f40a0-564">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f40a0-565">1.1</span><span class="sxs-lookup"><span data-stu-id="f40a0-565">1.1</span></span>|
|[<span data-ttu-id="f40a0-566">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="f40a0-566">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f40a0-567">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="f40a0-567">ReadWriteItem</span></span>|
|[<span data-ttu-id="f40a0-568">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="f40a0-568">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f40a0-569">撰写</span><span class="sxs-lookup"><span data-stu-id="f40a0-569">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="f40a0-570">示例</span><span class="sxs-lookup"><span data-stu-id="f40a0-570">Example</span></span>

```
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

####  <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="f40a0-571">addItemAttachmentAsync (itemId，attachmentName，[选项] [回调])</span><span class="sxs-lookup"><span data-stu-id="f40a0-571">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="f40a0-572">将 Exchange 项目（如邮件）作为附件添加到邮件或约会。</span><span class="sxs-lookup"><span data-stu-id="f40a0-572">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="f40a0-p134">`addItemAttachmentAsync` 方法将包含指定 Exchange 标识符的项目附加到撰写窗体中的项目。如果指定一个回调方法，此方法使用 `asyncResult` 参数调用，该参数包含一个附件标识符或代码，指示附加项目过程中出现的任何错误。可以使用 `options` 参数将状态信息传递给回调方法（如果需要）。</span><span class="sxs-lookup"><span data-stu-id="f40a0-p134">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="f40a0-576">你随后可以将该标识符与 [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) 方法一同使用，以删除同一个会话中的附件。</span><span class="sxs-lookup"><span data-stu-id="f40a0-576">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="f40a0-577">如果 Office 外接程序在 Outlook Web App 中运行，则 `addItemAttachmentAsync` 方法可以将项目附加到项目（正在编辑的项目除外）中；然而，不支持也不建议这样做。</span><span class="sxs-lookup"><span data-stu-id="f40a0-577">If your Office add-in is running in Outlook Web App, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="f40a0-578">参数：</span><span class="sxs-lookup"><span data-stu-id="f40a0-578">Parameters:</span></span>

|<span data-ttu-id="f40a0-579">名称</span><span class="sxs-lookup"><span data-stu-id="f40a0-579">Name</span></span>| <span data-ttu-id="f40a0-580">类型</span><span class="sxs-lookup"><span data-stu-id="f40a0-580">Type</span></span>| <span data-ttu-id="f40a0-581">属性</span><span class="sxs-lookup"><span data-stu-id="f40a0-581">Attributes</span></span>| <span data-ttu-id="f40a0-582">说明</span><span class="sxs-lookup"><span data-stu-id="f40a0-582">Description</span></span>|
|---|---|---|---|
|`itemId`| <span data-ttu-id="f40a0-583">String</span><span class="sxs-lookup"><span data-stu-id="f40a0-583">String</span></span>||<span data-ttu-id="f40a0-p135">要附加项目的 Exchange 标识符。最大长度为 100 个字符。</span><span class="sxs-lookup"><span data-stu-id="f40a0-p135">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="f40a0-586">String</span><span class="sxs-lookup"><span data-stu-id="f40a0-586">String</span></span>||<span data-ttu-id="f40a0-p136">要附加项目的主题。最大长度为 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="f40a0-p136">The sujbect of the item to be attached. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="f40a0-589">Object</span><span class="sxs-lookup"><span data-stu-id="f40a0-589">Object</span></span>| <span data-ttu-id="f40a0-590">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="f40a0-590">&lt;optional&gt;</span></span>|<span data-ttu-id="f40a0-591">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="f40a0-591">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="f40a0-592">Object</span><span class="sxs-lookup"><span data-stu-id="f40a0-592">Object</span></span>| <span data-ttu-id="f40a0-593">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="f40a0-593">&lt;optional&gt;</span></span>|<span data-ttu-id="f40a0-594">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="f40a0-594">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="f40a0-595">function</span><span class="sxs-lookup"><span data-stu-id="f40a0-595">function</span></span>| <span data-ttu-id="f40a0-596">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="f40a0-596">&lt;optional&gt;</span></span>|<span data-ttu-id="f40a0-597">方法完成后，使用单个参数 `asyncResult` （一个 [`AsyncResult`](/javascript/api/office/office.asyncresult)   对象）调用在 `callback`  参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="f40a0-597">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="f40a0-598">如果成功，附件标识符将在 `asyncResult.value` 属性中提供。</span><span class="sxs-lookup"><span data-stu-id="f40a0-598">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="f40a0-599">如果添加附件失败，`asyncResult` 对象将包含一个提供错误说明的 `Error` 对象。</span><span class="sxs-lookup"><span data-stu-id="f40a0-599">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="f40a0-600">错误</span><span class="sxs-lookup"><span data-stu-id="f40a0-600">Errors</span></span>

| <span data-ttu-id="f40a0-601">错误代码</span><span class="sxs-lookup"><span data-stu-id="f40a0-601">Error code</span></span> | <span data-ttu-id="f40a0-602">说明</span><span class="sxs-lookup"><span data-stu-id="f40a0-602">Description</span></span> |
|------------|-------------|
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="f40a0-603">邮件或约会具有的附件过多。</span><span class="sxs-lookup"><span data-stu-id="f40a0-603">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="f40a0-604">要求</span><span class="sxs-lookup"><span data-stu-id="f40a0-604">Requirements</span></span>

|<span data-ttu-id="f40a0-605">要求</span><span class="sxs-lookup"><span data-stu-id="f40a0-605">Requirement</span></span>| <span data-ttu-id="f40a0-606">值</span><span class="sxs-lookup"><span data-stu-id="f40a0-606">Value</span></span>|
|---|---|
|[<span data-ttu-id="f40a0-607">最低邮箱要求集版本</span><span class="sxs-lookup"><span data-stu-id="f40a0-607">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f40a0-608">1.1</span><span class="sxs-lookup"><span data-stu-id="f40a0-608">1.1</span></span>|
|[<span data-ttu-id="f40a0-609">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="f40a0-609">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f40a0-610">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="f40a0-610">ReadWriteItem</span></span>|
|[<span data-ttu-id="f40a0-611">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="f40a0-611">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f40a0-612">撰写</span><span class="sxs-lookup"><span data-stu-id="f40a0-612">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="f40a0-613">示例</span><span class="sxs-lookup"><span data-stu-id="f40a0-613">Example</span></span>

<span data-ttu-id="f40a0-614">以下示例将现有的 Outlook 项目添加为名为 `My Attachment` 的附件。</span><span class="sxs-lookup"><span data-stu-id="f40a0-614">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

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

####  <a name="close"></a><span data-ttu-id="f40a0-615">close()</span><span class="sxs-lookup"><span data-stu-id="f40a0-615">close()</span></span>

<span data-ttu-id="f40a0-616">关闭当前正在撰写的项目。</span><span class="sxs-lookup"><span data-stu-id="f40a0-616">Closes the current item that is being composed.</span></span>

<span data-ttu-id="f40a0-p137">`close` 方法的行为取决于要撰写的项目的当前状态。如果项目具有未保存的更改，客户端将提示用户保存、放弃或取消关闭操作。</span><span class="sxs-lookup"><span data-stu-id="f40a0-p137">The behavior of the `close` method depends on the current state of the item being composed. If the item has unsaved changes, the client prompts the user to save, discard, or cancel the close action.</span></span>

> [!NOTE]
> <span data-ttu-id="f40a0-619">在 Outlook 网页版中，如果是约会项，并之前用`saveAsync` 保存过，会提示用户保存、放弃或取消，即使该项上一次保存后并未有任何更改。</span><span class="sxs-lookup"><span data-stu-id="f40a0-619">In Outlook on the web, if the item is an appointment and it has previously been saved using `saveAsync`, the user is prompted to save, discard, or cancel even if no changes have occurred since the item was last saved.</span></span>

<span data-ttu-id="f40a0-620">在 Outlook 桌面客户端中，如果邮件是内联答复，`close` 方法不起作用。</span><span class="sxs-lookup"><span data-stu-id="f40a0-620">In the Outlook desktop client, if the message is an inline reply, the `close` method has no effect.</span></span>

##### <a name="requirements"></a><span data-ttu-id="f40a0-621">要求</span><span class="sxs-lookup"><span data-stu-id="f40a0-621">Requirements</span></span>

|<span data-ttu-id="f40a0-622">要求</span><span class="sxs-lookup"><span data-stu-id="f40a0-622">Requirement</span></span>| <span data-ttu-id="f40a0-623">值</span><span class="sxs-lookup"><span data-stu-id="f40a0-623">Value</span></span>|
|---|---|
|[<span data-ttu-id="f40a0-624">最低邮箱要求集版本</span><span class="sxs-lookup"><span data-stu-id="f40a0-624">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f40a0-625">1.3</span><span class="sxs-lookup"><span data-stu-id="f40a0-625">1.3</span></span>|
|[<span data-ttu-id="f40a0-626">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="f40a0-626">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f40a0-627">Restricted</span><span class="sxs-lookup"><span data-stu-id="f40a0-627">Restricted</span></span>|
|[<span data-ttu-id="f40a0-628">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="f40a0-628">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f40a0-629">撰写</span><span class="sxs-lookup"><span data-stu-id="f40a0-629">Compose</span></span>|

#### <a name="displayreplyallformformdata"></a><span data-ttu-id="f40a0-630">displayReplyAllForm(formData)</span><span class="sxs-lookup"><span data-stu-id="f40a0-630">displayReplyAllForm(formData)</span></span>

<span data-ttu-id="f40a0-631">显示答复窗体，其中包括所选邮件的发件人和所有收件人或所选约会的组织者和所有与会者。</span><span class="sxs-lookup"><span data-stu-id="f40a0-631">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="f40a0-632">在 Outlook for iOS 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="f40a0-632">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="f40a0-633">在 Outlook Web App 中，答复窗体显示为包含 3 列视图的弹出式窗体以及包含 2 列或 1 列视图的弹出式窗体。</span><span class="sxs-lookup"><span data-stu-id="f40a0-633">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="f40a0-634">如果任意字符串参数超出其限制， `displayReplyAllForm` 将引发异常。</span><span class="sxs-lookup"><span data-stu-id="f40a0-634">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="f40a0-p138">当在 `formData.attachments` 参数中指定附件时，Outlook 和 Outlook Web App 尝试下载所有附件并将其附加到答复窗体。如果无法添加任何附件，则在窗体 UI 中显示错误。如果这不可能，则不引发错误消息。</span><span class="sxs-lookup"><span data-stu-id="f40a0-p138">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="f40a0-638">参数：</span><span class="sxs-lookup"><span data-stu-id="f40a0-638">Parameters:</span></span>

|<span data-ttu-id="f40a0-639">名称</span><span class="sxs-lookup"><span data-stu-id="f40a0-639">Name</span></span>| <span data-ttu-id="f40a0-640">类型</span><span class="sxs-lookup"><span data-stu-id="f40a0-640">Type</span></span>| <span data-ttu-id="f40a0-641">说明</span><span class="sxs-lookup"><span data-stu-id="f40a0-641">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="f40a0-642">字符串 | 对象</span><span class="sxs-lookup"><span data-stu-id="f40a0-642">String &#124; Object</span></span>| |<span data-ttu-id="f40a0-p139">一个包含文本和 HTML 且代表答复窗体正文的字符串。字符串限制为 32 KB。</span><span class="sxs-lookup"><span data-stu-id="f40a0-p139">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="f40a0-645">**OR**</span><span class="sxs-lookup"><span data-stu-id="f40a0-645">**OR**</span></span><br/><span data-ttu-id="f40a0-p140">包含正文或附件数据和回调函数的对象。对象定义如下。</span><span class="sxs-lookup"><span data-stu-id="f40a0-p140">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="f40a0-648">String</span><span class="sxs-lookup"><span data-stu-id="f40a0-648">String</span></span> | <span data-ttu-id="f40a0-649">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="f40a0-649">&lt;optional&gt;</span></span> | <span data-ttu-id="f40a0-p141">一个包含文本和 HTML 且表示答复窗体正文的字符串。字符串限制为 32 KB。</span><span class="sxs-lookup"><span data-stu-id="f40a0-p141">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="f40a0-652">数组。&lt;对象&gt;</span><span class="sxs-lookup"><span data-stu-id="f40a0-652">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="f40a0-653">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="f40a0-653">&lt;optional&gt;</span></span> | <span data-ttu-id="f40a0-654">JSON 对象（是文件或项目附件）的数组。</span><span class="sxs-lookup"><span data-stu-id="f40a0-654">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="f40a0-655">String</span><span class="sxs-lookup"><span data-stu-id="f40a0-655">String</span></span> | | <span data-ttu-id="f40a0-p142">指示附件的类型。必须是文件附件的 `file` 或项目附件的 `item` 。</span><span class="sxs-lookup"><span data-stu-id="f40a0-p142">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="f40a0-658">String</span><span class="sxs-lookup"><span data-stu-id="f40a0-658">String</span></span> | | <span data-ttu-id="f40a0-659">一个包含附件的名称的字符串，最多包含 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="f40a0-659">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="f40a0-660">String</span><span class="sxs-lookup"><span data-stu-id="f40a0-660">String</span></span> | | <span data-ttu-id="f40a0-p143">仅在将 `type` 设置为 `file` 时使用。文件位置的 URI。</span><span class="sxs-lookup"><span data-stu-id="f40a0-p143">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="f40a0-663">String</span><span class="sxs-lookup"><span data-stu-id="f40a0-663">String</span></span> | | <span data-ttu-id="f40a0-p144">仅在将 `type` 设置为 `item` 时使用。附件的 EWS 项目 ID。字符串最长为 100 个字符。</span><span class="sxs-lookup"><span data-stu-id="f40a0-p144">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="f40a0-667">function</span><span class="sxs-lookup"><span data-stu-id="f40a0-667">function</span></span> | <span data-ttu-id="f40a0-668">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="f40a0-668">&lt;optional&gt;</span></span> | <span data-ttu-id="f40a0-669">方法完成后，使用单个参数 `asyncResult`（一个 [AsyncResult](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="f40a0-669">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="f40a0-670">要求</span><span class="sxs-lookup"><span data-stu-id="f40a0-670">Requirements</span></span>

|<span data-ttu-id="f40a0-671">要求</span><span class="sxs-lookup"><span data-stu-id="f40a0-671">Requirement</span></span>| <span data-ttu-id="f40a0-672">值</span><span class="sxs-lookup"><span data-stu-id="f40a0-672">Value</span></span>|
|---|---|
|[<span data-ttu-id="f40a0-673">最低邮箱要求集版本</span><span class="sxs-lookup"><span data-stu-id="f40a0-673">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f40a0-674">1.0</span><span class="sxs-lookup"><span data-stu-id="f40a0-674">1.0</span></span>|
|[<span data-ttu-id="f40a0-675">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="f40a0-675">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f40a0-676">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f40a0-676">ReadItem</span></span>|
|[<span data-ttu-id="f40a0-677">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="f40a0-677">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f40a0-678">阅读</span><span class="sxs-lookup"><span data-stu-id="f40a0-678">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="f40a0-679">示例</span><span class="sxs-lookup"><span data-stu-id="f40a0-679">Examples</span></span>

<span data-ttu-id="f40a0-680">以下代码将一个字符串传递到 `displayReplyAllForm` 函数。</span><span class="sxs-lookup"><span data-stu-id="f40a0-680">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="f40a0-681">使用空白正文答复。</span><span class="sxs-lookup"><span data-stu-id="f40a0-681">Reply with an empty body.</span></span>

```
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="f40a0-682">仅使用正文答复。</span><span class="sxs-lookup"><span data-stu-id="f40a0-682">Reply with just a body.</span></span>

```
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="f40a0-683">使用正文和文件附件答复。</span><span class="sxs-lookup"><span data-stu-id="f40a0-683">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="f40a0-684">使用正文和项目附件答复。</span><span class="sxs-lookup"><span data-stu-id="f40a0-684">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="f40a0-685">使用正文、文件附件、项目附件和回调答复。</span><span class="sxs-lookup"><span data-stu-id="f40a0-685">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="displayreplyformformdata"></a><span data-ttu-id="f40a0-686">displayReplyForm(formData)</span><span class="sxs-lookup"><span data-stu-id="f40a0-686">displayReplyForm(formData)</span></span>

<span data-ttu-id="f40a0-687">显示答复窗体，其中仅包括所选邮件的发件人或所选约会的组织者。</span><span class="sxs-lookup"><span data-stu-id="f40a0-687">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="f40a0-688">在 Outlook for iOS 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="f40a0-688">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="f40a0-689">在 Outlook Web App 中，答复窗体显示为包含 3 列视图的弹出式窗体以及包含 2 列或 1 列视图的弹出式窗体。</span><span class="sxs-lookup"><span data-stu-id="f40a0-689">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="f40a0-690">如果任意字符串参数超出其限制， `displayReplyForm` 将引发异常。</span><span class="sxs-lookup"><span data-stu-id="f40a0-690">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="f40a0-p145">当在 `formData.attachments` 参数中指定附件时，Outlook 和 Outlook Web App 尝试下载所有附件并将其附加到答复窗体。如果无法添加任何附件，则在窗体 UI 中显示错误。如果这不可能，则不引发错误消息。</span><span class="sxs-lookup"><span data-stu-id="f40a0-p145">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="f40a0-694">参数：</span><span class="sxs-lookup"><span data-stu-id="f40a0-694">Parameters:</span></span>

|<span data-ttu-id="f40a0-695">名称</span><span class="sxs-lookup"><span data-stu-id="f40a0-695">Name</span></span>| <span data-ttu-id="f40a0-696">类型</span><span class="sxs-lookup"><span data-stu-id="f40a0-696">Type</span></span>| <span data-ttu-id="f40a0-697">说明</span><span class="sxs-lookup"><span data-stu-id="f40a0-697">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="f40a0-698">字符串 | 对象</span><span class="sxs-lookup"><span data-stu-id="f40a0-698">String &#124; Object</span></span>| | <span data-ttu-id="f40a0-p146">一个包含文本和 HTML 且代表答复窗体正文的字符串。字符串限制为 32 KB。</span><span class="sxs-lookup"><span data-stu-id="f40a0-p146">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="f40a0-701">**OR**</span><span class="sxs-lookup"><span data-stu-id="f40a0-701">**OR**</span></span><br/><span data-ttu-id="f40a0-p147">包含正文或附件数据和回调函数的对象。对象定义如下。</span><span class="sxs-lookup"><span data-stu-id="f40a0-p147">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="f40a0-704">String</span><span class="sxs-lookup"><span data-stu-id="f40a0-704">String</span></span> | <span data-ttu-id="f40a0-705">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="f40a0-705">&lt;optional&gt;</span></span> | <span data-ttu-id="f40a0-p148">一个包含文本和 HTML 且表示答复窗体正文的字符串。字符串限制为 32 KB。</span><span class="sxs-lookup"><span data-stu-id="f40a0-p148">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="f40a0-708">数组。&lt;对象&gt;</span><span class="sxs-lookup"><span data-stu-id="f40a0-708">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="f40a0-709">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="f40a0-709">&lt;optional&gt;</span></span> | <span data-ttu-id="f40a0-710">JSON 对象（是文件或项目附件）的数组。</span><span class="sxs-lookup"><span data-stu-id="f40a0-710">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="f40a0-711">String</span><span class="sxs-lookup"><span data-stu-id="f40a0-711">String</span></span> | | <span data-ttu-id="f40a0-p149">指示附件的类型。必须是文件附件的 `file` 或项目附件的 `item` 。</span><span class="sxs-lookup"><span data-stu-id="f40a0-p149">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="f40a0-714">String</span><span class="sxs-lookup"><span data-stu-id="f40a0-714">String</span></span> | | <span data-ttu-id="f40a0-715">一个包含附件的名称的字符串，最多包含 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="f40a0-715">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="f40a0-716">String</span><span class="sxs-lookup"><span data-stu-id="f40a0-716">String</span></span> | | <span data-ttu-id="f40a0-p150">仅在将 `type` 设置为 `file` 时使用。文件位置的 URI。</span><span class="sxs-lookup"><span data-stu-id="f40a0-p150">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="f40a0-719">String</span><span class="sxs-lookup"><span data-stu-id="f40a0-719">String</span></span> | | <span data-ttu-id="f40a0-p151">仅在将 `type` 设置为 `item` 时使用。附件的 EWS 项目 ID。字符串最长为 100 个字符。</span><span class="sxs-lookup"><span data-stu-id="f40a0-p151">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="f40a0-723">function</span><span class="sxs-lookup"><span data-stu-id="f40a0-723">function</span></span> | <span data-ttu-id="f40a0-724">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="f40a0-724">&lt;optional&gt;</span></span> | <span data-ttu-id="f40a0-725">方法完成后，使用单个参数 `asyncResult`（一个 [AsyncResult](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="f40a0-725">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="f40a0-726">要求</span><span class="sxs-lookup"><span data-stu-id="f40a0-726">Requirements</span></span>

|<span data-ttu-id="f40a0-727">要求</span><span class="sxs-lookup"><span data-stu-id="f40a0-727">Requirement</span></span>| <span data-ttu-id="f40a0-728">值</span><span class="sxs-lookup"><span data-stu-id="f40a0-728">Value</span></span>|
|---|---|
|[<span data-ttu-id="f40a0-729">最低邮箱要求集版本</span><span class="sxs-lookup"><span data-stu-id="f40a0-729">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f40a0-730">1.0</span><span class="sxs-lookup"><span data-stu-id="f40a0-730">1.0</span></span>|
|[<span data-ttu-id="f40a0-731">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="f40a0-731">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f40a0-732">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f40a0-732">ReadItem</span></span>|
|[<span data-ttu-id="f40a0-733">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="f40a0-733">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f40a0-734">阅读</span><span class="sxs-lookup"><span data-stu-id="f40a0-734">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="f40a0-735">示例</span><span class="sxs-lookup"><span data-stu-id="f40a0-735">Examples</span></span>

<span data-ttu-id="f40a0-736">以下代码将一个字符串传递到 `displayReplyForm` 函数。</span><span class="sxs-lookup"><span data-stu-id="f40a0-736">The following code passes a string to the `displayReplyForm` function.</span></span>

```
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="f40a0-737">使用空白正文答复。</span><span class="sxs-lookup"><span data-stu-id="f40a0-737">Reply with an empty body.</span></span>

```
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="f40a0-738">仅使用正文答复。</span><span class="sxs-lookup"><span data-stu-id="f40a0-738">Reply with just a body.</span></span>

```
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="f40a0-739">使用正文和文件附件答复。</span><span class="sxs-lookup"><span data-stu-id="f40a0-739">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="f40a0-740">使用正文和项目附件答复。</span><span class="sxs-lookup"><span data-stu-id="f40a0-740">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="f40a0-741">使用正文、文件附件、项目附件和回调答复。</span><span class="sxs-lookup"><span data-stu-id="f40a0-741">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="getentities--entitiesjavascriptapioutlook14officeentities"></a><span data-ttu-id="f40a0-742">getEntities() → {[实体](/javascript/api/outlook_1_4/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="f40a0-742">getEntities() → {[Entities](/javascript/api/outlook_1_4/office.entities)}</span></span>

<span data-ttu-id="f40a0-743">获取在所选项正文中找到的实体。</span><span class="sxs-lookup"><span data-stu-id="f40a0-743">Gets the entities found in the selected item.</span></span>

> [!NOTE]
> <span data-ttu-id="f40a0-744">在 Outlook for iOS 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="f40a0-744">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="f40a0-745">要求</span><span class="sxs-lookup"><span data-stu-id="f40a0-745">Requirements</span></span>

|<span data-ttu-id="f40a0-746">要求</span><span class="sxs-lookup"><span data-stu-id="f40a0-746">Requirement</span></span>| <span data-ttu-id="f40a0-747">值</span><span class="sxs-lookup"><span data-stu-id="f40a0-747">Value</span></span>|
|---|---|
|[<span data-ttu-id="f40a0-748">最低邮箱要求集版本</span><span class="sxs-lookup"><span data-stu-id="f40a0-748">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f40a0-749">1.0</span><span class="sxs-lookup"><span data-stu-id="f40a0-749">1.0</span></span>|
|[<span data-ttu-id="f40a0-750">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="f40a0-750">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f40a0-751">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f40a0-751">ReadItem</span></span>|
|[<span data-ttu-id="f40a0-752">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="f40a0-752">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f40a0-753">阅读</span><span class="sxs-lookup"><span data-stu-id="f40a0-753">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="f40a0-754">返回：</span><span class="sxs-lookup"><span data-stu-id="f40a0-754">Returns:</span></span>

<span data-ttu-id="f40a0-755">类型：[Entities](/javascript/api/outlook_1_4/office.entities)</span><span class="sxs-lookup"><span data-stu-id="f40a0-755">Type: [Entities](/javascript/api/outlook_1_4/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="f40a0-756">示例</span><span class="sxs-lookup"><span data-stu-id="f40a0-756">Example</span></span>

<span data-ttu-id="f40a0-757">以下示例访问当前项正文中的联系人实体。</span><span class="sxs-lookup"><span data-stu-id="f40a0-757">The following example accesses the contacts entities on the current item.</span></span>

```
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlook14officecontactmeetingsuggestionjavascriptapioutlook14officemeetingsuggestionphonenumberjavascriptapioutlook14officephonenumbertasksuggestionjavascriptapioutlook14officetasksuggestion"></a><span data-ttu-id="f40a0-758">getEntitiesByType(entityType) → (nullable)  {数组。 <(String|[联系人](/javascript/api/outlook_1_4/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_4/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_4/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_4/office.tasksuggestion)) >}</span><span class="sxs-lookup"><span data-stu-id="f40a0-758">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_4/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_4/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_4/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_4/office.tasksuggestion))>}</span></span>

<span data-ttu-id="f40a0-759">获取所选项目正文中找到的指定实体类型的所有实体的数组。</span><span class="sxs-lookup"><span data-stu-id="f40a0-759">Gets an array of all the entities of the specified entity type found in the selected item.</span></span>

> [!NOTE]
> <span data-ttu-id="f40a0-760">在 Outlook for iOS 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="f40a0-760">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="f40a0-761">参数：</span><span class="sxs-lookup"><span data-stu-id="f40a0-761">Parameters:</span></span>

|<span data-ttu-id="f40a0-762">名称</span><span class="sxs-lookup"><span data-stu-id="f40a0-762">Name</span></span>| <span data-ttu-id="f40a0-763">类型</span><span class="sxs-lookup"><span data-stu-id="f40a0-763">Type</span></span>| <span data-ttu-id="f40a0-764">说明</span><span class="sxs-lookup"><span data-stu-id="f40a0-764">Description</span></span>|
|---|---|---|
|`entityType`| [<span data-ttu-id="f40a0-765">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="f40a0-765">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook_1_4/office.mailboxenums.entitytype)|<span data-ttu-id="f40a0-766">EntityType 枚举值之一。</span><span class="sxs-lookup"><span data-stu-id="f40a0-766">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f40a0-767">要求</span><span class="sxs-lookup"><span data-stu-id="f40a0-767">Requirements</span></span>

|<span data-ttu-id="f40a0-768">要求</span><span class="sxs-lookup"><span data-stu-id="f40a0-768">Requirement</span></span>| <span data-ttu-id="f40a0-769">值</span><span class="sxs-lookup"><span data-stu-id="f40a0-769">Value</span></span>|
|---|---|
|[<span data-ttu-id="f40a0-770">最低邮箱要求集版本</span><span class="sxs-lookup"><span data-stu-id="f40a0-770">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f40a0-771">1.0</span><span class="sxs-lookup"><span data-stu-id="f40a0-771">1.0</span></span>|
|[<span data-ttu-id="f40a0-772">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="f40a0-772">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f40a0-773">Restricted</span><span class="sxs-lookup"><span data-stu-id="f40a0-773">Restricted</span></span>|
|[<span data-ttu-id="f40a0-774">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="f40a0-774">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f40a0-775">阅读</span><span class="sxs-lookup"><span data-stu-id="f40a0-775">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="f40a0-776">返回：</span><span class="sxs-lookup"><span data-stu-id="f40a0-776">Returns:</span></span>

<span data-ttu-id="f40a0-777">如果在 `entityType` 中传递的值不是 `EntityType` 枚举的有效成员，该方法返回 null。</span><span class="sxs-lookup"><span data-stu-id="f40a0-777">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="f40a0-778">如果指定类型的任何实体都不存在于该项目上，该方法将返回空数组。</span><span class="sxs-lookup"><span data-stu-id="f40a0-778">If no entities of the specified type are present on the item, the method returns an empty array.</span></span> <span data-ttu-id="f40a0-779">否则，返回数组中的对象类型取决于 `entityType` 参数中请求实体的类型。</span><span class="sxs-lookup"><span data-stu-id="f40a0-779">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="f40a0-780">当使用此方法的最低权限级别**受限**时，某些实体类型需要 **ReadItem** 才能进行访问，如下表中所指定。</span><span class="sxs-lookup"><span data-stu-id="f40a0-780">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

| <span data-ttu-id="f40a0-781">值对应于 `entityType`</span><span class="sxs-lookup"><span data-stu-id="f40a0-781">Value of `entityType`</span></span> | <span data-ttu-id="f40a0-782">返回的数组中对象的类型</span><span class="sxs-lookup"><span data-stu-id="f40a0-782">Type of objects in returned array</span></span> | <span data-ttu-id="f40a0-783">所需权限级别</span><span class="sxs-lookup"><span data-stu-id="f40a0-783">Required Permission Level</span></span> |
| --- | --- | --- |
| `Address` | <span data-ttu-id="f40a0-784">String</span><span class="sxs-lookup"><span data-stu-id="f40a0-784">String</span></span> | <span data-ttu-id="f40a0-785">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="f40a0-785">**Restricted**</span></span> |
| `Contact` | <span data-ttu-id="f40a0-786">联系人</span><span class="sxs-lookup"><span data-stu-id="f40a0-786">Contact</span></span> | <span data-ttu-id="f40a0-787">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="f40a0-787">**ReadItem**</span></span> |
| `EmailAddress` | <span data-ttu-id="f40a0-788">String</span><span class="sxs-lookup"><span data-stu-id="f40a0-788">String</span></span> | <span data-ttu-id="f40a0-789">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="f40a0-789">**ReadItem**</span></span> |
| `MeetingSuggestion` | <span data-ttu-id="f40a0-790">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="f40a0-790">MeetingSuggestion</span></span> | <span data-ttu-id="f40a0-791">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="f40a0-791">**ReadItem**</span></span> |
| `PhoneNumber` | <span data-ttu-id="f40a0-792">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="f40a0-792">PhoneNumber</span></span> | <span data-ttu-id="f40a0-793">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="f40a0-793">**Restricted**</span></span> |
| `TaskSuggestion` | <span data-ttu-id="f40a0-794">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="f40a0-794">TaskSuggestion</span></span> | <span data-ttu-id="f40a0-795">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="f40a0-795">**ReadItem**</span></span> |
| `URL` | <span data-ttu-id="f40a0-796">String</span><span class="sxs-lookup"><span data-stu-id="f40a0-796">String</span></span> | <span data-ttu-id="f40a0-797">**Restricted**</span><span class="sxs-lookup"><span data-stu-id="f40a0-797">**Restricted**</span></span> |

<span data-ttu-id="f40a0-798">类型：数组.<(字符串|[Contact](/javascript/api/outlook_1_4/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_4/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_4/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_4/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="f40a0-798">Type: Array.<(String|[Contact](/javascript/api/outlook_1_4/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_4/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_4/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_4/office.tasksuggestion))></span></span>

##### <a name="example"></a><span data-ttu-id="f40a0-799">示例</span><span class="sxs-lookup"><span data-stu-id="f40a0-799">Example</span></span>

<span data-ttu-id="f40a0-800">以下示例显示了如何访问代表当前项正文中邮政地址的字符串数组。</span><span class="sxs-lookup"><span data-stu-id="f40a0-800">The following example shows how to access an array of strings that represent postal addresses in the subject or body of the current item.</span></span>

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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlook14officecontactmeetingsuggestionjavascriptapioutlook14officemeetingsuggestionphonenumberjavascriptapioutlook14officephonenumbertasksuggestionjavascriptapioutlook14officetasksuggestion"></a><span data-ttu-id="f40a0-801">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_4/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_4/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_4/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_4/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="f40a0-801">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_4/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_4/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_4/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_4/office.tasksuggestion))>}</span></span>

<span data-ttu-id="f40a0-802">返回传递清单 XML 文件中定义的命名筛选器所选项中的已知实体。</span><span class="sxs-lookup"><span data-stu-id="f40a0-802">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="f40a0-803">在 Outlook for iOS 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="f40a0-803">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="f40a0-804">`getFilteredEntitiesByName` 方法返回与具有指定 `FilterName` 元素值的清单 XML 文件中的 [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) 规则元素中定义的规则表达式相匹配的实体。</span><span class="sxs-lookup"><span data-stu-id="f40a0-804">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="f40a0-805">参数：</span><span class="sxs-lookup"><span data-stu-id="f40a0-805">Parameters:</span></span>

|<span data-ttu-id="f40a0-806">名称</span><span class="sxs-lookup"><span data-stu-id="f40a0-806">Name</span></span>| <span data-ttu-id="f40a0-807">类型</span><span class="sxs-lookup"><span data-stu-id="f40a0-807">Type</span></span>| <span data-ttu-id="f40a0-808">说明</span><span class="sxs-lookup"><span data-stu-id="f40a0-808">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="f40a0-809">String</span><span class="sxs-lookup"><span data-stu-id="f40a0-809">String</span></span>|<span data-ttu-id="f40a0-810">定义筛选器匹配的 `ItemHasKnownEntity` 规则元素的名称。</span><span class="sxs-lookup"><span data-stu-id="f40a0-810">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f40a0-811">要求</span><span class="sxs-lookup"><span data-stu-id="f40a0-811">Requirements</span></span>

|<span data-ttu-id="f40a0-812">要求</span><span class="sxs-lookup"><span data-stu-id="f40a0-812">Requirement</span></span>| <span data-ttu-id="f40a0-813">值</span><span class="sxs-lookup"><span data-stu-id="f40a0-813">Value</span></span>|
|---|---|
|[<span data-ttu-id="f40a0-814">最低邮箱要求集版本</span><span class="sxs-lookup"><span data-stu-id="f40a0-814">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f40a0-815">1.0</span><span class="sxs-lookup"><span data-stu-id="f40a0-815">1.0</span></span>|
|[<span data-ttu-id="f40a0-816">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="f40a0-816">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f40a0-817">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f40a0-817">ReadItem</span></span>|
|[<span data-ttu-id="f40a0-818">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="f40a0-818">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f40a0-819">阅读</span><span class="sxs-lookup"><span data-stu-id="f40a0-819">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="f40a0-820">返回：</span><span class="sxs-lookup"><span data-stu-id="f40a0-820">Returns:</span></span>

<span data-ttu-id="f40a0-p153">如果清单中 `ItemHasKnownEntity`  元素没有匹配 `FilterName` 参数的 `name` 元素值，则该方法返回 `null`。如果 `name` 参数匹配清单中的 `ItemHasKnownEntity` 元素，但在当前匹配的项目中没有实体，则该方法返回一个空数组。</span><span class="sxs-lookup"><span data-stu-id="f40a0-p153">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="f40a0-823">类型：Array.<(String|[Contact](/javascript/api/outlook_1_4/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_4/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_4/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_4/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="f40a0-823">Type: Array.<(String|[Contact](/javascript/api/outlook_1_4/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_4/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_4/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_4/office.tasksuggestion))></span></span>

#### <a name="getregexmatches--object"></a><span data-ttu-id="f40a0-824">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="f40a0-824">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="f40a0-825">返回所选项目中与在清单 XML 文件中定义的正则表达式相匹配的字符串值。</span><span class="sxs-lookup"><span data-stu-id="f40a0-825">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="f40a0-826">在 Outlook for iOS 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="f40a0-826">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="f40a0-p154">`getRegExMatches` 方法返回匹配在清单 XML 文件中的每个 `ItemHasRegularExpressionMatch` 或 `ItemHasKnownEntity` 规则元素中定义的正则表达式的字符串。对于 `ItemHasRegularExpressionMatch` 规则，匹配字符串必须发生在该规则指定的项目的属性中。`PropertyName` 简单类型定义支持的属性。</span><span class="sxs-lookup"><span data-stu-id="f40a0-p154">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="f40a0-830">例如，考虑一个加载项具有以下 `Rule` 元素：</span><span class="sxs-lookup"><span data-stu-id="f40a0-830">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="f40a0-831">从 `getRegExMatches` 返回的对象应有两个属性：`fruits` 和 `veggies`。</span><span class="sxs-lookup"><span data-stu-id="f40a0-831">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="f40a0-p155">如果在项目的正文属性上指定 `ItemHasRegularExpressionMatch` 规则，则正则表达式应进一步筛选正文，不应尝试返回该项目的整个正文。使用正则表达式（如 `.*`）获取项目的整个正文并不总是返回预期的结果。而该使用 [`Body.getAsync`](/javascript/api/outlook_1_4/office.body#getasync-coerciontype--options--callback-) 方法检索整个正文。</span><span class="sxs-lookup"><span data-stu-id="f40a0-p155">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook_1_4/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="f40a0-835">要求</span><span class="sxs-lookup"><span data-stu-id="f40a0-835">Requirements</span></span>

|<span data-ttu-id="f40a0-836">要求</span><span class="sxs-lookup"><span data-stu-id="f40a0-836">Requirement</span></span>| <span data-ttu-id="f40a0-837">值</span><span class="sxs-lookup"><span data-stu-id="f40a0-837">Value</span></span>|
|---|---|
|[<span data-ttu-id="f40a0-838">最低邮箱要求集版本</span><span class="sxs-lookup"><span data-stu-id="f40a0-838">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f40a0-839">1.0</span><span class="sxs-lookup"><span data-stu-id="f40a0-839">1.0</span></span>|
|[<span data-ttu-id="f40a0-840">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="f40a0-840">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f40a0-841">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f40a0-841">ReadItem</span></span>|
|[<span data-ttu-id="f40a0-842">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="f40a0-842">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f40a0-843">阅读</span><span class="sxs-lookup"><span data-stu-id="f40a0-843">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="f40a0-844">返回：</span><span class="sxs-lookup"><span data-stu-id="f40a0-844">Returns:</span></span>

<span data-ttu-id="f40a0-p156">一个包含与在清单 XML 文件中定义的正则表达式的字符串数组相匹配的对象。每个数组的名称等于匹配 `ItemHasRegularExpressionMatch` 规则的 `RegExName` 属性或匹配 `ItemHasKnownEntity` 规则的 `FilterName` 属性的相应值。</span><span class="sxs-lookup"><span data-stu-id="f40a0-p156">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<dl class="param-type"><span data-ttu-id="f40a0-847">

<dt>类型</dt>

</span><span class="sxs-lookup"><span data-stu-id="f40a0-847">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="f40a0-848">Object</span><span class="sxs-lookup"><span data-stu-id="f40a0-848">Object</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="f40a0-849">示例</span><span class="sxs-lookup"><span data-stu-id="f40a0-849">Example</span></span>

<span data-ttu-id="f40a0-850">以下示例显示了如何访问正则表达式 <rule>元素 `fruits` 和 `veggies` 的匹配项的数组，这些元素在清单中指定。</rule></span><span class="sxs-lookup"><span data-stu-id="f40a0-850">The following example shows how to access the array of matches for the regular expression <rule>elements `fruits` and `veggies`, which are specified in the manifest.</rule></span></span>

```
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veges = allMatches.veggies;
```

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="f40a0-851">getRegExMatchesByName(name) → (可空类型) {数组.< 字符串 >}</span><span class="sxs-lookup"><span data-stu-id="f40a0-851">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="f40a0-852">返回匹配在清单 XML 文件中定义的命名正则表达式的所选项目中的字符串值。</span><span class="sxs-lookup"><span data-stu-id="f40a0-852">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="f40a0-853">在 Outlook for iOS 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="f40a0-853">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="f40a0-854">`getRegExMatchesByName` 方法返回匹配在具有指定 `RegExName` 元素值的清单 XML 文件中的 `ItemHasRegularExpressionMatch` 规则元素中定义的正则表达式的字符串。</span><span class="sxs-lookup"><span data-stu-id="f40a0-854">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="f40a0-p157">如果在项目正文属性上指定 `ItemHasRegularExpressionMatch` 规则，则正则表达式应进一步筛选正文，不应尝试返回该项目的整个正文。使用正则表达式（如 `.*`）获取项目的整个正文并不总是返回预期的结果。</span><span class="sxs-lookup"><span data-stu-id="f40a0-p157">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="f40a0-857">参数：</span><span class="sxs-lookup"><span data-stu-id="f40a0-857">Parameters:</span></span>

|<span data-ttu-id="f40a0-858">名称</span><span class="sxs-lookup"><span data-stu-id="f40a0-858">Name</span></span>| <span data-ttu-id="f40a0-859">类型</span><span class="sxs-lookup"><span data-stu-id="f40a0-859">Type</span></span>| <span data-ttu-id="f40a0-860">说明</span><span class="sxs-lookup"><span data-stu-id="f40a0-860">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="f40a0-861">String</span><span class="sxs-lookup"><span data-stu-id="f40a0-861">String</span></span>|<span data-ttu-id="f40a0-862">定义筛选器匹配的 `ItemHasRegularExpressionMatch` 规则元素的名称。</span><span class="sxs-lookup"><span data-stu-id="f40a0-862">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f40a0-863">要求</span><span class="sxs-lookup"><span data-stu-id="f40a0-863">Requirements</span></span>

|<span data-ttu-id="f40a0-864">要求</span><span class="sxs-lookup"><span data-stu-id="f40a0-864">Requirement</span></span>| <span data-ttu-id="f40a0-865">值</span><span class="sxs-lookup"><span data-stu-id="f40a0-865">Value</span></span>|
|---|---|
|[<span data-ttu-id="f40a0-866">最低邮箱要求集版本</span><span class="sxs-lookup"><span data-stu-id="f40a0-866">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f40a0-867">1.0</span><span class="sxs-lookup"><span data-stu-id="f40a0-867">1.0</span></span>|
|[<span data-ttu-id="f40a0-868">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="f40a0-868">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f40a0-869">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f40a0-869">ReadItem</span></span>|
|[<span data-ttu-id="f40a0-870">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="f40a0-870">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f40a0-871">阅读</span><span class="sxs-lookup"><span data-stu-id="f40a0-871">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="f40a0-872">返回：</span><span class="sxs-lookup"><span data-stu-id="f40a0-872">Returns:</span></span>

<span data-ttu-id="f40a0-873">一个包含与在清单 XML 文件中定义的正则表达式的字符串相匹配的数组。</span><span class="sxs-lookup"><span data-stu-id="f40a0-873">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<dl class="param-type"><span data-ttu-id="f40a0-874">

<dt>类型</dt>

</span><span class="sxs-lookup"><span data-stu-id="f40a0-874">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="f40a0-875">数组。 < 字符串 ></span><span class="sxs-lookup"><span data-stu-id="f40a0-875">Array.< String ></span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="f40a0-876">示例</span><span class="sxs-lookup"><span data-stu-id="f40a0-876">Example</span></span>

```
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

####  <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="f40a0-877">getSelectedDataAsync (coercionType，[选项] 回调) → {字符串}</span><span class="sxs-lookup"><span data-stu-id="f40a0-877">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="f40a0-878">以异步方式返回邮件的主题或正文中选定的数据。</span><span class="sxs-lookup"><span data-stu-id="f40a0-878">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="f40a0-p158">如果没有选定内容，但光标位于正文或主题中，此方法将会为所选数据返回 null。如果选定的是字段，而不是正文或主题，则此方法返回 `InvalidSelection` 错误。</span><span class="sxs-lookup"><span data-stu-id="f40a0-p158">If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="f40a0-881">参数：</span><span class="sxs-lookup"><span data-stu-id="f40a0-881">Parameters:</span></span>

|<span data-ttu-id="f40a0-882">名称</span><span class="sxs-lookup"><span data-stu-id="f40a0-882">Name</span></span>| <span data-ttu-id="f40a0-883">类型</span><span class="sxs-lookup"><span data-stu-id="f40a0-883">Type</span></span>| <span data-ttu-id="f40a0-884">属性</span><span class="sxs-lookup"><span data-stu-id="f40a0-884">Attributes</span></span>| <span data-ttu-id="f40a0-885">说明</span><span class="sxs-lookup"><span data-stu-id="f40a0-885">Description</span></span>|
|---|---|---|---|
|`coercionType`| [<span data-ttu-id="f40a0-886">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="f40a0-886">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="f40a0-p159">请求数据的格式。如果为 Text，则此方法返回纯文本作为字符串，删除任何显示的 HTML 标记。如果为 HTML，则此方法返回所选文本，不论是纯文本还是 HTML。</span><span class="sxs-lookup"><span data-stu-id="f40a0-p159">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`| <span data-ttu-id="f40a0-890">Object</span><span class="sxs-lookup"><span data-stu-id="f40a0-890">Object</span></span>| <span data-ttu-id="f40a0-891">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="f40a0-891">&lt;optional&gt;</span></span>|<span data-ttu-id="f40a0-892">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="f40a0-892">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="f40a0-893">Object</span><span class="sxs-lookup"><span data-stu-id="f40a0-893">Object</span></span>| <span data-ttu-id="f40a0-894">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="f40a0-894">&lt;optional&gt;</span></span>|<span data-ttu-id="f40a0-895">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="f40a0-895">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="f40a0-896">函数</span><span class="sxs-lookup"><span data-stu-id="f40a0-896">function</span></span>||<span data-ttu-id="f40a0-897">方法完成后，使用单个参数 `callback`（一个 [`asyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `AsyncResult` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="f40a0-897">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="f40a0-898">若要从回调方法访问所选数据，请调用 `asyncResult.value.data`。</span><span class="sxs-lookup"><span data-stu-id="f40a0-898">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="f40a0-899">若要访问选定内容的源属性，请调用 `asyncResult.value.sourceProperty`，这将为 `body` 或 `subject`。</span><span class="sxs-lookup"><span data-stu-id="f40a0-899">To access the source property that the selection comes from, call , which will be either  or .|</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f40a0-900">要求</span><span class="sxs-lookup"><span data-stu-id="f40a0-900">Requirements</span></span>

|<span data-ttu-id="f40a0-901">要求</span><span class="sxs-lookup"><span data-stu-id="f40a0-901">Requirement</span></span>| <span data-ttu-id="f40a0-902">值</span><span class="sxs-lookup"><span data-stu-id="f40a0-902">Value</span></span>|
|---|---|
|[<span data-ttu-id="f40a0-903">最低邮箱要求集版本</span><span class="sxs-lookup"><span data-stu-id="f40a0-903">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f40a0-904">1.2</span><span class="sxs-lookup"><span data-stu-id="f40a0-904">1.2</span></span>|
|[<span data-ttu-id="f40a0-905">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="f40a0-905">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f40a0-906">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="f40a0-906">ReadWriteItem</span></span>|
|[<span data-ttu-id="f40a0-907">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="f40a0-907">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f40a0-908">撰写</span><span class="sxs-lookup"><span data-stu-id="f40a0-908">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="f40a0-909">返回：</span><span class="sxs-lookup"><span data-stu-id="f40a0-909">Returns:</span></span>

<span data-ttu-id="f40a0-910">作为字符串的所选数据的格式由 `coercionType` 确定。</span><span class="sxs-lookup"><span data-stu-id="f40a0-910">The selected data as a string with format determined by `coercionType`.</span></span>

<dl class="param-type"><span data-ttu-id="f40a0-911">

<dt>类型</dt>

</span><span class="sxs-lookup"><span data-stu-id="f40a0-911">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="f40a0-912">String</span><span class="sxs-lookup"><span data-stu-id="f40a0-912">String</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="f40a0-913">示例</span><span class="sxs-lookup"><span data-stu-id="f40a0-913">Example</span></span>

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

####  <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="f40a0-914">loadCustomPropertiesAsync （回调、 [userContext]）</span><span class="sxs-lookup"><span data-stu-id="f40a0-914">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="f40a0-915">为所选项目的加载项异步加载自定义属性。</span><span class="sxs-lookup"><span data-stu-id="f40a0-915">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="f40a0-p161">自定义属性基于每个应用、每个项目存储为键/值对。此方法在回调中返回 `CustomProperties` 对象，该回调提供访问特定于当前项目和当前外接程序的自定义属性的方法。自定义属性未在项目上加密，因此这不应用作安全存储。</span><span class="sxs-lookup"><span data-stu-id="f40a0-p161">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="f40a0-919">参数：</span><span class="sxs-lookup"><span data-stu-id="f40a0-919">Parameters:</span></span>

|<span data-ttu-id="f40a0-920">名称</span><span class="sxs-lookup"><span data-stu-id="f40a0-920">Name</span></span>| <span data-ttu-id="f40a0-921">类型</span><span class="sxs-lookup"><span data-stu-id="f40a0-921">Type</span></span>| <span data-ttu-id="f40a0-922">属性</span><span class="sxs-lookup"><span data-stu-id="f40a0-922">Attributes</span></span>| <span data-ttu-id="f40a0-923">说明</span><span class="sxs-lookup"><span data-stu-id="f40a0-923">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="f40a0-924">函数</span><span class="sxs-lookup"><span data-stu-id="f40a0-924">function</span></span>||<span data-ttu-id="f40a0-925">此方法完成时，用单个参数调用 `callback` 参数中传递的函数，`asyncResult`，这是一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象。</span><span class="sxs-lookup"><span data-stu-id="f40a0-925">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="f40a0-926">自定义属性作为 `asyncResult.value` 属性中的 [`CustomProperties`](/javascript/api/outlook_1_4/office.customproperties) 对象来提供。</span><span class="sxs-lookup"><span data-stu-id="f40a0-926">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook_1_4/office.customproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="f40a0-927">该对象可用于获取、设置和删除项目中的自定义属性，并将针对自定义属性集的更改保存回服务器。</span><span class="sxs-lookup"><span data-stu-id="f40a0-927">The custom properties are provided as a CustomProperties object in the asyncResult.value property. This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`| <span data-ttu-id="f40a0-928">Object</span><span class="sxs-lookup"><span data-stu-id="f40a0-928">Object</span></span>| <span data-ttu-id="f40a0-929">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="f40a0-929">&lt;optional&gt;</span></span>|<span data-ttu-id="f40a0-930">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="f40a0-930">Developers can provide any object they wish to access in the callback method.</span></span> <span data-ttu-id="f40a0-931">可以通过回调函数的 `asyncResult.asyncContext` 属性访问该对象。</span><span class="sxs-lookup"><span data-stu-id="f40a0-931">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.|</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f40a0-932">要求</span><span class="sxs-lookup"><span data-stu-id="f40a0-932">Requirements</span></span>

|<span data-ttu-id="f40a0-933">要求</span><span class="sxs-lookup"><span data-stu-id="f40a0-933">Requirement</span></span>| <span data-ttu-id="f40a0-934">值</span><span class="sxs-lookup"><span data-stu-id="f40a0-934">Value</span></span>|
|---|---|
|[<span data-ttu-id="f40a0-935">最低邮箱要求集版本</span><span class="sxs-lookup"><span data-stu-id="f40a0-935">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f40a0-936">1.0</span><span class="sxs-lookup"><span data-stu-id="f40a0-936">1.0</span></span>|
|[<span data-ttu-id="f40a0-937">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="f40a0-937">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f40a0-938">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f40a0-938">ReadItem</span></span>|
|[<span data-ttu-id="f40a0-939">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="f40a0-939">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f40a0-940">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="f40a0-940">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="f40a0-941">示例</span><span class="sxs-lookup"><span data-stu-id="f40a0-941">Example</span></span>

<span data-ttu-id="f40a0-p164">以下代码示例显示了如何使用 `loadCustomPropertiesAsync` 方法异步加载特定于当前项目的自定义属性。该示例还显示了如何使用 `CustomProperties.saveAsync` 方法将这些属性重新保存到服务器。加载自定义属性后，该代码示例将使用 `CustomProperties.get` 方法读取自定义属性 `myProp`，使用 `CustomProperties.set` 方法写入自定义属性 `otherProp`，最后调用 `saveAsync` 方法保存这些自定义属性。</span><span class="sxs-lookup"><span data-stu-id="f40a0-p164">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

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

####  <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="f40a0-945">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="f40a0-945">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="f40a0-946">将附件从邮件或约会中删除。</span><span class="sxs-lookup"><span data-stu-id="f40a0-946">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="f40a0-p165">`removeAttachmentAsync` 方法删除项目中带指定标识符的附件。最佳做法是，仅当同一个邮件应用程序在同一会话中添加了一个附件时，你才应使用该附件标识符来删除该附件。在 Outlook Web App 和适用于设备的 OWA 中，附件标识符只在同一个会话中才有效。当用户关闭应用程序，或者如果用户开始在内嵌窗体中撰写，并在随后弹出的内嵌窗体中继续在单独的窗口撰写时，会话即结束。</span><span class="sxs-lookup"><span data-stu-id="f40a0-p165">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item. As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session. In Outlook Web App and OWA for Devices, the attachment identifier is valid only within the same session. A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="f40a0-951">参数：</span><span class="sxs-lookup"><span data-stu-id="f40a0-951">Parameters:</span></span>

|<span data-ttu-id="f40a0-952">名称</span><span class="sxs-lookup"><span data-stu-id="f40a0-952">Name</span></span>| <span data-ttu-id="f40a0-953">类型</span><span class="sxs-lookup"><span data-stu-id="f40a0-953">Type</span></span>| <span data-ttu-id="f40a0-954">属性</span><span class="sxs-lookup"><span data-stu-id="f40a0-954">Attributes</span></span>| <span data-ttu-id="f40a0-955">说明</span><span class="sxs-lookup"><span data-stu-id="f40a0-955">Description</span></span>|
|---|---|---|---|
|`attachmentId`| <span data-ttu-id="f40a0-956">String</span><span class="sxs-lookup"><span data-stu-id="f40a0-956">String</span></span>||<span data-ttu-id="f40a0-p166">要删除的附件的标识符。字符串的最大长度为 100 个字符。</span><span class="sxs-lookup"><span data-stu-id="f40a0-p166">The identifier of the attachment to remove. The maximum length of the string is 100 characters.</span></span>|
|`options`| <span data-ttu-id="f40a0-959">Object</span><span class="sxs-lookup"><span data-stu-id="f40a0-959">Object</span></span>| <span data-ttu-id="f40a0-960">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="f40a0-960">&lt;optional&gt;</span></span>|<span data-ttu-id="f40a0-961">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="f40a0-961">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="f40a0-962">Object</span><span class="sxs-lookup"><span data-stu-id="f40a0-962">Object</span></span>| <span data-ttu-id="f40a0-963">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="f40a0-963">&lt;optional&gt;</span></span>|<span data-ttu-id="f40a0-964">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="f40a0-964">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="f40a0-965">function</span><span class="sxs-lookup"><span data-stu-id="f40a0-965">function</span></span>| <span data-ttu-id="f40a0-966">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="f40a0-966">&lt;optional&gt;</span></span>|<span data-ttu-id="f40a0-967">此方法完成时，用单个参数调用 `callback` 参数中传递的函数， `asyncResult` ，这是一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象。</span><span class="sxs-lookup"><span data-stu-id="f40a0-967">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="f40a0-968">如果删除附件失败，`asyncResult.error` 属性将包含一个说明失败原因的错误代码。</span><span class="sxs-lookup"><span data-stu-id="f40a0-968">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="f40a0-969">错误</span><span class="sxs-lookup"><span data-stu-id="f40a0-969">Errors</span></span>

| <span data-ttu-id="f40a0-970">错误代码</span><span class="sxs-lookup"><span data-stu-id="f40a0-970">Error code</span></span> | <span data-ttu-id="f40a0-971">说明</span><span class="sxs-lookup"><span data-stu-id="f40a0-971">Description</span></span> |
|------------|-------------|
| `InvalidAttachmentId` | <span data-ttu-id="f40a0-972">附件标识符不存在。</span><span class="sxs-lookup"><span data-stu-id="f40a0-972">The attachment identifier does not exist.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="f40a0-973">要求</span><span class="sxs-lookup"><span data-stu-id="f40a0-973">Requirements</span></span>

|<span data-ttu-id="f40a0-974">要求</span><span class="sxs-lookup"><span data-stu-id="f40a0-974">Requirement</span></span>| <span data-ttu-id="f40a0-975">值</span><span class="sxs-lookup"><span data-stu-id="f40a0-975">Value</span></span>|
|---|---|
|[<span data-ttu-id="f40a0-976">最低邮箱要求集版本</span><span class="sxs-lookup"><span data-stu-id="f40a0-976">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f40a0-977">1.1</span><span class="sxs-lookup"><span data-stu-id="f40a0-977">1.1</span></span>|
|[<span data-ttu-id="f40a0-978">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="f40a0-978">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f40a0-979">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="f40a0-979">ReadWriteItem</span></span>|
|[<span data-ttu-id="f40a0-980">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="f40a0-980">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f40a0-981">撰写</span><span class="sxs-lookup"><span data-stu-id="f40a0-981">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="f40a0-982">示例</span><span class="sxs-lookup"><span data-stu-id="f40a0-982">Example</span></span>

<span data-ttu-id="f40a0-983">以下代码删除包含标识符 '0' 的附件。</span><span class="sxs-lookup"><span data-stu-id="f40a0-983">The following code removes an attachment with an identifier of '0'.</span></span>

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

####  <a name="saveasyncoptions-callback"></a><span data-ttu-id="f40a0-984">saveAsync ([选项] 回调)</span><span class="sxs-lookup"><span data-stu-id="f40a0-984">saveAsync([options], callback)</span></span>

<span data-ttu-id="f40a0-985">异步保存项目。</span><span class="sxs-lookup"><span data-stu-id="f40a0-985">Asynchronously saves an item.</span></span>

<span data-ttu-id="f40a0-p167">调用时，此方法将当前邮件保存为草稿，并通过回调方法返回项目 ID。在 Outlook Web App 或 Outlook 联机模式下，该项目被保存到服务器中。在 Outlook 缓存模式下，该项目被保存到本地缓存中。</span><span class="sxs-lookup"><span data-stu-id="f40a0-p167">When invoked, this method saves the current message as a draft and returns the item id via the callback method. In Outlook Web App or Outlook in online mode, the item is saved to the server. In Outlook in cached mode, the item is saved to the local cache.</span></span>

> [!NOTE]
> <span data-ttu-id="f40a0-989">如果加载项调用 `saveAsync` 中的项目在撰写模式下才能获取 `itemId` 若要使用 EWS 或 REST API，请注意，缓存模式 Outlook 时，可能需要一些时间才能将项目实际同步到服务器。</span><span class="sxs-lookup"><span data-stu-id="f40a0-989">Note: If your add-in calls `saveAsync` on an item in compose mode in order to get an `itemId` to use with EWS or the REST API, be aware that when Outlook is in cached mode, it may take some time before the item is actually synced to the server. Until the item is synced, using the  will return an error.</span></span> <span data-ttu-id="f40a0-990">直到该项目同步，使用 `itemId` 将返回错误。</span><span class="sxs-lookup"><span data-stu-id="f40a0-990">Until the item is synced, using the `itemId` will return an error.</span></span>

<span data-ttu-id="f40a0-p169">由于约会没有草稿状态，如果以撰写模式在约会中调用 `saveAsync`，则该项将被保存为用户日历中的正常约会。对于之前未保存过的新约会，则不会发送邀请。保存现有约会将向添加或删除的与会者发送更新。</span><span class="sxs-lookup"><span data-stu-id="f40a0-p169">Since appointments have no draft state, if `saveAsync` is called on an appointment in compose mode, the item will be saved as a normal appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent. Saving an existing appointment will send an update to added or removed attendees.</span></span>

> [!NOTE]
> <span data-ttu-id="f40a0-994">以下客户端在约会上的撰写模式下具有 `saveAsync` 的不同行为：</span><span class="sxs-lookup"><span data-stu-id="f40a0-994">The following clients have different behavior for `saveAsync` on appointments in compose mode:</span></span>
>
> - <span data-ttu-id="f40a0-995">Mac Outlook 在会议的撰写模式中不支持 `saveAsync` 。</span><span class="sxs-lookup"><span data-stu-id="f40a0-995">Note: Mac Outlook does not support `saveAsync` on a meeting in compose mode. Calling  on a meeting in Mac Outlook will return an error.</span></span> <span data-ttu-id="f40a0-996">在Mac Outlook 中的会议上调用 `saveAsync` ，则将返回错误。</span><span class="sxs-lookup"><span data-stu-id="f40a0-996">Note: Mac Outlook does not support  on a meeting in compose mode. Calling `saveAsync` on a meeting in Mac Outlook will return an error.</span></span>
> - <span data-ttu-id="f40a0-997">当 `saveAsync` 在撰写模式调用约会时，Outlook 网页版总会发送一个邀请或更新。</span><span class="sxs-lookup"><span data-stu-id="f40a0-997">Outlook on the web always sends an invitation or update when `saveAsync` is called on an appointment in compose mode.</span></span>

##### <a name="parameters"></a><span data-ttu-id="f40a0-998">参数：</span><span class="sxs-lookup"><span data-stu-id="f40a0-998">Parameters:</span></span>

|<span data-ttu-id="f40a0-999">名称</span><span class="sxs-lookup"><span data-stu-id="f40a0-999">Name</span></span>| <span data-ttu-id="f40a0-1000">类型</span><span class="sxs-lookup"><span data-stu-id="f40a0-1000">Type</span></span>| <span data-ttu-id="f40a0-1001">属性</span><span class="sxs-lookup"><span data-stu-id="f40a0-1001">Attributes</span></span>| <span data-ttu-id="f40a0-1002">说明</span><span class="sxs-lookup"><span data-stu-id="f40a0-1002">Description</span></span>|
|---|---|---|---|
|`options`| <span data-ttu-id="f40a0-1003">Object</span><span class="sxs-lookup"><span data-stu-id="f40a0-1003">Object</span></span>| <span data-ttu-id="f40a0-1004">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="f40a0-1004">&lt;optional&gt;</span></span>|<span data-ttu-id="f40a0-1005">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="f40a0-1005">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="f40a0-1006">Object</span><span class="sxs-lookup"><span data-stu-id="f40a0-1006">Object</span></span>| <span data-ttu-id="f40a0-1007">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="f40a0-1007">&lt;optional&gt;</span></span>|<span data-ttu-id="f40a0-1008">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="f40a0-1008">Developers can provide any object they wish to access in the callback method.</span></span>||
|`callback`| <span data-ttu-id="f40a0-1009">函数</span><span class="sxs-lookup"><span data-stu-id="f40a0-1009">function</span></span>||<span data-ttu-id="f40a0-1010">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="f40a0-1010">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="f40a0-1011">如果成功，该项目标识符在 `asyncResult.value` 属性中提供。</span><span class="sxs-lookup"><span data-stu-id="f40a0-1011">On success, the item identifier is provided in the `asyncResult.value` property.|</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f40a0-1012">要求</span><span class="sxs-lookup"><span data-stu-id="f40a0-1012">Requirements</span></span>

|<span data-ttu-id="f40a0-1013">要求</span><span class="sxs-lookup"><span data-stu-id="f40a0-1013">Requirement</span></span>| <span data-ttu-id="f40a0-1014">值</span><span class="sxs-lookup"><span data-stu-id="f40a0-1014">Value</span></span>|
|---|---|
|[<span data-ttu-id="f40a0-1015">最低邮箱要求集版本</span><span class="sxs-lookup"><span data-stu-id="f40a0-1015">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f40a0-1016">1.3</span><span class="sxs-lookup"><span data-stu-id="f40a0-1016">1.3</span></span>|
|[<span data-ttu-id="f40a0-1017">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="f40a0-1017">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f40a0-1018">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="f40a0-1018">ReadWriteItem</span></span>|
|[<span data-ttu-id="f40a0-1019">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="f40a0-1019">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f40a0-1020">撰写</span><span class="sxs-lookup"><span data-stu-id="f40a0-1020">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="f40a0-1021">示例</span><span class="sxs-lookup"><span data-stu-id="f40a0-1021">Examples</span></span>

```
Office.context.mailbox.item.saveAsync(
  function callback(result) {
    // Process the result
  });
```

<span data-ttu-id="f40a0-p171">下面是传递给回调函数的 `result` 参数示例。`value` 属性包含的该项的项目 ID。</span><span class="sxs-lookup"><span data-stu-id="f40a0-p171">The following is an example of the `result` parameter passed to the callback function. The `value` property contains the item ID of the item.</span></span>

```
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

####  <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="f40a0-1024">setSelectedDataAsync (数据，[选项]，回调)</span><span class="sxs-lookup"><span data-stu-id="f40a0-1024">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="f40a0-1025">以异步方式将数据插入到邮件的正文或主题中。</span><span class="sxs-lookup"><span data-stu-id="f40a0-1025">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="f40a0-p172">`setSelectedDataAsync` 方法将指定字符串插入到项目主题或正文的光标位置，或者，如果在编辑器中已选择文本，则该方法将替换选择的文本。如果光标不在正文或主题字段中，则返回错误。插入之后，光标会位于插入内容的末尾。</span><span class="sxs-lookup"><span data-stu-id="f40a0-p172">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="f40a0-1029">参数：</span><span class="sxs-lookup"><span data-stu-id="f40a0-1029">Parameters:</span></span>

|<span data-ttu-id="f40a0-1030">名称</span><span class="sxs-lookup"><span data-stu-id="f40a0-1030">Name</span></span>| <span data-ttu-id="f40a0-1031">类型</span><span class="sxs-lookup"><span data-stu-id="f40a0-1031">Type</span></span>| <span data-ttu-id="f40a0-1032">属性</span><span class="sxs-lookup"><span data-stu-id="f40a0-1032">Attributes</span></span>| <span data-ttu-id="f40a0-1033">说明</span><span class="sxs-lookup"><span data-stu-id="f40a0-1033">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="f40a0-1034">String</span><span class="sxs-lookup"><span data-stu-id="f40a0-1034">String</span></span>||<span data-ttu-id="f40a0-p173">要插入的数据。数据不得超过 1,000,000 个字符。如果传入的数据超过 1,000,000 个字符，则会引发 `ArgumentOutOfRange` 异常。</span><span class="sxs-lookup"><span data-stu-id="f40a0-p173">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`| <span data-ttu-id="f40a0-1038">Object</span><span class="sxs-lookup"><span data-stu-id="f40a0-1038">Object</span></span>| <span data-ttu-id="f40a0-1039">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="f40a0-1039">&lt;optional&gt;</span></span>|<span data-ttu-id="f40a0-1040">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="f40a0-1040">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="f40a0-1041">Object</span><span class="sxs-lookup"><span data-stu-id="f40a0-1041">Object</span></span>| <span data-ttu-id="f40a0-1042">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="f40a0-1042">&lt;optional&gt;</span></span>|<span data-ttu-id="f40a0-1043">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="f40a0-1043">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`| [<span data-ttu-id="f40a0-1044">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="f40a0-1044">Office.CoercionType</span></span>](office.md#coerciontype-string)| <span data-ttu-id="f40a0-1045">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="f40a0-1045">&lt;optional&gt;</span></span>|<span data-ttu-id="f40a0-p174">如果是 `text`，则在 Outlook Web App 和 Outlook 中应用当前样式。如果该字段是 HTML 编辑器，则仅插入文本数据，即使数据为 HTML。</span><span class="sxs-lookup"><span data-stu-id="f40a0-p174">If `text`, the current style is applied in Outlook Web App and Outlook. If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="f40a0-p175">如果 `html` 和该字段支持 HTML（主题不支持），则在 Outlook Web App 中应用当前样式，而在 Outlook 中应用默认样式。如果该字段是文本字段，则返回 `InvalidDataFormat` 错误。</span><span class="sxs-lookup"><span data-stu-id="f40a0-p175">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook Web App and the default style is applied in Outlook. If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="f40a0-1050">如果未设置 `coercionType`，则结果取决于该字段：如果该字段是 HTML，则使用 HTML；如果该字段是文本，则使用纯文本。</span><span class="sxs-lookup"><span data-stu-id="f40a0-1050">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`| <span data-ttu-id="f40a0-1051">function</span><span class="sxs-lookup"><span data-stu-id="f40a0-1051">function</span></span>||<span data-ttu-id="f40a0-1052">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="f40a0-1052">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="f40a0-1053">要求</span><span class="sxs-lookup"><span data-stu-id="f40a0-1053">Requirements</span></span>

|<span data-ttu-id="f40a0-1054">要求</span><span class="sxs-lookup"><span data-stu-id="f40a0-1054">Requirement</span></span>| <span data-ttu-id="f40a0-1055">值</span><span class="sxs-lookup"><span data-stu-id="f40a0-1055">Value</span></span>|
|---|---|
|[<span data-ttu-id="f40a0-1056">最低的邮箱要求集版本</span><span class="sxs-lookup"><span data-stu-id="f40a0-1056">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f40a0-1057">1.2</span><span class="sxs-lookup"><span data-stu-id="f40a0-1057">1.2</span></span>|
|[<span data-ttu-id="f40a0-1058">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="f40a0-1058">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f40a0-1059">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="f40a0-1059">ReadWriteItem</span></span>|
|[<span data-ttu-id="f40a0-1060">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="f40a0-1060">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f40a0-1061">撰写</span><span class="sxs-lookup"><span data-stu-id="f40a0-1061">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="f40a0-1062">示例</span><span class="sxs-lookup"><span data-stu-id="f40a0-1062">Example</span></span>

```
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```