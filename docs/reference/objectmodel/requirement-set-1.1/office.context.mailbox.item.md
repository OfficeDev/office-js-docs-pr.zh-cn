---
title: "\"context\"-\"邮箱\"。项目-要求集1。1"
description: ''
ms.date: 02/15/2019
localization_priority: Normal
ms.openlocfilehash: 5a43029a64c63dec3d48136ffe0a9c3c76e18b6c
ms.sourcegitcommit: f26778b596b6b022814c39601485ff676ed4e2fa
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 02/16/2019
ms.locfileid: "30068159"
---
# <a name="item"></a><span data-ttu-id="3ee75-102">item</span><span class="sxs-lookup"><span data-stu-id="3ee75-102">item</span></span>

### <span data-ttu-id="3ee75-p101">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md). item</span><span class="sxs-lookup"><span data-stu-id="3ee75-p101">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md). item</span></span>

<span data-ttu-id="3ee75-p102">`item` 命名空间用于访问当前选定的邮件、会议请求或约会。可以通过使用 [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlook11officemailboxenumsitemtype) 属性确定 `item` 的类型。</span><span class="sxs-lookup"><span data-stu-id="3ee75-p102">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlook11officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="3ee75-107">Requirements</span><span class="sxs-lookup"><span data-stu-id="3ee75-107">Requirements</span></span>

|<span data-ttu-id="3ee75-108">要求</span><span class="sxs-lookup"><span data-stu-id="3ee75-108">Requirement</span></span>| <span data-ttu-id="3ee75-109">值</span><span class="sxs-lookup"><span data-stu-id="3ee75-109">Value</span></span>|
|---|---|
|[<span data-ttu-id="3ee75-110">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="3ee75-110">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3ee75-111">1.0</span><span class="sxs-lookup"><span data-stu-id="3ee75-111">1.0</span></span>|
|[<span data-ttu-id="3ee75-112">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="3ee75-112">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3ee75-113">受限</span><span class="sxs-lookup"><span data-stu-id="3ee75-113">Restricted</span></span>|
|[<span data-ttu-id="3ee75-114">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="3ee75-114">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="3ee75-115">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="3ee75-115">Compose or Read</span></span>|

### <a name="example"></a><span data-ttu-id="3ee75-116">示例</span><span class="sxs-lookup"><span data-stu-id="3ee75-116">Example</span></span>

<span data-ttu-id="3ee75-117">以下 JavaScript 代码示例显示了如何访问 Outlook 中当前项目的 `subject` 属性。</span><span class="sxs-lookup"><span data-stu-id="3ee75-117">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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
};
```

### <a name="members"></a><span data-ttu-id="3ee75-118">成员</span><span class="sxs-lookup"><span data-stu-id="3ee75-118">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlook11officeattachmentdetails"></a><span data-ttu-id="3ee75-119">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_1/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="3ee75-119">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_1/office.attachmentdetails)></span></span>

<span data-ttu-id="3ee75-p103">获取项目的附件数组。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="3ee75-p103">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="3ee75-122">由于潜在的安全问题，某些类型的文件会受到 Outlook 阻止，并且不会返回。</span><span class="sxs-lookup"><span data-stu-id="3ee75-122">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="3ee75-123">如需了解更多信息，请参阅 [Outlook 中阻止的附件](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519)。</span><span class="sxs-lookup"><span data-stu-id="3ee75-123">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="3ee75-124">类型</span><span class="sxs-lookup"><span data-stu-id="3ee75-124">Type</span></span>

*   <span data-ttu-id="3ee75-125">Array.<[AttachmentDetails](/javascript/api/outlook_1_1/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="3ee75-125">Array.<[AttachmentDetails](/javascript/api/outlook_1_1/office.attachmentdetails)></span></span>

##### <a name="requirements"></a><span data-ttu-id="3ee75-126">Requirements</span><span class="sxs-lookup"><span data-stu-id="3ee75-126">Requirements</span></span>

|<span data-ttu-id="3ee75-127">要求</span><span class="sxs-lookup"><span data-stu-id="3ee75-127">Requirement</span></span>| <span data-ttu-id="3ee75-128">值</span><span class="sxs-lookup"><span data-stu-id="3ee75-128">Value</span></span>|
|---|---|
|[<span data-ttu-id="3ee75-129">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="3ee75-129">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3ee75-130">1.0</span><span class="sxs-lookup"><span data-stu-id="3ee75-130">1.0</span></span>|
|[<span data-ttu-id="3ee75-131">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="3ee75-131">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3ee75-132">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3ee75-132">ReadItem</span></span>|
|[<span data-ttu-id="3ee75-133">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="3ee75-133">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="3ee75-134">阅读</span><span class="sxs-lookup"><span data-stu-id="3ee75-134">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="3ee75-135">示例</span><span class="sxs-lookup"><span data-stu-id="3ee75-135">Example</span></span>

<span data-ttu-id="3ee75-136">以下代码使用当前项目上所有附件的详细信息构成 HTML 字符串。</span><span class="sxs-lookup"><span data-stu-id="3ee75-136">The following code builds an HTML string with details of all attachments on the current item.</span></span>

```javascript
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

####  <a name="bcc-recipientsjavascriptapioutlook11officerecipients"></a><span data-ttu-id="3ee75-137">bcc :[Recipients](/javascript/api/outlook_1_1/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="3ee75-137">bcc :[Recipients](/javascript/api/outlook_1_1/office.recipients)</span></span>

<span data-ttu-id="3ee75-138">获取对象，该对象提供用于获取或更新邮件的密件抄送 (Bcc) 行上的收件人的方法。</span><span class="sxs-lookup"><span data-stu-id="3ee75-138">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="3ee75-139">仅限撰写模式。</span><span class="sxs-lookup"><span data-stu-id="3ee75-139">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="3ee75-140">类型</span><span class="sxs-lookup"><span data-stu-id="3ee75-140">Type</span></span>

*   [<span data-ttu-id="3ee75-141">收件人</span><span class="sxs-lookup"><span data-stu-id="3ee75-141">Recipients</span></span>](/javascript/api/outlook_1_1/office.recipients)

##### <a name="requirements"></a><span data-ttu-id="3ee75-142">Requirements</span><span class="sxs-lookup"><span data-stu-id="3ee75-142">Requirements</span></span>

|<span data-ttu-id="3ee75-143">要求</span><span class="sxs-lookup"><span data-stu-id="3ee75-143">Requirement</span></span>| <span data-ttu-id="3ee75-144">值</span><span class="sxs-lookup"><span data-stu-id="3ee75-144">Value</span></span>|
|---|---|
|[<span data-ttu-id="3ee75-145">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="3ee75-145">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3ee75-146">1.1</span><span class="sxs-lookup"><span data-stu-id="3ee75-146">1.1</span></span>|
|[<span data-ttu-id="3ee75-147">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="3ee75-147">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3ee75-148">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3ee75-148">ReadItem</span></span>|
|[<span data-ttu-id="3ee75-149">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="3ee75-149">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="3ee75-150">撰写</span><span class="sxs-lookup"><span data-stu-id="3ee75-150">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="3ee75-151">示例</span><span class="sxs-lookup"><span data-stu-id="3ee75-151">Example</span></span>

```javascript
Office.context.mailbox.item.bcc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.bcc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.bcc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfBccRecipients = asyncResult.value;
}
```

####  <a name="body-bodyjavascriptapioutlook11officebody"></a><span data-ttu-id="3ee75-152">body :[Body](/javascript/api/outlook_1_1/office.body)</span><span class="sxs-lookup"><span data-stu-id="3ee75-152">body :[Body](/javascript/api/outlook_1_1/office.body)</span></span>

<span data-ttu-id="3ee75-153">获取一个提供用于处理项目正文的方法的对象。</span><span class="sxs-lookup"><span data-stu-id="3ee75-153">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="3ee75-154">类型</span><span class="sxs-lookup"><span data-stu-id="3ee75-154">Type</span></span>

*   [<span data-ttu-id="3ee75-155">Body</span><span class="sxs-lookup"><span data-stu-id="3ee75-155">Body</span></span>](/javascript/api/outlook_1_1/office.body)

##### <a name="requirements"></a><span data-ttu-id="3ee75-156">要求</span><span class="sxs-lookup"><span data-stu-id="3ee75-156">Requirements</span></span>

|<span data-ttu-id="3ee75-157">要求</span><span class="sxs-lookup"><span data-stu-id="3ee75-157">Requirement</span></span>| <span data-ttu-id="3ee75-158">值</span><span class="sxs-lookup"><span data-stu-id="3ee75-158">Value</span></span>|
|---|---|
|[<span data-ttu-id="3ee75-159">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="3ee75-159">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3ee75-160">1.1</span><span class="sxs-lookup"><span data-stu-id="3ee75-160">1.1</span></span>|
|[<span data-ttu-id="3ee75-161">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="3ee75-161">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3ee75-162">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3ee75-162">ReadItem</span></span>|
|[<span data-ttu-id="3ee75-163">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="3ee75-163">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="3ee75-164">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="3ee75-164">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="3ee75-165">示例</span><span class="sxs-lookup"><span data-stu-id="3ee75-165">Example</span></span>

<span data-ttu-id="3ee75-166">本示例获取纯文本格式的邮件正文。</span><span class="sxs-lookup"><span data-stu-id="3ee75-166">This example gets the body of the message in plain text.</span></span>

```javascript
Office.context.mailbox.item.body.getAsync(
  "text",
  { asyncContext: "This is passed to the callback" },
  function callback(result) {
    // Do something with the result.
  });

```

<span data-ttu-id="3ee75-167">下面是传递给回调函数的 result 参数的示例。</span><span class="sxs-lookup"><span data-stu-id="3ee75-167">The following is an example of the result parameter passed to the callback function.</span></span>

```json
{
  "value": "TEXT of whole body (including threads below)",
  "status": "succeeded",
  "asyncContext": "This is passed to the callback"
}
```

####  <a name="cc-arrayemailaddressdetailsjavascriptapioutlook11officeemailaddressdetailsrecipientsjavascriptapioutlook11officerecipients"></a><span data-ttu-id="3ee75-168">cc :Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_1/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="3ee75-168">cc :Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_1/office.recipients)</span></span>

<span data-ttu-id="3ee75-169">提供对邮件的抄送 (Cc) 收件人的访问权限。</span><span class="sxs-lookup"><span data-stu-id="3ee75-169">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="3ee75-170">对象的类型和访问级别取决于当前项目的模式。</span><span class="sxs-lookup"><span data-stu-id="3ee75-170">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="3ee75-171">阅读模式</span><span class="sxs-lookup"><span data-stu-id="3ee75-171">Read mode</span></span>

<span data-ttu-id="3ee75-p107">`cc` 属性返回包含邮件的**抄送**行上所列的每个收件人的 `EmailAddressDetails` 对象的数组。集合上限为 100 个成员。</span><span class="sxs-lookup"><span data-stu-id="3ee75-p107">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.</span></span>

```javascript
console.log(JSON.stringify(Office.context.mailbox.item.cc));
```

##### <a name="compose-mode"></a><span data-ttu-id="3ee75-174">撰写模式</span><span class="sxs-lookup"><span data-stu-id="3ee75-174">Compose mode</span></span>

<span data-ttu-id="3ee75-175">`cc` 属性返回一个 `Recipients` 对象，该对象提供用于获取或更新邮件的“**抄送**”行上收件人的方法。</span><span class="sxs-lookup"><span data-stu-id="3ee75-175">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span>

```javascript
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="3ee75-176">类型</span><span class="sxs-lookup"><span data-stu-id="3ee75-176">Type</span></span>

*   <span data-ttu-id="3ee75-177">Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_1/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="3ee75-177">Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_1/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="3ee75-178">要求</span><span class="sxs-lookup"><span data-stu-id="3ee75-178">Requirements</span></span>

|<span data-ttu-id="3ee75-179">要求</span><span class="sxs-lookup"><span data-stu-id="3ee75-179">Requirement</span></span>| <span data-ttu-id="3ee75-180">值</span><span class="sxs-lookup"><span data-stu-id="3ee75-180">Value</span></span>|
|---|---|
|[<span data-ttu-id="3ee75-181">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="3ee75-181">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3ee75-182">1.0</span><span class="sxs-lookup"><span data-stu-id="3ee75-182">1.0</span></span>|
|[<span data-ttu-id="3ee75-183">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="3ee75-183">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3ee75-184">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3ee75-184">ReadItem</span></span>|
|[<span data-ttu-id="3ee75-185">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="3ee75-185">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="3ee75-186">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="3ee75-186">Compose or Read</span></span>|

####  <a name="nullable-conversationid-string"></a><span data-ttu-id="3ee75-187">(nullable) conversationId :String</span><span class="sxs-lookup"><span data-stu-id="3ee75-187">(nullable) conversationId :String</span></span>

<span data-ttu-id="3ee75-188">获取包含特定消息的电子邮件会话的标识符。</span><span class="sxs-lookup"><span data-stu-id="3ee75-188">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="3ee75-p108">如果在阅读窗体或撰写窗体的回复中激活邮件应用程序，则此属性可以获得一个整数值。如果用户随后更改了回复邮件的主题（若发送回复），则该邮件的对话 ID 将改变且之前获取的值将不适用。</span><span class="sxs-lookup"><span data-stu-id="3ee75-p108">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="3ee75-p109">对于撰写窗体的新项目，此属性获得一个 null 值。如果用户设置一个主题并保存该项目，`conversationId` 属性将返回一个值。</span><span class="sxs-lookup"><span data-stu-id="3ee75-p109">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="3ee75-193">Type</span><span class="sxs-lookup"><span data-stu-id="3ee75-193">Type</span></span>

*   <span data-ttu-id="3ee75-194">String</span><span class="sxs-lookup"><span data-stu-id="3ee75-194">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="3ee75-195">要求</span><span class="sxs-lookup"><span data-stu-id="3ee75-195">Requirements</span></span>

|<span data-ttu-id="3ee75-196">要求</span><span class="sxs-lookup"><span data-stu-id="3ee75-196">Requirement</span></span>| <span data-ttu-id="3ee75-197">值</span><span class="sxs-lookup"><span data-stu-id="3ee75-197">Value</span></span>|
|---|---|
|[<span data-ttu-id="3ee75-198">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="3ee75-198">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3ee75-199">1.0</span><span class="sxs-lookup"><span data-stu-id="3ee75-199">1.0</span></span>|
|[<span data-ttu-id="3ee75-200">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="3ee75-200">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3ee75-201">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3ee75-201">ReadItem</span></span>|
|[<span data-ttu-id="3ee75-202">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="3ee75-202">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="3ee75-203">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="3ee75-203">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="3ee75-204">示例</span><span class="sxs-lookup"><span data-stu-id="3ee75-204">Example</span></span>

```javascript
var conversationId = Office.context.mailbox.item.conversationId;
console.log("conversationId: " + conversationId);
```

#### <a name="datetimecreated-date"></a><span data-ttu-id="3ee75-205">dateTimeCreated :Date</span><span class="sxs-lookup"><span data-stu-id="3ee75-205">dateTimeCreated :Date</span></span>

<span data-ttu-id="3ee75-p110">获取项目创建的日期和时间。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="3ee75-p110">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="3ee75-208">类型</span><span class="sxs-lookup"><span data-stu-id="3ee75-208">Type</span></span>

*   <span data-ttu-id="3ee75-209">日期</span><span class="sxs-lookup"><span data-stu-id="3ee75-209">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="3ee75-210">Requirements</span><span class="sxs-lookup"><span data-stu-id="3ee75-210">Requirements</span></span>

|<span data-ttu-id="3ee75-211">要求</span><span class="sxs-lookup"><span data-stu-id="3ee75-211">Requirement</span></span>| <span data-ttu-id="3ee75-212">值</span><span class="sxs-lookup"><span data-stu-id="3ee75-212">Value</span></span>|
|---|---|
|[<span data-ttu-id="3ee75-213">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="3ee75-213">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3ee75-214">1.0</span><span class="sxs-lookup"><span data-stu-id="3ee75-214">1.0</span></span>|
|[<span data-ttu-id="3ee75-215">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="3ee75-215">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3ee75-216">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3ee75-216">ReadItem</span></span>|
|[<span data-ttu-id="3ee75-217">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="3ee75-217">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="3ee75-218">阅读</span><span class="sxs-lookup"><span data-stu-id="3ee75-218">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="3ee75-219">示例</span><span class="sxs-lookup"><span data-stu-id="3ee75-219">Example</span></span>

```javascript
var dateTimeCreated = Office.context.mailbox.item.dateTimeCreated;
console.log("Date and time created: " + dateTimeCreated);
```

#### <a name="datetimemodified-date"></a><span data-ttu-id="3ee75-220">dateTimeModified :Date</span><span class="sxs-lookup"><span data-stu-id="3ee75-220">dateTimeModified :Date</span></span>

<span data-ttu-id="3ee75-p111">获取项目最近一次修改的日期和时间。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="3ee75-p111">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="3ee75-223">在 Outlook for iOS 或 Outlook for Android 中不支持此成员。</span><span class="sxs-lookup"><span data-stu-id="3ee75-223">This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="type"></a><span data-ttu-id="3ee75-224">类型</span><span class="sxs-lookup"><span data-stu-id="3ee75-224">Type</span></span>

*   <span data-ttu-id="3ee75-225">日期</span><span class="sxs-lookup"><span data-stu-id="3ee75-225">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="3ee75-226">Requirements</span><span class="sxs-lookup"><span data-stu-id="3ee75-226">Requirements</span></span>

|<span data-ttu-id="3ee75-227">要求</span><span class="sxs-lookup"><span data-stu-id="3ee75-227">Requirement</span></span>| <span data-ttu-id="3ee75-228">值</span><span class="sxs-lookup"><span data-stu-id="3ee75-228">Value</span></span>|
|---|---|
|[<span data-ttu-id="3ee75-229">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="3ee75-229">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3ee75-230">1.0</span><span class="sxs-lookup"><span data-stu-id="3ee75-230">1.0</span></span>|
|[<span data-ttu-id="3ee75-231">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="3ee75-231">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3ee75-232">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3ee75-232">ReadItem</span></span>|
|[<span data-ttu-id="3ee75-233">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="3ee75-233">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="3ee75-234">阅读</span><span class="sxs-lookup"><span data-stu-id="3ee75-234">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="3ee75-235">示例</span><span class="sxs-lookup"><span data-stu-id="3ee75-235">Example</span></span>

```javascript
var dateTimeModified = Office.context.mailbox.item.dateTimeModified;
console.log("Date and time modified: " + dateTimeModified);
```

####  <a name="end-datetimejavascriptapioutlook11officetime"></a><span data-ttu-id="3ee75-236">end :Date|[Time](/javascript/api/outlook_1_1/office.time)</span><span class="sxs-lookup"><span data-stu-id="3ee75-236">end :Date|[Time](/javascript/api/outlook_1_1/office.time)</span></span>

<span data-ttu-id="3ee75-237">获取或设置约会结束的日期和时间。</span><span class="sxs-lookup"><span data-stu-id="3ee75-237">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="3ee75-p112">将 `end` 属性表示为协调世界时 (UTC) 的日期和时间值。可使用 [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook11officelocalclienttime) 方法将 end 属性值转换为客户端的本地日期和时间。</span><span class="sxs-lookup"><span data-stu-id="3ee75-p112">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook11officelocalclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="3ee75-240">阅读模式</span><span class="sxs-lookup"><span data-stu-id="3ee75-240">Read mode</span></span>

<span data-ttu-id="3ee75-241">`end` 属性返回 `Date` 对象。</span><span class="sxs-lookup"><span data-stu-id="3ee75-241">The `end` property returns a `Date` object.</span></span>

```javascript
var end = Office.context.mailbox.item.end;
console.log("Appointment end: " + end);
```

##### <a name="compose-mode"></a><span data-ttu-id="3ee75-242">撰写模式</span><span class="sxs-lookup"><span data-stu-id="3ee75-242">Compose mode</span></span>

<span data-ttu-id="3ee75-243">`end` 属性返回 `Time` 对象。</span><span class="sxs-lookup"><span data-stu-id="3ee75-243">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="3ee75-244">使用 [`Time.setAsync`](/javascript/api/outlook_1_1/office.time#setasync-datetime--options--callback-) 方法设置结束时间时，应使用 [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) 方法将客户端的本地时间转换为服务器的 UTC。</span><span class="sxs-lookup"><span data-stu-id="3ee75-244">When you use the [`Time.setAsync`](/javascript/api/outlook_1_1/office.time#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="3ee75-245">下面的示例使用[`setAsync`](/javascript/api/outlook_1_1/office.time#setasync-datetime--options--callback-) `Time`对象的方法设置约会的结束时间。</span><span class="sxs-lookup"><span data-stu-id="3ee75-245">The following example sets the end time of an appointment by using the [`setAsync`](/javascript/api/outlook_1_1/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

```javascript
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

##### <a name="type"></a><span data-ttu-id="3ee75-246">类型</span><span class="sxs-lookup"><span data-stu-id="3ee75-246">Type</span></span>

*   <span data-ttu-id="3ee75-247">Date | [Time](/javascript/api/outlook_1_1/office.time)</span><span class="sxs-lookup"><span data-stu-id="3ee75-247">Date | [Time](/javascript/api/outlook_1_1/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="3ee75-248">Requirements</span><span class="sxs-lookup"><span data-stu-id="3ee75-248">Requirements</span></span>

|<span data-ttu-id="3ee75-249">要求</span><span class="sxs-lookup"><span data-stu-id="3ee75-249">Requirement</span></span>| <span data-ttu-id="3ee75-250">值</span><span class="sxs-lookup"><span data-stu-id="3ee75-250">Value</span></span>|
|---|---|
|[<span data-ttu-id="3ee75-251">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="3ee75-251">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3ee75-252">1.0</span><span class="sxs-lookup"><span data-stu-id="3ee75-252">1.0</span></span>|
|[<span data-ttu-id="3ee75-253">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="3ee75-253">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3ee75-254">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3ee75-254">ReadItem</span></span>|
|[<span data-ttu-id="3ee75-255">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="3ee75-255">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="3ee75-256">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="3ee75-256">Compose or Read</span></span>|

#### <a name="from-emailaddressdetailsjavascriptapioutlook11officeemailaddressdetails"></a><span data-ttu-id="3ee75-257">from :[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="3ee75-257">from :[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)</span></span>

<span data-ttu-id="3ee75-p113">获取邮件发件人的电子邮件地址。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="3ee75-p113">Gets the email address of the sender of a message. Read mode only.</span></span>

<span data-ttu-id="3ee75-p114">`from` 和 [`sender`](#sender-emailaddressdetailsjavascriptapioutlook11officeemailaddressdetails) 属性表示同一个人，邮件由代理人发送的除外。在此情况下，`from` 属性表示代理程序，而 sender 属性表示代理人。</span><span class="sxs-lookup"><span data-stu-id="3ee75-p114">The `from` and [`sender`](#sender-emailaddressdetailsjavascriptapioutlook11officeemailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="3ee75-262">`from` 属性中 `EmailAddressDetails` 对象的 `recipientType` 属性为 `undefined`。</span><span class="sxs-lookup"><span data-stu-id="3ee75-262">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="3ee75-263">类型</span><span class="sxs-lookup"><span data-stu-id="3ee75-263">Type</span></span>

*   [<span data-ttu-id="3ee75-264">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="3ee75-264">EmailAddressDetails</span></span>](/javascript/api/outlook_1_1/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="3ee75-265">要求</span><span class="sxs-lookup"><span data-stu-id="3ee75-265">Requirements</span></span>

|<span data-ttu-id="3ee75-266">要求</span><span class="sxs-lookup"><span data-stu-id="3ee75-266">Requirement</span></span>| <span data-ttu-id="3ee75-267">值</span><span class="sxs-lookup"><span data-stu-id="3ee75-267">Value</span></span>|
|---|---|
|[<span data-ttu-id="3ee75-268">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="3ee75-268">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3ee75-269">1.0</span><span class="sxs-lookup"><span data-stu-id="3ee75-269">1.0</span></span>|
|[<span data-ttu-id="3ee75-270">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="3ee75-270">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3ee75-271">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3ee75-271">ReadItem</span></span>|
|[<span data-ttu-id="3ee75-272">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="3ee75-272">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="3ee75-273">阅读</span><span class="sxs-lookup"><span data-stu-id="3ee75-273">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="3ee75-274">示例</span><span class="sxs-lookup"><span data-stu-id="3ee75-274">Example</span></span>

```javascript
var from = Office.context.mailbox.item.from;
console.log("From " + from);
```

#### <a name="internetmessageid-string"></a><span data-ttu-id="3ee75-275">internetMessageId :String</span><span class="sxs-lookup"><span data-stu-id="3ee75-275">internetMessageId :String</span></span>

<span data-ttu-id="3ee75-p115">获取电子邮件的 Internet 消息标识符。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="3ee75-p115">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="3ee75-278">Type</span><span class="sxs-lookup"><span data-stu-id="3ee75-278">Type</span></span>

*   <span data-ttu-id="3ee75-279">String</span><span class="sxs-lookup"><span data-stu-id="3ee75-279">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="3ee75-280">要求</span><span class="sxs-lookup"><span data-stu-id="3ee75-280">Requirements</span></span>

|<span data-ttu-id="3ee75-281">要求</span><span class="sxs-lookup"><span data-stu-id="3ee75-281">Requirement</span></span>| <span data-ttu-id="3ee75-282">值</span><span class="sxs-lookup"><span data-stu-id="3ee75-282">Value</span></span>|
|---|---|
|[<span data-ttu-id="3ee75-283">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="3ee75-283">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3ee75-284">1.0</span><span class="sxs-lookup"><span data-stu-id="3ee75-284">1.0</span></span>|
|[<span data-ttu-id="3ee75-285">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="3ee75-285">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3ee75-286">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3ee75-286">ReadItem</span></span>|
|[<span data-ttu-id="3ee75-287">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="3ee75-287">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="3ee75-288">阅读</span><span class="sxs-lookup"><span data-stu-id="3ee75-288">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="3ee75-289">示例</span><span class="sxs-lookup"><span data-stu-id="3ee75-289">Example</span></span>

```javascript
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

#### <a name="itemclass-string"></a><span data-ttu-id="3ee75-290">itemClass :String</span><span class="sxs-lookup"><span data-stu-id="3ee75-290">itemClass :String</span></span>

<span data-ttu-id="3ee75-p116">获取选定项目的 Exchange Web 服务项目类。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="3ee75-p116">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="3ee75-p117">`itemClass` 属性指定所选项目的邮件类别。以下是邮件或约会项目的默认邮件类别。</span><span class="sxs-lookup"><span data-stu-id="3ee75-p117">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

| <span data-ttu-id="3ee75-295">类型</span><span class="sxs-lookup"><span data-stu-id="3ee75-295">Type</span></span> | <span data-ttu-id="3ee75-296">说明</span><span class="sxs-lookup"><span data-stu-id="3ee75-296">Description</span></span> | <span data-ttu-id="3ee75-297">项目类</span><span class="sxs-lookup"><span data-stu-id="3ee75-297">item class</span></span> |
| --- | --- | --- |
| <span data-ttu-id="3ee75-298">约会项目</span><span class="sxs-lookup"><span data-stu-id="3ee75-298">Appointment items</span></span> | <span data-ttu-id="3ee75-299">这些是项目类 `IPM.Appointment` 或 `IPM.Appointment.Occurrence` 的日历项目。</span><span class="sxs-lookup"><span data-stu-id="3ee75-299">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurrence`.</span></span> | `IPM.Appointment`<br />`IPM.Appointment.Occurrence` |
| <span data-ttu-id="3ee75-300">邮件项目</span><span class="sxs-lookup"><span data-stu-id="3ee75-300">Message items</span></span> | <span data-ttu-id="3ee75-301">这些项目包括具有默认邮件类别 `IPM.Note` 的电子邮件，以及使用 `IPM.Schedule.Meeting` 作为基础邮件类别的会议请求、响应和取消。</span><span class="sxs-lookup"><span data-stu-id="3ee75-301">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span> | `IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled` |

<span data-ttu-id="3ee75-302">你可以创建用于扩展默认邮件类别的自定义邮件类别，例如，自定义约会邮件类别 `IPM.Appointment.Contoso`。</span><span class="sxs-lookup"><span data-stu-id="3ee75-302">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="3ee75-303">Type</span><span class="sxs-lookup"><span data-stu-id="3ee75-303">Type</span></span>

*   <span data-ttu-id="3ee75-304">String</span><span class="sxs-lookup"><span data-stu-id="3ee75-304">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="3ee75-305">要求</span><span class="sxs-lookup"><span data-stu-id="3ee75-305">Requirements</span></span>

|<span data-ttu-id="3ee75-306">要求</span><span class="sxs-lookup"><span data-stu-id="3ee75-306">Requirement</span></span>| <span data-ttu-id="3ee75-307">值</span><span class="sxs-lookup"><span data-stu-id="3ee75-307">Value</span></span>|
|---|---|
|[<span data-ttu-id="3ee75-308">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="3ee75-308">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3ee75-309">1.0</span><span class="sxs-lookup"><span data-stu-id="3ee75-309">1.0</span></span>|
|[<span data-ttu-id="3ee75-310">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="3ee75-310">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3ee75-311">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3ee75-311">ReadItem</span></span>|
|[<span data-ttu-id="3ee75-312">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="3ee75-312">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="3ee75-313">阅读</span><span class="sxs-lookup"><span data-stu-id="3ee75-313">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="3ee75-314">示例</span><span class="sxs-lookup"><span data-stu-id="3ee75-314">Example</span></span>

```javascript
var itemClass = Office.context.mailbox.item.itemClass;
console.log("Item class: " + itemClass);
```

#### <a name="nullable-itemid-string"></a><span data-ttu-id="3ee75-315">(nullable) itemId :String</span><span class="sxs-lookup"><span data-stu-id="3ee75-315">(nullable) itemId :String</span></span>

<span data-ttu-id="3ee75-p118">获取当前项目的 Exchange Web 服务项目标识符。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="3ee75-p118">Gets the Exchange Web Services item identifier for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="3ee75-318">`itemId` 属性返回的标识符与 Exchange Web 服务项目标识符相同。</span><span class="sxs-lookup"><span data-stu-id="3ee75-318">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="3ee75-319">`itemId` 属性与 Outlook 条目 ID 或 Outlook REST API 使用的 ID 不同。</span><span class="sxs-lookup"><span data-stu-id="3ee75-319">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="3ee75-320">使用此值进行 REST API 调用前，应使用 `Office.context.mailbox.convertToRestId`（可在要求集 1.3 的开头部分中找到）对它进行转换。</span><span class="sxs-lookup"><span data-stu-id="3ee75-320">Before making REST API calls using this value, it should be converted using `Office.context.mailbox.convertToRestId`, which is available starting in requirement set 1.3.</span></span> <span data-ttu-id="3ee75-321">有关详细信息，请参阅[从 Outlook 加载项使用 Outlook REST API](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id)。</span><span class="sxs-lookup"><span data-stu-id="3ee75-321">For more details, see [Use the Outlook REST APIs from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

##### <a name="type"></a><span data-ttu-id="3ee75-322">Type</span><span class="sxs-lookup"><span data-stu-id="3ee75-322">Type</span></span>

*   <span data-ttu-id="3ee75-323">String</span><span class="sxs-lookup"><span data-stu-id="3ee75-323">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="3ee75-324">要求</span><span class="sxs-lookup"><span data-stu-id="3ee75-324">Requirements</span></span>

|<span data-ttu-id="3ee75-325">要求</span><span class="sxs-lookup"><span data-stu-id="3ee75-325">Requirement</span></span>| <span data-ttu-id="3ee75-326">值</span><span class="sxs-lookup"><span data-stu-id="3ee75-326">Value</span></span>|
|---|---|
|[<span data-ttu-id="3ee75-327">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="3ee75-327">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3ee75-328">1.0</span><span class="sxs-lookup"><span data-stu-id="3ee75-328">1.0</span></span>|
|[<span data-ttu-id="3ee75-329">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="3ee75-329">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3ee75-330">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3ee75-330">ReadItem</span></span>|
|[<span data-ttu-id="3ee75-331">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="3ee75-331">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="3ee75-332">阅读</span><span class="sxs-lookup"><span data-stu-id="3ee75-332">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="3ee75-333">示例</span><span class="sxs-lookup"><span data-stu-id="3ee75-333">Example</span></span>

<span data-ttu-id="3ee75-p120">以下代码检查项目标识符是否存在。如果 `itemId` 属性返回 `null` 或 `undefined`，则将项目保存到存储，并从异步结果中获取项目标识符。</span><span class="sxs-lookup"><span data-stu-id="3ee75-p120">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

```javascript
var itemId = Office.context.mailbox.item.itemId;
if (itemId === null || itemId == undefined) {
  Office.context.mailbox.item.saveAsync(function(result) {
    itemId = result.value;
  });
}
```

####  <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlook11officemailboxenumsitemtype"></a><span data-ttu-id="3ee75-336">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_1/office.mailboxenums.itemtype)</span><span class="sxs-lookup"><span data-stu-id="3ee75-336">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_1/office.mailboxenums.itemtype)</span></span>

<span data-ttu-id="3ee75-337">获取实例表示的项的类型。</span><span class="sxs-lookup"><span data-stu-id="3ee75-337">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="3ee75-338">`itemType` 属性返回其中一个 `ItemType` 枚举值，指示 `item` 对象实例是邮件还是约会。</span><span class="sxs-lookup"><span data-stu-id="3ee75-338">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="3ee75-339">类型</span><span class="sxs-lookup"><span data-stu-id="3ee75-339">Type</span></span>

*   [<span data-ttu-id="3ee75-340">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="3ee75-340">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook_1_1/office.mailboxenums.itemtype)

##### <a name="requirements"></a><span data-ttu-id="3ee75-341">要求</span><span class="sxs-lookup"><span data-stu-id="3ee75-341">Requirements</span></span>

|<span data-ttu-id="3ee75-342">要求</span><span class="sxs-lookup"><span data-stu-id="3ee75-342">Requirement</span></span>| <span data-ttu-id="3ee75-343">值</span><span class="sxs-lookup"><span data-stu-id="3ee75-343">Value</span></span>|
|---|---|
|[<span data-ttu-id="3ee75-344">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="3ee75-344">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3ee75-345">1.0</span><span class="sxs-lookup"><span data-stu-id="3ee75-345">1.0</span></span>|
|[<span data-ttu-id="3ee75-346">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="3ee75-346">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3ee75-347">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3ee75-347">ReadItem</span></span>|
|[<span data-ttu-id="3ee75-348">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="3ee75-348">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="3ee75-349">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="3ee75-349">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="3ee75-350">示例</span><span class="sxs-lookup"><span data-stu-id="3ee75-350">Example</span></span>

```javascript
if (Office.context.mailbox.item.itemType === Office.MailboxEnums.ItemType.Message) {
  // Do something.
} else {
  // Do something else.
}
```

####  <a name="location-stringlocationjavascriptapioutlook11officelocation"></a><span data-ttu-id="3ee75-351">location :String|[Location](/javascript/api/outlook_1_1/office.location)</span><span class="sxs-lookup"><span data-stu-id="3ee75-351">location :String|[Location](/javascript/api/outlook_1_1/office.location)</span></span>

<span data-ttu-id="3ee75-352">获取或设置约会的位置。</span><span class="sxs-lookup"><span data-stu-id="3ee75-352">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="3ee75-353">阅读模式</span><span class="sxs-lookup"><span data-stu-id="3ee75-353">Read mode</span></span>

<span data-ttu-id="3ee75-354">`location` 属性返回一个包含约会位置的字符串。</span><span class="sxs-lookup"><span data-stu-id="3ee75-354">The `location` property returns a string that contains the location of the appointment.</span></span>

```javascript
var location = Office.context.mailbox.item.location;
console.log("location: " + location);
```

##### <a name="compose-mode"></a><span data-ttu-id="3ee75-355">撰写模式</span><span class="sxs-lookup"><span data-stu-id="3ee75-355">Compose mode</span></span>

<span data-ttu-id="3ee75-356">`location` 属性返回一个 `Location` 对象，该对象提供用于获取和设置约会位置的方法。</span><span class="sxs-lookup"><span data-stu-id="3ee75-356">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

```javascript
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="3ee75-357">类型</span><span class="sxs-lookup"><span data-stu-id="3ee75-357">Type</span></span>

*   <span data-ttu-id="3ee75-358">String | [Location](/javascript/api/outlook_1_1/office.location)</span><span class="sxs-lookup"><span data-stu-id="3ee75-358">String | [Location](/javascript/api/outlook_1_1/office.location)</span></span>

##### <a name="requirements"></a><span data-ttu-id="3ee75-359">要求</span><span class="sxs-lookup"><span data-stu-id="3ee75-359">Requirements</span></span>

|<span data-ttu-id="3ee75-360">要求</span><span class="sxs-lookup"><span data-stu-id="3ee75-360">Requirement</span></span>| <span data-ttu-id="3ee75-361">值</span><span class="sxs-lookup"><span data-stu-id="3ee75-361">Value</span></span>|
|---|---|
|[<span data-ttu-id="3ee75-362">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="3ee75-362">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3ee75-363">1.0</span><span class="sxs-lookup"><span data-stu-id="3ee75-363">1.0</span></span>|
|[<span data-ttu-id="3ee75-364">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="3ee75-364">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3ee75-365">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3ee75-365">ReadItem</span></span>|
|[<span data-ttu-id="3ee75-366">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="3ee75-366">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="3ee75-367">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="3ee75-367">Compose or Read</span></span>|

#### <a name="normalizedsubject-string"></a><span data-ttu-id="3ee75-368">normalizedSubject :String</span><span class="sxs-lookup"><span data-stu-id="3ee75-368">normalizedSubject :String</span></span>

<span data-ttu-id="3ee75-p121">获取删除了所有前缀（包括 `RE:` 和 `FWD:`）的项目主题。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="3ee75-p121">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="3ee75-p122">normalizedSubject 属性获取包含由电子邮件程序添加的任何标准前缀（如 `RE:` 和 `FW:`）的项目主题。若要获取包含完整前缀的项目主题，请使用 [`subject`](#subject-stringsubjectjavascriptapioutlook11officesubject) 属性。</span><span class="sxs-lookup"><span data-stu-id="3ee75-p122">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubjectjavascriptapioutlook11officesubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="3ee75-373">Type</span><span class="sxs-lookup"><span data-stu-id="3ee75-373">Type</span></span>

*   <span data-ttu-id="3ee75-374">String</span><span class="sxs-lookup"><span data-stu-id="3ee75-374">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="3ee75-375">要求</span><span class="sxs-lookup"><span data-stu-id="3ee75-375">Requirements</span></span>

|<span data-ttu-id="3ee75-376">要求</span><span class="sxs-lookup"><span data-stu-id="3ee75-376">Requirement</span></span>| <span data-ttu-id="3ee75-377">值</span><span class="sxs-lookup"><span data-stu-id="3ee75-377">Value</span></span>|
|---|---|
|[<span data-ttu-id="3ee75-378">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="3ee75-378">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3ee75-379">1.0</span><span class="sxs-lookup"><span data-stu-id="3ee75-379">1.0</span></span>|
|[<span data-ttu-id="3ee75-380">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="3ee75-380">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3ee75-381">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3ee75-381">ReadItem</span></span>|
|[<span data-ttu-id="3ee75-382">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="3ee75-382">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="3ee75-383">阅读</span><span class="sxs-lookup"><span data-stu-id="3ee75-383">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="3ee75-384">示例</span><span class="sxs-lookup"><span data-stu-id="3ee75-384">Example</span></span>

```javascript
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
console.log("Normalized subject: " + normalizedSubject);
```

####  <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlook11officeemailaddressdetailsrecipientsjavascriptapioutlook11officerecipients"></a><span data-ttu-id="3ee75-385">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_1/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="3ee75-385">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_1/office.recipients)</span></span>

<span data-ttu-id="3ee75-386">提供对事件的可选与会者的访问权限。</span><span class="sxs-lookup"><span data-stu-id="3ee75-386">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="3ee75-387">对象的类型和访问级别取决于当前项目的模式。</span><span class="sxs-lookup"><span data-stu-id="3ee75-387">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="3ee75-388">阅读模式</span><span class="sxs-lookup"><span data-stu-id="3ee75-388">Read mode</span></span>

<span data-ttu-id="3ee75-389">`optionalAttendees` 属性返回一个数组，其中包含每个可选与会者的 `EmailAddressDetails` 对象。</span><span class="sxs-lookup"><span data-stu-id="3ee75-389">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span>

```javascript
var optionalAttendees = Office.context.mailbox.item.optionalAttendees;
console.log("Optional attendees: " + JSON.stringify(optionalAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="3ee75-390">撰写模式</span><span class="sxs-lookup"><span data-stu-id="3ee75-390">Compose mode</span></span>

<span data-ttu-id="3ee75-391">`optionalAttendees` 属性返回一个 `Recipients` 对象，该对象提供用于获取或更新可选与会者的方法。</span><span class="sxs-lookup"><span data-stu-id="3ee75-391">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span>

```javascript
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="3ee75-392">类型</span><span class="sxs-lookup"><span data-stu-id="3ee75-392">Type</span></span>

*   <span data-ttu-id="3ee75-393">Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_1/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="3ee75-393">Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_1/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="3ee75-394">Requirements</span><span class="sxs-lookup"><span data-stu-id="3ee75-394">Requirements</span></span>

|<span data-ttu-id="3ee75-395">要求</span><span class="sxs-lookup"><span data-stu-id="3ee75-395">Requirement</span></span>| <span data-ttu-id="3ee75-396">值</span><span class="sxs-lookup"><span data-stu-id="3ee75-396">Value</span></span>|
|---|---|
|[<span data-ttu-id="3ee75-397">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="3ee75-397">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3ee75-398">1.0</span><span class="sxs-lookup"><span data-stu-id="3ee75-398">1.0</span></span>|
|[<span data-ttu-id="3ee75-399">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="3ee75-399">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3ee75-400">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3ee75-400">ReadItem</span></span>|
|[<span data-ttu-id="3ee75-401">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="3ee75-401">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="3ee75-402">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="3ee75-402">Compose or Read</span></span>|

#### <a name="organizer-emailaddressdetailsjavascriptapioutlook11officeemailaddressdetails"></a><span data-ttu-id="3ee75-403">organizer :[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="3ee75-403">organizer :[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)</span></span>

<span data-ttu-id="3ee75-p124">获取指定会议的会议组织者的电子邮件地址。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="3ee75-p124">Gets the email address of the meeting organizer for a specified meeting. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="3ee75-406">类型</span><span class="sxs-lookup"><span data-stu-id="3ee75-406">Type</span></span>

*   [<span data-ttu-id="3ee75-407">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="3ee75-407">EmailAddressDetails</span></span>](/javascript/api/outlook_1_1/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="3ee75-408">Requirements</span><span class="sxs-lookup"><span data-stu-id="3ee75-408">Requirements</span></span>

|<span data-ttu-id="3ee75-409">要求</span><span class="sxs-lookup"><span data-stu-id="3ee75-409">Requirement</span></span>| <span data-ttu-id="3ee75-410">值</span><span class="sxs-lookup"><span data-stu-id="3ee75-410">Value</span></span>|
|---|---|
|[<span data-ttu-id="3ee75-411">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="3ee75-411">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3ee75-412">1.0</span><span class="sxs-lookup"><span data-stu-id="3ee75-412">1.0</span></span>|
|[<span data-ttu-id="3ee75-413">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="3ee75-413">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3ee75-414">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3ee75-414">ReadItem</span></span>|
|[<span data-ttu-id="3ee75-415">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="3ee75-415">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="3ee75-416">阅读</span><span class="sxs-lookup"><span data-stu-id="3ee75-416">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="3ee75-417">示例</span><span class="sxs-lookup"><span data-stu-id="3ee75-417">Example</span></span>

```javascript
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
console.log("Organizer: " + organizerName + " (" + organizerAddress + ")");
```

####  <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlook11officeemailaddressdetailsrecipientsjavascriptapioutlook11officerecipients"></a><span data-ttu-id="3ee75-418">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_1/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="3ee75-418">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_1/office.recipients)</span></span>

<span data-ttu-id="3ee75-419">提供对事件的必需与会者的访问权限。</span><span class="sxs-lookup"><span data-stu-id="3ee75-419">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="3ee75-420">对象的类型和访问级别取决于当前项目的模式。</span><span class="sxs-lookup"><span data-stu-id="3ee75-420">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="3ee75-421">阅读模式</span><span class="sxs-lookup"><span data-stu-id="3ee75-421">Read mode</span></span>

<span data-ttu-id="3ee75-422">`requiredAttendees` 属性返回一个数组，其中包含每个必需与会者的 `EmailAddressDetails` 对象。</span><span class="sxs-lookup"><span data-stu-id="3ee75-422">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span>

```javascript
var requiredAttendees = Office.context.mailbox.item.requiredAttendees;
console.log("Required attendees: " + JSON.stringify(requiredAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="3ee75-423">撰写模式</span><span class="sxs-lookup"><span data-stu-id="3ee75-423">Compose mode</span></span>

<span data-ttu-id="3ee75-424">`requiredAttendees` 属性返回一个 `Recipients` 对象，该对象提供用于获取或更新必需与会者的方法。</span><span class="sxs-lookup"><span data-stu-id="3ee75-424">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span>

```javascript
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
  console.log(JSON.stringify(arrayOfRequiredAttendeesRecipients));
}
```

##### <a name="type"></a><span data-ttu-id="3ee75-425">类型</span><span class="sxs-lookup"><span data-stu-id="3ee75-425">Type</span></span>

*   <span data-ttu-id="3ee75-426">Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_1/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="3ee75-426">Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_1/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="3ee75-427">Requirements</span><span class="sxs-lookup"><span data-stu-id="3ee75-427">Requirements</span></span>

|<span data-ttu-id="3ee75-428">要求</span><span class="sxs-lookup"><span data-stu-id="3ee75-428">Requirement</span></span>| <span data-ttu-id="3ee75-429">值</span><span class="sxs-lookup"><span data-stu-id="3ee75-429">Value</span></span>|
|---|---|
|[<span data-ttu-id="3ee75-430">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="3ee75-430">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3ee75-431">1.0</span><span class="sxs-lookup"><span data-stu-id="3ee75-431">1.0</span></span>|
|[<span data-ttu-id="3ee75-432">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="3ee75-432">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3ee75-433">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3ee75-433">ReadItem</span></span>|
|[<span data-ttu-id="3ee75-434">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="3ee75-434">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="3ee75-435">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="3ee75-435">Compose or Read</span></span>|

#### <a name="sender-emailaddressdetailsjavascriptapioutlook11officeemailaddressdetails"></a><span data-ttu-id="3ee75-436">sender :[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="3ee75-436">sender :[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)</span></span>

<span data-ttu-id="3ee75-p126">获取电子邮件发件人的电子邮件地址。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="3ee75-p126">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="3ee75-p127">[`from`](#from-emailaddressdetailsjavascriptapioutlook11officeemailaddressdetails) 和 `sender` 属性表示同一个人，邮件由代理人发送的除外。在此情况下，`from` 属性表示代理程序，而 sender 属性表示代理人。</span><span class="sxs-lookup"><span data-stu-id="3ee75-p127">The [`from`](#from-emailaddressdetailsjavascriptapioutlook11officeemailaddressdetails) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="3ee75-441">`from` 属性中 `EmailAddressDetails` 对象的 `recipientType` 属性为 `undefined`。</span><span class="sxs-lookup"><span data-stu-id="3ee75-441">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="3ee75-442">类型</span><span class="sxs-lookup"><span data-stu-id="3ee75-442">Type</span></span>

*   [<span data-ttu-id="3ee75-443">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="3ee75-443">EmailAddressDetails</span></span>](/javascript/api/outlook_1_1/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="3ee75-444">Requirements</span><span class="sxs-lookup"><span data-stu-id="3ee75-444">Requirements</span></span>

|<span data-ttu-id="3ee75-445">要求</span><span class="sxs-lookup"><span data-stu-id="3ee75-445">Requirement</span></span>| <span data-ttu-id="3ee75-446">值</span><span class="sxs-lookup"><span data-stu-id="3ee75-446">Value</span></span>|
|---|---|
|[<span data-ttu-id="3ee75-447">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="3ee75-447">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3ee75-448">1.0</span><span class="sxs-lookup"><span data-stu-id="3ee75-448">1.0</span></span>|
|[<span data-ttu-id="3ee75-449">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="3ee75-449">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3ee75-450">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3ee75-450">ReadItem</span></span>|
|[<span data-ttu-id="3ee75-451">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="3ee75-451">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="3ee75-452">阅读</span><span class="sxs-lookup"><span data-stu-id="3ee75-452">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="3ee75-453">示例</span><span class="sxs-lookup"><span data-stu-id="3ee75-453">Example</span></span>

```javascript
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
console.log("Sender: " + senderName + " (" + senderAddress + ")");
```

####  <a name="start-datetimejavascriptapioutlook11officetime"></a><span data-ttu-id="3ee75-454">start :Date|[Time](/javascript/api/outlook_1_1/office.time)</span><span class="sxs-lookup"><span data-stu-id="3ee75-454">start :Date|[Time](/javascript/api/outlook_1_1/office.time)</span></span>

<span data-ttu-id="3ee75-455">获取或设置约会开始的日期和时间。</span><span class="sxs-lookup"><span data-stu-id="3ee75-455">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="3ee75-p128">将 `start` 属性表示为协调世界时 (UTC) 的日期和时间值。可使用 [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook11officelocalclienttime) 方法将该值转换为客户端的本地日期和时间。</span><span class="sxs-lookup"><span data-stu-id="3ee75-p128">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook11officelocalclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="3ee75-458">阅读模式</span><span class="sxs-lookup"><span data-stu-id="3ee75-458">Read mode</span></span>

<span data-ttu-id="3ee75-459">`start` 属性返回 `Date` 对象。</span><span class="sxs-lookup"><span data-stu-id="3ee75-459">The `start` property returns a `Date` object.</span></span>

```javascript
var start = Office.context.mailbox.item.start;
console.log("Appointment start: " + JSON.stringify(start));
```

##### <a name="compose-mode"></a><span data-ttu-id="3ee75-460">撰写模式</span><span class="sxs-lookup"><span data-stu-id="3ee75-460">Compose mode</span></span>

<span data-ttu-id="3ee75-461">`start` 属性返回 `Time` 对象。</span><span class="sxs-lookup"><span data-stu-id="3ee75-461">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="3ee75-462">使用 [`Time.setAsync`](/javascript/api/outlook_1_1/office.time#setasync-datetime--options--callback-) 方法设置开始时间时，应使用 [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) 方法将客户端的本地时间转换为服务器的 UTC。</span><span class="sxs-lookup"><span data-stu-id="3ee75-462">When you use the [`Time.setAsync`](/javascript/api/outlook_1_1/office.time#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="3ee75-463">以下示例通过使用 `Time` 对象的 [`setAsync`](/javascript/api/outlook_1_1/office.time#setasync-datetime--options--callback-) 方法，设置撰写模式下约会的开始时间。</span><span class="sxs-lookup"><span data-stu-id="3ee75-463">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_1/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

```javascript
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

##### <a name="type"></a><span data-ttu-id="3ee75-464">类型</span><span class="sxs-lookup"><span data-stu-id="3ee75-464">Type</span></span>

*   <span data-ttu-id="3ee75-465">Date | [Time](/javascript/api/outlook_1_1/office.time)</span><span class="sxs-lookup"><span data-stu-id="3ee75-465">Date | [Time](/javascript/api/outlook_1_1/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="3ee75-466">Requirements</span><span class="sxs-lookup"><span data-stu-id="3ee75-466">Requirements</span></span>

|<span data-ttu-id="3ee75-467">要求</span><span class="sxs-lookup"><span data-stu-id="3ee75-467">Requirement</span></span>| <span data-ttu-id="3ee75-468">值</span><span class="sxs-lookup"><span data-stu-id="3ee75-468">Value</span></span>|
|---|---|
|[<span data-ttu-id="3ee75-469">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="3ee75-469">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3ee75-470">1.0</span><span class="sxs-lookup"><span data-stu-id="3ee75-470">1.0</span></span>|
|[<span data-ttu-id="3ee75-471">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="3ee75-471">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3ee75-472">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3ee75-472">ReadItem</span></span>|
|[<span data-ttu-id="3ee75-473">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="3ee75-473">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="3ee75-474">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="3ee75-474">Compose or Read</span></span>|

####  <a name="subject-stringsubjectjavascriptapioutlook11officesubject"></a><span data-ttu-id="3ee75-475">subject :String|[Subject](/javascript/api/outlook_1_1/office.subject)</span><span class="sxs-lookup"><span data-stu-id="3ee75-475">subject :String|[Subject](/javascript/api/outlook_1_1/office.subject)</span></span>

<span data-ttu-id="3ee75-476">获取或设置显示在项目的主题字段中的说明。</span><span class="sxs-lookup"><span data-stu-id="3ee75-476">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="3ee75-477">`subject` 属性获取或设置由电子邮件服务器发送项目时的整个主题。</span><span class="sxs-lookup"><span data-stu-id="3ee75-477">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="3ee75-478">阅读模式</span><span class="sxs-lookup"><span data-stu-id="3ee75-478">Read mode</span></span>

<span data-ttu-id="3ee75-p129">`subject` 属性返回一个字符串。使用 [`normalizedSubject`](#normalizedsubject-string) 属性获取不带任何前导前缀（如 `RE:` 和 `FW:`）的主题。</span><span class="sxs-lookup"><span data-stu-id="3ee75-p129">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

```js
var subject = Office.context.mailbox.item.subject;
console.log(subject);
```

##### <a name="compose-mode"></a><span data-ttu-id="3ee75-481">撰写模式</span><span class="sxs-lookup"><span data-stu-id="3ee75-481">Compose mode</span></span>

<span data-ttu-id="3ee75-482">`subject` 属性返回一个 `Subject` 对象，该对象提供用于获取和设置主题的方法。</span><span class="sxs-lookup"><span data-stu-id="3ee75-482">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```javascript
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
  console.log(subject);
}
```

##### <a name="type"></a><span data-ttu-id="3ee75-483">类型</span><span class="sxs-lookup"><span data-stu-id="3ee75-483">Type</span></span>

*   <span data-ttu-id="3ee75-484">String | [Subject](/javascript/api/outlook_1_1/office.subject)</span><span class="sxs-lookup"><span data-stu-id="3ee75-484">String | [Subject](/javascript/api/outlook_1_1/office.subject)</span></span>

##### <a name="requirements"></a><span data-ttu-id="3ee75-485">要求</span><span class="sxs-lookup"><span data-stu-id="3ee75-485">Requirements</span></span>

|<span data-ttu-id="3ee75-486">要求</span><span class="sxs-lookup"><span data-stu-id="3ee75-486">Requirement</span></span>| <span data-ttu-id="3ee75-487">值</span><span class="sxs-lookup"><span data-stu-id="3ee75-487">Value</span></span>|
|---|---|
|[<span data-ttu-id="3ee75-488">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="3ee75-488">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3ee75-489">1.0</span><span class="sxs-lookup"><span data-stu-id="3ee75-489">1.0</span></span>|
|[<span data-ttu-id="3ee75-490">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="3ee75-490">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3ee75-491">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3ee75-491">ReadItem</span></span>|
|[<span data-ttu-id="3ee75-492">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="3ee75-492">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="3ee75-493">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="3ee75-493">Compose or Read</span></span>|

####  <a name="to-arrayemailaddressdetailsjavascriptapioutlook11officeemailaddressdetailsrecipientsjavascriptapioutlook11officerecipients"></a><span data-ttu-id="3ee75-494">to :Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_1/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="3ee75-494">to :Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_1/office.recipients)</span></span>

<span data-ttu-id="3ee75-495">提供对邮件的“**收件人**”行上的收件人的访问权限。</span><span class="sxs-lookup"><span data-stu-id="3ee75-495">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="3ee75-496">对象的类型和访问级别取决于当前项目的模式。</span><span class="sxs-lookup"><span data-stu-id="3ee75-496">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="3ee75-497">阅读模式</span><span class="sxs-lookup"><span data-stu-id="3ee75-497">Read mode</span></span>

<span data-ttu-id="3ee75-p131">`to` 属性返回包含邮件的**收件人**行上所列的每个收件人的 `EmailAddressDetails` 对象的数组。集合上限为 100 个成员。</span><span class="sxs-lookup"><span data-stu-id="3ee75-p131">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.</span></span>

```javascript
console.log(JSON.stringify(Office.context.mailbox.item.to));
```

##### <a name="compose-mode"></a><span data-ttu-id="3ee75-500">撰写模式</span><span class="sxs-lookup"><span data-stu-id="3ee75-500">Compose mode</span></span>

<span data-ttu-id="3ee75-501">`to` 属性返回一个 `Recipients` 对象，该对象提供用于获取或更新邮件的“**收件人**”行上收件人的方法。</span><span class="sxs-lookup"><span data-stu-id="3ee75-501">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span>

```javascript
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="3ee75-502">类型</span><span class="sxs-lookup"><span data-stu-id="3ee75-502">Type</span></span>

*   <span data-ttu-id="3ee75-503">Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_1/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="3ee75-503">Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_1/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="3ee75-504">Requirements</span><span class="sxs-lookup"><span data-stu-id="3ee75-504">Requirements</span></span>

|<span data-ttu-id="3ee75-505">要求</span><span class="sxs-lookup"><span data-stu-id="3ee75-505">Requirement</span></span>| <span data-ttu-id="3ee75-506">值</span><span class="sxs-lookup"><span data-stu-id="3ee75-506">Value</span></span>|
|---|---|
|[<span data-ttu-id="3ee75-507">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="3ee75-507">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3ee75-508">1.0</span><span class="sxs-lookup"><span data-stu-id="3ee75-508">1.0</span></span>|
|[<span data-ttu-id="3ee75-509">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="3ee75-509">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3ee75-510">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3ee75-510">ReadItem</span></span>|
|[<span data-ttu-id="3ee75-511">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="3ee75-511">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="3ee75-512">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="3ee75-512">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="3ee75-513">方法</span><span class="sxs-lookup"><span data-stu-id="3ee75-513">Methods</span></span>

####  <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="3ee75-514">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="3ee75-514">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="3ee75-515">将文件作为附件添加到邮件或约会。</span><span class="sxs-lookup"><span data-stu-id="3ee75-515">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="3ee75-516">`addFileAttachmentAsync` 方法在指定的 URI 上载文件并将其附加到撰写窗体中的项目。</span><span class="sxs-lookup"><span data-stu-id="3ee75-516">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="3ee75-517">你随后可以将该标识符与 [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) 方法一同使用，以删除同一个会话中的附件。</span><span class="sxs-lookup"><span data-stu-id="3ee75-517">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="3ee75-518">Parameters</span><span class="sxs-lookup"><span data-stu-id="3ee75-518">Parameters</span></span>

|<span data-ttu-id="3ee75-519">名称</span><span class="sxs-lookup"><span data-stu-id="3ee75-519">Name</span></span>| <span data-ttu-id="3ee75-520">类型</span><span class="sxs-lookup"><span data-stu-id="3ee75-520">Type</span></span>| <span data-ttu-id="3ee75-521">属性</span><span class="sxs-lookup"><span data-stu-id="3ee75-521">Attributes</span></span>| <span data-ttu-id="3ee75-522">说明</span><span class="sxs-lookup"><span data-stu-id="3ee75-522">Description</span></span>|
|---|---|---|---|
|`uri`| <span data-ttu-id="3ee75-523">String</span><span class="sxs-lookup"><span data-stu-id="3ee75-523">String</span></span>||<span data-ttu-id="3ee75-p132">提供附加到邮件或约会的文件的位置的 URI。最大长度为 2048 个字符。</span><span class="sxs-lookup"><span data-stu-id="3ee75-p132">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="3ee75-526">字符串</span><span class="sxs-lookup"><span data-stu-id="3ee75-526">String</span></span>||<span data-ttu-id="3ee75-p133">在附件上载过程中显示的附件名称。最大长度为 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="3ee75-p133">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="3ee75-529">Object</span><span class="sxs-lookup"><span data-stu-id="3ee75-529">Object</span></span>| <span data-ttu-id="3ee75-530">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="3ee75-530">&lt;optional&gt;</span></span>|<span data-ttu-id="3ee75-531">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="3ee75-531">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="3ee75-532">对象</span><span class="sxs-lookup"><span data-stu-id="3ee75-532">Object</span></span>| <span data-ttu-id="3ee75-533">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="3ee75-533">&lt;optional&gt;</span></span>|<span data-ttu-id="3ee75-534">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="3ee75-534">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="3ee75-535">函数</span><span class="sxs-lookup"><span data-stu-id="3ee75-535">function</span></span>| <span data-ttu-id="3ee75-536">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="3ee75-536">&lt;optional&gt;</span></span>|<span data-ttu-id="3ee75-537">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="3ee75-537">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="3ee75-538">如果成功，附件标识符将在 `asyncResult.value` 属性中提供。</span><span class="sxs-lookup"><span data-stu-id="3ee75-538">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="3ee75-539">如果上传附件失败，`asyncResult` 对象将包含一个提供错误说明的 `Error` 对象。</span><span class="sxs-lookup"><span data-stu-id="3ee75-539">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="3ee75-540">错误</span><span class="sxs-lookup"><span data-stu-id="3ee75-540">Errors</span></span>

| <span data-ttu-id="3ee75-541">错误代码</span><span class="sxs-lookup"><span data-stu-id="3ee75-541">Error code</span></span> | <span data-ttu-id="3ee75-542">说明</span><span class="sxs-lookup"><span data-stu-id="3ee75-542">Description</span></span> |
|------------|-------------|
| `AttachmentSizeExceeded` | <span data-ttu-id="3ee75-543">附件大小超过了允许的大小。</span><span class="sxs-lookup"><span data-stu-id="3ee75-543">The attachment is larger than allowed.</span></span> |
| `FileTypeNotSupported` | <span data-ttu-id="3ee75-544">该附件的扩展名不是允许的扩展名。</span><span class="sxs-lookup"><span data-stu-id="3ee75-544">The attachment has an extension that is not allowed.</span></span> |
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="3ee75-545">邮件或约会具有的附件过多。</span><span class="sxs-lookup"><span data-stu-id="3ee75-545">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="3ee75-546">要求</span><span class="sxs-lookup"><span data-stu-id="3ee75-546">Requirements</span></span>

|<span data-ttu-id="3ee75-547">要求</span><span class="sxs-lookup"><span data-stu-id="3ee75-547">Requirement</span></span>| <span data-ttu-id="3ee75-548">值</span><span class="sxs-lookup"><span data-stu-id="3ee75-548">Value</span></span>|
|---|---|
|[<span data-ttu-id="3ee75-549">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="3ee75-549">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3ee75-550">1.1</span><span class="sxs-lookup"><span data-stu-id="3ee75-550">1.1</span></span>|
|[<span data-ttu-id="3ee75-551">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="3ee75-551">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3ee75-552">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="3ee75-552">ReadWriteItem</span></span>|
|[<span data-ttu-id="3ee75-553">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="3ee75-553">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="3ee75-554">撰写</span><span class="sxs-lookup"><span data-stu-id="3ee75-554">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="3ee75-555">示例</span><span class="sxs-lookup"><span data-stu-id="3ee75-555">Example</span></span>

```javascript
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

####  <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="3ee75-556">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="3ee75-556">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="3ee75-557">将 Exchange 项目（如邮件）作为附件添加到邮件或约会。</span><span class="sxs-lookup"><span data-stu-id="3ee75-557">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="3ee75-p134">`addItemAttachmentAsync` 方法将包含指定 Exchange 标识符的项目附加到撰写窗体中的项目。如果指定一个回调方法，此方法使用 `asyncResult` 参数调用，该参数包含一个附件标识符或代码，指示附加项目过程中出现的任何错误。可以使用 `options` 参数将状态信息传递给回调方法（如果需要）。</span><span class="sxs-lookup"><span data-stu-id="3ee75-p134">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="3ee75-561">随后可以将该标识符与 [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) 方法一同使用，以删除同一个会话中的附件。</span><span class="sxs-lookup"><span data-stu-id="3ee75-561">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="3ee75-562">如果 Office 加载项在 Outlook Web App 中运行，则 `addItemAttachmentAsync` 方法可以将项目附加到项目（正在编辑的项目除外）中；然而，不支持也不建议这样做。</span><span class="sxs-lookup"><span data-stu-id="3ee75-562">If your Office Add-in is running in Outlook Web App, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="3ee75-563">Parameters</span><span class="sxs-lookup"><span data-stu-id="3ee75-563">Parameters</span></span>

|<span data-ttu-id="3ee75-564">名称</span><span class="sxs-lookup"><span data-stu-id="3ee75-564">Name</span></span>| <span data-ttu-id="3ee75-565">类型</span><span class="sxs-lookup"><span data-stu-id="3ee75-565">Type</span></span>| <span data-ttu-id="3ee75-566">属性</span><span class="sxs-lookup"><span data-stu-id="3ee75-566">Attributes</span></span>| <span data-ttu-id="3ee75-567">说明</span><span class="sxs-lookup"><span data-stu-id="3ee75-567">Description</span></span>|
|---|---|---|---|
|`itemId`| <span data-ttu-id="3ee75-568">String</span><span class="sxs-lookup"><span data-stu-id="3ee75-568">String</span></span>||<span data-ttu-id="3ee75-p135">要附加的项目的 Exchange 标识符。最大长度为 100 个字符。</span><span class="sxs-lookup"><span data-stu-id="3ee75-p135">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="3ee75-571">String</span><span class="sxs-lookup"><span data-stu-id="3ee75-571">String</span></span>||<span data-ttu-id="3ee75-572">要附加的项目的主题。</span><span class="sxs-lookup"><span data-stu-id="3ee75-572">The subject of the item to be attached.</span></span> <span data-ttu-id="3ee75-573">最大长度为 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="3ee75-573">The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="3ee75-574">对象</span><span class="sxs-lookup"><span data-stu-id="3ee75-574">Object</span></span>| <span data-ttu-id="3ee75-575">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="3ee75-575">&lt;optional&gt;</span></span>|<span data-ttu-id="3ee75-576">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="3ee75-576">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="3ee75-577">对象</span><span class="sxs-lookup"><span data-stu-id="3ee75-577">Object</span></span>| <span data-ttu-id="3ee75-578">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="3ee75-578">&lt;optional&gt;</span></span>|<span data-ttu-id="3ee75-579">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="3ee75-579">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="3ee75-580">函数</span><span class="sxs-lookup"><span data-stu-id="3ee75-580">function</span></span>| <span data-ttu-id="3ee75-581">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="3ee75-581">&lt;optional&gt;</span></span>|<span data-ttu-id="3ee75-582">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="3ee75-582">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="3ee75-583">如果成功，附件标识符将在 `asyncResult.value` 属性中提供。</span><span class="sxs-lookup"><span data-stu-id="3ee75-583">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="3ee75-584">如果添加附件失败，`asyncResult` 对象将包含一个提供错误说明的 `Error` 对象。</span><span class="sxs-lookup"><span data-stu-id="3ee75-584">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="3ee75-585">错误</span><span class="sxs-lookup"><span data-stu-id="3ee75-585">Errors</span></span>

| <span data-ttu-id="3ee75-586">错误代码</span><span class="sxs-lookup"><span data-stu-id="3ee75-586">Error code</span></span> | <span data-ttu-id="3ee75-587">说明</span><span class="sxs-lookup"><span data-stu-id="3ee75-587">Description</span></span> |
|------------|-------------|
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="3ee75-588">邮件或约会具有的附件过多。</span><span class="sxs-lookup"><span data-stu-id="3ee75-588">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="3ee75-589">Requirements</span><span class="sxs-lookup"><span data-stu-id="3ee75-589">Requirements</span></span>

|<span data-ttu-id="3ee75-590">要求</span><span class="sxs-lookup"><span data-stu-id="3ee75-590">Requirement</span></span>| <span data-ttu-id="3ee75-591">值</span><span class="sxs-lookup"><span data-stu-id="3ee75-591">Value</span></span>|
|---|---|
|[<span data-ttu-id="3ee75-592">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="3ee75-592">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3ee75-593">1.1</span><span class="sxs-lookup"><span data-stu-id="3ee75-593">1.1</span></span>|
|[<span data-ttu-id="3ee75-594">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="3ee75-594">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3ee75-595">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="3ee75-595">ReadWriteItem</span></span>|
|[<span data-ttu-id="3ee75-596">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="3ee75-596">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="3ee75-597">撰写</span><span class="sxs-lookup"><span data-stu-id="3ee75-597">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="3ee75-598">示例</span><span class="sxs-lookup"><span data-stu-id="3ee75-598">Example</span></span>

<span data-ttu-id="3ee75-599">以下示例将现有的 Outlook 项目添加为名为 `My Attachment` 的附件。</span><span class="sxs-lookup"><span data-stu-id="3ee75-599">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

```javascript
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

#### <a name="displayreplyallformformdata-callback"></a><span data-ttu-id="3ee75-600">displayReplyAllForm (formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="3ee75-600">displayReplyAllForm(formData, [callback])</span></span>

<span data-ttu-id="3ee75-601">显示答复窗体，其中包括所选邮件的发件人和所有收件人或所选约会的组织者和所有与会者。</span><span class="sxs-lookup"><span data-stu-id="3ee75-601">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="3ee75-602">在 Outlook for iOS 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="3ee75-602">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="3ee75-603">在 Outlook Web App 中，答复窗体显示为包含 3 列视图的弹出式窗体以及包含 2 列或 1 列视图的弹出式窗体。</span><span class="sxs-lookup"><span data-stu-id="3ee75-603">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="3ee75-604">如果任意字符串参数超出其限制，`displayReplyAllForm` 将引发异常。</span><span class="sxs-lookup"><span data-stu-id="3ee75-604">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

> [!NOTE]
> <span data-ttu-id="3ee75-605">要求集 1.1 不支持 `displayReplyAllForm` 在调用中包括附件的功能。</span><span class="sxs-lookup"><span data-stu-id="3ee75-605">The ability to include attachments in the call to `displayReplyAllForm` is not supported in requirement set 1.1.</span></span> <span data-ttu-id="3ee75-606">附件支持已添加到要求集 1.2 及以上的 `displayReplyAllForm` 中。</span><span class="sxs-lookup"><span data-stu-id="3ee75-606">Attachment support was added to `displayReplyAllForm` in requirement set 1.2 and above.</span></span>

##### <a name="parameters"></a><span data-ttu-id="3ee75-607">Parameters</span><span class="sxs-lookup"><span data-stu-id="3ee75-607">Parameters</span></span>

|<span data-ttu-id="3ee75-608">名称</span><span class="sxs-lookup"><span data-stu-id="3ee75-608">Name</span></span>| <span data-ttu-id="3ee75-609">说明</span><span class="sxs-lookup"><span data-stu-id="3ee75-609">Type</span></span>| <span data-ttu-id="3ee75-610">说明</span><span class="sxs-lookup"><span data-stu-id="3ee75-610">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="3ee75-611">字符串 &#124; 对象</span><span class="sxs-lookup"><span data-stu-id="3ee75-611">String &#124; Object</span></span>| |<span data-ttu-id="3ee75-p138">一个包含文本和 HTML 且表示答复窗体的正文的字符串。字符串限制为 32 KB。</span><span class="sxs-lookup"><span data-stu-id="3ee75-p138">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="3ee75-614">**或**</span><span class="sxs-lookup"><span data-stu-id="3ee75-614">**OR**</span></span><br/><span data-ttu-id="3ee75-p139">包含正文或附件数据和回调函数的对象。对象定义如下。</span><span class="sxs-lookup"><span data-stu-id="3ee75-p139">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="3ee75-617">字符串</span><span class="sxs-lookup"><span data-stu-id="3ee75-617">String</span></span> | <span data-ttu-id="3ee75-618">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="3ee75-618">&lt;optional&gt;</span></span> | <span data-ttu-id="3ee75-p140">一个包含文本和 HTML 且表示答复窗体的正文的字符串。字符串限制为 32 KB。</span><span class="sxs-lookup"><span data-stu-id="3ee75-p140">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `callback` | <span data-ttu-id="3ee75-621">函数</span><span class="sxs-lookup"><span data-stu-id="3ee75-621">function</span></span> | <span data-ttu-id="3ee75-622">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="3ee75-622">&lt;optional&gt;</span></span> | <span data-ttu-id="3ee75-623">方法完成后，使用单个参数 `asyncResult`（一个 [AsyncResult](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="3ee75-623">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="3ee75-624">要求</span><span class="sxs-lookup"><span data-stu-id="3ee75-624">Requirements</span></span>

|<span data-ttu-id="3ee75-625">要求</span><span class="sxs-lookup"><span data-stu-id="3ee75-625">Requirement</span></span>| <span data-ttu-id="3ee75-626">值</span><span class="sxs-lookup"><span data-stu-id="3ee75-626">Value</span></span>|
|---|---|
|[<span data-ttu-id="3ee75-627">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="3ee75-627">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3ee75-628">1.0</span><span class="sxs-lookup"><span data-stu-id="3ee75-628">1.0</span></span>|
|[<span data-ttu-id="3ee75-629">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="3ee75-629">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3ee75-630">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3ee75-630">ReadItem</span></span>|
|[<span data-ttu-id="3ee75-631">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="3ee75-631">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="3ee75-632">阅读</span><span class="sxs-lookup"><span data-stu-id="3ee75-632">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="3ee75-633">示例</span><span class="sxs-lookup"><span data-stu-id="3ee75-633">Examples</span></span>

<span data-ttu-id="3ee75-634">以下代码将一个字符串传递到 `displayReplyAllForm` 函数。</span><span class="sxs-lookup"><span data-stu-id="3ee75-634">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="3ee75-635">使用空白正文答复。</span><span class="sxs-lookup"><span data-stu-id="3ee75-635">Reply with an empty body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="3ee75-636">仅使用正文答复。</span><span class="sxs-lookup"><span data-stu-id="3ee75-636">Reply with just a body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="3ee75-637">使用正文和回调答复。</span><span class="sxs-lookup"><span data-stu-id="3ee75-637">Reply with a body and a callback.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi',
  'callback' : function(asyncResult)
  {
    console.log(asyncResult.value);
  }
});
```

#### <a name="displayreplyformformdata-callback"></a><span data-ttu-id="3ee75-638">displayReplyForm (formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="3ee75-638">displayReplyForm(formData, [callback])</span></span>

<span data-ttu-id="3ee75-639">显示答复窗体，其中仅包括所选邮件的发件人或所选约会的组织者。</span><span class="sxs-lookup"><span data-stu-id="3ee75-639">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="3ee75-640">在 Outlook for iOS 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="3ee75-640">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="3ee75-641">在 Outlook Web App 中，答复窗体显示为包含 3 列视图的弹出式窗体以及包含 2 列或 1 列视图的弹出式窗体。</span><span class="sxs-lookup"><span data-stu-id="3ee75-641">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="3ee75-642">如果任意字符串参数超出其限制，`displayReplyForm` 将引发异常。</span><span class="sxs-lookup"><span data-stu-id="3ee75-642">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

> [!NOTE]
> <span data-ttu-id="3ee75-643">要求集 1.1 不支持 `displayReplyForm` 在调用中包括附件的功能。</span><span class="sxs-lookup"><span data-stu-id="3ee75-643">The ability to include attachments in the call to `displayReplyForm` is not supported in requirement set 1.1.</span></span> <span data-ttu-id="3ee75-644">附件支持已添加到要求集 1.2 及以上的 `displayReplyForm` 中。</span><span class="sxs-lookup"><span data-stu-id="3ee75-644">Attachment support was added to `displayReplyForm` in requirement set 1.2 and above.</span></span>

##### <a name="parameters"></a><span data-ttu-id="3ee75-645">Parameters</span><span class="sxs-lookup"><span data-stu-id="3ee75-645">Parameters</span></span>

|<span data-ttu-id="3ee75-646">名称</span><span class="sxs-lookup"><span data-stu-id="3ee75-646">Name</span></span>| <span data-ttu-id="3ee75-647">说明</span><span class="sxs-lookup"><span data-stu-id="3ee75-647">Type</span></span>| <span data-ttu-id="3ee75-648">说明</span><span class="sxs-lookup"><span data-stu-id="3ee75-648">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="3ee75-649">字符串 &#124; 对象</span><span class="sxs-lookup"><span data-stu-id="3ee75-649">String &#124; Object</span></span>| | <span data-ttu-id="3ee75-p142">一个包含文本和 HTML 且表示答复窗体的正文的字符串。字符串限制为 32 KB。</span><span class="sxs-lookup"><span data-stu-id="3ee75-p142">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="3ee75-652">**或**</span><span class="sxs-lookup"><span data-stu-id="3ee75-652">**OR**</span></span><br/><span data-ttu-id="3ee75-p143">包含正文或附件数据和回调函数的对象。对象定义如下。</span><span class="sxs-lookup"><span data-stu-id="3ee75-p143">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="3ee75-655">字符串</span><span class="sxs-lookup"><span data-stu-id="3ee75-655">String</span></span> | <span data-ttu-id="3ee75-656">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="3ee75-656">&lt;optional&gt;</span></span> | <span data-ttu-id="3ee75-p144">一个包含文本和 HTML 且表示答复窗体的正文的字符串。字符串限制为 32 KB。</span><span class="sxs-lookup"><span data-stu-id="3ee75-p144">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `callback` | <span data-ttu-id="3ee75-659">函数</span><span class="sxs-lookup"><span data-stu-id="3ee75-659">function</span></span> | <span data-ttu-id="3ee75-660">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="3ee75-660">&lt;optional&gt;</span></span> | <span data-ttu-id="3ee75-661">方法完成后，使用单个参数 `asyncResult`（一个 [AsyncResult](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="3ee75-661">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="3ee75-662">要求</span><span class="sxs-lookup"><span data-stu-id="3ee75-662">Requirements</span></span>

|<span data-ttu-id="3ee75-663">要求</span><span class="sxs-lookup"><span data-stu-id="3ee75-663">Requirement</span></span>| <span data-ttu-id="3ee75-664">值</span><span class="sxs-lookup"><span data-stu-id="3ee75-664">Value</span></span>|
|---|---|
|[<span data-ttu-id="3ee75-665">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="3ee75-665">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3ee75-666">1.0</span><span class="sxs-lookup"><span data-stu-id="3ee75-666">1.0</span></span>|
|[<span data-ttu-id="3ee75-667">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="3ee75-667">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3ee75-668">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3ee75-668">ReadItem</span></span>|
|[<span data-ttu-id="3ee75-669">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="3ee75-669">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="3ee75-670">阅读</span><span class="sxs-lookup"><span data-stu-id="3ee75-670">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="3ee75-671">示例</span><span class="sxs-lookup"><span data-stu-id="3ee75-671">Examples</span></span>

<span data-ttu-id="3ee75-672">以下代码将一个字符串传递到 `displayReplyForm` 函数。</span><span class="sxs-lookup"><span data-stu-id="3ee75-672">The following code passes a string to the `displayReplyForm` function.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="3ee75-673">使用空白正文答复。</span><span class="sxs-lookup"><span data-stu-id="3ee75-673">Reply with an empty body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="3ee75-674">仅使用正文答复。</span><span class="sxs-lookup"><span data-stu-id="3ee75-674">Reply with just a body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="3ee75-675">使用正文和回调答复。</span><span class="sxs-lookup"><span data-stu-id="3ee75-675">Reply with a body and a callback.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi',
  'callback' : function(asyncResult)
  {
    console.log(asyncResult.value);
  }
});
```

#### <a name="getentities--entitiesjavascriptapioutlook11officeentities"></a><span data-ttu-id="3ee75-676">getEntities() → {[Entities](/javascript/api/outlook_1_1/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="3ee75-676">getEntities() → {[Entities](/javascript/api/outlook_1_1/office.entities)}</span></span>

<span data-ttu-id="3ee75-677">获取在所选项目的正文中找到的实体。</span><span class="sxs-lookup"><span data-stu-id="3ee75-677">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="3ee75-678">在 Outlook for iOS 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="3ee75-678">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="3ee75-679">Requirements</span><span class="sxs-lookup"><span data-stu-id="3ee75-679">Requirements</span></span>

|<span data-ttu-id="3ee75-680">要求</span><span class="sxs-lookup"><span data-stu-id="3ee75-680">Requirement</span></span>| <span data-ttu-id="3ee75-681">值</span><span class="sxs-lookup"><span data-stu-id="3ee75-681">Value</span></span>|
|---|---|
|[<span data-ttu-id="3ee75-682">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="3ee75-682">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3ee75-683">1.0</span><span class="sxs-lookup"><span data-stu-id="3ee75-683">1.0</span></span>|
|[<span data-ttu-id="3ee75-684">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="3ee75-684">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3ee75-685">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3ee75-685">ReadItem</span></span>|
|[<span data-ttu-id="3ee75-686">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="3ee75-686">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="3ee75-687">阅读</span><span class="sxs-lookup"><span data-stu-id="3ee75-687">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="3ee75-688">返回：</span><span class="sxs-lookup"><span data-stu-id="3ee75-688">Returns:</span></span>

<span data-ttu-id="3ee75-689">类型：[Entities](/javascript/api/outlook_1_1/office.entities)</span><span class="sxs-lookup"><span data-stu-id="3ee75-689">Type: [Entities](/javascript/api/outlook_1_1/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="3ee75-690">示例</span><span class="sxs-lookup"><span data-stu-id="3ee75-690">Example</span></span>

<span data-ttu-id="3ee75-691">以下示例访问当前项目的正文中的联系人实体。</span><span class="sxs-lookup"><span data-stu-id="3ee75-691">The following example accesses the contacts entities in the current item's body.</span></span>

```javascript
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlook11officecontactmeetingsuggestionjavascriptapioutlook11officemeetingsuggestionphonenumberjavascriptapioutlook11officephonenumbertasksuggestionjavascriptapioutlook11officetasksuggestion"></a><span data-ttu-id="3ee75-692">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_1/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_1/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_1/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_1/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="3ee75-692">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_1/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_1/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_1/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_1/office.tasksuggestion))>}</span></span>

<span data-ttu-id="3ee75-693">获取所选项目的正文中找到的指定实体类型的所有实体的数组。</span><span class="sxs-lookup"><span data-stu-id="3ee75-693">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="3ee75-694">在 Outlook for iOS 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="3ee75-694">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="3ee75-695">Parameters</span><span class="sxs-lookup"><span data-stu-id="3ee75-695">Parameters</span></span>

|<span data-ttu-id="3ee75-696">名称</span><span class="sxs-lookup"><span data-stu-id="3ee75-696">Name</span></span>| <span data-ttu-id="3ee75-697">类型</span><span class="sxs-lookup"><span data-stu-id="3ee75-697">Type</span></span>| <span data-ttu-id="3ee75-698">说明</span><span class="sxs-lookup"><span data-stu-id="3ee75-698">Description</span></span>|
|---|---|---|
|`entityType`| [<span data-ttu-id="3ee75-699">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="3ee75-699">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook_1_1/office.MailboxEnums.entitytype)|<span data-ttu-id="3ee75-700">EntityType 枚举值之一。</span><span class="sxs-lookup"><span data-stu-id="3ee75-700">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="3ee75-701">Requirements</span><span class="sxs-lookup"><span data-stu-id="3ee75-701">Requirements</span></span>

|<span data-ttu-id="3ee75-702">要求</span><span class="sxs-lookup"><span data-stu-id="3ee75-702">Requirement</span></span>| <span data-ttu-id="3ee75-703">值</span><span class="sxs-lookup"><span data-stu-id="3ee75-703">Value</span></span>|
|---|---|
|[<span data-ttu-id="3ee75-704">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="3ee75-704">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3ee75-705">1.0</span><span class="sxs-lookup"><span data-stu-id="3ee75-705">1.0</span></span>|
|[<span data-ttu-id="3ee75-706">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="3ee75-706">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3ee75-707">受限</span><span class="sxs-lookup"><span data-stu-id="3ee75-707">Restricted</span></span>|
|[<span data-ttu-id="3ee75-708">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="3ee75-708">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="3ee75-709">阅读</span><span class="sxs-lookup"><span data-stu-id="3ee75-709">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="3ee75-710">返回：</span><span class="sxs-lookup"><span data-stu-id="3ee75-710">Returns:</span></span>

<span data-ttu-id="3ee75-711">如果在 `entityType` 中传递的值不是 `EntityType` 枚举的有效成员，该方法返回 null。</span><span class="sxs-lookup"><span data-stu-id="3ee75-711">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="3ee75-712">如果指定类型的任何实体都不存在于该项目的正文中，该方法将返回空数组。</span><span class="sxs-lookup"><span data-stu-id="3ee75-712">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="3ee75-713">否则，返回的数组中对象的类型取决于 `entityType` 参数中请求实体的类型。</span><span class="sxs-lookup"><span data-stu-id="3ee75-713">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="3ee75-714">当使用此方法的最低权限级别**受限**时，某些实体类型需要 **ReadItem** 才能进行访问，如下表中所指定。</span><span class="sxs-lookup"><span data-stu-id="3ee75-714">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

| <span data-ttu-id="3ee75-715">`entityType` 的值</span><span class="sxs-lookup"><span data-stu-id="3ee75-715">Value of `entityType`</span></span> | <span data-ttu-id="3ee75-716">返回的数组中对象的类型</span><span class="sxs-lookup"><span data-stu-id="3ee75-716">Type of objects in returned array</span></span> | <span data-ttu-id="3ee75-717">所需权限级别</span><span class="sxs-lookup"><span data-stu-id="3ee75-717">Required Permission Level</span></span> |
| --- | --- | --- |
| `Address` | <span data-ttu-id="3ee75-718">字符串</span><span class="sxs-lookup"><span data-stu-id="3ee75-718">String</span></span> | <span data-ttu-id="3ee75-719">**受限**</span><span class="sxs-lookup"><span data-stu-id="3ee75-719">**Restricted**</span></span> |
| `Contact` | <span data-ttu-id="3ee75-720">Contact</span><span class="sxs-lookup"><span data-stu-id="3ee75-720">Contact</span></span> | <span data-ttu-id="3ee75-721">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="3ee75-721">**ReadItem**</span></span> |
| `EmailAddress` | <span data-ttu-id="3ee75-722">String</span><span class="sxs-lookup"><span data-stu-id="3ee75-722">String</span></span> | <span data-ttu-id="3ee75-723">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="3ee75-723">**ReadItem**</span></span> |
| `MeetingSuggestion` | <span data-ttu-id="3ee75-724">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="3ee75-724">MeetingSuggestion</span></span> | <span data-ttu-id="3ee75-725">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="3ee75-725">**ReadItem**</span></span> |
| `PhoneNumber` | <span data-ttu-id="3ee75-726">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="3ee75-726">PhoneNumber</span></span> | <span data-ttu-id="3ee75-727">**受限**</span><span class="sxs-lookup"><span data-stu-id="3ee75-727">**Restricted**</span></span> |
| `TaskSuggestion` | <span data-ttu-id="3ee75-728">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="3ee75-728">TaskSuggestion</span></span> | <span data-ttu-id="3ee75-729">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="3ee75-729">**ReadItem**</span></span> |
| `URL` | <span data-ttu-id="3ee75-730">String</span><span class="sxs-lookup"><span data-stu-id="3ee75-730">String</span></span> | <span data-ttu-id="3ee75-731">**受限**</span><span class="sxs-lookup"><span data-stu-id="3ee75-731">**Restricted**</span></span> |

<span data-ttu-id="3ee75-732">类型：Array.<(String|[Contact](/javascript/api/outlook_1_1/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_1/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_1/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_1/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="3ee75-732">Type:  Array.<(String|[Contact](/javascript/api/outlook_1_1/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_1/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_1/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_1/office.tasksuggestion))></span></span>


##### <a name="example"></a><span data-ttu-id="3ee75-733">示例</span><span class="sxs-lookup"><span data-stu-id="3ee75-733">Example</span></span>

<span data-ttu-id="3ee75-734">以下示例显示了如何访问表示当前项目的正文中的邮政地址的字符串数组。</span><span class="sxs-lookup"><span data-stu-id="3ee75-734">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlook11officecontactmeetingsuggestionjavascriptapioutlook11officemeetingsuggestionphonenumberjavascriptapioutlook11officephonenumbertasksuggestionjavascriptapioutlook11officetasksuggestion"></a><span data-ttu-id="3ee75-735">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_1/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_1/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_1/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_1/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="3ee75-735">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_1/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_1/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_1/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_1/office.tasksuggestion))>}</span></span>

<span data-ttu-id="3ee75-736">返回传递清单 XML 文件中定义的命名筛选器的所选项目中的已知实体。</span><span class="sxs-lookup"><span data-stu-id="3ee75-736">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="3ee75-737">在 Outlook for iOS 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="3ee75-737">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="3ee75-738">`getFilteredEntitiesByName` 方法返回匹配在具有指定 `FilterName` 元素值的清单 XML 文件中的 [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) 规则元素中定义的正则表达式的实体。</span><span class="sxs-lookup"><span data-stu-id="3ee75-738">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="3ee75-739">Parameters</span><span class="sxs-lookup"><span data-stu-id="3ee75-739">Parameters</span></span>

|<span data-ttu-id="3ee75-740">名称</span><span class="sxs-lookup"><span data-stu-id="3ee75-740">Name</span></span>| <span data-ttu-id="3ee75-741">类型</span><span class="sxs-lookup"><span data-stu-id="3ee75-741">Type</span></span>| <span data-ttu-id="3ee75-742">说明</span><span class="sxs-lookup"><span data-stu-id="3ee75-742">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="3ee75-743">字符串</span><span class="sxs-lookup"><span data-stu-id="3ee75-743">String</span></span>|<span data-ttu-id="3ee75-744">定义筛选器匹配的 `ItemHasKnownEntity` 规则元素的名称。</span><span class="sxs-lookup"><span data-stu-id="3ee75-744">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="3ee75-745">Requirements</span><span class="sxs-lookup"><span data-stu-id="3ee75-745">Requirements</span></span>

|<span data-ttu-id="3ee75-746">要求</span><span class="sxs-lookup"><span data-stu-id="3ee75-746">Requirement</span></span>| <span data-ttu-id="3ee75-747">值</span><span class="sxs-lookup"><span data-stu-id="3ee75-747">Value</span></span>|
|---|---|
|[<span data-ttu-id="3ee75-748">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="3ee75-748">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3ee75-749">1.0</span><span class="sxs-lookup"><span data-stu-id="3ee75-749">1.0</span></span>|
|[<span data-ttu-id="3ee75-750">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="3ee75-750">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3ee75-751">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3ee75-751">ReadItem</span></span>|
|[<span data-ttu-id="3ee75-752">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="3ee75-752">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="3ee75-753">阅读</span><span class="sxs-lookup"><span data-stu-id="3ee75-753">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="3ee75-754">返回：</span><span class="sxs-lookup"><span data-stu-id="3ee75-754">Returns:</span></span>

<span data-ttu-id="3ee75-p146">如果具有匹配 `name` 参数的 `FilterName` 元素值的清单中没有任何 `ItemHasKnownEntity` 元素，则该方法返回 `null`。如果 `name` 参数匹配清单中的 `ItemHasKnownEntity` 元素，但在匹配的当前项目中没有实体，则该方法返回一个空数组。</span><span class="sxs-lookup"><span data-stu-id="3ee75-p146">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>


<span data-ttu-id="3ee75-757">类型：Array.<(String|[Contact](/javascript/api/outlook_1_1/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_1/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_1/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_1/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="3ee75-757">Type: Array.<(String|[Contact](/javascript/api/outlook_1_1/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_1/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_1/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_1/office.tasksuggestion))></span></span>


#### <a name="getregexmatches--object"></a><span data-ttu-id="3ee75-758">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="3ee75-758">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="3ee75-759">返回所选项目中匹配在清单 XML 文件中定义的正则表达式的字符串值。</span><span class="sxs-lookup"><span data-stu-id="3ee75-759">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="3ee75-760">在 Outlook for iOS 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="3ee75-760">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="3ee75-p147">`getRegExMatches` 方法返回匹配在清单 XML 文件中的每个 `ItemHasRegularExpressionMatch` 或 `ItemHasKnownEntity` 规则元素中定义的正则表达式的字符串。对于 `ItemHasRegularExpressionMatch` 规则，匹配字符串必须发生在该规则指定的项目的属性中。`PropertyName` 简单类型定义支持的属性。</span><span class="sxs-lookup"><span data-stu-id="3ee75-p147">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="3ee75-764">例如，考虑一个外接程序清单具有以下 `Rule` 元素：</span><span class="sxs-lookup"><span data-stu-id="3ee75-764">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="3ee75-765">从 `getRegExMatches` 返回的对象应有两个属性：`fruits` 和 `veggies`。</span><span class="sxs-lookup"><span data-stu-id="3ee75-765">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

> [!NOTE]
> <span data-ttu-id="3ee75-p148">如果在项目的正文属性上指定 `ItemHasRegularExpressionMatch` 规则，则正则表达式应进一步筛选正文并应尝试返回该项目的整个正文。使用正则表达式（如 `.*`）获取项目的整个正文并不总是返回预期的结果。</span><span class="sxs-lookup"><span data-stu-id="3ee75-p148">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="requirements"></a><span data-ttu-id="3ee75-768">要求</span><span class="sxs-lookup"><span data-stu-id="3ee75-768">Requirements</span></span>

|<span data-ttu-id="3ee75-769">要求</span><span class="sxs-lookup"><span data-stu-id="3ee75-769">Requirement</span></span>| <span data-ttu-id="3ee75-770">值</span><span class="sxs-lookup"><span data-stu-id="3ee75-770">Value</span></span>|
|---|---|
|[<span data-ttu-id="3ee75-771">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="3ee75-771">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3ee75-772">1.0</span><span class="sxs-lookup"><span data-stu-id="3ee75-772">1.0</span></span>|
|[<span data-ttu-id="3ee75-773">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="3ee75-773">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3ee75-774">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3ee75-774">ReadItem</span></span>|
|[<span data-ttu-id="3ee75-775">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="3ee75-775">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="3ee75-776">阅读</span><span class="sxs-lookup"><span data-stu-id="3ee75-776">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="3ee75-777">返回：</span><span class="sxs-lookup"><span data-stu-id="3ee75-777">Returns:</span></span>

<span data-ttu-id="3ee75-p149">一个包含匹配在清单 XML 文件中定义的正则表达式的字符串数组的对象。每个数组的名称等于匹配 `ItemHasRegularExpressionMatch` 规则的 `RegExName` 属性或匹配 `ItemHasKnownEntity` 规则的 `FilterName` 属性的相应值。</span><span class="sxs-lookup"><span data-stu-id="3ee75-p149">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<dl class="param-type"><span data-ttu-id="3ee75-780">

<dt>
类型</dt>


</span><span class="sxs-lookup"><span data-stu-id="3ee75-780">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="3ee75-781">对象</span><span class="sxs-lookup"><span data-stu-id="3ee75-781">Object</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="3ee75-782">示例</span><span class="sxs-lookup"><span data-stu-id="3ee75-782">Example</span></span>

<span data-ttu-id="3ee75-783">以下示例显示了如何访问正则表达式 <rule> 元素 `fruits` 和 `veggies` 的匹配项的数组，这些元素在清单中指定。</rule></span><span class="sxs-lookup"><span data-stu-id="3ee75-783">The following example shows how to access the array of matches for the regular expression <rule>elements `fruits` and `veggies`, which are specified in the manifest.</rule></span></span>

```javascript
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veggies = allMatches.veggies;
```

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="3ee75-784">getRegExMatchesByName(name) → (nullable) {Array.<String>}</span><span class="sxs-lookup"><span data-stu-id="3ee75-784">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="3ee75-785">返回匹配在清单 XML 文件中定义的命名正则表达式的所选项目中的字符串值。</span><span class="sxs-lookup"><span data-stu-id="3ee75-785">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="3ee75-786">在 Outlook for iOS 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="3ee75-786">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="3ee75-787">`getRegExMatchesByName` 方法返回匹配在具有指定 `RegExName` 元素值的清单 XML 文件中的 `ItemHasRegularExpressionMatch` 规则元素中定义的正则表达式的字符串。</span><span class="sxs-lookup"><span data-stu-id="3ee75-787">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="3ee75-p150">如果在项目的正文属性上指定 `ItemHasRegularExpressionMatch` 规则，则正则表达式应进一步筛选正文，不应尝试返回该项目的整个正文。使用正则表达式（如 `.*`）获取项目的整个正文并不总是返回预期的结果。</span><span class="sxs-lookup"><span data-stu-id="3ee75-p150">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="3ee75-790">Parameters</span><span class="sxs-lookup"><span data-stu-id="3ee75-790">Parameters</span></span>

|<span data-ttu-id="3ee75-791">名称</span><span class="sxs-lookup"><span data-stu-id="3ee75-791">Name</span></span>| <span data-ttu-id="3ee75-792">说明</span><span class="sxs-lookup"><span data-stu-id="3ee75-792">Type</span></span>| <span data-ttu-id="3ee75-793">描述</span><span class="sxs-lookup"><span data-stu-id="3ee75-793">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="3ee75-794">字符串</span><span class="sxs-lookup"><span data-stu-id="3ee75-794">String</span></span>|<span data-ttu-id="3ee75-795">定义筛选器匹配的 `ItemHasRegularExpressionMatch` 规则元素的名称。</span><span class="sxs-lookup"><span data-stu-id="3ee75-795">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="3ee75-796">Requirements</span><span class="sxs-lookup"><span data-stu-id="3ee75-796">Requirements</span></span>

|<span data-ttu-id="3ee75-797">要求</span><span class="sxs-lookup"><span data-stu-id="3ee75-797">Requirement</span></span>| <span data-ttu-id="3ee75-798">值</span><span class="sxs-lookup"><span data-stu-id="3ee75-798">Value</span></span>|
|---|---|
|[<span data-ttu-id="3ee75-799">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="3ee75-799">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3ee75-800">1.0</span><span class="sxs-lookup"><span data-stu-id="3ee75-800">1.0</span></span>|
|[<span data-ttu-id="3ee75-801">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="3ee75-801">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3ee75-802">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3ee75-802">ReadItem</span></span>|
|[<span data-ttu-id="3ee75-803">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="3ee75-803">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="3ee75-804">阅读</span><span class="sxs-lookup"><span data-stu-id="3ee75-804">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="3ee75-805">返回：</span><span class="sxs-lookup"><span data-stu-id="3ee75-805">Returns:</span></span>

<span data-ttu-id="3ee75-806">一个包含匹配在清单 XML 文件中定义的正则表达式的字符串的数组。</span><span class="sxs-lookup"><span data-stu-id="3ee75-806">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<dl class="param-type"><span data-ttu-id="3ee75-807">

<dt>类型</dt>

</span><span class="sxs-lookup"><span data-stu-id="3ee75-807">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="3ee75-808">Array.< String ></span><span class="sxs-lookup"><span data-stu-id="3ee75-808">Array.< String ></span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="3ee75-809">示例</span><span class="sxs-lookup"><span data-stu-id="3ee75-809">Example</span></span>

```javascript
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

####  <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="3ee75-810">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="3ee75-810">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="3ee75-811">异步加载所选项目上此外接程序的自定义属性。</span><span class="sxs-lookup"><span data-stu-id="3ee75-811">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="3ee75-p151">自定义属性基于每个应用、每个项目存储为键/值对。此方法在回调中返回 `CustomProperties` 对象，该回调提供访问特定于当前项目和当前外接程序的自定义属性的方法。自定义属性未在项目上加密，因此这不应用作安全存储。</span><span class="sxs-lookup"><span data-stu-id="3ee75-p151">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="3ee75-815">Parameters</span><span class="sxs-lookup"><span data-stu-id="3ee75-815">Parameters</span></span>

|<span data-ttu-id="3ee75-816">名称</span><span class="sxs-lookup"><span data-stu-id="3ee75-816">Name</span></span>| <span data-ttu-id="3ee75-817">类型</span><span class="sxs-lookup"><span data-stu-id="3ee75-817">Type</span></span>| <span data-ttu-id="3ee75-818">属性</span><span class="sxs-lookup"><span data-stu-id="3ee75-818">Attributes</span></span>| <span data-ttu-id="3ee75-819">说明</span><span class="sxs-lookup"><span data-stu-id="3ee75-819">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="3ee75-820">函数</span><span class="sxs-lookup"><span data-stu-id="3ee75-820">function</span></span>||<span data-ttu-id="3ee75-821">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="3ee75-821">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="3ee75-822">自定义属性作为 `asyncResult.value` 属性中的 [`CustomProperties`](/javascript/api/outlook_1_1/office.customproperties) 对象提供。</span><span class="sxs-lookup"><span data-stu-id="3ee75-822">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook_1_1/office.customproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="3ee75-823">此对象可用于获取、设置以及从项目中删除自定义属性，并将自定义属性集的更改重新保存到服务器。</span><span class="sxs-lookup"><span data-stu-id="3ee75-823">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`| <span data-ttu-id="3ee75-824">对象</span><span class="sxs-lookup"><span data-stu-id="3ee75-824">Object</span></span>| <span data-ttu-id="3ee75-825">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="3ee75-825">&lt;optional&gt;</span></span>|<span data-ttu-id="3ee75-826">开发人员可以提供他们想要在回调函数中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="3ee75-826">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="3ee75-827">此对象可以通过回调函数中的 `asyncResult.asyncContext` 属性进行访问。</span><span class="sxs-lookup"><span data-stu-id="3ee75-827">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="3ee75-828">要求</span><span class="sxs-lookup"><span data-stu-id="3ee75-828">Requirements</span></span>

|<span data-ttu-id="3ee75-829">要求</span><span class="sxs-lookup"><span data-stu-id="3ee75-829">Requirement</span></span>| <span data-ttu-id="3ee75-830">值</span><span class="sxs-lookup"><span data-stu-id="3ee75-830">Value</span></span>|
|---|---|
|[<span data-ttu-id="3ee75-831">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="3ee75-831">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3ee75-832">1.0</span><span class="sxs-lookup"><span data-stu-id="3ee75-832">1.0</span></span>|
|[<span data-ttu-id="3ee75-833">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="3ee75-833">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3ee75-834">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3ee75-834">ReadItem</span></span>|
|[<span data-ttu-id="3ee75-835">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="3ee75-835">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="3ee75-836">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="3ee75-836">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="3ee75-837">示例</span><span class="sxs-lookup"><span data-stu-id="3ee75-837">Example</span></span>

<span data-ttu-id="3ee75-p154">以下代码示例显示了如何使用 `loadCustomPropertiesAsync` 方法异步加载特定于当前项目的自定义属性。该示例还显示了如何使用 `CustomProperties.saveAsync` 方法将这些属性重新保存到服务器。加载自定义属性后，该代码示例将使用 `CustomProperties.get` 方法读取自定义属性 `myProp`，使用 `CustomProperties.set` 方法写入自定义属性 `otherProp`，最后调用 `saveAsync` 方法保存这些自定义属性。</span><span class="sxs-lookup"><span data-stu-id="3ee75-p154">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

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

####  <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="3ee75-841">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="3ee75-841">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="3ee75-842">将附件从邮件或约会中删除。</span><span class="sxs-lookup"><span data-stu-id="3ee75-842">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="3ee75-p155">`removeAttachmentAsync` 方法删除项目中带指定标识符的附件。最佳做法是，仅当同一个邮件应用程序在同一会话中添加了一个附件时，你才应使用该附件标识符来删除该附件。在 Outlook Web App 和适用于设备的 OWA 中，附件标识符只在同一个会话中才有效。当用户关闭应用程序，或者如果用户开始在内嵌窗体中撰写，并在随后弹出的内嵌窗体中继续在单独的窗口撰写时，会话即结束。</span><span class="sxs-lookup"><span data-stu-id="3ee75-p155">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item. As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session. In Outlook Web App and OWA for Devices, the attachment identifier is valid only within the same session. A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="3ee75-847">Parameters</span><span class="sxs-lookup"><span data-stu-id="3ee75-847">Parameters</span></span>

|<span data-ttu-id="3ee75-848">名称</span><span class="sxs-lookup"><span data-stu-id="3ee75-848">Name</span></span>| <span data-ttu-id="3ee75-849">类型</span><span class="sxs-lookup"><span data-stu-id="3ee75-849">Type</span></span>| <span data-ttu-id="3ee75-850">属性</span><span class="sxs-lookup"><span data-stu-id="3ee75-850">Attributes</span></span>| <span data-ttu-id="3ee75-851">说明</span><span class="sxs-lookup"><span data-stu-id="3ee75-851">Description</span></span>|
|---|---|---|---|
|`attachmentId`| <span data-ttu-id="3ee75-852">字符串</span><span class="sxs-lookup"><span data-stu-id="3ee75-852">String</span></span>||<span data-ttu-id="3ee75-853">要删除的附件的标识符。</span><span class="sxs-lookup"><span data-stu-id="3ee75-853">The identifier of the attachment to remove.</span></span>|
|`options`| <span data-ttu-id="3ee75-854">对象</span><span class="sxs-lookup"><span data-stu-id="3ee75-854">Object</span></span>| <span data-ttu-id="3ee75-855">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="3ee75-855">&lt;optional&gt;</span></span>|<span data-ttu-id="3ee75-856">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="3ee75-856">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="3ee75-857">对象</span><span class="sxs-lookup"><span data-stu-id="3ee75-857">Object</span></span>| <span data-ttu-id="3ee75-858">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="3ee75-858">&lt;optional&gt;</span></span>|<span data-ttu-id="3ee75-859">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="3ee75-859">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="3ee75-860">函数</span><span class="sxs-lookup"><span data-stu-id="3ee75-860">function</span></span>| <span data-ttu-id="3ee75-861">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="3ee75-861">&lt;optional&gt;</span></span>|<span data-ttu-id="3ee75-862">方法完成后，使用单个参数 `callback`（一个 [`asyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `AsyncResult` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="3ee75-862">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="3ee75-863">如果删除附件失败，`asyncResult.error` 属性将包含一个说明失败原因的错误代码。</span><span class="sxs-lookup"><span data-stu-id="3ee75-863">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="3ee75-864">错误</span><span class="sxs-lookup"><span data-stu-id="3ee75-864">Errors</span></span>

| <span data-ttu-id="3ee75-865">错误代码</span><span class="sxs-lookup"><span data-stu-id="3ee75-865">Error code</span></span> | <span data-ttu-id="3ee75-866">说明</span><span class="sxs-lookup"><span data-stu-id="3ee75-866">Description</span></span> |
|------------|-------------|
| `InvalidAttachmentId` | <span data-ttu-id="3ee75-867">附件标识符不存在。</span><span class="sxs-lookup"><span data-stu-id="3ee75-867">The attachment identifier does not exist.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="3ee75-868">Requirements</span><span class="sxs-lookup"><span data-stu-id="3ee75-868">Requirements</span></span>

|<span data-ttu-id="3ee75-869">要求</span><span class="sxs-lookup"><span data-stu-id="3ee75-869">Requirement</span></span>| <span data-ttu-id="3ee75-870">值</span><span class="sxs-lookup"><span data-stu-id="3ee75-870">Value</span></span>|
|---|---|
|[<span data-ttu-id="3ee75-871">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="3ee75-871">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3ee75-872">1.1</span><span class="sxs-lookup"><span data-stu-id="3ee75-872">1.1</span></span>|
|[<span data-ttu-id="3ee75-873">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="3ee75-873">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3ee75-874">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="3ee75-874">ReadWriteItem</span></span>|
|[<span data-ttu-id="3ee75-875">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="3ee75-875">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="3ee75-876">撰写</span><span class="sxs-lookup"><span data-stu-id="3ee75-876">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="3ee75-877">示例</span><span class="sxs-lookup"><span data-stu-id="3ee75-877">Example</span></span>

<span data-ttu-id="3ee75-878">以下代码删除包含标识符 '0' 的附件。</span><span class="sxs-lookup"><span data-stu-id="3ee75-878">The following code removes an attachment with an identifier of '0'.</span></span>

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
