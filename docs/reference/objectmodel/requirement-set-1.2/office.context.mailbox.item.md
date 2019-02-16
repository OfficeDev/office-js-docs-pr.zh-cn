---
title: "\"context\"-\"邮箱\"。项目-要求集1。2"
description: ''
ms.date: 02/15/2019
localization_priority: Normal
ms.openlocfilehash: c0b956cac0410ef7d8e8e0d59a69e221e29c540a
ms.sourcegitcommit: f26778b596b6b022814c39601485ff676ed4e2fa
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 02/16/2019
ms.locfileid: "30068040"
---
# <a name="item"></a><span data-ttu-id="4a511-102">item</span><span class="sxs-lookup"><span data-stu-id="4a511-102">item</span></span>

### <span data-ttu-id="4a511-p101">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md). item</span><span class="sxs-lookup"><span data-stu-id="4a511-p101">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md). item</span></span>

<span data-ttu-id="4a511-p102">`item` 命名空间用于访问当前选定的邮件、会议请求或约会。可以通过使用 [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlook12officemailboxenumsitemtype) 属性确定 `item` 的类型。</span><span class="sxs-lookup"><span data-stu-id="4a511-p102">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlook12officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="4a511-107">Requirements</span><span class="sxs-lookup"><span data-stu-id="4a511-107">Requirements</span></span>

|<span data-ttu-id="4a511-108">要求</span><span class="sxs-lookup"><span data-stu-id="4a511-108">Requirement</span></span>| <span data-ttu-id="4a511-109">值</span><span class="sxs-lookup"><span data-stu-id="4a511-109">Value</span></span>|
|---|---|
|[<span data-ttu-id="4a511-110">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="4a511-110">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4a511-111">1.0</span><span class="sxs-lookup"><span data-stu-id="4a511-111">1.0</span></span>|
|[<span data-ttu-id="4a511-112">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="4a511-112">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4a511-113">受限</span><span class="sxs-lookup"><span data-stu-id="4a511-113">Restricted</span></span>|
|[<span data-ttu-id="4a511-114">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="4a511-114">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4a511-115">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="4a511-115">Compose or Read</span></span>|

### <a name="example"></a><span data-ttu-id="4a511-116">示例</span><span class="sxs-lookup"><span data-stu-id="4a511-116">Example</span></span>

<span data-ttu-id="4a511-117">以下 JavaScript 代码示例显示了如何访问 Outlook 中当前项目的 `subject` 属性。</span><span class="sxs-lookup"><span data-stu-id="4a511-117">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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

### <a name="members"></a><span data-ttu-id="4a511-118">成员</span><span class="sxs-lookup"><span data-stu-id="4a511-118">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlook12officeattachmentdetails"></a><span data-ttu-id="4a511-119">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_2/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="4a511-119">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_2/office.attachmentdetails)></span></span>

<span data-ttu-id="4a511-p103">获取项目的附件数组。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="4a511-p103">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="4a511-122">由于潜在的安全问题，某些类型的文件会受到 Outlook 阻止，并且不会返回。</span><span class="sxs-lookup"><span data-stu-id="4a511-122">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="4a511-123">如需了解更多信息，请参阅 [Outlook 中阻止的附件](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519)。</span><span class="sxs-lookup"><span data-stu-id="4a511-123">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="4a511-124">类型</span><span class="sxs-lookup"><span data-stu-id="4a511-124">Type</span></span>

*   <span data-ttu-id="4a511-125">Array.<[AttachmentDetails](/javascript/api/outlook_1_2/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="4a511-125">Array.<[AttachmentDetails](/javascript/api/outlook_1_2/office.attachmentdetails)></span></span>

##### <a name="requirements"></a><span data-ttu-id="4a511-126">Requirements</span><span class="sxs-lookup"><span data-stu-id="4a511-126">Requirements</span></span>

|<span data-ttu-id="4a511-127">要求</span><span class="sxs-lookup"><span data-stu-id="4a511-127">Requirement</span></span>| <span data-ttu-id="4a511-128">值</span><span class="sxs-lookup"><span data-stu-id="4a511-128">Value</span></span>|
|---|---|
|[<span data-ttu-id="4a511-129">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="4a511-129">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4a511-130">1.0</span><span class="sxs-lookup"><span data-stu-id="4a511-130">1.0</span></span>|
|[<span data-ttu-id="4a511-131">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="4a511-131">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4a511-132">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4a511-132">ReadItem</span></span>|
|[<span data-ttu-id="4a511-133">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="4a511-133">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4a511-134">阅读</span><span class="sxs-lookup"><span data-stu-id="4a511-134">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4a511-135">示例</span><span class="sxs-lookup"><span data-stu-id="4a511-135">Example</span></span>

<span data-ttu-id="4a511-136">以下代码使用当前项目上所有附件的详细信息构成 HTML 字符串。</span><span class="sxs-lookup"><span data-stu-id="4a511-136">The following code builds an HTML string with details of all attachments on the current item.</span></span>

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

####  <a name="bcc-recipientsjavascriptapioutlook12officerecipients"></a><span data-ttu-id="4a511-137">bcc :[Recipients](/javascript/api/outlook_1_2/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="4a511-137">bcc :[Recipients](/javascript/api/outlook_1_2/office.recipients)</span></span>

<span data-ttu-id="4a511-138">获取对象，该对象提供用于获取或更新邮件的密件抄送 (Bcc) 行上的收件人的方法。</span><span class="sxs-lookup"><span data-stu-id="4a511-138">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="4a511-139">仅限撰写模式。</span><span class="sxs-lookup"><span data-stu-id="4a511-139">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="4a511-140">类型</span><span class="sxs-lookup"><span data-stu-id="4a511-140">Type</span></span>

*   [<span data-ttu-id="4a511-141">收件人</span><span class="sxs-lookup"><span data-stu-id="4a511-141">Recipients</span></span>](/javascript/api/outlook_1_2/office.recipients)

##### <a name="requirements"></a><span data-ttu-id="4a511-142">Requirements</span><span class="sxs-lookup"><span data-stu-id="4a511-142">Requirements</span></span>

|<span data-ttu-id="4a511-143">要求</span><span class="sxs-lookup"><span data-stu-id="4a511-143">Requirement</span></span>| <span data-ttu-id="4a511-144">值</span><span class="sxs-lookup"><span data-stu-id="4a511-144">Value</span></span>|
|---|---|
|[<span data-ttu-id="4a511-145">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="4a511-145">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4a511-146">1.1</span><span class="sxs-lookup"><span data-stu-id="4a511-146">1.1</span></span>|
|[<span data-ttu-id="4a511-147">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="4a511-147">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4a511-148">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4a511-148">ReadItem</span></span>|
|[<span data-ttu-id="4a511-149">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="4a511-149">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4a511-150">撰写</span><span class="sxs-lookup"><span data-stu-id="4a511-150">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="4a511-151">示例</span><span class="sxs-lookup"><span data-stu-id="4a511-151">Example</span></span>

```javascript
Office.context.mailbox.item.bcc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.bcc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.bcc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfBccRecipients = asyncResult.value;
}
```

####  <a name="body-bodyjavascriptapioutlook12officebody"></a><span data-ttu-id="4a511-152">body :[Body](/javascript/api/outlook_1_2/office.body)</span><span class="sxs-lookup"><span data-stu-id="4a511-152">body :[Body](/javascript/api/outlook_1_2/office.body)</span></span>

<span data-ttu-id="4a511-153">获取一个提供用于处理项目正文的方法的对象。</span><span class="sxs-lookup"><span data-stu-id="4a511-153">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="4a511-154">类型</span><span class="sxs-lookup"><span data-stu-id="4a511-154">Type</span></span>

*   [<span data-ttu-id="4a511-155">Body</span><span class="sxs-lookup"><span data-stu-id="4a511-155">Body</span></span>](/javascript/api/outlook_1_2/office.body)

##### <a name="requirements"></a><span data-ttu-id="4a511-156">要求</span><span class="sxs-lookup"><span data-stu-id="4a511-156">Requirements</span></span>

|<span data-ttu-id="4a511-157">要求</span><span class="sxs-lookup"><span data-stu-id="4a511-157">Requirement</span></span>| <span data-ttu-id="4a511-158">值</span><span class="sxs-lookup"><span data-stu-id="4a511-158">Value</span></span>|
|---|---|
|[<span data-ttu-id="4a511-159">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="4a511-159">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4a511-160">1.1</span><span class="sxs-lookup"><span data-stu-id="4a511-160">1.1</span></span>|
|[<span data-ttu-id="4a511-161">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="4a511-161">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4a511-162">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4a511-162">ReadItem</span></span>|
|[<span data-ttu-id="4a511-163">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="4a511-163">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4a511-164">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="4a511-164">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4a511-165">示例</span><span class="sxs-lookup"><span data-stu-id="4a511-165">Example</span></span>

<span data-ttu-id="4a511-166">本示例获取纯文本格式的邮件正文。</span><span class="sxs-lookup"><span data-stu-id="4a511-166">This example gets the body of the message in plain text.</span></span>

```javascript
Office.context.mailbox.item.body.getAsync(
  "text",
  { asyncContext: "This is passed to the callback" },
  function callback(result) {
    // Do something with the result.
  });

```

<span data-ttu-id="4a511-167">下面是传递给回调函数的 result 参数的示例。</span><span class="sxs-lookup"><span data-stu-id="4a511-167">The following is an example of the result parameter passed to the callback function.</span></span>

```json
{
  "value": "TEXT of whole body (including threads below)",
  "status": "succeeded",
  "asyncContext": "This is passed to the callback"
}
```

####  <a name="cc-arrayemailaddressdetailsjavascriptapioutlook12officeemailaddressdetailsrecipientsjavascriptapioutlook12officerecipients"></a><span data-ttu-id="4a511-168">cc :Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_2/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="4a511-168">cc :Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_2/office.recipients)</span></span>

<span data-ttu-id="4a511-169">提供对邮件的抄送 (Cc) 收件人的访问权限。</span><span class="sxs-lookup"><span data-stu-id="4a511-169">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="4a511-170">对象的类型和访问级别取决于当前项目的模式。</span><span class="sxs-lookup"><span data-stu-id="4a511-170">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="4a511-171">阅读模式</span><span class="sxs-lookup"><span data-stu-id="4a511-171">Read mode</span></span>

<span data-ttu-id="4a511-p107">`cc` 属性返回包含邮件的**抄送**行上所列的每个收件人的 `EmailAddressDetails` 对象的数组。集合上限为 100 个成员。</span><span class="sxs-lookup"><span data-stu-id="4a511-p107">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.</span></span>

```javascript
console.log(JSON.stringify(Office.context.mailbox.item.cc));
```

##### <a name="compose-mode"></a><span data-ttu-id="4a511-174">撰写模式</span><span class="sxs-lookup"><span data-stu-id="4a511-174">Compose mode</span></span>

<span data-ttu-id="4a511-175">`cc` 属性返回一个 `Recipients` 对象，该对象提供用于获取或更新邮件的“**抄送**”行上收件人的方法。</span><span class="sxs-lookup"><span data-stu-id="4a511-175">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span>

```javascript
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="4a511-176">类型</span><span class="sxs-lookup"><span data-stu-id="4a511-176">Type</span></span>

*   <span data-ttu-id="4a511-177">Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_2/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="4a511-177">Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_2/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="4a511-178">要求</span><span class="sxs-lookup"><span data-stu-id="4a511-178">Requirements</span></span>

|<span data-ttu-id="4a511-179">要求</span><span class="sxs-lookup"><span data-stu-id="4a511-179">Requirement</span></span>| <span data-ttu-id="4a511-180">值</span><span class="sxs-lookup"><span data-stu-id="4a511-180">Value</span></span>|
|---|---|
|[<span data-ttu-id="4a511-181">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="4a511-181">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4a511-182">1.0</span><span class="sxs-lookup"><span data-stu-id="4a511-182">1.0</span></span>|
|[<span data-ttu-id="4a511-183">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="4a511-183">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4a511-184">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4a511-184">ReadItem</span></span>|
|[<span data-ttu-id="4a511-185">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="4a511-185">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4a511-186">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="4a511-186">Compose or Read</span></span>|

####  <a name="nullable-conversationid-string"></a><span data-ttu-id="4a511-187">(nullable) conversationId :String</span><span class="sxs-lookup"><span data-stu-id="4a511-187">(nullable) conversationId :String</span></span>

<span data-ttu-id="4a511-188">获取包含特定消息的电子邮件会话的标识符。</span><span class="sxs-lookup"><span data-stu-id="4a511-188">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="4a511-p108">如果在阅读窗体或撰写窗体的回复中激活邮件应用程序，则此属性可以获得一个整数值。如果用户随后更改了回复邮件的主题（若发送回复），则该邮件的对话 ID 将改变且之前获取的值将不适用。</span><span class="sxs-lookup"><span data-stu-id="4a511-p108">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="4a511-p109">对于撰写窗体的新项目，此属性获得一个 null 值。如果用户设置一个主题并保存该项目，`conversationId` 属性将返回一个值。</span><span class="sxs-lookup"><span data-stu-id="4a511-p109">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="4a511-193">Type</span><span class="sxs-lookup"><span data-stu-id="4a511-193">Type</span></span>

*   <span data-ttu-id="4a511-194">String</span><span class="sxs-lookup"><span data-stu-id="4a511-194">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="4a511-195">要求</span><span class="sxs-lookup"><span data-stu-id="4a511-195">Requirements</span></span>

|<span data-ttu-id="4a511-196">要求</span><span class="sxs-lookup"><span data-stu-id="4a511-196">Requirement</span></span>| <span data-ttu-id="4a511-197">值</span><span class="sxs-lookup"><span data-stu-id="4a511-197">Value</span></span>|
|---|---|
|[<span data-ttu-id="4a511-198">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="4a511-198">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4a511-199">1.0</span><span class="sxs-lookup"><span data-stu-id="4a511-199">1.0</span></span>|
|[<span data-ttu-id="4a511-200">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="4a511-200">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4a511-201">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4a511-201">ReadItem</span></span>|
|[<span data-ttu-id="4a511-202">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="4a511-202">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4a511-203">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="4a511-203">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4a511-204">示例</span><span class="sxs-lookup"><span data-stu-id="4a511-204">Example</span></span>

```javascript
var conversationId = Office.context.mailbox.item.conversationId;
console.log("conversationId: " + conversationId);
```

#### <a name="datetimecreated-date"></a><span data-ttu-id="4a511-205">dateTimeCreated :Date</span><span class="sxs-lookup"><span data-stu-id="4a511-205">dateTimeCreated :Date</span></span>

<span data-ttu-id="4a511-p110">获取项目创建的日期和时间。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="4a511-p110">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="4a511-208">类型</span><span class="sxs-lookup"><span data-stu-id="4a511-208">Type</span></span>

*   <span data-ttu-id="4a511-209">日期</span><span class="sxs-lookup"><span data-stu-id="4a511-209">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="4a511-210">Requirements</span><span class="sxs-lookup"><span data-stu-id="4a511-210">Requirements</span></span>

|<span data-ttu-id="4a511-211">要求</span><span class="sxs-lookup"><span data-stu-id="4a511-211">Requirement</span></span>| <span data-ttu-id="4a511-212">值</span><span class="sxs-lookup"><span data-stu-id="4a511-212">Value</span></span>|
|---|---|
|[<span data-ttu-id="4a511-213">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="4a511-213">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4a511-214">1.0</span><span class="sxs-lookup"><span data-stu-id="4a511-214">1.0</span></span>|
|[<span data-ttu-id="4a511-215">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="4a511-215">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4a511-216">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4a511-216">ReadItem</span></span>|
|[<span data-ttu-id="4a511-217">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="4a511-217">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4a511-218">阅读</span><span class="sxs-lookup"><span data-stu-id="4a511-218">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4a511-219">示例</span><span class="sxs-lookup"><span data-stu-id="4a511-219">Example</span></span>

```javascript
var dateTimeCreated = Office.context.mailbox.item.dateTimeCreated;
console.log("Date and time created: " + dateTimeCreated);
```

#### <a name="datetimemodified-date"></a><span data-ttu-id="4a511-220">dateTimeModified :Date</span><span class="sxs-lookup"><span data-stu-id="4a511-220">dateTimeModified :Date</span></span>

<span data-ttu-id="4a511-p111">获取项目最近一次修改的日期和时间。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="4a511-p111">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="4a511-223">在 Outlook for iOS 或 Outlook for Android 中不支持此成员。</span><span class="sxs-lookup"><span data-stu-id="4a511-223">This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="type"></a><span data-ttu-id="4a511-224">类型</span><span class="sxs-lookup"><span data-stu-id="4a511-224">Type</span></span>

*   <span data-ttu-id="4a511-225">日期</span><span class="sxs-lookup"><span data-stu-id="4a511-225">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="4a511-226">Requirements</span><span class="sxs-lookup"><span data-stu-id="4a511-226">Requirements</span></span>

|<span data-ttu-id="4a511-227">要求</span><span class="sxs-lookup"><span data-stu-id="4a511-227">Requirement</span></span>| <span data-ttu-id="4a511-228">值</span><span class="sxs-lookup"><span data-stu-id="4a511-228">Value</span></span>|
|---|---|
|[<span data-ttu-id="4a511-229">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="4a511-229">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4a511-230">1.0</span><span class="sxs-lookup"><span data-stu-id="4a511-230">1.0</span></span>|
|[<span data-ttu-id="4a511-231">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="4a511-231">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4a511-232">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4a511-232">ReadItem</span></span>|
|[<span data-ttu-id="4a511-233">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="4a511-233">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4a511-234">阅读</span><span class="sxs-lookup"><span data-stu-id="4a511-234">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4a511-235">示例</span><span class="sxs-lookup"><span data-stu-id="4a511-235">Example</span></span>

```javascript
var dateTimeModified = Office.context.mailbox.item.dateTimeModified;
console.log("Date and time modified: " + dateTimeModified);
```

####  <a name="end-datetimejavascriptapioutlook12officetime"></a><span data-ttu-id="4a511-236">end :Date|[Time](/javascript/api/outlook_1_2/office.time)</span><span class="sxs-lookup"><span data-stu-id="4a511-236">end :Date|[Time](/javascript/api/outlook_1_2/office.time)</span></span>

<span data-ttu-id="4a511-237">获取或设置约会结束的日期和时间。</span><span class="sxs-lookup"><span data-stu-id="4a511-237">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="4a511-p112">将 `end` 属性表示为协调世界时 (UTC) 的日期和时间值。可使用 [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook12officelocalclienttime) 方法将 end 属性值转换为客户端的本地日期和时间。</span><span class="sxs-lookup"><span data-stu-id="4a511-p112">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook12officelocalclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="4a511-240">阅读模式</span><span class="sxs-lookup"><span data-stu-id="4a511-240">Read mode</span></span>

<span data-ttu-id="4a511-241">`end` 属性返回 `Date` 对象。</span><span class="sxs-lookup"><span data-stu-id="4a511-241">The `end` property returns a `Date` object.</span></span>

```javascript
var end = Office.context.mailbox.item.end;
console.log("Appointment end: " + end);
```

##### <a name="compose-mode"></a><span data-ttu-id="4a511-242">撰写模式</span><span class="sxs-lookup"><span data-stu-id="4a511-242">Compose mode</span></span>

<span data-ttu-id="4a511-243">`end` 属性返回 `Time` 对象。</span><span class="sxs-lookup"><span data-stu-id="4a511-243">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="4a511-244">使用 [`Time.setAsync`](/javascript/api/outlook_1_2/office.time#setasync-datetime--options--callback-) 方法设置结束时间时，应使用 [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) 方法将客户端的本地时间转换为服务器的 UTC。</span><span class="sxs-lookup"><span data-stu-id="4a511-244">When you use the [`Time.setAsync`](/javascript/api/outlook_1_2/office.time#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="4a511-245">下面的示例使用[`setAsync`](/javascript/api/outlook_1_2/office.time#setasync-datetime--options--callback-) `Time`对象的方法设置约会的结束时间。</span><span class="sxs-lookup"><span data-stu-id="4a511-245">The following example sets the end time of an appointment by using the [`setAsync`](/javascript/api/outlook_1_2/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="4a511-246">类型</span><span class="sxs-lookup"><span data-stu-id="4a511-246">Type</span></span>

*   <span data-ttu-id="4a511-247">Date | [Time](/javascript/api/outlook_1_2/office.time)</span><span class="sxs-lookup"><span data-stu-id="4a511-247">Date | [Time](/javascript/api/outlook_1_2/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="4a511-248">Requirements</span><span class="sxs-lookup"><span data-stu-id="4a511-248">Requirements</span></span>

|<span data-ttu-id="4a511-249">要求</span><span class="sxs-lookup"><span data-stu-id="4a511-249">Requirement</span></span>| <span data-ttu-id="4a511-250">值</span><span class="sxs-lookup"><span data-stu-id="4a511-250">Value</span></span>|
|---|---|
|[<span data-ttu-id="4a511-251">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="4a511-251">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4a511-252">1.0</span><span class="sxs-lookup"><span data-stu-id="4a511-252">1.0</span></span>|
|[<span data-ttu-id="4a511-253">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="4a511-253">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4a511-254">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4a511-254">ReadItem</span></span>|
|[<span data-ttu-id="4a511-255">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="4a511-255">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4a511-256">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="4a511-256">Compose or Read</span></span>|

#### <a name="from-emailaddressdetailsjavascriptapioutlook12officeemailaddressdetails"></a><span data-ttu-id="4a511-257">from :[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="4a511-257">from :[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)</span></span>

<span data-ttu-id="4a511-p113">获取邮件发件人的电子邮件地址。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="4a511-p113">Gets the email address of the sender of a message. Read mode only.</span></span>

<span data-ttu-id="4a511-p114">`from` 和 [`sender`](#sender-emailaddressdetailsjavascriptapioutlook12officeemailaddressdetails) 属性表示同一个人，邮件由代理人发送的除外。在此情况下，`from` 属性表示代理程序，而 sender 属性表示代理人。</span><span class="sxs-lookup"><span data-stu-id="4a511-p114">The `from` and [`sender`](#sender-emailaddressdetailsjavascriptapioutlook12officeemailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="4a511-262">`from` 属性中 `EmailAddressDetails` 对象的 `recipientType` 属性为 `undefined`。</span><span class="sxs-lookup"><span data-stu-id="4a511-262">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="4a511-263">类型</span><span class="sxs-lookup"><span data-stu-id="4a511-263">Type</span></span>

*   [<span data-ttu-id="4a511-264">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="4a511-264">EmailAddressDetails</span></span>](/javascript/api/outlook_1_2/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="4a511-265">要求</span><span class="sxs-lookup"><span data-stu-id="4a511-265">Requirements</span></span>

|<span data-ttu-id="4a511-266">要求</span><span class="sxs-lookup"><span data-stu-id="4a511-266">Requirement</span></span>| <span data-ttu-id="4a511-267">值</span><span class="sxs-lookup"><span data-stu-id="4a511-267">Value</span></span>|
|---|---|
|[<span data-ttu-id="4a511-268">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="4a511-268">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4a511-269">1.0</span><span class="sxs-lookup"><span data-stu-id="4a511-269">1.0</span></span>|
|[<span data-ttu-id="4a511-270">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="4a511-270">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4a511-271">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4a511-271">ReadItem</span></span>|
|[<span data-ttu-id="4a511-272">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="4a511-272">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4a511-273">阅读</span><span class="sxs-lookup"><span data-stu-id="4a511-273">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4a511-274">示例</span><span class="sxs-lookup"><span data-stu-id="4a511-274">Example</span></span>

```javascript
var from = Office.context.mailbox.item.from;
console.log("From " + from);
```

#### <a name="internetmessageid-string"></a><span data-ttu-id="4a511-275">internetMessageId :String</span><span class="sxs-lookup"><span data-stu-id="4a511-275">internetMessageId :String</span></span>

<span data-ttu-id="4a511-p115">获取电子邮件的 Internet 消息标识符。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="4a511-p115">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="4a511-278">Type</span><span class="sxs-lookup"><span data-stu-id="4a511-278">Type</span></span>

*   <span data-ttu-id="4a511-279">String</span><span class="sxs-lookup"><span data-stu-id="4a511-279">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="4a511-280">要求</span><span class="sxs-lookup"><span data-stu-id="4a511-280">Requirements</span></span>

|<span data-ttu-id="4a511-281">要求</span><span class="sxs-lookup"><span data-stu-id="4a511-281">Requirement</span></span>| <span data-ttu-id="4a511-282">值</span><span class="sxs-lookup"><span data-stu-id="4a511-282">Value</span></span>|
|---|---|
|[<span data-ttu-id="4a511-283">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="4a511-283">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4a511-284">1.0</span><span class="sxs-lookup"><span data-stu-id="4a511-284">1.0</span></span>|
|[<span data-ttu-id="4a511-285">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="4a511-285">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4a511-286">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4a511-286">ReadItem</span></span>|
|[<span data-ttu-id="4a511-287">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="4a511-287">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4a511-288">阅读</span><span class="sxs-lookup"><span data-stu-id="4a511-288">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4a511-289">示例</span><span class="sxs-lookup"><span data-stu-id="4a511-289">Example</span></span>

```javascript
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

#### <a name="itemclass-string"></a><span data-ttu-id="4a511-290">itemClass :String</span><span class="sxs-lookup"><span data-stu-id="4a511-290">itemClass :String</span></span>

<span data-ttu-id="4a511-p116">获取选定项目的 Exchange Web 服务项目类。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="4a511-p116">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="4a511-p117">`itemClass` 属性指定所选项目的邮件类别。以下是邮件或约会项目的默认邮件类别。</span><span class="sxs-lookup"><span data-stu-id="4a511-p117">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

| <span data-ttu-id="4a511-295">类型</span><span class="sxs-lookup"><span data-stu-id="4a511-295">Type</span></span> | <span data-ttu-id="4a511-296">说明</span><span class="sxs-lookup"><span data-stu-id="4a511-296">Description</span></span> | <span data-ttu-id="4a511-297">项目类</span><span class="sxs-lookup"><span data-stu-id="4a511-297">item class</span></span> |
| --- | --- | --- |
| <span data-ttu-id="4a511-298">约会项目</span><span class="sxs-lookup"><span data-stu-id="4a511-298">Appointment items</span></span> | <span data-ttu-id="4a511-299">这些是项目类 `IPM.Appointment` 或 `IPM.Appointment.Occurrence` 的日历项目。</span><span class="sxs-lookup"><span data-stu-id="4a511-299">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurrence`.</span></span> | `IPM.Appointment`<br />`IPM.Appointment.Occurrence` |
| <span data-ttu-id="4a511-300">邮件项目</span><span class="sxs-lookup"><span data-stu-id="4a511-300">Message items</span></span> | <span data-ttu-id="4a511-301">这些项目包括具有默认邮件类别 `IPM.Note` 的电子邮件，以及使用 `IPM.Schedule.Meeting` 作为基础邮件类别的会议请求、响应和取消。</span><span class="sxs-lookup"><span data-stu-id="4a511-301">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span> | `IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled` |

<span data-ttu-id="4a511-302">你可以创建用于扩展默认邮件类别的自定义邮件类别，例如，自定义约会邮件类别 `IPM.Appointment.Contoso`。</span><span class="sxs-lookup"><span data-stu-id="4a511-302">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="4a511-303">Type</span><span class="sxs-lookup"><span data-stu-id="4a511-303">Type</span></span>

*   <span data-ttu-id="4a511-304">String</span><span class="sxs-lookup"><span data-stu-id="4a511-304">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="4a511-305">要求</span><span class="sxs-lookup"><span data-stu-id="4a511-305">Requirements</span></span>

|<span data-ttu-id="4a511-306">要求</span><span class="sxs-lookup"><span data-stu-id="4a511-306">Requirement</span></span>| <span data-ttu-id="4a511-307">值</span><span class="sxs-lookup"><span data-stu-id="4a511-307">Value</span></span>|
|---|---|
|[<span data-ttu-id="4a511-308">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="4a511-308">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4a511-309">1.0</span><span class="sxs-lookup"><span data-stu-id="4a511-309">1.0</span></span>|
|[<span data-ttu-id="4a511-310">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="4a511-310">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4a511-311">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4a511-311">ReadItem</span></span>|
|[<span data-ttu-id="4a511-312">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="4a511-312">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4a511-313">阅读</span><span class="sxs-lookup"><span data-stu-id="4a511-313">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4a511-314">示例</span><span class="sxs-lookup"><span data-stu-id="4a511-314">Example</span></span>

```javascript
var itemClass = Office.context.mailbox.item.itemClass;
console.log("Item class: " + itemClass);
```

#### <a name="nullable-itemid-string"></a><span data-ttu-id="4a511-315">(nullable) itemId :String</span><span class="sxs-lookup"><span data-stu-id="4a511-315">(nullable) itemId :String</span></span>

<span data-ttu-id="4a511-p118">获取当前项目的 Exchange Web 服务项目标识符。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="4a511-p118">Gets the Exchange Web Services item identifier for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="4a511-318">`itemId` 属性返回的标识符与 Exchange Web 服务项目标识符相同。</span><span class="sxs-lookup"><span data-stu-id="4a511-318">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="4a511-319">`itemId` 属性与 Outlook 条目 ID 或 Outlook REST API 使用的 ID 不同。</span><span class="sxs-lookup"><span data-stu-id="4a511-319">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="4a511-320">使用此值进行 REST API 调用前，应使用 `Office.context.mailbox.convertToRestId`（可在要求集 1.3 的开头部分中找到）对它进行转换。</span><span class="sxs-lookup"><span data-stu-id="4a511-320">Before making REST API calls using this value, it should be converted using `Office.context.mailbox.convertToRestId`, which is available starting in requirement set 1.3.</span></span> <span data-ttu-id="4a511-321">有关详细信息，请参阅[从 Outlook 加载项使用 Outlook REST API](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id)。</span><span class="sxs-lookup"><span data-stu-id="4a511-321">For more details, see [Use the Outlook REST APIs from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

##### <a name="type"></a><span data-ttu-id="4a511-322">Type</span><span class="sxs-lookup"><span data-stu-id="4a511-322">Type</span></span>

*   <span data-ttu-id="4a511-323">String</span><span class="sxs-lookup"><span data-stu-id="4a511-323">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="4a511-324">要求</span><span class="sxs-lookup"><span data-stu-id="4a511-324">Requirements</span></span>

|<span data-ttu-id="4a511-325">要求</span><span class="sxs-lookup"><span data-stu-id="4a511-325">Requirement</span></span>| <span data-ttu-id="4a511-326">值</span><span class="sxs-lookup"><span data-stu-id="4a511-326">Value</span></span>|
|---|---|
|[<span data-ttu-id="4a511-327">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="4a511-327">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4a511-328">1.0</span><span class="sxs-lookup"><span data-stu-id="4a511-328">1.0</span></span>|
|[<span data-ttu-id="4a511-329">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="4a511-329">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4a511-330">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4a511-330">ReadItem</span></span>|
|[<span data-ttu-id="4a511-331">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="4a511-331">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4a511-332">阅读</span><span class="sxs-lookup"><span data-stu-id="4a511-332">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4a511-333">示例</span><span class="sxs-lookup"><span data-stu-id="4a511-333">Example</span></span>

<span data-ttu-id="4a511-p120">以下代码检查项目标识符是否存在。如果 `itemId` 属性返回 `null` 或 `undefined`，则将项目保存到存储，并从异步结果中获取项目标识符。</span><span class="sxs-lookup"><span data-stu-id="4a511-p120">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

```javascript
var itemId = Office.context.mailbox.item.itemId;
if (itemId === null || itemId == undefined) {
  Office.context.mailbox.item.saveAsync(function(result) {
    itemId = result.value;
  });
}
```

####  <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlook12officemailboxenumsitemtype"></a><span data-ttu-id="4a511-336">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_2/office.mailboxenums.itemtype)</span><span class="sxs-lookup"><span data-stu-id="4a511-336">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_2/office.mailboxenums.itemtype)</span></span>

<span data-ttu-id="4a511-337">获取实例表示的项的类型。</span><span class="sxs-lookup"><span data-stu-id="4a511-337">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="4a511-338">`itemType` 属性返回其中一个 `ItemType` 枚举值，指示 `item` 对象实例是邮件还是约会。</span><span class="sxs-lookup"><span data-stu-id="4a511-338">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="4a511-339">类型</span><span class="sxs-lookup"><span data-stu-id="4a511-339">Type</span></span>

*   [<span data-ttu-id="4a511-340">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="4a511-340">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook_1_2/office.mailboxenums.itemtype)

##### <a name="requirements"></a><span data-ttu-id="4a511-341">要求</span><span class="sxs-lookup"><span data-stu-id="4a511-341">Requirements</span></span>

|<span data-ttu-id="4a511-342">要求</span><span class="sxs-lookup"><span data-stu-id="4a511-342">Requirement</span></span>| <span data-ttu-id="4a511-343">值</span><span class="sxs-lookup"><span data-stu-id="4a511-343">Value</span></span>|
|---|---|
|[<span data-ttu-id="4a511-344">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="4a511-344">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4a511-345">1.0</span><span class="sxs-lookup"><span data-stu-id="4a511-345">1.0</span></span>|
|[<span data-ttu-id="4a511-346">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="4a511-346">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4a511-347">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4a511-347">ReadItem</span></span>|
|[<span data-ttu-id="4a511-348">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="4a511-348">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4a511-349">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="4a511-349">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4a511-350">示例</span><span class="sxs-lookup"><span data-stu-id="4a511-350">Example</span></span>

```javascript
if (Office.context.mailbox.item.itemType === Office.MailboxEnums.ItemType.Message) {
  // Do something.
} else {
  // Do something else.
}
```

####  <a name="location-stringlocationjavascriptapioutlook12officelocation"></a><span data-ttu-id="4a511-351">location :String|[Location](/javascript/api/outlook_1_2/office.location)</span><span class="sxs-lookup"><span data-stu-id="4a511-351">location :String|[Location](/javascript/api/outlook_1_2/office.location)</span></span>

<span data-ttu-id="4a511-352">获取或设置约会的位置。</span><span class="sxs-lookup"><span data-stu-id="4a511-352">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="4a511-353">阅读模式</span><span class="sxs-lookup"><span data-stu-id="4a511-353">Read mode</span></span>

<span data-ttu-id="4a511-354">`location` 属性返回一个包含约会位置的字符串。</span><span class="sxs-lookup"><span data-stu-id="4a511-354">The `location` property returns a string that contains the location of the appointment.</span></span>

```javascript
var location = Office.context.mailbox.item.location;
console.log("location: " + location);
```

##### <a name="compose-mode"></a><span data-ttu-id="4a511-355">撰写模式</span><span class="sxs-lookup"><span data-stu-id="4a511-355">Compose mode</span></span>

<span data-ttu-id="4a511-356">`location` 属性返回一个 `Location` 对象，该对象提供用于获取和设置约会位置的方法。</span><span class="sxs-lookup"><span data-stu-id="4a511-356">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

```javascript
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="4a511-357">类型</span><span class="sxs-lookup"><span data-stu-id="4a511-357">Type</span></span>

*   <span data-ttu-id="4a511-358">String | [Location](/javascript/api/outlook_1_2/office.location)</span><span class="sxs-lookup"><span data-stu-id="4a511-358">String | [Location](/javascript/api/outlook_1_2/office.location)</span></span>

##### <a name="requirements"></a><span data-ttu-id="4a511-359">要求</span><span class="sxs-lookup"><span data-stu-id="4a511-359">Requirements</span></span>

|<span data-ttu-id="4a511-360">要求</span><span class="sxs-lookup"><span data-stu-id="4a511-360">Requirement</span></span>| <span data-ttu-id="4a511-361">值</span><span class="sxs-lookup"><span data-stu-id="4a511-361">Value</span></span>|
|---|---|
|[<span data-ttu-id="4a511-362">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="4a511-362">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4a511-363">1.0</span><span class="sxs-lookup"><span data-stu-id="4a511-363">1.0</span></span>|
|[<span data-ttu-id="4a511-364">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="4a511-364">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4a511-365">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4a511-365">ReadItem</span></span>|
|[<span data-ttu-id="4a511-366">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="4a511-366">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4a511-367">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="4a511-367">Compose or Read</span></span>|

#### <a name="normalizedsubject-string"></a><span data-ttu-id="4a511-368">normalizedSubject :String</span><span class="sxs-lookup"><span data-stu-id="4a511-368">normalizedSubject :String</span></span>

<span data-ttu-id="4a511-p121">获取删除了所有前缀（包括 `RE:` 和 `FWD:`）的项目主题。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="4a511-p121">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="4a511-p122">normalizedSubject 属性获取包含由电子邮件程序添加的任何标准前缀（如 `RE:` 和 `FW:`）的项目主题。若要获取包含完整前缀的项目主题，请使用 [`subject`](#subject-stringsubjectjavascriptapioutlook12officesubject) 属性。</span><span class="sxs-lookup"><span data-stu-id="4a511-p122">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubjectjavascriptapioutlook12officesubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="4a511-373">Type</span><span class="sxs-lookup"><span data-stu-id="4a511-373">Type</span></span>

*   <span data-ttu-id="4a511-374">String</span><span class="sxs-lookup"><span data-stu-id="4a511-374">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="4a511-375">要求</span><span class="sxs-lookup"><span data-stu-id="4a511-375">Requirements</span></span>

|<span data-ttu-id="4a511-376">要求</span><span class="sxs-lookup"><span data-stu-id="4a511-376">Requirement</span></span>| <span data-ttu-id="4a511-377">值</span><span class="sxs-lookup"><span data-stu-id="4a511-377">Value</span></span>|
|---|---|
|[<span data-ttu-id="4a511-378">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="4a511-378">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4a511-379">1.0</span><span class="sxs-lookup"><span data-stu-id="4a511-379">1.0</span></span>|
|[<span data-ttu-id="4a511-380">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="4a511-380">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4a511-381">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4a511-381">ReadItem</span></span>|
|[<span data-ttu-id="4a511-382">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="4a511-382">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4a511-383">阅读</span><span class="sxs-lookup"><span data-stu-id="4a511-383">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4a511-384">示例</span><span class="sxs-lookup"><span data-stu-id="4a511-384">Example</span></span>

```javascript
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
console.log("Normalized subject: " + normalizedSubject);
```

####  <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlook12officeemailaddressdetailsrecipientsjavascriptapioutlook12officerecipients"></a><span data-ttu-id="4a511-385">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_2/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="4a511-385">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_2/office.recipients)</span></span>

<span data-ttu-id="4a511-386">提供对事件的可选与会者的访问权限。</span><span class="sxs-lookup"><span data-stu-id="4a511-386">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="4a511-387">对象的类型和访问级别取决于当前项目的模式。</span><span class="sxs-lookup"><span data-stu-id="4a511-387">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="4a511-388">阅读模式</span><span class="sxs-lookup"><span data-stu-id="4a511-388">Read mode</span></span>

<span data-ttu-id="4a511-389">`optionalAttendees` 属性返回一个数组，其中包含每个可选与会者的 `EmailAddressDetails` 对象。</span><span class="sxs-lookup"><span data-stu-id="4a511-389">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span>

```javascript
var optionalAttendees = Office.context.mailbox.item.optionalAttendees;
console.log("Optional attendees: " + JSON.stringify(optionalAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="4a511-390">撰写模式</span><span class="sxs-lookup"><span data-stu-id="4a511-390">Compose mode</span></span>

<span data-ttu-id="4a511-391">`optionalAttendees` 属性返回一个 `Recipients` 对象，该对象提供用于获取或更新可选与会者的方法。</span><span class="sxs-lookup"><span data-stu-id="4a511-391">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span>

```javascript
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="4a511-392">类型</span><span class="sxs-lookup"><span data-stu-id="4a511-392">Type</span></span>

*   <span data-ttu-id="4a511-393">Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_2/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="4a511-393">Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_2/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="4a511-394">Requirements</span><span class="sxs-lookup"><span data-stu-id="4a511-394">Requirements</span></span>

|<span data-ttu-id="4a511-395">要求</span><span class="sxs-lookup"><span data-stu-id="4a511-395">Requirement</span></span>| <span data-ttu-id="4a511-396">值</span><span class="sxs-lookup"><span data-stu-id="4a511-396">Value</span></span>|
|---|---|
|[<span data-ttu-id="4a511-397">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="4a511-397">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4a511-398">1.0</span><span class="sxs-lookup"><span data-stu-id="4a511-398">1.0</span></span>|
|[<span data-ttu-id="4a511-399">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="4a511-399">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4a511-400">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4a511-400">ReadItem</span></span>|
|[<span data-ttu-id="4a511-401">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="4a511-401">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4a511-402">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="4a511-402">Compose or Read</span></span>|

#### <a name="organizer-emailaddressdetailsjavascriptapioutlook12officeemailaddressdetails"></a><span data-ttu-id="4a511-403">organizer :[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="4a511-403">organizer :[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)</span></span>

<span data-ttu-id="4a511-p124">获取指定会议的会议组织者的电子邮件地址。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="4a511-p124">Gets the email address of the meeting organizer for a specified meeting. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="4a511-406">类型</span><span class="sxs-lookup"><span data-stu-id="4a511-406">Type</span></span>

*   [<span data-ttu-id="4a511-407">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="4a511-407">EmailAddressDetails</span></span>](/javascript/api/outlook_1_2/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="4a511-408">Requirements</span><span class="sxs-lookup"><span data-stu-id="4a511-408">Requirements</span></span>

|<span data-ttu-id="4a511-409">要求</span><span class="sxs-lookup"><span data-stu-id="4a511-409">Requirement</span></span>| <span data-ttu-id="4a511-410">值</span><span class="sxs-lookup"><span data-stu-id="4a511-410">Value</span></span>|
|---|---|
|[<span data-ttu-id="4a511-411">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="4a511-411">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4a511-412">1.0</span><span class="sxs-lookup"><span data-stu-id="4a511-412">1.0</span></span>|
|[<span data-ttu-id="4a511-413">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="4a511-413">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4a511-414">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4a511-414">ReadItem</span></span>|
|[<span data-ttu-id="4a511-415">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="4a511-415">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4a511-416">阅读</span><span class="sxs-lookup"><span data-stu-id="4a511-416">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4a511-417">示例</span><span class="sxs-lookup"><span data-stu-id="4a511-417">Example</span></span>

```javascript
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
console.log("Organizer: " + organizerName + " (" + organizerAddress + ")");
```

####  <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlook12officeemailaddressdetailsrecipientsjavascriptapioutlook12officerecipients"></a><span data-ttu-id="4a511-418">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_2/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="4a511-418">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_2/office.recipients)</span></span>

<span data-ttu-id="4a511-419">提供对事件的必需与会者的访问权限。</span><span class="sxs-lookup"><span data-stu-id="4a511-419">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="4a511-420">对象的类型和访问级别取决于当前项目的模式。</span><span class="sxs-lookup"><span data-stu-id="4a511-420">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="4a511-421">阅读模式</span><span class="sxs-lookup"><span data-stu-id="4a511-421">Read mode</span></span>

<span data-ttu-id="4a511-422">`requiredAttendees` 属性返回一个数组，其中包含每个必需与会者的 `EmailAddressDetails` 对象。</span><span class="sxs-lookup"><span data-stu-id="4a511-422">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span>

```javascript
var requiredAttendees = Office.context.mailbox.item.requiredAttendees;
console.log("Required attendees: " + JSON.stringify(requiredAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="4a511-423">撰写模式</span><span class="sxs-lookup"><span data-stu-id="4a511-423">Compose mode</span></span>

<span data-ttu-id="4a511-424">`requiredAttendees` 属性返回一个 `Recipients` 对象，该对象提供用于获取或更新必需与会者的方法。</span><span class="sxs-lookup"><span data-stu-id="4a511-424">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span>

```javascript
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
  console.log(JSON.stringify(arrayOfRequiredAttendeesRecipients));
}
```

##### <a name="type"></a><span data-ttu-id="4a511-425">类型</span><span class="sxs-lookup"><span data-stu-id="4a511-425">Type</span></span>

*   <span data-ttu-id="4a511-426">Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_2/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="4a511-426">Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_2/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="4a511-427">Requirements</span><span class="sxs-lookup"><span data-stu-id="4a511-427">Requirements</span></span>

|<span data-ttu-id="4a511-428">要求</span><span class="sxs-lookup"><span data-stu-id="4a511-428">Requirement</span></span>| <span data-ttu-id="4a511-429">值</span><span class="sxs-lookup"><span data-stu-id="4a511-429">Value</span></span>|
|---|---|
|[<span data-ttu-id="4a511-430">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="4a511-430">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4a511-431">1.0</span><span class="sxs-lookup"><span data-stu-id="4a511-431">1.0</span></span>|
|[<span data-ttu-id="4a511-432">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="4a511-432">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4a511-433">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4a511-433">ReadItem</span></span>|
|[<span data-ttu-id="4a511-434">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="4a511-434">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4a511-435">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="4a511-435">Compose or Read</span></span>|

#### <a name="sender-emailaddressdetailsjavascriptapioutlook12officeemailaddressdetails"></a><span data-ttu-id="4a511-436">sender :[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="4a511-436">sender :[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)</span></span>

<span data-ttu-id="4a511-p126">获取电子邮件发件人的电子邮件地址。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="4a511-p126">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="4a511-p127">[`from`](#from-emailaddressdetailsjavascriptapioutlook12officeemailaddressdetails) 和 `sender` 属性表示同一个人，邮件由代理人发送的除外。在此情况下，`from` 属性表示代理程序，而 sender 属性表示代理人。</span><span class="sxs-lookup"><span data-stu-id="4a511-p127">The [`from`](#from-emailaddressdetailsjavascriptapioutlook12officeemailaddressdetails) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="4a511-441">`sender` 属性中 `EmailAddressDetails` 对象的 `recipientType` 属性为 `undefined`。</span><span class="sxs-lookup"><span data-stu-id="4a511-441">The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="4a511-442">类型</span><span class="sxs-lookup"><span data-stu-id="4a511-442">Type</span></span>

*   [<span data-ttu-id="4a511-443">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="4a511-443">EmailAddressDetails</span></span>](/javascript/api/outlook_1_2/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="4a511-444">Requirements</span><span class="sxs-lookup"><span data-stu-id="4a511-444">Requirements</span></span>

|<span data-ttu-id="4a511-445">要求</span><span class="sxs-lookup"><span data-stu-id="4a511-445">Requirement</span></span>| <span data-ttu-id="4a511-446">值</span><span class="sxs-lookup"><span data-stu-id="4a511-446">Value</span></span>|
|---|---|
|[<span data-ttu-id="4a511-447">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="4a511-447">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4a511-448">1.0</span><span class="sxs-lookup"><span data-stu-id="4a511-448">1.0</span></span>|
|[<span data-ttu-id="4a511-449">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="4a511-449">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4a511-450">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4a511-450">ReadItem</span></span>|
|[<span data-ttu-id="4a511-451">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="4a511-451">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4a511-452">阅读</span><span class="sxs-lookup"><span data-stu-id="4a511-452">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4a511-453">示例</span><span class="sxs-lookup"><span data-stu-id="4a511-453">Example</span></span>

```javascript
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
console.log("Sender: " + senderName + " (" + senderAddress + ")");
```

####  <a name="start-datetimejavascriptapioutlook12officetime"></a><span data-ttu-id="4a511-454">start :Date|[Time](/javascript/api/outlook_1_2/office.time)</span><span class="sxs-lookup"><span data-stu-id="4a511-454">start :Date|[Time](/javascript/api/outlook_1_2/office.time)</span></span>

<span data-ttu-id="4a511-455">获取或设置约会开始的日期和时间。</span><span class="sxs-lookup"><span data-stu-id="4a511-455">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="4a511-p128">将 `start` 属性表示为协调世界时 (UTC) 的日期和时间值。可使用 [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook12officelocalclienttime) 方法将该值转换为客户端的本地日期和时间。</span><span class="sxs-lookup"><span data-stu-id="4a511-p128">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook12officelocalclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="4a511-458">阅读模式</span><span class="sxs-lookup"><span data-stu-id="4a511-458">Read mode</span></span>

<span data-ttu-id="4a511-459">`start` 属性返回 `Date` 对象。</span><span class="sxs-lookup"><span data-stu-id="4a511-459">The `start` property returns a `Date` object.</span></span>

```javascript
var start = Office.context.mailbox.item.start;
console.log("Appointment start: " + JSON.stringify(start));
```

##### <a name="compose-mode"></a><span data-ttu-id="4a511-460">撰写模式</span><span class="sxs-lookup"><span data-stu-id="4a511-460">Compose mode</span></span>

<span data-ttu-id="4a511-461">`start` 属性返回 `Time` 对象。</span><span class="sxs-lookup"><span data-stu-id="4a511-461">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="4a511-462">使用 [`Time.setAsync`](/javascript/api/outlook_1_2/office.time#setasync-datetime--options--callback-) 方法设置开始时间时，应使用 [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) 方法将客户端的本地时间转换为服务器的 UTC。</span><span class="sxs-lookup"><span data-stu-id="4a511-462">When you use the [`Time.setAsync`](/javascript/api/outlook_1_2/office.time#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>
<span data-ttu-id="4a511-463">以下示例通过使用 `Time` 对象的 [`setAsync`](/javascript/api/outlook_1_2/office.time#setasync-datetime--options--callback-) 方法，设置撰写模式下约会的开始时间。</span><span class="sxs-lookup"><span data-stu-id="4a511-463">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_2/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="4a511-464">类型</span><span class="sxs-lookup"><span data-stu-id="4a511-464">Type</span></span>

*   <span data-ttu-id="4a511-465">Date | [Time](/javascript/api/outlook_1_2/office.time)</span><span class="sxs-lookup"><span data-stu-id="4a511-465">Date | [Time](/javascript/api/outlook_1_2/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="4a511-466">Requirements</span><span class="sxs-lookup"><span data-stu-id="4a511-466">Requirements</span></span>

|<span data-ttu-id="4a511-467">要求</span><span class="sxs-lookup"><span data-stu-id="4a511-467">Requirement</span></span>| <span data-ttu-id="4a511-468">值</span><span class="sxs-lookup"><span data-stu-id="4a511-468">Value</span></span>|
|---|---|
|[<span data-ttu-id="4a511-469">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="4a511-469">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4a511-470">1.0</span><span class="sxs-lookup"><span data-stu-id="4a511-470">1.0</span></span>|
|[<span data-ttu-id="4a511-471">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="4a511-471">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4a511-472">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4a511-472">ReadItem</span></span>|
|[<span data-ttu-id="4a511-473">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="4a511-473">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4a511-474">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="4a511-474">Compose or Read</span></span>|

####  <a name="subject-stringsubjectjavascriptapioutlook12officesubject"></a><span data-ttu-id="4a511-475">subject :String|[Subject](/javascript/api/outlook_1_2/office.subject)</span><span class="sxs-lookup"><span data-stu-id="4a511-475">subject :String|[Subject](/javascript/api/outlook_1_2/office.subject)</span></span>

<span data-ttu-id="4a511-476">获取或设置显示在项目的主题字段中的说明。</span><span class="sxs-lookup"><span data-stu-id="4a511-476">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="4a511-477">`subject` 属性获取或设置由电子邮件服务器发送项目时的整个主题。</span><span class="sxs-lookup"><span data-stu-id="4a511-477">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="4a511-478">阅读模式</span><span class="sxs-lookup"><span data-stu-id="4a511-478">Read mode</span></span>

<span data-ttu-id="4a511-p130">`subject` 属性返回一个字符串。使用 [`normalizedSubject`](#normalizedsubject-string) 属性获取不带任何前导前缀（如 `RE:` 和 `FW:`）的主题。</span><span class="sxs-lookup"><span data-stu-id="4a511-p130">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

```javascript
var subject = Office.context.mailbox.item.subject;
console.log(subject);
```

##### <a name="compose-mode"></a><span data-ttu-id="4a511-481">撰写模式</span><span class="sxs-lookup"><span data-stu-id="4a511-481">Compose mode</span></span>

<span data-ttu-id="4a511-482">`subject` 属性返回一个 `Subject` 对象，该对象提供用于获取和设置主题的方法。</span><span class="sxs-lookup"><span data-stu-id="4a511-482">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```javascript
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
  console.log(subject);
}
```

##### <a name="type"></a><span data-ttu-id="4a511-483">类型</span><span class="sxs-lookup"><span data-stu-id="4a511-483">Type</span></span>

*   <span data-ttu-id="4a511-484">String | [Subject](/javascript/api/outlook_1_2/office.subject)</span><span class="sxs-lookup"><span data-stu-id="4a511-484">String | [Subject](/javascript/api/outlook_1_2/office.subject)</span></span>

##### <a name="requirements"></a><span data-ttu-id="4a511-485">要求</span><span class="sxs-lookup"><span data-stu-id="4a511-485">Requirements</span></span>

|<span data-ttu-id="4a511-486">要求</span><span class="sxs-lookup"><span data-stu-id="4a511-486">Requirement</span></span>| <span data-ttu-id="4a511-487">值</span><span class="sxs-lookup"><span data-stu-id="4a511-487">Value</span></span>|
|---|---|
|[<span data-ttu-id="4a511-488">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="4a511-488">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4a511-489">1.0</span><span class="sxs-lookup"><span data-stu-id="4a511-489">1.0</span></span>|
|[<span data-ttu-id="4a511-490">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="4a511-490">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4a511-491">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4a511-491">ReadItem</span></span>|
|[<span data-ttu-id="4a511-492">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="4a511-492">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4a511-493">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="4a511-493">Compose or Read</span></span>|

####  <a name="to-arrayemailaddressdetailsjavascriptapioutlook12officeemailaddressdetailsrecipientsjavascriptapioutlook12officerecipients"></a><span data-ttu-id="4a511-494">to :Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_2/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="4a511-494">to :Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_2/office.recipients)</span></span>

<span data-ttu-id="4a511-495">提供对邮件的“**收件人**”行上的收件人的访问权限。</span><span class="sxs-lookup"><span data-stu-id="4a511-495">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="4a511-496">对象的类型和访问级别取决于当前项目的模式。</span><span class="sxs-lookup"><span data-stu-id="4a511-496">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="4a511-497">阅读模式</span><span class="sxs-lookup"><span data-stu-id="4a511-497">Read mode</span></span>

<span data-ttu-id="4a511-p132">`to` 属性返回包含邮件的**收件人**行上所列的每个收件人的 `EmailAddressDetails` 对象的数组。集合上限为 100 个成员。</span><span class="sxs-lookup"><span data-stu-id="4a511-p132">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.</span></span>

```javascript
console.log(JSON.stringify(Office.context.mailbox.item.to));
```

##### <a name="compose-mode"></a><span data-ttu-id="4a511-500">撰写模式</span><span class="sxs-lookup"><span data-stu-id="4a511-500">Compose mode</span></span>

<span data-ttu-id="4a511-501">`to` 属性返回一个 `Recipients` 对象，该对象提供用于获取或更新邮件的“**收件人**”行上收件人的方法。</span><span class="sxs-lookup"><span data-stu-id="4a511-501">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span>

```javascript
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="4a511-502">类型</span><span class="sxs-lookup"><span data-stu-id="4a511-502">Type</span></span>

*   <span data-ttu-id="4a511-503">Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_2/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="4a511-503">Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_2/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="4a511-504">Requirements</span><span class="sxs-lookup"><span data-stu-id="4a511-504">Requirements</span></span>

|<span data-ttu-id="4a511-505">要求</span><span class="sxs-lookup"><span data-stu-id="4a511-505">Requirement</span></span>| <span data-ttu-id="4a511-506">值</span><span class="sxs-lookup"><span data-stu-id="4a511-506">Value</span></span>|
|---|---|
|[<span data-ttu-id="4a511-507">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="4a511-507">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4a511-508">1.0</span><span class="sxs-lookup"><span data-stu-id="4a511-508">1.0</span></span>|
|[<span data-ttu-id="4a511-509">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="4a511-509">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4a511-510">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4a511-510">ReadItem</span></span>|
|[<span data-ttu-id="4a511-511">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="4a511-511">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4a511-512">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="4a511-512">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="4a511-513">方法</span><span class="sxs-lookup"><span data-stu-id="4a511-513">Methods</span></span>

####  <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="4a511-514">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="4a511-514">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="4a511-515">将文件作为附件添加到邮件或约会。</span><span class="sxs-lookup"><span data-stu-id="4a511-515">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="4a511-516">`addFileAttachmentAsync` 方法在指定的 URI 上载文件并将其附加到撰写窗体中的项目。</span><span class="sxs-lookup"><span data-stu-id="4a511-516">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="4a511-517">你随后可以将该标识符与 [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) 方法一同使用，以删除同一个会话中的附件。</span><span class="sxs-lookup"><span data-stu-id="4a511-517">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4a511-518">Parameters</span><span class="sxs-lookup"><span data-stu-id="4a511-518">Parameters</span></span>

|<span data-ttu-id="4a511-519">名称</span><span class="sxs-lookup"><span data-stu-id="4a511-519">Name</span></span>| <span data-ttu-id="4a511-520">类型</span><span class="sxs-lookup"><span data-stu-id="4a511-520">Type</span></span>| <span data-ttu-id="4a511-521">属性</span><span class="sxs-lookup"><span data-stu-id="4a511-521">Attributes</span></span>| <span data-ttu-id="4a511-522">说明</span><span class="sxs-lookup"><span data-stu-id="4a511-522">Description</span></span>|
|---|---|---|---|
|`uri`| <span data-ttu-id="4a511-523">String</span><span class="sxs-lookup"><span data-stu-id="4a511-523">String</span></span>||<span data-ttu-id="4a511-p133">提供附加到邮件或约会的文件的位置的 URI。最大长度为 2048 个字符。</span><span class="sxs-lookup"><span data-stu-id="4a511-p133">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="4a511-526">字符串</span><span class="sxs-lookup"><span data-stu-id="4a511-526">String</span></span>||<span data-ttu-id="4a511-p134">在附件上载过程中显示的附件名称。最大长度为 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="4a511-p134">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="4a511-529">Object</span><span class="sxs-lookup"><span data-stu-id="4a511-529">Object</span></span>| <span data-ttu-id="4a511-530">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="4a511-530">&lt;optional&gt;</span></span>|<span data-ttu-id="4a511-531">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="4a511-531">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="4a511-532">对象</span><span class="sxs-lookup"><span data-stu-id="4a511-532">Object</span></span>| <span data-ttu-id="4a511-533">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="4a511-533">&lt;optional&gt;</span></span>|<span data-ttu-id="4a511-534">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="4a511-534">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="4a511-535">函数</span><span class="sxs-lookup"><span data-stu-id="4a511-535">function</span></span>| <span data-ttu-id="4a511-536">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="4a511-536">&lt;optional&gt;</span></span>|<span data-ttu-id="4a511-537">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="4a511-537">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="4a511-538">如果成功，附件标识符将在 `asyncResult.value` 属性中提供。</span><span class="sxs-lookup"><span data-stu-id="4a511-538">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="4a511-539">如果上传附件失败，`asyncResult` 对象将包含一个提供错误说明的 `Error` 对象。</span><span class="sxs-lookup"><span data-stu-id="4a511-539">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="4a511-540">错误</span><span class="sxs-lookup"><span data-stu-id="4a511-540">Errors</span></span>

| <span data-ttu-id="4a511-541">错误代码</span><span class="sxs-lookup"><span data-stu-id="4a511-541">Error code</span></span> | <span data-ttu-id="4a511-542">说明</span><span class="sxs-lookup"><span data-stu-id="4a511-542">Description</span></span> |
|------------|-------------|
| `AttachmentSizeExceeded` | <span data-ttu-id="4a511-543">附件大小超过了允许的大小。</span><span class="sxs-lookup"><span data-stu-id="4a511-543">The attachment is larger than allowed.</span></span> |
| `FileTypeNotSupported` | <span data-ttu-id="4a511-544">该附件的扩展名不是允许的扩展名。</span><span class="sxs-lookup"><span data-stu-id="4a511-544">The attachment has an extension that is not allowed.</span></span> |
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="4a511-545">邮件或约会具有的附件过多。</span><span class="sxs-lookup"><span data-stu-id="4a511-545">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="4a511-546">要求</span><span class="sxs-lookup"><span data-stu-id="4a511-546">Requirements</span></span>

|<span data-ttu-id="4a511-547">要求</span><span class="sxs-lookup"><span data-stu-id="4a511-547">Requirement</span></span>| <span data-ttu-id="4a511-548">值</span><span class="sxs-lookup"><span data-stu-id="4a511-548">Value</span></span>|
|---|---|
|[<span data-ttu-id="4a511-549">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="4a511-549">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4a511-550">1.1</span><span class="sxs-lookup"><span data-stu-id="4a511-550">1.1</span></span>|
|[<span data-ttu-id="4a511-551">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="4a511-551">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4a511-552">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="4a511-552">ReadWriteItem</span></span>|
|[<span data-ttu-id="4a511-553">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="4a511-553">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4a511-554">撰写</span><span class="sxs-lookup"><span data-stu-id="4a511-554">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="4a511-555">示例</span><span class="sxs-lookup"><span data-stu-id="4a511-555">Example</span></span>

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

####  <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="4a511-556">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="4a511-556">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="4a511-557">将 Exchange 项目（如邮件）作为附件添加到邮件或约会。</span><span class="sxs-lookup"><span data-stu-id="4a511-557">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="4a511-p135">`addItemAttachmentAsync` 方法将包含指定 Exchange 标识符的项目附加到撰写窗体中的项目。如果指定一个回调方法，此方法使用 `asyncResult` 参数调用，该参数包含一个附件标识符或代码，指示附加项目过程中出现的任何错误。可以使用 `options` 参数将状态信息传递给回调方法（如果需要）。</span><span class="sxs-lookup"><span data-stu-id="4a511-p135">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="4a511-561">随后可以将该标识符与 [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) 方法一同使用，以删除同一个会话中的附件。</span><span class="sxs-lookup"><span data-stu-id="4a511-561">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="4a511-562">如果 Office 加载项在 Outlook Web App 中运行，则 `addItemAttachmentAsync` 方法可以将项目附加到项目（正在编辑的项目除外）中；然而，不支持也不建议这样做。</span><span class="sxs-lookup"><span data-stu-id="4a511-562">If your Office Add-in is running in Outlook Web App, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4a511-563">Parameters</span><span class="sxs-lookup"><span data-stu-id="4a511-563">Parameters</span></span>

|<span data-ttu-id="4a511-564">名称</span><span class="sxs-lookup"><span data-stu-id="4a511-564">Name</span></span>| <span data-ttu-id="4a511-565">类型</span><span class="sxs-lookup"><span data-stu-id="4a511-565">Type</span></span>| <span data-ttu-id="4a511-566">属性</span><span class="sxs-lookup"><span data-stu-id="4a511-566">Attributes</span></span>| <span data-ttu-id="4a511-567">说明</span><span class="sxs-lookup"><span data-stu-id="4a511-567">Description</span></span>|
|---|---|---|---|
|`itemId`| <span data-ttu-id="4a511-568">字符串</span><span class="sxs-lookup"><span data-stu-id="4a511-568">String</span></span>||<span data-ttu-id="4a511-p136">要附加的项目的 Exchange 标识符。最大长度为 100 个字符。</span><span class="sxs-lookup"><span data-stu-id="4a511-p136">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="4a511-571">String</span><span class="sxs-lookup"><span data-stu-id="4a511-571">String</span></span>||<span data-ttu-id="4a511-572">要附加的项目的主题。</span><span class="sxs-lookup"><span data-stu-id="4a511-572">The subject of the item to be attached.</span></span> <span data-ttu-id="4a511-573">最大长度为 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="4a511-573">The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="4a511-574">对象</span><span class="sxs-lookup"><span data-stu-id="4a511-574">Object</span></span>| <span data-ttu-id="4a511-575">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="4a511-575">&lt;optional&gt;</span></span>|<span data-ttu-id="4a511-576">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="4a511-576">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="4a511-577">对象</span><span class="sxs-lookup"><span data-stu-id="4a511-577">Object</span></span>| <span data-ttu-id="4a511-578">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="4a511-578">&lt;optional&gt;</span></span>|<span data-ttu-id="4a511-579">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="4a511-579">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="4a511-580">函数</span><span class="sxs-lookup"><span data-stu-id="4a511-580">function</span></span>| <span data-ttu-id="4a511-581">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="4a511-581">&lt;optional&gt;</span></span>|<span data-ttu-id="4a511-582">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="4a511-582">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="4a511-583">如果成功，附件标识符将在 `asyncResult.value` 属性中提供。</span><span class="sxs-lookup"><span data-stu-id="4a511-583">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="4a511-584">如果添加附件失败，`asyncResult` 对象将包含一个提供错误说明的 `Error` 对象。</span><span class="sxs-lookup"><span data-stu-id="4a511-584">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="4a511-585">错误</span><span class="sxs-lookup"><span data-stu-id="4a511-585">Errors</span></span>

| <span data-ttu-id="4a511-586">错误代码</span><span class="sxs-lookup"><span data-stu-id="4a511-586">Error code</span></span> | <span data-ttu-id="4a511-587">说明</span><span class="sxs-lookup"><span data-stu-id="4a511-587">Description</span></span> |
|------------|-------------|
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="4a511-588">邮件或约会具有的附件过多。</span><span class="sxs-lookup"><span data-stu-id="4a511-588">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="4a511-589">Requirements</span><span class="sxs-lookup"><span data-stu-id="4a511-589">Requirements</span></span>

|<span data-ttu-id="4a511-590">要求</span><span class="sxs-lookup"><span data-stu-id="4a511-590">Requirement</span></span>| <span data-ttu-id="4a511-591">值</span><span class="sxs-lookup"><span data-stu-id="4a511-591">Value</span></span>|
|---|---|
|[<span data-ttu-id="4a511-592">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="4a511-592">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4a511-593">1.1</span><span class="sxs-lookup"><span data-stu-id="4a511-593">1.1</span></span>|
|[<span data-ttu-id="4a511-594">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="4a511-594">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4a511-595">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="4a511-595">ReadWriteItem</span></span>|
|[<span data-ttu-id="4a511-596">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="4a511-596">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4a511-597">撰写</span><span class="sxs-lookup"><span data-stu-id="4a511-597">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="4a511-598">示例</span><span class="sxs-lookup"><span data-stu-id="4a511-598">Example</span></span>

<span data-ttu-id="4a511-599">以下示例将现有的 Outlook 项目添加为名为 `My Attachment` 的附件。</span><span class="sxs-lookup"><span data-stu-id="4a511-599">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

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

#### <a name="displayreplyallformformdata-callback"></a><span data-ttu-id="4a511-600">displayReplyAllForm (formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="4a511-600">displayReplyAllForm(formData, [callback])</span></span>

<span data-ttu-id="4a511-601">显示答复窗体，其中包括所选邮件的发件人和所有收件人或所选约会的组织者和所有与会者。</span><span class="sxs-lookup"><span data-stu-id="4a511-601">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="4a511-602">在 Outlook for iOS 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="4a511-602">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="4a511-603">在 Outlook Web App 中，答复窗体显示为包含 3 列视图的弹出式窗体以及包含 2 列或 1 列视图的弹出式窗体。</span><span class="sxs-lookup"><span data-stu-id="4a511-603">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="4a511-604">如果任意字符串参数超出其限制，`displayReplyAllForm` 将引发异常。</span><span class="sxs-lookup"><span data-stu-id="4a511-604">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="4a511-p138">当在 `formData.attachments` 参数中指定附件时，Outlook 和 Outlook Web App 尝试下载所有附件并将其附加到答复窗体。如果无法添加任何附件，则在窗体 UI 中显示错误。如果这不可能，则不引发错误消息。</span><span class="sxs-lookup"><span data-stu-id="4a511-p138">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4a511-608">Parameters</span><span class="sxs-lookup"><span data-stu-id="4a511-608">Parameters</span></span>

|<span data-ttu-id="4a511-609">名称</span><span class="sxs-lookup"><span data-stu-id="4a511-609">Name</span></span>| <span data-ttu-id="4a511-610">说明</span><span class="sxs-lookup"><span data-stu-id="4a511-610">Type</span></span>| <span data-ttu-id="4a511-611">说明</span><span class="sxs-lookup"><span data-stu-id="4a511-611">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="4a511-612">字符串 &#124; 对象</span><span class="sxs-lookup"><span data-stu-id="4a511-612">String &#124; Object</span></span>| |<span data-ttu-id="4a511-p139">一个包含文本和 HTML 且表示答复窗体的正文的字符串。字符串限制为 32 KB。</span><span class="sxs-lookup"><span data-stu-id="4a511-p139">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="4a511-615">**或**</span><span class="sxs-lookup"><span data-stu-id="4a511-615">**OR**</span></span><br/><span data-ttu-id="4a511-p140">包含正文或附件数据和回调函数的对象。对象定义如下。</span><span class="sxs-lookup"><span data-stu-id="4a511-p140">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="4a511-618">String</span><span class="sxs-lookup"><span data-stu-id="4a511-618">String</span></span> | <span data-ttu-id="4a511-619">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="4a511-619">&lt;optional&gt;</span></span> | <span data-ttu-id="4a511-p141">一个包含文本和 HTML 且表示答复窗体的正文的字符串。字符串限制为 32 KB。</span><span class="sxs-lookup"><span data-stu-id="4a511-p141">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="4a511-622">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="4a511-622">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="4a511-623">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="4a511-623">&lt;optional&gt;</span></span> | <span data-ttu-id="4a511-624">JSON 对象（是文件或项目附件）的数组。</span><span class="sxs-lookup"><span data-stu-id="4a511-624">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="4a511-625">String</span><span class="sxs-lookup"><span data-stu-id="4a511-625">String</span></span> | | <span data-ttu-id="4a511-p142">指示附件的类型。必须是文件附件的 `file` 或项目附件的 `item`。</span><span class="sxs-lookup"><span data-stu-id="4a511-p142">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="4a511-628">字符串</span><span class="sxs-lookup"><span data-stu-id="4a511-628">String</span></span> | | <span data-ttu-id="4a511-629">一个包含附件的名称的字符串，最多包含 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="4a511-629">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="4a511-630">String</span><span class="sxs-lookup"><span data-stu-id="4a511-630">String</span></span> | | <span data-ttu-id="4a511-p143">仅在将 `type` 设置为 `file` 时使用。文件的位置的 URI。</span><span class="sxs-lookup"><span data-stu-id="4a511-p143">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="4a511-633">字符串</span><span class="sxs-lookup"><span data-stu-id="4a511-633">String</span></span> | | <span data-ttu-id="4a511-p144">仅在将 `type` 设置为 `item` 时使用。附件的 EWS 项目 ID。字符串最长为 100 个字符。</span><span class="sxs-lookup"><span data-stu-id="4a511-p144">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="4a511-637">函数</span><span class="sxs-lookup"><span data-stu-id="4a511-637">function</span></span> | <span data-ttu-id="4a511-638">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="4a511-638">&lt;optional&gt;</span></span> | <span data-ttu-id="4a511-639">方法完成后，使用单个参数 `asyncResult`（一个 [AsyncResult](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="4a511-639">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="4a511-640">要求</span><span class="sxs-lookup"><span data-stu-id="4a511-640">Requirements</span></span>

|<span data-ttu-id="4a511-641">要求</span><span class="sxs-lookup"><span data-stu-id="4a511-641">Requirement</span></span>| <span data-ttu-id="4a511-642">值</span><span class="sxs-lookup"><span data-stu-id="4a511-642">Value</span></span>|
|---|---|
|[<span data-ttu-id="4a511-643">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="4a511-643">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4a511-644">1.0</span><span class="sxs-lookup"><span data-stu-id="4a511-644">1.0</span></span>|
|[<span data-ttu-id="4a511-645">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="4a511-645">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4a511-646">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4a511-646">ReadItem</span></span>|
|[<span data-ttu-id="4a511-647">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="4a511-647">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4a511-648">阅读</span><span class="sxs-lookup"><span data-stu-id="4a511-648">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="4a511-649">示例</span><span class="sxs-lookup"><span data-stu-id="4a511-649">Examples</span></span>

<span data-ttu-id="4a511-650">以下代码将一个字符串传递到 `displayReplyAllForm` 函数。</span><span class="sxs-lookup"><span data-stu-id="4a511-650">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="4a511-651">使用空白正文答复。</span><span class="sxs-lookup"><span data-stu-id="4a511-651">Reply with an empty body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="4a511-652">仅使用正文答复。</span><span class="sxs-lookup"><span data-stu-id="4a511-652">Reply with just a body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="4a511-653">使用正文和文件附件答复。</span><span class="sxs-lookup"><span data-stu-id="4a511-653">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="4a511-654">使用正文和项目附件答复。</span><span class="sxs-lookup"><span data-stu-id="4a511-654">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="4a511-655">使用正文、文件附件、项目附件和回调答复。</span><span class="sxs-lookup"><span data-stu-id="4a511-655">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="displayreplyformformdata-callback"></a><span data-ttu-id="4a511-656">displayReplyForm (formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="4a511-656">displayReplyForm(formData, [callback])</span></span>

<span data-ttu-id="4a511-657">显示答复窗体，其中仅包括所选邮件的发件人或所选约会的组织者。</span><span class="sxs-lookup"><span data-stu-id="4a511-657">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="4a511-658">在 Outlook for iOS 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="4a511-658">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="4a511-659">在 Outlook Web App 中，答复窗体显示为包含 3 列视图的弹出式窗体以及包含 2 列或 1 列视图的弹出式窗体。</span><span class="sxs-lookup"><span data-stu-id="4a511-659">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="4a511-660">如果任意字符串参数超出其限制，`displayReplyForm` 将引发异常。</span><span class="sxs-lookup"><span data-stu-id="4a511-660">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="4a511-p145">当在 `formData.attachments` 参数中指定附件时，Outlook 和 Outlook Web App 尝试下载所有附件并将其附加到答复窗体。如果无法添加任何附件，则在窗体 UI 中显示错误。如果这不可能，则不引发错误消息。</span><span class="sxs-lookup"><span data-stu-id="4a511-p145">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4a511-664">Parameters</span><span class="sxs-lookup"><span data-stu-id="4a511-664">Parameters</span></span>

|<span data-ttu-id="4a511-665">名称</span><span class="sxs-lookup"><span data-stu-id="4a511-665">Name</span></span>| <span data-ttu-id="4a511-666">说明</span><span class="sxs-lookup"><span data-stu-id="4a511-666">Type</span></span>| <span data-ttu-id="4a511-667">说明</span><span class="sxs-lookup"><span data-stu-id="4a511-667">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="4a511-668">字符串 &#124; 对象</span><span class="sxs-lookup"><span data-stu-id="4a511-668">String &#124; Object</span></span>| | <span data-ttu-id="4a511-p146">一个包含文本和 HTML 且表示答复窗体的正文的字符串。字符串限制为 32 KB。</span><span class="sxs-lookup"><span data-stu-id="4a511-p146">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="4a511-671">**或**</span><span class="sxs-lookup"><span data-stu-id="4a511-671">**OR**</span></span><br/><span data-ttu-id="4a511-p147">包含正文或附件数据和回调函数的对象。对象定义如下。</span><span class="sxs-lookup"><span data-stu-id="4a511-p147">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="4a511-674">String</span><span class="sxs-lookup"><span data-stu-id="4a511-674">String</span></span> | <span data-ttu-id="4a511-675">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="4a511-675">&lt;optional&gt;</span></span> | <span data-ttu-id="4a511-p148">一个包含文本和 HTML 且表示答复窗体的正文的字符串。字符串限制为 32 KB。</span><span class="sxs-lookup"><span data-stu-id="4a511-p148">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="4a511-678">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="4a511-678">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="4a511-679">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="4a511-679">&lt;optional&gt;</span></span> | <span data-ttu-id="4a511-680">JSON 对象（是文件或项目附件）的数组。</span><span class="sxs-lookup"><span data-stu-id="4a511-680">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="4a511-681">String</span><span class="sxs-lookup"><span data-stu-id="4a511-681">String</span></span> | | <span data-ttu-id="4a511-p149">指示附件的类型。必须是文件附件的 `file` 或项目附件的 `item`。</span><span class="sxs-lookup"><span data-stu-id="4a511-p149">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="4a511-684">字符串</span><span class="sxs-lookup"><span data-stu-id="4a511-684">String</span></span> | | <span data-ttu-id="4a511-685">一个包含附件的名称的字符串，最多包含 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="4a511-685">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="4a511-686">String</span><span class="sxs-lookup"><span data-stu-id="4a511-686">String</span></span> | | <span data-ttu-id="4a511-p150">仅在将 `type` 设置为 `file` 时使用。文件的位置的 URI。</span><span class="sxs-lookup"><span data-stu-id="4a511-p150">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="4a511-689">字符串</span><span class="sxs-lookup"><span data-stu-id="4a511-689">String</span></span> | | <span data-ttu-id="4a511-p151">仅在将 `type` 设置为 `item` 时使用。附件的 EWS 项目 ID。字符串最长为 100 个字符。</span><span class="sxs-lookup"><span data-stu-id="4a511-p151">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="4a511-693">函数</span><span class="sxs-lookup"><span data-stu-id="4a511-693">function</span></span> | <span data-ttu-id="4a511-694">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="4a511-694">&lt;optional&gt;</span></span> | <span data-ttu-id="4a511-695">方法完成后，使用单个参数 `asyncResult`（一个 [AsyncResult](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="4a511-695">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="4a511-696">Requirements</span><span class="sxs-lookup"><span data-stu-id="4a511-696">Requirements</span></span>

|<span data-ttu-id="4a511-697">要求</span><span class="sxs-lookup"><span data-stu-id="4a511-697">Requirement</span></span>| <span data-ttu-id="4a511-698">值</span><span class="sxs-lookup"><span data-stu-id="4a511-698">Value</span></span>|
|---|---|
|[<span data-ttu-id="4a511-699">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="4a511-699">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4a511-700">1.0</span><span class="sxs-lookup"><span data-stu-id="4a511-700">1.0</span></span>|
|[<span data-ttu-id="4a511-701">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="4a511-701">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4a511-702">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4a511-702">ReadItem</span></span>|
|[<span data-ttu-id="4a511-703">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="4a511-703">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4a511-704">阅读</span><span class="sxs-lookup"><span data-stu-id="4a511-704">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="4a511-705">示例</span><span class="sxs-lookup"><span data-stu-id="4a511-705">Examples</span></span>

<span data-ttu-id="4a511-706">以下代码将一个字符串传递到 `displayReplyForm` 函数。</span><span class="sxs-lookup"><span data-stu-id="4a511-706">The following code passes a string to the `displayReplyForm` function.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="4a511-707">使用空白正文答复。</span><span class="sxs-lookup"><span data-stu-id="4a511-707">Reply with an empty body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="4a511-708">仅使用正文答复。</span><span class="sxs-lookup"><span data-stu-id="4a511-708">Reply with just a body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="4a511-709">使用正文和文件附件答复。</span><span class="sxs-lookup"><span data-stu-id="4a511-709">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="4a511-710">使用正文和项目附件答复。</span><span class="sxs-lookup"><span data-stu-id="4a511-710">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="4a511-711">使用正文、文件附件、项目附件和回调答复。</span><span class="sxs-lookup"><span data-stu-id="4a511-711">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="getentities--entitiesjavascriptapioutlook12officeentities"></a><span data-ttu-id="4a511-712">getEntities() → {[Entities](/javascript/api/outlook_1_2/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="4a511-712">getEntities() → {[Entities](/javascript/api/outlook_1_2/office.entities)}</span></span>

<span data-ttu-id="4a511-713">获取在所选项目的正文中找到的实体。</span><span class="sxs-lookup"><span data-stu-id="4a511-713">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="4a511-714">在 Outlook for iOS 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="4a511-714">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="4a511-715">要求</span><span class="sxs-lookup"><span data-stu-id="4a511-715">Requirements</span></span>

|<span data-ttu-id="4a511-716">要求</span><span class="sxs-lookup"><span data-stu-id="4a511-716">Requirement</span></span>| <span data-ttu-id="4a511-717">值</span><span class="sxs-lookup"><span data-stu-id="4a511-717">Value</span></span>|
|---|---|
|[<span data-ttu-id="4a511-718">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="4a511-718">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4a511-719">1.0</span><span class="sxs-lookup"><span data-stu-id="4a511-719">1.0</span></span>|
|[<span data-ttu-id="4a511-720">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="4a511-720">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4a511-721">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4a511-721">ReadItem</span></span>|
|[<span data-ttu-id="4a511-722">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="4a511-722">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4a511-723">阅读</span><span class="sxs-lookup"><span data-stu-id="4a511-723">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="4a511-724">返回：</span><span class="sxs-lookup"><span data-stu-id="4a511-724">Returns:</span></span>

<span data-ttu-id="4a511-725">类型：[Entities](/javascript/api/outlook_1_2/office.entities)</span><span class="sxs-lookup"><span data-stu-id="4a511-725">Type: [Entities](/javascript/api/outlook_1_2/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="4a511-726">示例</span><span class="sxs-lookup"><span data-stu-id="4a511-726">Example</span></span>

<span data-ttu-id="4a511-727">以下示例访问当前项目的正文中的联系人实体。</span><span class="sxs-lookup"><span data-stu-id="4a511-727">The following example accesses the contacts entities in the current item's body.</span></span>

```javascript
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlook12officecontactmeetingsuggestionjavascriptapioutlook12officemeetingsuggestionphonenumberjavascriptapioutlook12officephonenumbertasksuggestionjavascriptapioutlook12officetasksuggestion"></a><span data-ttu-id="4a511-728">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_2/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_2/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_2/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_2/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="4a511-728">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_2/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_2/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_2/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_2/office.tasksuggestion))>}</span></span>

<span data-ttu-id="4a511-729">获取所选项目的正文中找到的指定实体类型的所有实体的数组。</span><span class="sxs-lookup"><span data-stu-id="4a511-729">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="4a511-730">在 Outlook for iOS 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="4a511-730">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4a511-731">Parameters</span><span class="sxs-lookup"><span data-stu-id="4a511-731">Parameters</span></span>

|<span data-ttu-id="4a511-732">名称</span><span class="sxs-lookup"><span data-stu-id="4a511-732">Name</span></span>| <span data-ttu-id="4a511-733">类型</span><span class="sxs-lookup"><span data-stu-id="4a511-733">Type</span></span>| <span data-ttu-id="4a511-734">说明</span><span class="sxs-lookup"><span data-stu-id="4a511-734">Description</span></span>|
|---|---|---|
|`entityType`| [<span data-ttu-id="4a511-735">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="4a511-735">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook_1_2/office.mailboxenums.entitytype)|<span data-ttu-id="4a511-736">EntityType 枚举值之一。</span><span class="sxs-lookup"><span data-stu-id="4a511-736">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4a511-737">Requirements</span><span class="sxs-lookup"><span data-stu-id="4a511-737">Requirements</span></span>

|<span data-ttu-id="4a511-738">要求</span><span class="sxs-lookup"><span data-stu-id="4a511-738">Requirement</span></span>| <span data-ttu-id="4a511-739">值</span><span class="sxs-lookup"><span data-stu-id="4a511-739">Value</span></span>|
|---|---|
|[<span data-ttu-id="4a511-740">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="4a511-740">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4a511-741">1.0</span><span class="sxs-lookup"><span data-stu-id="4a511-741">1.0</span></span>|
|[<span data-ttu-id="4a511-742">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="4a511-742">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4a511-743">受限</span><span class="sxs-lookup"><span data-stu-id="4a511-743">Restricted</span></span>|
|[<span data-ttu-id="4a511-744">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="4a511-744">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4a511-745">阅读</span><span class="sxs-lookup"><span data-stu-id="4a511-745">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="4a511-746">返回：</span><span class="sxs-lookup"><span data-stu-id="4a511-746">Returns:</span></span>

<span data-ttu-id="4a511-747">如果在 `entityType` 中传递的值不是 `EntityType` 枚举的有效成员，该方法返回 null。</span><span class="sxs-lookup"><span data-stu-id="4a511-747">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="4a511-748">如果指定类型的任何实体都不存在于该项目的正文中，该方法将返回空数组。</span><span class="sxs-lookup"><span data-stu-id="4a511-748">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="4a511-749">否则，返回的数组中对象的类型取决于 `entityType` 参数中请求实体的类型。</span><span class="sxs-lookup"><span data-stu-id="4a511-749">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="4a511-750">当使用此方法的最低权限级别**受限**时，某些实体类型需要 **ReadItem** 才能进行访问，如下表中所指定。</span><span class="sxs-lookup"><span data-stu-id="4a511-750">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

| <span data-ttu-id="4a511-751">`entityType` 的值</span><span class="sxs-lookup"><span data-stu-id="4a511-751">Value of `entityType`</span></span> | <span data-ttu-id="4a511-752">返回的数组中对象的类型</span><span class="sxs-lookup"><span data-stu-id="4a511-752">Type of objects in returned array</span></span> | <span data-ttu-id="4a511-753">所需权限级别</span><span class="sxs-lookup"><span data-stu-id="4a511-753">Required Permission Level</span></span> |
| --- | --- | --- |
| `Address` | <span data-ttu-id="4a511-754">String</span><span class="sxs-lookup"><span data-stu-id="4a511-754">String</span></span> | <span data-ttu-id="4a511-755">**受限**</span><span class="sxs-lookup"><span data-stu-id="4a511-755">**Restricted**</span></span> |
| `Contact` | <span data-ttu-id="4a511-756">Contact</span><span class="sxs-lookup"><span data-stu-id="4a511-756">Contact</span></span> | <span data-ttu-id="4a511-757">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="4a511-757">**ReadItem**</span></span> |
| `EmailAddress` | <span data-ttu-id="4a511-758">String</span><span class="sxs-lookup"><span data-stu-id="4a511-758">String</span></span> | <span data-ttu-id="4a511-759">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="4a511-759">**ReadItem**</span></span> |
| `MeetingSuggestion` | <span data-ttu-id="4a511-760">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="4a511-760">MeetingSuggestion</span></span> | <span data-ttu-id="4a511-761">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="4a511-761">**ReadItem**</span></span> |
| `PhoneNumber` | <span data-ttu-id="4a511-762">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="4a511-762">PhoneNumber</span></span> | <span data-ttu-id="4a511-763">**受限**</span><span class="sxs-lookup"><span data-stu-id="4a511-763">**Restricted**</span></span> |
| `TaskSuggestion` | <span data-ttu-id="4a511-764">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="4a511-764">TaskSuggestion</span></span> | <span data-ttu-id="4a511-765">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="4a511-765">**ReadItem**</span></span> |
| `URL` | <span data-ttu-id="4a511-766">String</span><span class="sxs-lookup"><span data-stu-id="4a511-766">String</span></span> | <span data-ttu-id="4a511-767">**受限**</span><span class="sxs-lookup"><span data-stu-id="4a511-767">**Restricted**</span></span> |

<span data-ttu-id="4a511-768">类型：Array.<(String|[Contact](/javascript/api/outlook_1_2/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_2/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_2/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_2/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="4a511-768">Type: Array.<(String|[Contact](/javascript/api/outlook_1_2/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_2/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_2/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_2/office.tasksuggestion))></span></span>

##### <a name="example"></a><span data-ttu-id="4a511-769">示例</span><span class="sxs-lookup"><span data-stu-id="4a511-769">Example</span></span>

<span data-ttu-id="4a511-770">以下示例显示了如何访问表示当前项目的正文中的邮政地址的字符串数组。</span><span class="sxs-lookup"><span data-stu-id="4a511-770">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlook12officecontactmeetingsuggestionjavascriptapioutlook12officemeetingsuggestionphonenumberjavascriptapioutlook12officephonenumbertasksuggestionjavascriptapioutlook12officetasksuggestion"></a><span data-ttu-id="4a511-771">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_2/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_2/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_2/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_2/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="4a511-771">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_2/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_2/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_2/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_2/office.tasksuggestion))>}</span></span>

<span data-ttu-id="4a511-772">返回传递清单 XML 文件中定义的命名筛选器的所选项目中的已知实体。</span><span class="sxs-lookup"><span data-stu-id="4a511-772">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="4a511-773">在 Outlook for iOS 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="4a511-773">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="4a511-774">`getFilteredEntitiesByName` 方法返回匹配在具有指定 `FilterName` 元素值的清单 XML 文件中的 [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) 规则元素中定义的正则表达式的实体。</span><span class="sxs-lookup"><span data-stu-id="4a511-774">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4a511-775">Parameters</span><span class="sxs-lookup"><span data-stu-id="4a511-775">Parameters</span></span>

|<span data-ttu-id="4a511-776">名称</span><span class="sxs-lookup"><span data-stu-id="4a511-776">Name</span></span>| <span data-ttu-id="4a511-777">类型</span><span class="sxs-lookup"><span data-stu-id="4a511-777">Type</span></span>| <span data-ttu-id="4a511-778">说明</span><span class="sxs-lookup"><span data-stu-id="4a511-778">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="4a511-779">字符串</span><span class="sxs-lookup"><span data-stu-id="4a511-779">String</span></span>|<span data-ttu-id="4a511-780">定义筛选器匹配的 `ItemHasKnownEntity` 规则元素的名称。</span><span class="sxs-lookup"><span data-stu-id="4a511-780">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4a511-781">Requirements</span><span class="sxs-lookup"><span data-stu-id="4a511-781">Requirements</span></span>

|<span data-ttu-id="4a511-782">要求</span><span class="sxs-lookup"><span data-stu-id="4a511-782">Requirement</span></span>| <span data-ttu-id="4a511-783">值</span><span class="sxs-lookup"><span data-stu-id="4a511-783">Value</span></span>|
|---|---|
|[<span data-ttu-id="4a511-784">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="4a511-784">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4a511-785">1.0</span><span class="sxs-lookup"><span data-stu-id="4a511-785">1.0</span></span>|
|[<span data-ttu-id="4a511-786">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="4a511-786">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4a511-787">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4a511-787">ReadItem</span></span>|
|[<span data-ttu-id="4a511-788">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="4a511-788">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4a511-789">阅读</span><span class="sxs-lookup"><span data-stu-id="4a511-789">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="4a511-790">返回：</span><span class="sxs-lookup"><span data-stu-id="4a511-790">Returns:</span></span>

<span data-ttu-id="4a511-p153">如果具有匹配 `name` 参数的 `FilterName` 元素值的清单中没有任何 `ItemHasKnownEntity` 元素，则该方法返回 `null`。如果 `name` 参数匹配清单中的 `ItemHasKnownEntity` 元素，但在匹配的当前项目中没有实体，则该方法返回一个空数组。</span><span class="sxs-lookup"><span data-stu-id="4a511-p153">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="4a511-793">类型：Array.<(String|[Contact](/javascript/api/outlook_1_2/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_2/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_2/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_2/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="4a511-793">Type: Array.<(String|[Contact](/javascript/api/outlook_1_2/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_2/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_2/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_2/office.tasksuggestion))></span></span>

#### <a name="getregexmatches--object"></a><span data-ttu-id="4a511-794">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="4a511-794">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="4a511-795">返回所选项目中匹配在清单 XML 文件中定义的正则表达式的字符串值。</span><span class="sxs-lookup"><span data-stu-id="4a511-795">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="4a511-796">在 Outlook for iOS 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="4a511-796">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="4a511-p154">`getRegExMatches` 方法返回匹配在清单 XML 文件中的每个 `ItemHasRegularExpressionMatch` 或 `ItemHasKnownEntity` 规则元素中定义的正则表达式的字符串。对于 `ItemHasRegularExpressionMatch` 规则，匹配字符串必须发生在该规则指定的项目的属性中。`PropertyName` 简单类型定义支持的属性。</span><span class="sxs-lookup"><span data-stu-id="4a511-p154">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="4a511-800">例如，考虑一个外接程序清单具有以下 `Rule` 元素：</span><span class="sxs-lookup"><span data-stu-id="4a511-800">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="4a511-801">从 `getRegExMatches` 返回的对象应有两个属性：`fruits` 和 `veggies`。</span><span class="sxs-lookup"><span data-stu-id="4a511-801">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

> [!NOTE]
> <span data-ttu-id="4a511-p155">如果在项目的正文属性上指定 `ItemHasRegularExpressionMatch` 规则，则正则表达式应进一步筛选正文并应尝试返回该项目的整个正文。使用正则表达式（如 `.*`）获取项目的整个正文并不总是返回预期的结果。</span><span class="sxs-lookup"><span data-stu-id="4a511-p155">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="requirements"></a><span data-ttu-id="4a511-804">Requirements</span><span class="sxs-lookup"><span data-stu-id="4a511-804">Requirements</span></span>

|<span data-ttu-id="4a511-805">要求</span><span class="sxs-lookup"><span data-stu-id="4a511-805">Requirement</span></span>| <span data-ttu-id="4a511-806">值</span><span class="sxs-lookup"><span data-stu-id="4a511-806">Value</span></span>|
|---|---|
|[<span data-ttu-id="4a511-807">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="4a511-807">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4a511-808">1.0</span><span class="sxs-lookup"><span data-stu-id="4a511-808">1.0</span></span>|
|[<span data-ttu-id="4a511-809">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="4a511-809">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4a511-810">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4a511-810">ReadItem</span></span>|
|[<span data-ttu-id="4a511-811">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="4a511-811">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4a511-812">阅读</span><span class="sxs-lookup"><span data-stu-id="4a511-812">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="4a511-813">返回：</span><span class="sxs-lookup"><span data-stu-id="4a511-813">Returns:</span></span>

<span data-ttu-id="4a511-p156">一个包含匹配在清单 XML 文件中定义的正则表达式的字符串数组的对象。每个数组的名称等于匹配 `ItemHasRegularExpressionMatch` 规则的 `RegExName` 属性或匹配 `ItemHasKnownEntity` 规则的 `FilterName` 属性的相应值。</span><span class="sxs-lookup"><span data-stu-id="4a511-p156">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<dl class="param-type"><span data-ttu-id="4a511-816">

<dt>
类型</dt>


</span><span class="sxs-lookup"><span data-stu-id="4a511-816">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="4a511-817">对象</span><span class="sxs-lookup"><span data-stu-id="4a511-817">Object</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="4a511-818">示例</span><span class="sxs-lookup"><span data-stu-id="4a511-818">Example</span></span>

<span data-ttu-id="4a511-819">以下示例显示了如何访问正则表达式 <rule> 元素 `fruits` 和 `veggies` 的匹配项的数组，这些元素在清单中指定。</rule></span><span class="sxs-lookup"><span data-stu-id="4a511-819">The following example shows how to access the array of matches for the regular expression <rule>elements `fruits` and `veggies`, which are specified in the manifest.</rule></span></span>

```javascript
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veggies = allMatches.veggies;
```

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="4a511-820">getRegExMatchesByName(name) → (nullable) {Array.<String>}</span><span class="sxs-lookup"><span data-stu-id="4a511-820">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="4a511-821">返回匹配在清单 XML 文件中定义的命名正则表达式的所选项目中的字符串值。</span><span class="sxs-lookup"><span data-stu-id="4a511-821">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="4a511-822">在 Outlook for iOS 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="4a511-822">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="4a511-823">`getRegExMatchesByName` 方法返回匹配在具有指定 `RegExName` 元素值的清单 XML 文件中的 `ItemHasRegularExpressionMatch` 规则元素中定义的正则表达式的字符串。</span><span class="sxs-lookup"><span data-stu-id="4a511-823">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="4a511-p157">如果在项目的正文属性上指定 `ItemHasRegularExpressionMatch` 规则，则正则表达式应进一步筛选正文，不应尝试返回该项目的整个正文。使用正则表达式（如 `.*`）获取项目的整个正文并不总是返回预期的结果。</span><span class="sxs-lookup"><span data-stu-id="4a511-p157">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4a511-826">Parameters</span><span class="sxs-lookup"><span data-stu-id="4a511-826">Parameters</span></span>

|<span data-ttu-id="4a511-827">名称</span><span class="sxs-lookup"><span data-stu-id="4a511-827">Name</span></span>| <span data-ttu-id="4a511-828">说明</span><span class="sxs-lookup"><span data-stu-id="4a511-828">Type</span></span>| <span data-ttu-id="4a511-829">描述</span><span class="sxs-lookup"><span data-stu-id="4a511-829">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="4a511-830">字符串</span><span class="sxs-lookup"><span data-stu-id="4a511-830">String</span></span>|<span data-ttu-id="4a511-831">定义筛选器匹配的 `ItemHasRegularExpressionMatch` 规则元素的名称。</span><span class="sxs-lookup"><span data-stu-id="4a511-831">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4a511-832">Requirements</span><span class="sxs-lookup"><span data-stu-id="4a511-832">Requirements</span></span>

|<span data-ttu-id="4a511-833">要求</span><span class="sxs-lookup"><span data-stu-id="4a511-833">Requirement</span></span>| <span data-ttu-id="4a511-834">值</span><span class="sxs-lookup"><span data-stu-id="4a511-834">Value</span></span>|
|---|---|
|[<span data-ttu-id="4a511-835">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="4a511-835">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4a511-836">1.0</span><span class="sxs-lookup"><span data-stu-id="4a511-836">1.0</span></span>|
|[<span data-ttu-id="4a511-837">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="4a511-837">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4a511-838">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4a511-838">ReadItem</span></span>|
|[<span data-ttu-id="4a511-839">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="4a511-839">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4a511-840">阅读</span><span class="sxs-lookup"><span data-stu-id="4a511-840">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="4a511-841">返回：</span><span class="sxs-lookup"><span data-stu-id="4a511-841">Returns:</span></span>

<span data-ttu-id="4a511-842">一个包含匹配在清单 XML 文件中定义的正则表达式的字符串的数组。</span><span class="sxs-lookup"><span data-stu-id="4a511-842">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<dl class="param-type"><span data-ttu-id="4a511-843">

<dt>类型</dt>

</span><span class="sxs-lookup"><span data-stu-id="4a511-843">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="4a511-844">Array.< String ></span><span class="sxs-lookup"><span data-stu-id="4a511-844">Array.< String ></span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="4a511-845">示例</span><span class="sxs-lookup"><span data-stu-id="4a511-845">Example</span></span>

```javascript
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

####  <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="4a511-846">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="4a511-846">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="4a511-847">以异步方式返回邮件的主题或正文中选定的数据。</span><span class="sxs-lookup"><span data-stu-id="4a511-847">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="4a511-p158">如果没有选定内容，但光标位于正文或主题中，此方法将会为所选数据返回 null。如果选定的是字段，而不是正文或主题，则此方法返回 `InvalidSelection` 错误。</span><span class="sxs-lookup"><span data-stu-id="4a511-p158">If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4a511-850">Parameters</span><span class="sxs-lookup"><span data-stu-id="4a511-850">Parameters</span></span>

|<span data-ttu-id="4a511-851">名称</span><span class="sxs-lookup"><span data-stu-id="4a511-851">Name</span></span>| <span data-ttu-id="4a511-852">类型</span><span class="sxs-lookup"><span data-stu-id="4a511-852">Type</span></span>| <span data-ttu-id="4a511-853">属性</span><span class="sxs-lookup"><span data-stu-id="4a511-853">Attributes</span></span>| <span data-ttu-id="4a511-854">说明</span><span class="sxs-lookup"><span data-stu-id="4a511-854">Description</span></span>|
|---|---|---|---|
|`coercionType`| [<span data-ttu-id="4a511-855">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="4a511-855">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="4a511-p159">请求数据的格式。如果为文本，则此方法返回纯文本作为字符串，删除任何显示的 HTML 标记。如果为 HTML，则此方法返回所选文本，不论是纯文本还是 HTML。</span><span class="sxs-lookup"><span data-stu-id="4a511-p159">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`| <span data-ttu-id="4a511-859">对象</span><span class="sxs-lookup"><span data-stu-id="4a511-859">Object</span></span>| <span data-ttu-id="4a511-860">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="4a511-860">&lt;optional&gt;</span></span>|<span data-ttu-id="4a511-861">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="4a511-861">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="4a511-862">对象</span><span class="sxs-lookup"><span data-stu-id="4a511-862">Object</span></span>| <span data-ttu-id="4a511-863">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="4a511-863">&lt;optional&gt;</span></span>|<span data-ttu-id="4a511-864">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="4a511-864">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="4a511-865">函数</span><span class="sxs-lookup"><span data-stu-id="4a511-865">function</span></span>||<span data-ttu-id="4a511-866">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="4a511-866">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="4a511-867">若要从回调方法访问所选数据，请调用 `asyncResult.value.data`。</span><span class="sxs-lookup"><span data-stu-id="4a511-867">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="4a511-868">若要访问选定内容的源属性，请调用 `asyncResult.value.sourceProperty`，这将为 `body` 或 `subject`。</span><span class="sxs-lookup"><span data-stu-id="4a511-868">To access the source property that the selection comes from, call `asyncResult.value.sourceProperty`, which will be either `body` or `subject`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4a511-869">要求</span><span class="sxs-lookup"><span data-stu-id="4a511-869">Requirements</span></span>

|<span data-ttu-id="4a511-870">要求</span><span class="sxs-lookup"><span data-stu-id="4a511-870">Requirement</span></span>| <span data-ttu-id="4a511-871">值</span><span class="sxs-lookup"><span data-stu-id="4a511-871">Value</span></span>|
|---|---|
|[<span data-ttu-id="4a511-872">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="4a511-872">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4a511-873">1.2</span><span class="sxs-lookup"><span data-stu-id="4a511-873">1.2</span></span>|
|[<span data-ttu-id="4a511-874">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="4a511-874">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4a511-875">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="4a511-875">ReadWriteItem</span></span>|
|[<span data-ttu-id="4a511-876">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="4a511-876">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4a511-877">撰写</span><span class="sxs-lookup"><span data-stu-id="4a511-877">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="4a511-878">返回：</span><span class="sxs-lookup"><span data-stu-id="4a511-878">Returns:</span></span>

<span data-ttu-id="4a511-879">作为字符串的所选数据的格式由 `coercionType` 确定。</span><span class="sxs-lookup"><span data-stu-id="4a511-879">The selected data as a string with format determined by `coercionType`.</span></span>

<dl class="param-type"><span data-ttu-id="4a511-880">

<dt>
类型</dt>


</span><span class="sxs-lookup"><span data-stu-id="4a511-880">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="4a511-881">String</span><span class="sxs-lookup"><span data-stu-id="4a511-881">String</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="4a511-882">示例</span><span class="sxs-lookup"><span data-stu-id="4a511-882">Example</span></span>

```javascript
// Get selected data.
Office.initialize = function () {
  Office.context.mailbox.item.getSelectedDataAsync(Office.CoercionType.Text, {}, getCallback);
};

function getCallback(asyncResult) {
  var text = asyncResult.value.data;
  var prop = asyncResult.value.sourceProperty;

  Office.context.mailbox.item.setSelectedDataAsync('Setting ' + prop + ': ' + text, {}, setCallback);
}

function setCallback(asyncResult) {
  // Check for errors.
}
```

####  <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="4a511-883">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="4a511-883">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="4a511-884">异步加载所选项目上此外接程序的自定义属性。</span><span class="sxs-lookup"><span data-stu-id="4a511-884">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="4a511-p161">自定义属性基于每个应用、每个项目存储为键/值对。此方法在回调中返回 `CustomProperties` 对象，该回调提供访问特定于当前项目和当前外接程序的自定义属性的方法。自定义属性未在项目上加密，因此这不应用作安全存储。</span><span class="sxs-lookup"><span data-stu-id="4a511-p161">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4a511-888">Parameters</span><span class="sxs-lookup"><span data-stu-id="4a511-888">Parameters</span></span>

|<span data-ttu-id="4a511-889">名称</span><span class="sxs-lookup"><span data-stu-id="4a511-889">Name</span></span>| <span data-ttu-id="4a511-890">类型</span><span class="sxs-lookup"><span data-stu-id="4a511-890">Type</span></span>| <span data-ttu-id="4a511-891">属性</span><span class="sxs-lookup"><span data-stu-id="4a511-891">Attributes</span></span>| <span data-ttu-id="4a511-892">说明</span><span class="sxs-lookup"><span data-stu-id="4a511-892">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="4a511-893">函数</span><span class="sxs-lookup"><span data-stu-id="4a511-893">function</span></span>||<span data-ttu-id="4a511-894">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="4a511-894">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="4a511-895">自定义属性作为 `asyncResult.value` 属性中的 [`CustomProperties`](/javascript/api/outlook_1_2/office.customproperties) 对象提供。</span><span class="sxs-lookup"><span data-stu-id="4a511-895">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook_1_2/office.customproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="4a511-896">此对象可用于获取、设置以及从项目中删除自定义属性，并将自定义属性集的更改重新保存到服务器。</span><span class="sxs-lookup"><span data-stu-id="4a511-896">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`| <span data-ttu-id="4a511-897">对象</span><span class="sxs-lookup"><span data-stu-id="4a511-897">Object</span></span>| <span data-ttu-id="4a511-898">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="4a511-898">&lt;optional&gt;</span></span>|<span data-ttu-id="4a511-899">开发人员可以提供他们想要在回调函数中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="4a511-899">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="4a511-900">此对象可以通过回调函数中的 `asyncResult.asyncContext` 属性进行访问。</span><span class="sxs-lookup"><span data-stu-id="4a511-900">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4a511-901">要求</span><span class="sxs-lookup"><span data-stu-id="4a511-901">Requirements</span></span>

|<span data-ttu-id="4a511-902">要求</span><span class="sxs-lookup"><span data-stu-id="4a511-902">Requirement</span></span>| <span data-ttu-id="4a511-903">值</span><span class="sxs-lookup"><span data-stu-id="4a511-903">Value</span></span>|
|---|---|
|[<span data-ttu-id="4a511-904">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="4a511-904">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4a511-905">1.0</span><span class="sxs-lookup"><span data-stu-id="4a511-905">1.0</span></span>|
|[<span data-ttu-id="4a511-906">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="4a511-906">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4a511-907">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4a511-907">ReadItem</span></span>|
|[<span data-ttu-id="4a511-908">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="4a511-908">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4a511-909">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="4a511-909">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4a511-910">示例</span><span class="sxs-lookup"><span data-stu-id="4a511-910">Example</span></span>

<span data-ttu-id="4a511-p164">以下代码示例显示了如何使用 `loadCustomPropertiesAsync` 方法异步加载特定于当前项目的自定义属性。该示例还显示了如何使用 `CustomProperties.saveAsync` 方法将这些属性重新保存到服务器。加载自定义属性后，该代码示例将使用 `CustomProperties.get` 方法读取自定义属性 `myProp`，使用 `CustomProperties.set` 方法写入自定义属性 `otherProp`，最后调用 `saveAsync` 方法保存这些自定义属性。</span><span class="sxs-lookup"><span data-stu-id="4a511-p164">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

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

####  <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="4a511-914">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="4a511-914">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="4a511-915">将附件从邮件或约会中删除。</span><span class="sxs-lookup"><span data-stu-id="4a511-915">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="4a511-p165">`removeAttachmentAsync` 方法删除项目中带指定标识符的附件。最佳做法是，仅当同一个邮件应用程序在同一会话中添加了一个附件时，你才应使用该附件标识符来删除该附件。在 Outlook Web App 和适用于设备的 OWA 中，附件标识符只在同一个会话中才有效。当用户关闭应用程序，或者如果用户开始在内嵌窗体中撰写，并在随后弹出的内嵌窗体中继续在单独的窗口撰写时，会话即结束。</span><span class="sxs-lookup"><span data-stu-id="4a511-p165">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item. As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session. In Outlook Web App and OWA for Devices, the attachment identifier is valid only within the same session. A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4a511-920">Parameters</span><span class="sxs-lookup"><span data-stu-id="4a511-920">Parameters</span></span>

|<span data-ttu-id="4a511-921">名称</span><span class="sxs-lookup"><span data-stu-id="4a511-921">Name</span></span>| <span data-ttu-id="4a511-922">类型</span><span class="sxs-lookup"><span data-stu-id="4a511-922">Type</span></span>| <span data-ttu-id="4a511-923">属性</span><span class="sxs-lookup"><span data-stu-id="4a511-923">Attributes</span></span>| <span data-ttu-id="4a511-924">说明</span><span class="sxs-lookup"><span data-stu-id="4a511-924">Description</span></span>|
|---|---|---|---|
|`attachmentId`| <span data-ttu-id="4a511-925">字符串</span><span class="sxs-lookup"><span data-stu-id="4a511-925">String</span></span>||<span data-ttu-id="4a511-926">要删除的附件的标识符。</span><span class="sxs-lookup"><span data-stu-id="4a511-926">The identifier of the attachment to remove.</span></span>|
|`options`| <span data-ttu-id="4a511-927">对象</span><span class="sxs-lookup"><span data-stu-id="4a511-927">Object</span></span>| <span data-ttu-id="4a511-928">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="4a511-928">&lt;optional&gt;</span></span>|<span data-ttu-id="4a511-929">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="4a511-929">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="4a511-930">对象</span><span class="sxs-lookup"><span data-stu-id="4a511-930">Object</span></span>| <span data-ttu-id="4a511-931">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="4a511-931">&lt;optional&gt;</span></span>|<span data-ttu-id="4a511-932">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="4a511-932">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="4a511-933">函数</span><span class="sxs-lookup"><span data-stu-id="4a511-933">function</span></span>| <span data-ttu-id="4a511-934">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="4a511-934">&lt;optional&gt;</span></span>|<span data-ttu-id="4a511-935">方法完成后，使用单个参数 `callback`（一个 [`asyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `AsyncResult` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="4a511-935">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="4a511-936">如果删除附件失败，`asyncResult.error` 属性将包含一个说明失败原因的错误代码。</span><span class="sxs-lookup"><span data-stu-id="4a511-936">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="4a511-937">错误</span><span class="sxs-lookup"><span data-stu-id="4a511-937">Errors</span></span>

| <span data-ttu-id="4a511-938">错误代码</span><span class="sxs-lookup"><span data-stu-id="4a511-938">Error code</span></span> | <span data-ttu-id="4a511-939">说明</span><span class="sxs-lookup"><span data-stu-id="4a511-939">Description</span></span> |
|------------|-------------|
| `InvalidAttachmentId` | <span data-ttu-id="4a511-940">附件标识符不存在。</span><span class="sxs-lookup"><span data-stu-id="4a511-940">The attachment identifier does not exist.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="4a511-941">要求</span><span class="sxs-lookup"><span data-stu-id="4a511-941">Requirements</span></span>

|<span data-ttu-id="4a511-942">要求</span><span class="sxs-lookup"><span data-stu-id="4a511-942">Requirement</span></span>| <span data-ttu-id="4a511-943">值</span><span class="sxs-lookup"><span data-stu-id="4a511-943">Value</span></span>|
|---|---|
|[<span data-ttu-id="4a511-944">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="4a511-944">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4a511-945">1.1</span><span class="sxs-lookup"><span data-stu-id="4a511-945">1.1</span></span>|
|[<span data-ttu-id="4a511-946">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="4a511-946">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4a511-947">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="4a511-947">ReadWriteItem</span></span>|
|[<span data-ttu-id="4a511-948">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="4a511-948">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4a511-949">撰写</span><span class="sxs-lookup"><span data-stu-id="4a511-949">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="4a511-950">示例</span><span class="sxs-lookup"><span data-stu-id="4a511-950">Example</span></span>

<span data-ttu-id="4a511-951">以下代码删除包含标识符 '0' 的附件。</span><span class="sxs-lookup"><span data-stu-id="4a511-951">The following code removes an attachment with an identifier of '0'.</span></span>

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

####  <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="4a511-952">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="4a511-952">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="4a511-953">以异步方式将数据插入到邮件的正文或主题中。</span><span class="sxs-lookup"><span data-stu-id="4a511-953">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="4a511-p166">`setSelectedDataAsync` 方法将指定字符串插入到项目主题或正文的光标位置，或者，如果在编辑器中已选择文本，则该方法将替换选择的文本。如果光标不在正文或主题字段中，则返回错误。插入之后，光标会位于插入内容的末尾。</span><span class="sxs-lookup"><span data-stu-id="4a511-p166">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4a511-957">Parameters</span><span class="sxs-lookup"><span data-stu-id="4a511-957">Parameters</span></span>

|<span data-ttu-id="4a511-958">名称</span><span class="sxs-lookup"><span data-stu-id="4a511-958">Name</span></span>| <span data-ttu-id="4a511-959">类型</span><span class="sxs-lookup"><span data-stu-id="4a511-959">Type</span></span>| <span data-ttu-id="4a511-960">属性</span><span class="sxs-lookup"><span data-stu-id="4a511-960">Attributes</span></span>| <span data-ttu-id="4a511-961">说明</span><span class="sxs-lookup"><span data-stu-id="4a511-961">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="4a511-962">字符串</span><span class="sxs-lookup"><span data-stu-id="4a511-962">String</span></span>||<span data-ttu-id="4a511-p167">要插入的数据。数据不得超过 1,000,000 个字符。如果传入的数据超过 1,000,000 个字符，则会引发 `ArgumentOutOfRange` 异常。</span><span class="sxs-lookup"><span data-stu-id="4a511-p167">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`| <span data-ttu-id="4a511-966">Object</span><span class="sxs-lookup"><span data-stu-id="4a511-966">Object</span></span>| <span data-ttu-id="4a511-967">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="4a511-967">&lt;optional&gt;</span></span>|<span data-ttu-id="4a511-968">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="4a511-968">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="4a511-969">对象</span><span class="sxs-lookup"><span data-stu-id="4a511-969">Object</span></span>| <span data-ttu-id="4a511-970">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="4a511-970">&lt;optional&gt;</span></span>|<span data-ttu-id="4a511-971">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="4a511-971">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`| [<span data-ttu-id="4a511-972">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="4a511-972">Office.CoercionType</span></span>](office.md#coerciontype-string)| <span data-ttu-id="4a511-973">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="4a511-973">&lt;optional&gt;</span></span>|<span data-ttu-id="4a511-p168">如果为 `text`，则在 Outlook Web App 和 Outlook 中应用当前样式。如果该字段是 HTML 编辑器，则仅插入文本数据，即使数据为 HTML。</span><span class="sxs-lookup"><span data-stu-id="4a511-p168">If `text`, the current style is applied in Outlook Web App and Outlook. If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="4a511-p169">如果 `html` 和该字段支持 HTML（主题不支持），则在 Outlook Web App 中应用当前样式，而在 Outlook 中应用默认样式。如果该字段是文本字段，则返回 `InvalidDataFormat` 错误。</span><span class="sxs-lookup"><span data-stu-id="4a511-p169">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook Web App and the default style is applied in Outlook. If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="4a511-978">如果未设置 `coercionType`，则结果取决于该字段：如果该字段是 HTML，则使用 HTML；如果该字段是文本，则使用纯文本。</span><span class="sxs-lookup"><span data-stu-id="4a511-978">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`| <span data-ttu-id="4a511-979">function</span><span class="sxs-lookup"><span data-stu-id="4a511-979">function</span></span>||<span data-ttu-id="4a511-980">方法完成后，使用单个参数 `callback`（一个 [`asyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `AsyncResult` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="4a511-980">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="4a511-981">Requirements</span><span class="sxs-lookup"><span data-stu-id="4a511-981">Requirements</span></span>

|<span data-ttu-id="4a511-982">要求</span><span class="sxs-lookup"><span data-stu-id="4a511-982">Requirement</span></span>| <span data-ttu-id="4a511-983">值</span><span class="sxs-lookup"><span data-stu-id="4a511-983">Value</span></span>|
|---|---|
|[<span data-ttu-id="4a511-984">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="4a511-984">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4a511-985">1.2</span><span class="sxs-lookup"><span data-stu-id="4a511-985">1.2</span></span>|
|[<span data-ttu-id="4a511-986">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="4a511-986">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4a511-987">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="4a511-987">ReadWriteItem</span></span>|
|[<span data-ttu-id="4a511-988">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="4a511-988">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4a511-989">撰写</span><span class="sxs-lookup"><span data-stu-id="4a511-989">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="4a511-990">示例</span><span class="sxs-lookup"><span data-stu-id="4a511-990">Example</span></span>

```javascript
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```
