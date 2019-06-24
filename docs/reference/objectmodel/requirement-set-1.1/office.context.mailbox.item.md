---
title: "\"Context\"-\"邮箱\"。项目-要求集1。1"
description: ''
ms.date: 06/20/2019
localization_priority: Normal
ms.openlocfilehash: 09956256f856db6a8bea3053ecb64a2ac2a5fcd9
ms.sourcegitcommit: 382e2735a1295da914f2bfc38883e518070cec61
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 06/21/2019
ms.locfileid: "35128466"
---
# <a name="item"></a><span data-ttu-id="02660-102">item</span><span class="sxs-lookup"><span data-stu-id="02660-102">item</span></span>

### <span data-ttu-id="02660-p101">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md). 项目</span><span class="sxs-lookup"><span data-stu-id="02660-p101">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md). item</span></span>

<span data-ttu-id="02660-p102">`item` 命名空间用于访问当前选定的邮件、会议请求或约会。可以通过使用 [itemType](#itemtype-officemailboxenumsitemtype) 属性确定 `item` 的类型。</span><span class="sxs-lookup"><span data-stu-id="02660-p102">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="02660-107">Requirements</span><span class="sxs-lookup"><span data-stu-id="02660-107">Requirements</span></span>

|<span data-ttu-id="02660-108">要求</span><span class="sxs-lookup"><span data-stu-id="02660-108">Requirement</span></span>| <span data-ttu-id="02660-109">值</span><span class="sxs-lookup"><span data-stu-id="02660-109">Value</span></span>|
|---|---|
|[<span data-ttu-id="02660-110">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="02660-110">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="02660-111">1.0</span><span class="sxs-lookup"><span data-stu-id="02660-111">1.0</span></span>|
|[<span data-ttu-id="02660-112">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="02660-112">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="02660-113">受限</span><span class="sxs-lookup"><span data-stu-id="02660-113">Restricted</span></span>|
|[<span data-ttu-id="02660-114">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="02660-114">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="02660-115">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="02660-115">Compose or Read</span></span>|

### <a name="example"></a><span data-ttu-id="02660-116">示例</span><span class="sxs-lookup"><span data-stu-id="02660-116">Example</span></span>

<span data-ttu-id="02660-117">以下 JavaScript 代码示例显示了如何访问 Outlook 中当前项目的 `subject` 属性。</span><span class="sxs-lookup"><span data-stu-id="02660-117">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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

### <a name="members"></a><span data-ttu-id="02660-118">成员</span><span class="sxs-lookup"><span data-stu-id="02660-118">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlook11officeattachmentdetails"></a><span data-ttu-id="02660-119">附件: Array. <[AttachmentDetails](/javascript/api/outlook_1_1/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="02660-119">attachments: Array.<[AttachmentDetails](/javascript/api/outlook_1_1/office.attachmentdetails)></span></span>

<span data-ttu-id="02660-p103">获取项目的附件数组。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="02660-p103">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="02660-122">由于潜在的安全问题，某些类型的文件会受到 Outlook 阻止，并且不会返回。</span><span class="sxs-lookup"><span data-stu-id="02660-122">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="02660-123">如需了解更多信息，请参阅 [Outlook 中阻止的附件](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519)。</span><span class="sxs-lookup"><span data-stu-id="02660-123">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="02660-124">类型</span><span class="sxs-lookup"><span data-stu-id="02660-124">Type</span></span>

*   <span data-ttu-id="02660-125">Array.<[AttachmentDetails](/javascript/api/outlook_1_1/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="02660-125">Array.<[AttachmentDetails](/javascript/api/outlook_1_1/office.attachmentdetails)></span></span>

##### <a name="requirements"></a><span data-ttu-id="02660-126">要求</span><span class="sxs-lookup"><span data-stu-id="02660-126">Requirements</span></span>

|<span data-ttu-id="02660-127">要求</span><span class="sxs-lookup"><span data-stu-id="02660-127">Requirement</span></span>| <span data-ttu-id="02660-128">值</span><span class="sxs-lookup"><span data-stu-id="02660-128">Value</span></span>|
|---|---|
|[<span data-ttu-id="02660-129">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="02660-129">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="02660-130">1.0</span><span class="sxs-lookup"><span data-stu-id="02660-130">1.0</span></span>|
|[<span data-ttu-id="02660-131">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="02660-131">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="02660-132">ReadItem</span><span class="sxs-lookup"><span data-stu-id="02660-132">ReadItem</span></span>|
|[<span data-ttu-id="02660-133">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="02660-133">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="02660-134">阅读</span><span class="sxs-lookup"><span data-stu-id="02660-134">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="02660-135">示例</span><span class="sxs-lookup"><span data-stu-id="02660-135">Example</span></span>

<span data-ttu-id="02660-136">以下代码使用当前项目上所有附件的详细信息构成 HTML 字符串。</span><span class="sxs-lookup"><span data-stu-id="02660-136">The following code builds an HTML string with details of all attachments on the current item.</span></span>

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

#### <a name="bcc-recipientsjavascriptapioutlook11officerecipients"></a><span data-ttu-id="02660-137">密件抄送:[收件人](/javascript/api/outlook_1_1/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="02660-137">bcc: [Recipients](/javascript/api/outlook_1_1/office.recipients)</span></span>

<span data-ttu-id="02660-138">获取对象，该对象提供用于获取或更新邮件的密件抄送 (Bcc) 行上的收件人的方法。</span><span class="sxs-lookup"><span data-stu-id="02660-138">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="02660-139">仅限撰写模式。</span><span class="sxs-lookup"><span data-stu-id="02660-139">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="02660-140">类型</span><span class="sxs-lookup"><span data-stu-id="02660-140">Type</span></span>

*   [<span data-ttu-id="02660-141">收件人</span><span class="sxs-lookup"><span data-stu-id="02660-141">Recipients</span></span>](/javascript/api/outlook_1_1/office.recipients)

##### <a name="requirements"></a><span data-ttu-id="02660-142">要求</span><span class="sxs-lookup"><span data-stu-id="02660-142">Requirements</span></span>

|<span data-ttu-id="02660-143">要求</span><span class="sxs-lookup"><span data-stu-id="02660-143">Requirement</span></span>| <span data-ttu-id="02660-144">值</span><span class="sxs-lookup"><span data-stu-id="02660-144">Value</span></span>|
|---|---|
|[<span data-ttu-id="02660-145">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="02660-145">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="02660-146">1.1</span><span class="sxs-lookup"><span data-stu-id="02660-146">1.1</span></span>|
|[<span data-ttu-id="02660-147">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="02660-147">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="02660-148">ReadItem</span><span class="sxs-lookup"><span data-stu-id="02660-148">ReadItem</span></span>|
|[<span data-ttu-id="02660-149">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="02660-149">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="02660-150">撰写</span><span class="sxs-lookup"><span data-stu-id="02660-150">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="02660-151">示例</span><span class="sxs-lookup"><span data-stu-id="02660-151">Example</span></span>

```javascript
Office.context.mailbox.item.bcc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.bcc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.bcc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfBccRecipients = asyncResult.value;
}
```

#### <a name="body-bodyjavascriptapioutlook11officebody"></a><span data-ttu-id="02660-152">正文:[正文](/javascript/api/outlook_1_1/office.body)</span><span class="sxs-lookup"><span data-stu-id="02660-152">body: [Body](/javascript/api/outlook_1_1/office.body)</span></span>

<span data-ttu-id="02660-153">获取一个提供用于处理项目正文的方法的对象。</span><span class="sxs-lookup"><span data-stu-id="02660-153">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="02660-154">类型</span><span class="sxs-lookup"><span data-stu-id="02660-154">Type</span></span>

*   [<span data-ttu-id="02660-155">Body</span><span class="sxs-lookup"><span data-stu-id="02660-155">Body</span></span>](/javascript/api/outlook_1_1/office.body)

##### <a name="requirements"></a><span data-ttu-id="02660-156">要求</span><span class="sxs-lookup"><span data-stu-id="02660-156">Requirements</span></span>

|<span data-ttu-id="02660-157">要求</span><span class="sxs-lookup"><span data-stu-id="02660-157">Requirement</span></span>| <span data-ttu-id="02660-158">值</span><span class="sxs-lookup"><span data-stu-id="02660-158">Value</span></span>|
|---|---|
|[<span data-ttu-id="02660-159">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="02660-159">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="02660-160">1.1</span><span class="sxs-lookup"><span data-stu-id="02660-160">1.1</span></span>|
|[<span data-ttu-id="02660-161">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="02660-161">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="02660-162">ReadItem</span><span class="sxs-lookup"><span data-stu-id="02660-162">ReadItem</span></span>|
|[<span data-ttu-id="02660-163">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="02660-163">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="02660-164">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="02660-164">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="02660-165">示例</span><span class="sxs-lookup"><span data-stu-id="02660-165">Example</span></span>

<span data-ttu-id="02660-166">本示例获取纯文本格式的邮件正文。</span><span class="sxs-lookup"><span data-stu-id="02660-166">This example gets the body of the message in plain text.</span></span>

```javascript
Office.context.mailbox.item.body.getAsync(
  "text",
  { asyncContext: "This is passed to the callback" },
  function callback(result) {
    // Do something with the result.
  });

```

<span data-ttu-id="02660-167">以下是传递到回调函数的结果参数的示例。</span><span class="sxs-lookup"><span data-stu-id="02660-167">The following is an example of the result parameter passed to the callback function.</span></span>

```json
{
  "value": "TEXT of whole body (including threads below)",
  "status": "succeeded",
  "asyncContext": "This is passed to the callback"
}
```

#### <a name="cc-arrayemailaddressdetailsjavascriptapioutlook11officeemailaddressdetailsrecipientsjavascriptapioutlook11officerecipients"></a><span data-ttu-id="02660-168"><[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)>|[收件人](/javascript/api/outlook_1_1/office.recipients)的抄送: Array</span><span class="sxs-lookup"><span data-stu-id="02660-168">cc: Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_1/office.recipients)</span></span>

<span data-ttu-id="02660-169">提供对邮件的抄送 (Cc) 收件人的访问权限。</span><span class="sxs-lookup"><span data-stu-id="02660-169">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="02660-170">对象的类型和访问级别取决于当前项目的模式。</span><span class="sxs-lookup"><span data-stu-id="02660-170">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="02660-171">阅读模式</span><span class="sxs-lookup"><span data-stu-id="02660-171">Read mode</span></span>

<span data-ttu-id="02660-p107">`cc` 属性返回包含邮件的**抄送**行上所列的每个收件人的 `EmailAddressDetails` 对象的数组。集合上限为 100 个成员。</span><span class="sxs-lookup"><span data-stu-id="02660-p107">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.</span></span>

```javascript
console.log(JSON.stringify(Office.context.mailbox.item.cc));
```

##### <a name="compose-mode"></a><span data-ttu-id="02660-174">撰写模式</span><span class="sxs-lookup"><span data-stu-id="02660-174">Compose mode</span></span>

<span data-ttu-id="02660-175">`cc` 属性返回一个 `Recipients` 对象，该对象提供用于获取或更新邮件的“**抄送**”行上收件人的方法。</span><span class="sxs-lookup"><span data-stu-id="02660-175">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span>

```javascript
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="02660-176">类型</span><span class="sxs-lookup"><span data-stu-id="02660-176">Type</span></span>

*   <span data-ttu-id="02660-177">Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_1/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="02660-177">Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_1/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="02660-178">要求</span><span class="sxs-lookup"><span data-stu-id="02660-178">Requirements</span></span>

|<span data-ttu-id="02660-179">要求</span><span class="sxs-lookup"><span data-stu-id="02660-179">Requirement</span></span>| <span data-ttu-id="02660-180">值</span><span class="sxs-lookup"><span data-stu-id="02660-180">Value</span></span>|
|---|---|
|[<span data-ttu-id="02660-181">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="02660-181">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="02660-182">1.0</span><span class="sxs-lookup"><span data-stu-id="02660-182">1.0</span></span>|
|[<span data-ttu-id="02660-183">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="02660-183">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="02660-184">ReadItem</span><span class="sxs-lookup"><span data-stu-id="02660-184">ReadItem</span></span>|
|[<span data-ttu-id="02660-185">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="02660-185">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="02660-186">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="02660-186">Compose or Read</span></span>|

#### <a name="nullable-conversationid-string"></a><span data-ttu-id="02660-187">(可以为 null) conversationId: String</span><span class="sxs-lookup"><span data-stu-id="02660-187">(nullable) conversationId: String</span></span>

<span data-ttu-id="02660-188">获取包含特定消息的电子邮件会话的标识符。</span><span class="sxs-lookup"><span data-stu-id="02660-188">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="02660-p108">如果在阅读窗体或撰写窗体的回复中激活邮件应用程序，则此属性可以获得一个整数值。如果用户随后更改了回复邮件的主题（若发送回复），则该邮件的对话 ID 将改变且之前获取的值将不适用。</span><span class="sxs-lookup"><span data-stu-id="02660-p108">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="02660-p109">对于撰写窗体的新项目，此属性获得一个 null 值。如果用户设置一个主题并保存该项目，`conversationId` 属性将返回一个值。</span><span class="sxs-lookup"><span data-stu-id="02660-p109">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="02660-193">Type</span><span class="sxs-lookup"><span data-stu-id="02660-193">Type</span></span>

*   <span data-ttu-id="02660-194">String</span><span class="sxs-lookup"><span data-stu-id="02660-194">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="02660-195">要求</span><span class="sxs-lookup"><span data-stu-id="02660-195">Requirements</span></span>

|<span data-ttu-id="02660-196">要求</span><span class="sxs-lookup"><span data-stu-id="02660-196">Requirement</span></span>| <span data-ttu-id="02660-197">值</span><span class="sxs-lookup"><span data-stu-id="02660-197">Value</span></span>|
|---|---|
|[<span data-ttu-id="02660-198">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="02660-198">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="02660-199">1.0</span><span class="sxs-lookup"><span data-stu-id="02660-199">1.0</span></span>|
|[<span data-ttu-id="02660-200">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="02660-200">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="02660-201">ReadItem</span><span class="sxs-lookup"><span data-stu-id="02660-201">ReadItem</span></span>|
|[<span data-ttu-id="02660-202">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="02660-202">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="02660-203">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="02660-203">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="02660-204">示例</span><span class="sxs-lookup"><span data-stu-id="02660-204">Example</span></span>

```javascript
var conversationId = Office.context.mailbox.item.conversationId;
console.log("conversationId: " + conversationId);
```

#### <a name="datetimecreated-date"></a><span data-ttu-id="02660-205">dateTimeCreated: Date</span><span class="sxs-lookup"><span data-stu-id="02660-205">dateTimeCreated: Date</span></span>

<span data-ttu-id="02660-p110">获取项目创建的日期和时间。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="02660-p110">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="02660-208">类型</span><span class="sxs-lookup"><span data-stu-id="02660-208">Type</span></span>

*   <span data-ttu-id="02660-209">日期</span><span class="sxs-lookup"><span data-stu-id="02660-209">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="02660-210">要求</span><span class="sxs-lookup"><span data-stu-id="02660-210">Requirements</span></span>

|<span data-ttu-id="02660-211">要求</span><span class="sxs-lookup"><span data-stu-id="02660-211">Requirement</span></span>| <span data-ttu-id="02660-212">值</span><span class="sxs-lookup"><span data-stu-id="02660-212">Value</span></span>|
|---|---|
|[<span data-ttu-id="02660-213">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="02660-213">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="02660-214">1.0</span><span class="sxs-lookup"><span data-stu-id="02660-214">1.0</span></span>|
|[<span data-ttu-id="02660-215">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="02660-215">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="02660-216">ReadItem</span><span class="sxs-lookup"><span data-stu-id="02660-216">ReadItem</span></span>|
|[<span data-ttu-id="02660-217">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="02660-217">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="02660-218">阅读</span><span class="sxs-lookup"><span data-stu-id="02660-218">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="02660-219">示例</span><span class="sxs-lookup"><span data-stu-id="02660-219">Example</span></span>

```javascript
var dateTimeCreated = Office.context.mailbox.item.dateTimeCreated;
console.log("Date and time created: " + dateTimeCreated);
```

#### <a name="datetimemodified-date"></a><span data-ttu-id="02660-220">dateTimeModified: Date</span><span class="sxs-lookup"><span data-stu-id="02660-220">dateTimeModified: Date</span></span>

<span data-ttu-id="02660-221">获取项目最近一次修改的日期和时间。</span><span class="sxs-lookup"><span data-stu-id="02660-221">Gets the date and time that an item was last modified.</span></span> <span data-ttu-id="02660-222">仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="02660-222">Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="02660-223">IOS 或 Android 上的 Outlook 不支持此成员。</span><span class="sxs-lookup"><span data-stu-id="02660-223">This member is not supported in Outlook on iOS or Android.</span></span>

##### <a name="type"></a><span data-ttu-id="02660-224">类型</span><span class="sxs-lookup"><span data-stu-id="02660-224">Type</span></span>

*   <span data-ttu-id="02660-225">日期</span><span class="sxs-lookup"><span data-stu-id="02660-225">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="02660-226">要求</span><span class="sxs-lookup"><span data-stu-id="02660-226">Requirements</span></span>

|<span data-ttu-id="02660-227">要求</span><span class="sxs-lookup"><span data-stu-id="02660-227">Requirement</span></span>| <span data-ttu-id="02660-228">值</span><span class="sxs-lookup"><span data-stu-id="02660-228">Value</span></span>|
|---|---|
|[<span data-ttu-id="02660-229">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="02660-229">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="02660-230">1.0</span><span class="sxs-lookup"><span data-stu-id="02660-230">1.0</span></span>|
|[<span data-ttu-id="02660-231">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="02660-231">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="02660-232">ReadItem</span><span class="sxs-lookup"><span data-stu-id="02660-232">ReadItem</span></span>|
|[<span data-ttu-id="02660-233">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="02660-233">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="02660-234">阅读</span><span class="sxs-lookup"><span data-stu-id="02660-234">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="02660-235">示例</span><span class="sxs-lookup"><span data-stu-id="02660-235">Example</span></span>

```javascript
var dateTimeModified = Office.context.mailbox.item.dateTimeModified;
console.log("Date and time modified: " + dateTimeModified);
```

#### <a name="end-datetimejavascriptapioutlook11officetime"></a><span data-ttu-id="02660-236">结束: 日期 |[时间](/javascript/api/outlook_1_1/office.time)</span><span class="sxs-lookup"><span data-stu-id="02660-236">end: Date|[Time](/javascript/api/outlook_1_1/office.time)</span></span>

<span data-ttu-id="02660-237">获取或设置约会结束的日期和时间。</span><span class="sxs-lookup"><span data-stu-id="02660-237">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="02660-p112">将 `end` 属性表示为协调世界时 (UTC) 的日期和时间值。可使用 [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) 方法将 end 属性值转换为客户端的本地日期和时间。</span><span class="sxs-lookup"><span data-stu-id="02660-p112">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="02660-240">阅读模式</span><span class="sxs-lookup"><span data-stu-id="02660-240">Read mode</span></span>

<span data-ttu-id="02660-241">`end` 属性返回 `Date` 对象。</span><span class="sxs-lookup"><span data-stu-id="02660-241">The `end` property returns a `Date` object.</span></span>

```javascript
var end = Office.context.mailbox.item.end;
console.log("Appointment end: " + end);
```

##### <a name="compose-mode"></a><span data-ttu-id="02660-242">撰写模式</span><span class="sxs-lookup"><span data-stu-id="02660-242">Compose mode</span></span>

<span data-ttu-id="02660-243">`end` 属性返回 `Time` 对象。</span><span class="sxs-lookup"><span data-stu-id="02660-243">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="02660-244">使用 [`Time.setAsync`](/javascript/api/outlook_1_1/office.time#setasync-datetime--options--callback-) 方法设置结束时间时，应使用 [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) 方法将客户端的本地时间转换为服务器的 UTC。</span><span class="sxs-lookup"><span data-stu-id="02660-244">When you use the [`Time.setAsync`](/javascript/api/outlook_1_1/office.time#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="02660-245">以下示例使用 `Time` 对象的 [`setAsync`](/javascript/api/outlook_1_1/office.time#setasync-datetime--options--callback-) 方法设置约会的结束时间。</span><span class="sxs-lookup"><span data-stu-id="02660-245">The following example sets the end time of an appointment by using the [`setAsync`](/javascript/api/outlook_1_1/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="02660-246">类型</span><span class="sxs-lookup"><span data-stu-id="02660-246">Type</span></span>

*   <span data-ttu-id="02660-247">Date | [Time](/javascript/api/outlook_1_1/office.time)</span><span class="sxs-lookup"><span data-stu-id="02660-247">Date | [Time](/javascript/api/outlook_1_1/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="02660-248">要求</span><span class="sxs-lookup"><span data-stu-id="02660-248">Requirements</span></span>

|<span data-ttu-id="02660-249">要求</span><span class="sxs-lookup"><span data-stu-id="02660-249">Requirement</span></span>| <span data-ttu-id="02660-250">值</span><span class="sxs-lookup"><span data-stu-id="02660-250">Value</span></span>|
|---|---|
|[<span data-ttu-id="02660-251">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="02660-251">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="02660-252">1.0</span><span class="sxs-lookup"><span data-stu-id="02660-252">1.0</span></span>|
|[<span data-ttu-id="02660-253">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="02660-253">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="02660-254">ReadItem</span><span class="sxs-lookup"><span data-stu-id="02660-254">ReadItem</span></span>|
|[<span data-ttu-id="02660-255">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="02660-255">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="02660-256">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="02660-256">Compose or Read</span></span>|

#### <a name="from-emailaddressdetailsjavascriptapioutlook11officeemailaddressdetails"></a><span data-ttu-id="02660-257">发件人: [EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="02660-257">from: [EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)</span></span>

<span data-ttu-id="02660-p113">获取邮件发件人的电子邮件地址。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="02660-p113">Gets the email address of the sender of a message. Read mode only.</span></span>

<span data-ttu-id="02660-p114">`from` 和 [`sender`](#sender-emailaddressdetails) 属性表示同一个人，邮件由代理人发送的除外。在此情况下，`from` 属性表示代理程序，而 sender 属性表示代理人。</span><span class="sxs-lookup"><span data-stu-id="02660-p114">The `from` and [`sender`](#sender-emailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="02660-262">`from` 属性中 `EmailAddressDetails` 对象的 `recipientType` 属性为 `undefined`。</span><span class="sxs-lookup"><span data-stu-id="02660-262">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="02660-263">类型</span><span class="sxs-lookup"><span data-stu-id="02660-263">Type</span></span>

*   [<span data-ttu-id="02660-264">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="02660-264">EmailAddressDetails</span></span>](/javascript/api/outlook_1_1/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="02660-265">要求</span><span class="sxs-lookup"><span data-stu-id="02660-265">Requirements</span></span>

|<span data-ttu-id="02660-266">要求</span><span class="sxs-lookup"><span data-stu-id="02660-266">Requirement</span></span>| <span data-ttu-id="02660-267">值</span><span class="sxs-lookup"><span data-stu-id="02660-267">Value</span></span>|
|---|---|
|[<span data-ttu-id="02660-268">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="02660-268">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="02660-269">1.0</span><span class="sxs-lookup"><span data-stu-id="02660-269">1.0</span></span>|
|[<span data-ttu-id="02660-270">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="02660-270">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="02660-271">ReadItem</span><span class="sxs-lookup"><span data-stu-id="02660-271">ReadItem</span></span>|
|[<span data-ttu-id="02660-272">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="02660-272">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="02660-273">阅读</span><span class="sxs-lookup"><span data-stu-id="02660-273">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="02660-274">示例</span><span class="sxs-lookup"><span data-stu-id="02660-274">Example</span></span>

```javascript
var from = Office.context.mailbox.item.from;
console.log("From " + from);
```

#### <a name="internetmessageid-string"></a><span data-ttu-id="02660-275">internetMessageId: String</span><span class="sxs-lookup"><span data-stu-id="02660-275">internetMessageId: String</span></span>

<span data-ttu-id="02660-p115">获取电子邮件的 Internet 消息标识符。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="02660-p115">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="02660-278">Type</span><span class="sxs-lookup"><span data-stu-id="02660-278">Type</span></span>

*   <span data-ttu-id="02660-279">String</span><span class="sxs-lookup"><span data-stu-id="02660-279">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="02660-280">要求</span><span class="sxs-lookup"><span data-stu-id="02660-280">Requirements</span></span>

|<span data-ttu-id="02660-281">要求</span><span class="sxs-lookup"><span data-stu-id="02660-281">Requirement</span></span>| <span data-ttu-id="02660-282">值</span><span class="sxs-lookup"><span data-stu-id="02660-282">Value</span></span>|
|---|---|
|[<span data-ttu-id="02660-283">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="02660-283">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="02660-284">1.0</span><span class="sxs-lookup"><span data-stu-id="02660-284">1.0</span></span>|
|[<span data-ttu-id="02660-285">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="02660-285">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="02660-286">ReadItem</span><span class="sxs-lookup"><span data-stu-id="02660-286">ReadItem</span></span>|
|[<span data-ttu-id="02660-287">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="02660-287">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="02660-288">阅读</span><span class="sxs-lookup"><span data-stu-id="02660-288">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="02660-289">示例</span><span class="sxs-lookup"><span data-stu-id="02660-289">Example</span></span>

```javascript
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

#### <a name="itemclass-string"></a><span data-ttu-id="02660-290">itemClass: String</span><span class="sxs-lookup"><span data-stu-id="02660-290">itemClass: String</span></span>

<span data-ttu-id="02660-p116">获取选定项目的 Exchange Web 服务项目类。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="02660-p116">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="02660-p117">`itemClass` 属性指定所选项目的邮件类别。以下是邮件或约会项目的默认邮件类别。</span><span class="sxs-lookup"><span data-stu-id="02660-p117">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

| <span data-ttu-id="02660-295">类型</span><span class="sxs-lookup"><span data-stu-id="02660-295">Type</span></span> | <span data-ttu-id="02660-296">说明</span><span class="sxs-lookup"><span data-stu-id="02660-296">Description</span></span> | <span data-ttu-id="02660-297">项目类</span><span class="sxs-lookup"><span data-stu-id="02660-297">item class</span></span> |
| --- | --- | --- |
| <span data-ttu-id="02660-298">约会项目</span><span class="sxs-lookup"><span data-stu-id="02660-298">Appointment items</span></span> | <span data-ttu-id="02660-299">这些是项目类 `IPM.Appointment` 或 `IPM.Appointment.Occurrence` 的日历项目。</span><span class="sxs-lookup"><span data-stu-id="02660-299">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurrence`.</span></span> | `IPM.Appointment`<br />`IPM.Appointment.Occurrence` |
| <span data-ttu-id="02660-300">邮件项目</span><span class="sxs-lookup"><span data-stu-id="02660-300">Message items</span></span> | <span data-ttu-id="02660-301">这些项目包括具有默认邮件类别 `IPM.Note` 的电子邮件，以及使用 `IPM.Schedule.Meeting` 作为基础邮件类别的会议请求、响应和取消。</span><span class="sxs-lookup"><span data-stu-id="02660-301">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span> | `IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled` |

<span data-ttu-id="02660-302">你可以创建用于扩展默认邮件类别的自定义邮件类别，例如，自定义约会邮件类别 `IPM.Appointment.Contoso`。</span><span class="sxs-lookup"><span data-stu-id="02660-302">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="02660-303">类型</span><span class="sxs-lookup"><span data-stu-id="02660-303">Type</span></span>

*   <span data-ttu-id="02660-304">String</span><span class="sxs-lookup"><span data-stu-id="02660-304">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="02660-305">要求</span><span class="sxs-lookup"><span data-stu-id="02660-305">Requirements</span></span>

|<span data-ttu-id="02660-306">要求</span><span class="sxs-lookup"><span data-stu-id="02660-306">Requirement</span></span>| <span data-ttu-id="02660-307">值</span><span class="sxs-lookup"><span data-stu-id="02660-307">Value</span></span>|
|---|---|
|[<span data-ttu-id="02660-308">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="02660-308">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="02660-309">1.0</span><span class="sxs-lookup"><span data-stu-id="02660-309">1.0</span></span>|
|[<span data-ttu-id="02660-310">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="02660-310">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="02660-311">ReadItem</span><span class="sxs-lookup"><span data-stu-id="02660-311">ReadItem</span></span>|
|[<span data-ttu-id="02660-312">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="02660-312">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="02660-313">阅读</span><span class="sxs-lookup"><span data-stu-id="02660-313">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="02660-314">示例</span><span class="sxs-lookup"><span data-stu-id="02660-314">Example</span></span>

```javascript
var itemClass = Office.context.mailbox.item.itemClass;
console.log("Item class: " + itemClass);
```

#### <a name="nullable-itemid-string"></a><span data-ttu-id="02660-315">(可以为 null) itemId: String</span><span class="sxs-lookup"><span data-stu-id="02660-315">(nullable) itemId: String</span></span>

<span data-ttu-id="02660-316">获取当前项目的 Exchange Web 服务项目标识符。</span><span class="sxs-lookup"><span data-stu-id="02660-316">Gets the Exchange Web Services item identifier for the current item.</span></span> <span data-ttu-id="02660-317">仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="02660-317">Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="02660-318">`itemId` 属性返回的标识符与 Exchange Web 服务项目标识符相同。</span><span class="sxs-lookup"><span data-stu-id="02660-318">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="02660-319">`itemId` 属性与 Outlook 条目 ID 或 Outlook REST API 使用的 ID 不同。</span><span class="sxs-lookup"><span data-stu-id="02660-319">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="02660-320">在使用此值进行 REST API 调用之前, 应使用`Office.context.mailbox.convertToRestId`转换它, 这可从要求集1.3 中开始。</span><span class="sxs-lookup"><span data-stu-id="02660-320">Before making REST API calls using this value, it should be converted using `Office.context.mailbox.convertToRestId`, which is available starting in requirement set 1.3.</span></span> <span data-ttu-id="02660-321">有关详细信息，请参阅[从 Outlook 加载项使用 Outlook REST API](/outlook/add-ins/use-rest-api#get-the-item-id)。</span><span class="sxs-lookup"><span data-stu-id="02660-321">For more details, see [Use the Outlook REST APIs from an Outlook add-in](/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

##### <a name="type"></a><span data-ttu-id="02660-322">Type</span><span class="sxs-lookup"><span data-stu-id="02660-322">Type</span></span>

*   <span data-ttu-id="02660-323">String</span><span class="sxs-lookup"><span data-stu-id="02660-323">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="02660-324">要求</span><span class="sxs-lookup"><span data-stu-id="02660-324">Requirements</span></span>

|<span data-ttu-id="02660-325">要求</span><span class="sxs-lookup"><span data-stu-id="02660-325">Requirement</span></span>| <span data-ttu-id="02660-326">值</span><span class="sxs-lookup"><span data-stu-id="02660-326">Value</span></span>|
|---|---|
|[<span data-ttu-id="02660-327">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="02660-327">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="02660-328">1.0</span><span class="sxs-lookup"><span data-stu-id="02660-328">1.0</span></span>|
|[<span data-ttu-id="02660-329">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="02660-329">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="02660-330">ReadItem</span><span class="sxs-lookup"><span data-stu-id="02660-330">ReadItem</span></span>|
|[<span data-ttu-id="02660-331">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="02660-331">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="02660-332">阅读</span><span class="sxs-lookup"><span data-stu-id="02660-332">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="02660-333">示例</span><span class="sxs-lookup"><span data-stu-id="02660-333">Example</span></span>

<span data-ttu-id="02660-p120">以下代码检查项目标识符是否存在。如果 `itemId` 属性返回 `null` 或 `undefined`，则将项目保存到存储，并从异步结果中获取项目标识符。</span><span class="sxs-lookup"><span data-stu-id="02660-p120">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

```javascript
var itemId = Office.context.mailbox.item.itemId;
if (itemId === null || itemId == undefined) {
  Office.context.mailbox.item.saveAsync(function(result) {
    itemId = result.value;
  });
}
```

#### <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlook11officemailboxenumsitemtype"></a><span data-ttu-id="02660-336">itemType: [MailboxEnums](/javascript/api/outlook_1_1/office.mailboxenums.itemtype)</span><span class="sxs-lookup"><span data-stu-id="02660-336">itemType: [Office.MailboxEnums.ItemType](/javascript/api/outlook_1_1/office.mailboxenums.itemtype)</span></span>

<span data-ttu-id="02660-337">获取实例表示的项的类型。</span><span class="sxs-lookup"><span data-stu-id="02660-337">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="02660-338">`itemType` 属性返回其中一个 `ItemType` 枚举值，指示 `item` 对象实例是邮件还是约会。</span><span class="sxs-lookup"><span data-stu-id="02660-338">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="02660-339">类型</span><span class="sxs-lookup"><span data-stu-id="02660-339">Type</span></span>

*   [<span data-ttu-id="02660-340">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="02660-340">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook_1_1/office.mailboxenums.itemtype)

##### <a name="requirements"></a><span data-ttu-id="02660-341">要求</span><span class="sxs-lookup"><span data-stu-id="02660-341">Requirements</span></span>

|<span data-ttu-id="02660-342">要求</span><span class="sxs-lookup"><span data-stu-id="02660-342">Requirement</span></span>| <span data-ttu-id="02660-343">值</span><span class="sxs-lookup"><span data-stu-id="02660-343">Value</span></span>|
|---|---|
|[<span data-ttu-id="02660-344">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="02660-344">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="02660-345">1.0</span><span class="sxs-lookup"><span data-stu-id="02660-345">1.0</span></span>|
|[<span data-ttu-id="02660-346">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="02660-346">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="02660-347">ReadItem</span><span class="sxs-lookup"><span data-stu-id="02660-347">ReadItem</span></span>|
|[<span data-ttu-id="02660-348">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="02660-348">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="02660-349">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="02660-349">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="02660-350">示例</span><span class="sxs-lookup"><span data-stu-id="02660-350">Example</span></span>

```javascript
if (Office.context.mailbox.item.itemType === Office.MailboxEnums.ItemType.Message) {
  // Do something.
} else {
  // Do something else.
}
```

#### <a name="location-stringlocationjavascriptapioutlook11officelocation"></a><span data-ttu-id="02660-351">位置: 字符串 |[位置](/javascript/api/outlook_1_1/office.location)</span><span class="sxs-lookup"><span data-stu-id="02660-351">location: String|[Location](/javascript/api/outlook_1_1/office.location)</span></span>

<span data-ttu-id="02660-352">获取或设置约会的位置。</span><span class="sxs-lookup"><span data-stu-id="02660-352">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="02660-353">阅读模式</span><span class="sxs-lookup"><span data-stu-id="02660-353">Read mode</span></span>

<span data-ttu-id="02660-354">`location` 属性返回一个包含约会位置的字符串。</span><span class="sxs-lookup"><span data-stu-id="02660-354">The `location` property returns a string that contains the location of the appointment.</span></span>

```javascript
var location = Office.context.mailbox.item.location;
console.log("location: " + location);
```

##### <a name="compose-mode"></a><span data-ttu-id="02660-355">撰写模式</span><span class="sxs-lookup"><span data-stu-id="02660-355">Compose mode</span></span>

<span data-ttu-id="02660-356">`location` 属性返回一个 `Location` 对象，该对象提供用于获取和设置约会位置的方法。</span><span class="sxs-lookup"><span data-stu-id="02660-356">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

```javascript
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="02660-357">类型</span><span class="sxs-lookup"><span data-stu-id="02660-357">Type</span></span>

*   <span data-ttu-id="02660-358">String | [Location](/javascript/api/outlook_1_1/office.location)</span><span class="sxs-lookup"><span data-stu-id="02660-358">String | [Location](/javascript/api/outlook_1_1/office.location)</span></span>

##### <a name="requirements"></a><span data-ttu-id="02660-359">要求</span><span class="sxs-lookup"><span data-stu-id="02660-359">Requirements</span></span>

|<span data-ttu-id="02660-360">要求</span><span class="sxs-lookup"><span data-stu-id="02660-360">Requirement</span></span>| <span data-ttu-id="02660-361">值</span><span class="sxs-lookup"><span data-stu-id="02660-361">Value</span></span>|
|---|---|
|[<span data-ttu-id="02660-362">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="02660-362">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="02660-363">1.0</span><span class="sxs-lookup"><span data-stu-id="02660-363">1.0</span></span>|
|[<span data-ttu-id="02660-364">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="02660-364">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="02660-365">ReadItem</span><span class="sxs-lookup"><span data-stu-id="02660-365">ReadItem</span></span>|
|[<span data-ttu-id="02660-366">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="02660-366">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="02660-367">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="02660-367">Compose or Read</span></span>|

#### <a name="normalizedsubject-string"></a><span data-ttu-id="02660-368">normalizedSubject: String</span><span class="sxs-lookup"><span data-stu-id="02660-368">normalizedSubject: String</span></span>

<span data-ttu-id="02660-p121">获取删除了所有前缀（包括 `RE:` 和 `FWD:`）的项目主题。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="02660-p121">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="02660-p122">normalizedSubject 属性获取包含由电子邮件程序添加的任何标准前缀（如 `RE:` 和 `FW:`）的项目主题。若要获取包含完整前缀的项目主题，请使用 [`subject`](#subject-stringsubject) 属性。</span><span class="sxs-lookup"><span data-stu-id="02660-p122">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="02660-373">Type</span><span class="sxs-lookup"><span data-stu-id="02660-373">Type</span></span>

*   <span data-ttu-id="02660-374">String</span><span class="sxs-lookup"><span data-stu-id="02660-374">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="02660-375">要求</span><span class="sxs-lookup"><span data-stu-id="02660-375">Requirements</span></span>

|<span data-ttu-id="02660-376">要求</span><span class="sxs-lookup"><span data-stu-id="02660-376">Requirement</span></span>| <span data-ttu-id="02660-377">值</span><span class="sxs-lookup"><span data-stu-id="02660-377">Value</span></span>|
|---|---|
|[<span data-ttu-id="02660-378">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="02660-378">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="02660-379">1.0</span><span class="sxs-lookup"><span data-stu-id="02660-379">1.0</span></span>|
|[<span data-ttu-id="02660-380">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="02660-380">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="02660-381">ReadItem</span><span class="sxs-lookup"><span data-stu-id="02660-381">ReadItem</span></span>|
|[<span data-ttu-id="02660-382">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="02660-382">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="02660-383">阅读</span><span class="sxs-lookup"><span data-stu-id="02660-383">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="02660-384">示例</span><span class="sxs-lookup"><span data-stu-id="02660-384">Example</span></span>

```javascript
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
console.log("Normalized subject: " + normalizedSubject);
```

#### <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlook11officeemailaddressdetailsrecipientsjavascriptapioutlook11officerecipients"></a><span data-ttu-id="02660-385">optionalAttendees: <[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)>|[收件人](/javascript/api/outlook_1_1/office.recipients)的数组</span><span class="sxs-lookup"><span data-stu-id="02660-385">optionalAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_1/office.recipients)</span></span>

<span data-ttu-id="02660-386">提供对事件的可选与会者的访问权限。</span><span class="sxs-lookup"><span data-stu-id="02660-386">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="02660-387">对象的类型和访问级别取决于当前项目的模式。</span><span class="sxs-lookup"><span data-stu-id="02660-387">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="02660-388">阅读模式</span><span class="sxs-lookup"><span data-stu-id="02660-388">Read mode</span></span>

<span data-ttu-id="02660-389">`optionalAttendees` 属性返回一个数组，其中包含每个可选与会者的 `EmailAddressDetails` 对象。</span><span class="sxs-lookup"><span data-stu-id="02660-389">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span>

```javascript
var optionalAttendees = Office.context.mailbox.item.optionalAttendees;
console.log("Optional attendees: " + JSON.stringify(optionalAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="02660-390">撰写模式</span><span class="sxs-lookup"><span data-stu-id="02660-390">Compose mode</span></span>

<span data-ttu-id="02660-391">`optionalAttendees` 属性返回一个 `Recipients` 对象，该对象提供用于获取或更新可选与会者的方法。</span><span class="sxs-lookup"><span data-stu-id="02660-391">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span>

```javascript
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="02660-392">类型</span><span class="sxs-lookup"><span data-stu-id="02660-392">Type</span></span>

*   <span data-ttu-id="02660-393">Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_1/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="02660-393">Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_1/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="02660-394">要求</span><span class="sxs-lookup"><span data-stu-id="02660-394">Requirements</span></span>

|<span data-ttu-id="02660-395">要求</span><span class="sxs-lookup"><span data-stu-id="02660-395">Requirement</span></span>| <span data-ttu-id="02660-396">值</span><span class="sxs-lookup"><span data-stu-id="02660-396">Value</span></span>|
|---|---|
|[<span data-ttu-id="02660-397">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="02660-397">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="02660-398">1.0</span><span class="sxs-lookup"><span data-stu-id="02660-398">1.0</span></span>|
|[<span data-ttu-id="02660-399">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="02660-399">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="02660-400">ReadItem</span><span class="sxs-lookup"><span data-stu-id="02660-400">ReadItem</span></span>|
|[<span data-ttu-id="02660-401">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="02660-401">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="02660-402">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="02660-402">Compose or Read</span></span>|

#### <a name="organizer-emailaddressdetailsjavascriptapioutlook11officeemailaddressdetails"></a><span data-ttu-id="02660-403">组织者: [EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="02660-403">organizer: [EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)</span></span>

<span data-ttu-id="02660-p124">获取指定会议的会议组织者的电子邮件地址。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="02660-p124">Gets the email address of the meeting organizer for a specified meeting. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="02660-406">类型</span><span class="sxs-lookup"><span data-stu-id="02660-406">Type</span></span>

*   [<span data-ttu-id="02660-407">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="02660-407">EmailAddressDetails</span></span>](/javascript/api/outlook_1_1/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="02660-408">要求</span><span class="sxs-lookup"><span data-stu-id="02660-408">Requirements</span></span>

|<span data-ttu-id="02660-409">要求</span><span class="sxs-lookup"><span data-stu-id="02660-409">Requirement</span></span>| <span data-ttu-id="02660-410">值</span><span class="sxs-lookup"><span data-stu-id="02660-410">Value</span></span>|
|---|---|
|[<span data-ttu-id="02660-411">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="02660-411">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="02660-412">1.0</span><span class="sxs-lookup"><span data-stu-id="02660-412">1.0</span></span>|
|[<span data-ttu-id="02660-413">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="02660-413">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="02660-414">ReadItem</span><span class="sxs-lookup"><span data-stu-id="02660-414">ReadItem</span></span>|
|[<span data-ttu-id="02660-415">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="02660-415">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="02660-416">阅读</span><span class="sxs-lookup"><span data-stu-id="02660-416">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="02660-417">示例</span><span class="sxs-lookup"><span data-stu-id="02660-417">Example</span></span>

```javascript
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
console.log("Organizer: " + organizerName + " (" + organizerAddress + ")");
```

#### <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlook11officeemailaddressdetailsrecipientsjavascriptapioutlook11officerecipients"></a><span data-ttu-id="02660-418">requiredAttendees: <[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)>|[收件人](/javascript/api/outlook_1_1/office.recipients)的数组</span><span class="sxs-lookup"><span data-stu-id="02660-418">requiredAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_1/office.recipients)</span></span>

<span data-ttu-id="02660-419">提供对事件的必需与会者的访问权限。</span><span class="sxs-lookup"><span data-stu-id="02660-419">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="02660-420">对象的类型和访问级别取决于当前项目的模式。</span><span class="sxs-lookup"><span data-stu-id="02660-420">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="02660-421">阅读模式</span><span class="sxs-lookup"><span data-stu-id="02660-421">Read mode</span></span>

<span data-ttu-id="02660-422">`requiredAttendees` 属性返回一个数组，其中包含每个必需与会者的 `EmailAddressDetails` 对象。</span><span class="sxs-lookup"><span data-stu-id="02660-422">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span>

```javascript
var requiredAttendees = Office.context.mailbox.item.requiredAttendees;
console.log("Required attendees: " + JSON.stringify(requiredAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="02660-423">撰写模式</span><span class="sxs-lookup"><span data-stu-id="02660-423">Compose mode</span></span>

<span data-ttu-id="02660-424">`requiredAttendees` 属性返回一个 `Recipients` 对象，该对象提供用于获取或更新必需与会者的方法。</span><span class="sxs-lookup"><span data-stu-id="02660-424">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span>

```javascript
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
  console.log(JSON.stringify(arrayOfRequiredAttendeesRecipients));
}
```

##### <a name="type"></a><span data-ttu-id="02660-425">类型</span><span class="sxs-lookup"><span data-stu-id="02660-425">Type</span></span>

*   <span data-ttu-id="02660-426">Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_1/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="02660-426">Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_1/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="02660-427">要求</span><span class="sxs-lookup"><span data-stu-id="02660-427">Requirements</span></span>

|<span data-ttu-id="02660-428">要求</span><span class="sxs-lookup"><span data-stu-id="02660-428">Requirement</span></span>| <span data-ttu-id="02660-429">值</span><span class="sxs-lookup"><span data-stu-id="02660-429">Value</span></span>|
|---|---|
|[<span data-ttu-id="02660-430">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="02660-430">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="02660-431">1.0</span><span class="sxs-lookup"><span data-stu-id="02660-431">1.0</span></span>|
|[<span data-ttu-id="02660-432">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="02660-432">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="02660-433">ReadItem</span><span class="sxs-lookup"><span data-stu-id="02660-433">ReadItem</span></span>|
|[<span data-ttu-id="02660-434">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="02660-434">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="02660-435">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="02660-435">Compose or Read</span></span>|

#### <a name="sender-emailaddressdetailsjavascriptapioutlook11officeemailaddressdetails"></a><span data-ttu-id="02660-436">发件人: [EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="02660-436">sender: [EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)</span></span>

<span data-ttu-id="02660-p126">获取电子邮件发件人的电子邮件地址。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="02660-p126">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="02660-p127">[`from`](#from-emailaddressdetails) 和 `sender` 属性表示同一个人，邮件由代理人发送的除外。在此情况下，`from` 属性表示代理程序，而 sender 属性表示代理人。</span><span class="sxs-lookup"><span data-stu-id="02660-p127">The [`from`](#from-emailaddressdetails) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="02660-441">`from` 属性中 `EmailAddressDetails` 对象的 `recipientType` 属性为 `undefined`。</span><span class="sxs-lookup"><span data-stu-id="02660-441">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="02660-442">类型</span><span class="sxs-lookup"><span data-stu-id="02660-442">Type</span></span>

*   [<span data-ttu-id="02660-443">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="02660-443">EmailAddressDetails</span></span>](/javascript/api/outlook_1_1/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="02660-444">要求</span><span class="sxs-lookup"><span data-stu-id="02660-444">Requirements</span></span>

|<span data-ttu-id="02660-445">要求</span><span class="sxs-lookup"><span data-stu-id="02660-445">Requirement</span></span>| <span data-ttu-id="02660-446">值</span><span class="sxs-lookup"><span data-stu-id="02660-446">Value</span></span>|
|---|---|
|[<span data-ttu-id="02660-447">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="02660-447">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="02660-448">1.0</span><span class="sxs-lookup"><span data-stu-id="02660-448">1.0</span></span>|
|[<span data-ttu-id="02660-449">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="02660-449">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="02660-450">ReadItem</span><span class="sxs-lookup"><span data-stu-id="02660-450">ReadItem</span></span>|
|[<span data-ttu-id="02660-451">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="02660-451">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="02660-452">阅读</span><span class="sxs-lookup"><span data-stu-id="02660-452">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="02660-453">示例</span><span class="sxs-lookup"><span data-stu-id="02660-453">Example</span></span>

```javascript
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
console.log("Sender: " + senderName + " (" + senderAddress + ")");
```

#### <a name="start-datetimejavascriptapioutlook11officetime"></a><span data-ttu-id="02660-454">开始日期: 日期 |[时间](/javascript/api/outlook_1_1/office.time)</span><span class="sxs-lookup"><span data-stu-id="02660-454">start: Date|[Time](/javascript/api/outlook_1_1/office.time)</span></span>

<span data-ttu-id="02660-455">获取或设置约会开始的日期和时间。</span><span class="sxs-lookup"><span data-stu-id="02660-455">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="02660-p128">将 `start` 属性表示为协调世界时 (UTC) 的日期和时间值。可使用 [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) 方法将该值转换为客户端的本地日期和时间。</span><span class="sxs-lookup"><span data-stu-id="02660-p128">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="02660-458">阅读模式</span><span class="sxs-lookup"><span data-stu-id="02660-458">Read mode</span></span>

<span data-ttu-id="02660-459">`start` 属性返回 `Date` 对象。</span><span class="sxs-lookup"><span data-stu-id="02660-459">The `start` property returns a `Date` object.</span></span>

```javascript
var start = Office.context.mailbox.item.start;
console.log("Appointment start: " + JSON.stringify(start));
```

##### <a name="compose-mode"></a><span data-ttu-id="02660-460">撰写模式</span><span class="sxs-lookup"><span data-stu-id="02660-460">Compose mode</span></span>

<span data-ttu-id="02660-461">`start` 属性返回 `Time` 对象。</span><span class="sxs-lookup"><span data-stu-id="02660-461">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="02660-462">使用 [`Time.setAsync`](/javascript/api/outlook_1_1/office.time#setasync-datetime--options--callback-) 方法设置开始时间时，应使用 [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) 方法将客户端的本地时间转换为服务器的 UTC。</span><span class="sxs-lookup"><span data-stu-id="02660-462">When you use the [`Time.setAsync`](/javascript/api/outlook_1_1/office.time#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="02660-463">以下示例通过使用 `Time` 对象的 [`setAsync`](/javascript/api/outlook_1_1/office.time#setasync-datetime--options--callback-) 方法，设置撰写模式下约会的开始时间。</span><span class="sxs-lookup"><span data-stu-id="02660-463">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_1/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="02660-464">类型</span><span class="sxs-lookup"><span data-stu-id="02660-464">Type</span></span>

*   <span data-ttu-id="02660-465">Date | [Time](/javascript/api/outlook_1_1/office.time)</span><span class="sxs-lookup"><span data-stu-id="02660-465">Date | [Time](/javascript/api/outlook_1_1/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="02660-466">要求</span><span class="sxs-lookup"><span data-stu-id="02660-466">Requirements</span></span>

|<span data-ttu-id="02660-467">要求</span><span class="sxs-lookup"><span data-stu-id="02660-467">Requirement</span></span>| <span data-ttu-id="02660-468">值</span><span class="sxs-lookup"><span data-stu-id="02660-468">Value</span></span>|
|---|---|
|[<span data-ttu-id="02660-469">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="02660-469">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="02660-470">1.0</span><span class="sxs-lookup"><span data-stu-id="02660-470">1.0</span></span>|
|[<span data-ttu-id="02660-471">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="02660-471">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="02660-472">ReadItem</span><span class="sxs-lookup"><span data-stu-id="02660-472">ReadItem</span></span>|
|[<span data-ttu-id="02660-473">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="02660-473">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="02660-474">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="02660-474">Compose or Read</span></span>|

#### <a name="subject-stringsubjectjavascriptapioutlook11officesubject"></a><span data-ttu-id="02660-475">subject: String |[主题](/javascript/api/outlook_1_1/office.subject)</span><span class="sxs-lookup"><span data-stu-id="02660-475">subject: String|[Subject](/javascript/api/outlook_1_1/office.subject)</span></span>

<span data-ttu-id="02660-476">获取或设置显示在项目的主题字段中的说明。</span><span class="sxs-lookup"><span data-stu-id="02660-476">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="02660-477">`subject` 属性获取或设置由电子邮件服务器发送项目时的整个主题。</span><span class="sxs-lookup"><span data-stu-id="02660-477">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="02660-478">阅读模式</span><span class="sxs-lookup"><span data-stu-id="02660-478">Read mode</span></span>

<span data-ttu-id="02660-p129">`subject` 属性返回一个字符串。使用 [`normalizedSubject`](#normalizedsubject-string) 属性获取不带任何前导前缀（如 `RE:` 和 `FW:`）的主题。</span><span class="sxs-lookup"><span data-stu-id="02660-p129">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

```js
var subject = Office.context.mailbox.item.subject;
console.log(subject);
```

##### <a name="compose-mode"></a><span data-ttu-id="02660-481">撰写模式</span><span class="sxs-lookup"><span data-stu-id="02660-481">Compose mode</span></span>

<span data-ttu-id="02660-482">`subject` 属性返回一个 `Subject` 对象，该对象提供用于获取和设置主题的方法。</span><span class="sxs-lookup"><span data-stu-id="02660-482">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```javascript
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
  console.log(subject);
}
```

##### <a name="type"></a><span data-ttu-id="02660-483">类型</span><span class="sxs-lookup"><span data-stu-id="02660-483">Type</span></span>

*   <span data-ttu-id="02660-484">String | [Subject](/javascript/api/outlook_1_1/office.subject)</span><span class="sxs-lookup"><span data-stu-id="02660-484">String | [Subject](/javascript/api/outlook_1_1/office.subject)</span></span>

##### <a name="requirements"></a><span data-ttu-id="02660-485">要求</span><span class="sxs-lookup"><span data-stu-id="02660-485">Requirements</span></span>

|<span data-ttu-id="02660-486">要求</span><span class="sxs-lookup"><span data-stu-id="02660-486">Requirement</span></span>| <span data-ttu-id="02660-487">值</span><span class="sxs-lookup"><span data-stu-id="02660-487">Value</span></span>|
|---|---|
|[<span data-ttu-id="02660-488">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="02660-488">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="02660-489">1.0</span><span class="sxs-lookup"><span data-stu-id="02660-489">1.0</span></span>|
|[<span data-ttu-id="02660-490">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="02660-490">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="02660-491">ReadItem</span><span class="sxs-lookup"><span data-stu-id="02660-491">ReadItem</span></span>|
|[<span data-ttu-id="02660-492">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="02660-492">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="02660-493">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="02660-493">Compose or Read</span></span>|

#### <a name="to-arrayemailaddressdetailsjavascriptapioutlook11officeemailaddressdetailsrecipientsjavascriptapioutlook11officerecipients"></a><span data-ttu-id="02660-494">to: <[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)>|[收件人](/javascript/api/outlook_1_1/office.recipients)的数组</span><span class="sxs-lookup"><span data-stu-id="02660-494">to: Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_1/office.recipients)</span></span>

<span data-ttu-id="02660-495">提供对邮件的“**收件人**”行上的收件人的访问权限。</span><span class="sxs-lookup"><span data-stu-id="02660-495">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="02660-496">对象的类型和访问级别取决于当前项目的模式。</span><span class="sxs-lookup"><span data-stu-id="02660-496">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="02660-497">阅读模式</span><span class="sxs-lookup"><span data-stu-id="02660-497">Read mode</span></span>

<span data-ttu-id="02660-p131">`to` 属性返回包含邮件的**收件人**行上所列的每个收件人的 `EmailAddressDetails` 对象的数组。集合上限为 100 个成员。</span><span class="sxs-lookup"><span data-stu-id="02660-p131">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.</span></span>

```javascript
console.log(JSON.stringify(Office.context.mailbox.item.to));
```

##### <a name="compose-mode"></a><span data-ttu-id="02660-500">撰写模式</span><span class="sxs-lookup"><span data-stu-id="02660-500">Compose mode</span></span>

<span data-ttu-id="02660-501">`to` 属性返回一个 `Recipients` 对象，该对象提供用于获取或更新邮件的“**收件人**”行上收件人的方法。</span><span class="sxs-lookup"><span data-stu-id="02660-501">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span>

```javascript
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="02660-502">类型</span><span class="sxs-lookup"><span data-stu-id="02660-502">Type</span></span>

*   <span data-ttu-id="02660-503">Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_1/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="02660-503">Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_1/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="02660-504">要求</span><span class="sxs-lookup"><span data-stu-id="02660-504">Requirements</span></span>

|<span data-ttu-id="02660-505">要求</span><span class="sxs-lookup"><span data-stu-id="02660-505">Requirement</span></span>| <span data-ttu-id="02660-506">值</span><span class="sxs-lookup"><span data-stu-id="02660-506">Value</span></span>|
|---|---|
|[<span data-ttu-id="02660-507">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="02660-507">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="02660-508">1.0</span><span class="sxs-lookup"><span data-stu-id="02660-508">1.0</span></span>|
|[<span data-ttu-id="02660-509">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="02660-509">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="02660-510">ReadItem</span><span class="sxs-lookup"><span data-stu-id="02660-510">ReadItem</span></span>|
|[<span data-ttu-id="02660-511">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="02660-511">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="02660-512">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="02660-512">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="02660-513">方法</span><span class="sxs-lookup"><span data-stu-id="02660-513">Methods</span></span>

#### <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="02660-514">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="02660-514">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="02660-515">将文件作为附件添加到邮件或约会。</span><span class="sxs-lookup"><span data-stu-id="02660-515">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="02660-516">`addFileAttachmentAsync` 方法在指定的 URI 上载文件并将其附加到撰写窗体中的项目。</span><span class="sxs-lookup"><span data-stu-id="02660-516">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="02660-517">随后可以将该标识符与 [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) 方法一同使用，以删除同一个会话中的附件。</span><span class="sxs-lookup"><span data-stu-id="02660-517">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="02660-518">参数</span><span class="sxs-lookup"><span data-stu-id="02660-518">Parameters</span></span>

|<span data-ttu-id="02660-519">名称</span><span class="sxs-lookup"><span data-stu-id="02660-519">Name</span></span>| <span data-ttu-id="02660-520">类型</span><span class="sxs-lookup"><span data-stu-id="02660-520">Type</span></span>| <span data-ttu-id="02660-521">属性</span><span class="sxs-lookup"><span data-stu-id="02660-521">Attributes</span></span>| <span data-ttu-id="02660-522">说明</span><span class="sxs-lookup"><span data-stu-id="02660-522">Description</span></span>|
|---|---|---|---|
|`uri`| <span data-ttu-id="02660-523">String</span><span class="sxs-lookup"><span data-stu-id="02660-523">String</span></span>||<span data-ttu-id="02660-p132">提供附加到邮件或约会的文件的位置的 URI。最大长度为 2048 个字符。</span><span class="sxs-lookup"><span data-stu-id="02660-p132">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="02660-526">字符串</span><span class="sxs-lookup"><span data-stu-id="02660-526">String</span></span>||<span data-ttu-id="02660-p133">在附件上载过程中显示的附件名称。最大长度为 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="02660-p133">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="02660-529">Object</span><span class="sxs-lookup"><span data-stu-id="02660-529">Object</span></span>| <span data-ttu-id="02660-530">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="02660-530">&lt;optional&gt;</span></span>|<span data-ttu-id="02660-531">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="02660-531">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="02660-532">对象</span><span class="sxs-lookup"><span data-stu-id="02660-532">Object</span></span>| <span data-ttu-id="02660-533">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="02660-533">&lt;optional&gt;</span></span>|<span data-ttu-id="02660-534">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="02660-534">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="02660-535">函数</span><span class="sxs-lookup"><span data-stu-id="02660-535">function</span></span>| <span data-ttu-id="02660-536">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="02660-536">&lt;optional&gt;</span></span>|<span data-ttu-id="02660-537">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="02660-537">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="02660-538">如果成功，附件标识符将在 `asyncResult.value` 属性中提供。</span><span class="sxs-lookup"><span data-stu-id="02660-538">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="02660-539">如果上传附件失败，`asyncResult` 对象将包含一个提供错误说明的 `Error` 对象。</span><span class="sxs-lookup"><span data-stu-id="02660-539">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="02660-540">错误</span><span class="sxs-lookup"><span data-stu-id="02660-540">Errors</span></span>

| <span data-ttu-id="02660-541">错误代码</span><span class="sxs-lookup"><span data-stu-id="02660-541">Error code</span></span> | <span data-ttu-id="02660-542">说明</span><span class="sxs-lookup"><span data-stu-id="02660-542">Description</span></span> |
|------------|-------------|
| `AttachmentSizeExceeded` | <span data-ttu-id="02660-543">附件大小超过了允许的大小。</span><span class="sxs-lookup"><span data-stu-id="02660-543">The attachment is larger than allowed.</span></span> |
| `FileTypeNotSupported` | <span data-ttu-id="02660-544">该附件的扩展名不是允许的扩展名。</span><span class="sxs-lookup"><span data-stu-id="02660-544">The attachment has an extension that is not allowed.</span></span> |
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="02660-545">邮件或约会具有的附件过多。</span><span class="sxs-lookup"><span data-stu-id="02660-545">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="02660-546">要求</span><span class="sxs-lookup"><span data-stu-id="02660-546">Requirements</span></span>

|<span data-ttu-id="02660-547">要求</span><span class="sxs-lookup"><span data-stu-id="02660-547">Requirement</span></span>| <span data-ttu-id="02660-548">值</span><span class="sxs-lookup"><span data-stu-id="02660-548">Value</span></span>|
|---|---|
|[<span data-ttu-id="02660-549">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="02660-549">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="02660-550">1.1</span><span class="sxs-lookup"><span data-stu-id="02660-550">1.1</span></span>|
|[<span data-ttu-id="02660-551">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="02660-551">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="02660-552">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="02660-552">ReadWriteItem</span></span>|
|[<span data-ttu-id="02660-553">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="02660-553">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="02660-554">撰写</span><span class="sxs-lookup"><span data-stu-id="02660-554">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="02660-555">示例</span><span class="sxs-lookup"><span data-stu-id="02660-555">Example</span></span>

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

#### <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="02660-556">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="02660-556">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="02660-557">将 Exchange 项目（如邮件）作为附件添加到邮件或约会。</span><span class="sxs-lookup"><span data-stu-id="02660-557">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="02660-p134">`addItemAttachmentAsync` 方法将包含指定 Exchange 标识符的项目附加到撰写窗体中的项目。如果指定一个回调方法，此方法使用 `asyncResult` 参数调用，该参数包含一个附件标识符或代码，指示附加项目过程中出现的任何错误。可以使用 `options` 参数将状态信息传递给回调方法（如果需要）。</span><span class="sxs-lookup"><span data-stu-id="02660-p134">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="02660-561">随后可以将该标识符与 [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) 方法一同使用，以删除同一个会话中的附件。</span><span class="sxs-lookup"><span data-stu-id="02660-561">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="02660-562">如果 Office 外接程序在 web 上的 Outlook 中运行, 则该`addItemAttachmentAsync`方法可以将项目附加到您正在编辑的项目之外的项目中;但是, 不支持这种情况, 建议不要这样做。</span><span class="sxs-lookup"><span data-stu-id="02660-562">If your Office Add-in is running in Outlook on the web, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="02660-563">参数</span><span class="sxs-lookup"><span data-stu-id="02660-563">Parameters</span></span>

|<span data-ttu-id="02660-564">名称</span><span class="sxs-lookup"><span data-stu-id="02660-564">Name</span></span>| <span data-ttu-id="02660-565">类型</span><span class="sxs-lookup"><span data-stu-id="02660-565">Type</span></span>| <span data-ttu-id="02660-566">属性</span><span class="sxs-lookup"><span data-stu-id="02660-566">Attributes</span></span>| <span data-ttu-id="02660-567">说明</span><span class="sxs-lookup"><span data-stu-id="02660-567">Description</span></span>|
|---|---|---|---|
|`itemId`| <span data-ttu-id="02660-568">String</span><span class="sxs-lookup"><span data-stu-id="02660-568">String</span></span>||<span data-ttu-id="02660-p135">要附加的项目的 Exchange 标识符。最大长度为 100 个字符。</span><span class="sxs-lookup"><span data-stu-id="02660-p135">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="02660-571">String</span><span class="sxs-lookup"><span data-stu-id="02660-571">String</span></span>||<span data-ttu-id="02660-572">要附加的项目的主题。</span><span class="sxs-lookup"><span data-stu-id="02660-572">The subject of the item to be attached.</span></span> <span data-ttu-id="02660-573">最大长度为 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="02660-573">The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="02660-574">对象</span><span class="sxs-lookup"><span data-stu-id="02660-574">Object</span></span>| <span data-ttu-id="02660-575">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="02660-575">&lt;optional&gt;</span></span>|<span data-ttu-id="02660-576">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="02660-576">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="02660-577">对象</span><span class="sxs-lookup"><span data-stu-id="02660-577">Object</span></span>| <span data-ttu-id="02660-578">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="02660-578">&lt;optional&gt;</span></span>|<span data-ttu-id="02660-579">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="02660-579">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="02660-580">函数</span><span class="sxs-lookup"><span data-stu-id="02660-580">function</span></span>| <span data-ttu-id="02660-581">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="02660-581">&lt;optional&gt;</span></span>|<span data-ttu-id="02660-582">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="02660-582">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="02660-583">如果成功，附件标识符将在 `asyncResult.value` 属性中提供。</span><span class="sxs-lookup"><span data-stu-id="02660-583">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="02660-584">如果添加附件失败，`asyncResult` 对象将包含一个提供错误说明的 `Error` 对象。</span><span class="sxs-lookup"><span data-stu-id="02660-584">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="02660-585">错误</span><span class="sxs-lookup"><span data-stu-id="02660-585">Errors</span></span>

| <span data-ttu-id="02660-586">错误代码</span><span class="sxs-lookup"><span data-stu-id="02660-586">Error code</span></span> | <span data-ttu-id="02660-587">说明</span><span class="sxs-lookup"><span data-stu-id="02660-587">Description</span></span> |
|------------|-------------|
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="02660-588">邮件或约会具有的附件过多。</span><span class="sxs-lookup"><span data-stu-id="02660-588">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="02660-589">要求</span><span class="sxs-lookup"><span data-stu-id="02660-589">Requirements</span></span>

|<span data-ttu-id="02660-590">要求</span><span class="sxs-lookup"><span data-stu-id="02660-590">Requirement</span></span>| <span data-ttu-id="02660-591">值</span><span class="sxs-lookup"><span data-stu-id="02660-591">Value</span></span>|
|---|---|
|[<span data-ttu-id="02660-592">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="02660-592">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="02660-593">1.1</span><span class="sxs-lookup"><span data-stu-id="02660-593">1.1</span></span>|
|[<span data-ttu-id="02660-594">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="02660-594">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="02660-595">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="02660-595">ReadWriteItem</span></span>|
|[<span data-ttu-id="02660-596">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="02660-596">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="02660-597">撰写</span><span class="sxs-lookup"><span data-stu-id="02660-597">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="02660-598">示例</span><span class="sxs-lookup"><span data-stu-id="02660-598">Example</span></span>

<span data-ttu-id="02660-599">以下示例将现有的 Outlook 项目添加为名为 `My Attachment` 的附件。</span><span class="sxs-lookup"><span data-stu-id="02660-599">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

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

#### <a name="displayreplyallformformdata-callback"></a><span data-ttu-id="02660-600">displayReplyAllForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="02660-600">displayReplyAllForm(formData, [callback])</span></span>

<span data-ttu-id="02660-601">显示答复窗体，其中包括所选邮件的发件人和所有收件人或所选约会的组织者和所有与会者。</span><span class="sxs-lookup"><span data-stu-id="02660-601">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="02660-602">IOS 或 Android 上的 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="02660-602">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="02660-603">在 web 上的 Outlook 中, 答复窗体显示为3列视图中的弹出窗体和2列或1列视图中的弹出式窗体。</span><span class="sxs-lookup"><span data-stu-id="02660-603">In Outlook on the web, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="02660-604">如果任意字符串参数超出其限制，`displayReplyAllForm` 将引发异常。</span><span class="sxs-lookup"><span data-stu-id="02660-604">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

> [!NOTE]
> <span data-ttu-id="02660-605">要求集1.1 中不支持在呼叫`displayReplyAllForm`中包含附件的功能。</span><span class="sxs-lookup"><span data-stu-id="02660-605">The ability to include attachments in the call to `displayReplyAllForm` is not supported in requirement set 1.1.</span></span> <span data-ttu-id="02660-606">附件支持已添加到要求集 1.2 及以上的 `displayReplyAllForm` 中。</span><span class="sxs-lookup"><span data-stu-id="02660-606">Attachment support was added to `displayReplyAllForm` in requirement set 1.2 and above.</span></span>

##### <a name="parameters"></a><span data-ttu-id="02660-607">参数</span><span class="sxs-lookup"><span data-stu-id="02660-607">Parameters</span></span>

|<span data-ttu-id="02660-608">名称</span><span class="sxs-lookup"><span data-stu-id="02660-608">Name</span></span>| <span data-ttu-id="02660-609">类型</span><span class="sxs-lookup"><span data-stu-id="02660-609">Type</span></span>| <span data-ttu-id="02660-610">说明</span><span class="sxs-lookup"><span data-stu-id="02660-610">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="02660-611">字符串 &#124; 对象</span><span class="sxs-lookup"><span data-stu-id="02660-611">String &#124; Object</span></span>| |<span data-ttu-id="02660-p138">一个包含文本和 HTML 且表示答复窗体的正文的字符串。字符串限制为 32 KB。</span><span class="sxs-lookup"><span data-stu-id="02660-p138">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="02660-614">**或**</span><span class="sxs-lookup"><span data-stu-id="02660-614">**OR**</span></span><br/><span data-ttu-id="02660-p139">包含正文或附件数据和回调函数的对象。对象定义如下。</span><span class="sxs-lookup"><span data-stu-id="02660-p139">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="02660-617">字符串</span><span class="sxs-lookup"><span data-stu-id="02660-617">String</span></span> | <span data-ttu-id="02660-618">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="02660-618">&lt;optional&gt;</span></span> | <span data-ttu-id="02660-p140">一个包含文本和 HTML 且表示答复窗体的正文的字符串。字符串限制为 32 KB。</span><span class="sxs-lookup"><span data-stu-id="02660-p140">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `callback` | <span data-ttu-id="02660-621">函数</span><span class="sxs-lookup"><span data-stu-id="02660-621">function</span></span> | <span data-ttu-id="02660-622">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="02660-622">&lt;optional&gt;</span></span> | <span data-ttu-id="02660-623">方法完成后，使用单个参数 `asyncResult`（一个 [AsyncResult](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="02660-623">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="02660-624">要求</span><span class="sxs-lookup"><span data-stu-id="02660-624">Requirements</span></span>

|<span data-ttu-id="02660-625">要求</span><span class="sxs-lookup"><span data-stu-id="02660-625">Requirement</span></span>| <span data-ttu-id="02660-626">值</span><span class="sxs-lookup"><span data-stu-id="02660-626">Value</span></span>|
|---|---|
|[<span data-ttu-id="02660-627">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="02660-627">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="02660-628">1.0</span><span class="sxs-lookup"><span data-stu-id="02660-628">1.0</span></span>|
|[<span data-ttu-id="02660-629">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="02660-629">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="02660-630">ReadItem</span><span class="sxs-lookup"><span data-stu-id="02660-630">ReadItem</span></span>|
|[<span data-ttu-id="02660-631">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="02660-631">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="02660-632">阅读</span><span class="sxs-lookup"><span data-stu-id="02660-632">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="02660-633">示例</span><span class="sxs-lookup"><span data-stu-id="02660-633">Examples</span></span>

<span data-ttu-id="02660-634">以下代码将一个字符串传递到 `displayReplyAllForm` 函数。</span><span class="sxs-lookup"><span data-stu-id="02660-634">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="02660-635">使用空白正文答复。</span><span class="sxs-lookup"><span data-stu-id="02660-635">Reply with an empty body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="02660-636">仅使用正文答复。</span><span class="sxs-lookup"><span data-stu-id="02660-636">Reply with just a body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="02660-637">使用正文和回调答复。</span><span class="sxs-lookup"><span data-stu-id="02660-637">Reply with a body and a callback.</span></span>

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

#### <a name="displayreplyformformdata-callback"></a><span data-ttu-id="02660-638">displayReplyForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="02660-638">displayReplyForm(formData, [callback])</span></span>

<span data-ttu-id="02660-639">显示答复窗体，其中仅包括所选邮件的发件人或所选约会的组织者。</span><span class="sxs-lookup"><span data-stu-id="02660-639">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="02660-640">IOS 或 Android 上的 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="02660-640">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="02660-641">在 web 上的 Outlook 中, 答复窗体显示为3列视图中的弹出窗体和2列或1列视图中的弹出式窗体。</span><span class="sxs-lookup"><span data-stu-id="02660-641">In Outlook on the web, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="02660-642">如果任意字符串参数超出其限制，`displayReplyForm` 将引发异常。</span><span class="sxs-lookup"><span data-stu-id="02660-642">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

> [!NOTE]
> <span data-ttu-id="02660-643">要求集1.1 中不支持在呼叫`displayReplyForm`中包含附件的功能。</span><span class="sxs-lookup"><span data-stu-id="02660-643">The ability to include attachments in the call to `displayReplyForm` is not supported in requirement set 1.1.</span></span> <span data-ttu-id="02660-644">附件支持已添加到要求集 1.2 及以上的 `displayReplyForm` 中。</span><span class="sxs-lookup"><span data-stu-id="02660-644">Attachment support was added to `displayReplyForm` in requirement set 1.2 and above.</span></span>

##### <a name="parameters"></a><span data-ttu-id="02660-645">参数</span><span class="sxs-lookup"><span data-stu-id="02660-645">Parameters</span></span>

|<span data-ttu-id="02660-646">名称</span><span class="sxs-lookup"><span data-stu-id="02660-646">Name</span></span>| <span data-ttu-id="02660-647">类型</span><span class="sxs-lookup"><span data-stu-id="02660-647">Type</span></span>| <span data-ttu-id="02660-648">说明</span><span class="sxs-lookup"><span data-stu-id="02660-648">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="02660-649">字符串 &#124; 对象</span><span class="sxs-lookup"><span data-stu-id="02660-649">String &#124; Object</span></span>| | <span data-ttu-id="02660-p142">一个包含文本和 HTML 且表示答复窗体的正文的字符串。字符串限制为 32 KB。</span><span class="sxs-lookup"><span data-stu-id="02660-p142">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="02660-652">**或**</span><span class="sxs-lookup"><span data-stu-id="02660-652">**OR**</span></span><br/><span data-ttu-id="02660-p143">包含正文或附件数据和回调函数的对象。对象定义如下。</span><span class="sxs-lookup"><span data-stu-id="02660-p143">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="02660-655">字符串</span><span class="sxs-lookup"><span data-stu-id="02660-655">String</span></span> | <span data-ttu-id="02660-656">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="02660-656">&lt;optional&gt;</span></span> | <span data-ttu-id="02660-p144">一个包含文本和 HTML 且表示答复窗体的正文的字符串。字符串限制为 32 KB。</span><span class="sxs-lookup"><span data-stu-id="02660-p144">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `callback` | <span data-ttu-id="02660-659">函数</span><span class="sxs-lookup"><span data-stu-id="02660-659">function</span></span> | <span data-ttu-id="02660-660">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="02660-660">&lt;optional&gt;</span></span> | <span data-ttu-id="02660-661">方法完成后，使用单个参数 `asyncResult`（一个 [AsyncResult](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="02660-661">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="02660-662">要求</span><span class="sxs-lookup"><span data-stu-id="02660-662">Requirements</span></span>

|<span data-ttu-id="02660-663">要求</span><span class="sxs-lookup"><span data-stu-id="02660-663">Requirement</span></span>| <span data-ttu-id="02660-664">值</span><span class="sxs-lookup"><span data-stu-id="02660-664">Value</span></span>|
|---|---|
|[<span data-ttu-id="02660-665">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="02660-665">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="02660-666">1.0</span><span class="sxs-lookup"><span data-stu-id="02660-666">1.0</span></span>|
|[<span data-ttu-id="02660-667">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="02660-667">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="02660-668">ReadItem</span><span class="sxs-lookup"><span data-stu-id="02660-668">ReadItem</span></span>|
|[<span data-ttu-id="02660-669">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="02660-669">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="02660-670">阅读</span><span class="sxs-lookup"><span data-stu-id="02660-670">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="02660-671">示例</span><span class="sxs-lookup"><span data-stu-id="02660-671">Examples</span></span>

<span data-ttu-id="02660-672">以下代码将一个字符串传递到 `displayReplyForm` 函数。</span><span class="sxs-lookup"><span data-stu-id="02660-672">The following code passes a string to the `displayReplyForm` function.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="02660-673">使用空白正文答复。</span><span class="sxs-lookup"><span data-stu-id="02660-673">Reply with an empty body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="02660-674">仅使用正文答复。</span><span class="sxs-lookup"><span data-stu-id="02660-674">Reply with just a body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="02660-675">使用正文和回调答复。</span><span class="sxs-lookup"><span data-stu-id="02660-675">Reply with a body and a callback.</span></span>

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

#### <a name="getentities--entitiesjavascriptapioutlook11officeentities"></a><span data-ttu-id="02660-676">getEntities() → {[Entities](/javascript/api/outlook_1_1/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="02660-676">getEntities() → {[Entities](/javascript/api/outlook_1_1/office.entities)}</span></span>

<span data-ttu-id="02660-677">获取在所选项目的正文中找到的实体。</span><span class="sxs-lookup"><span data-stu-id="02660-677">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="02660-678">IOS 或 Android 上的 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="02660-678">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="02660-679">要求</span><span class="sxs-lookup"><span data-stu-id="02660-679">Requirements</span></span>

|<span data-ttu-id="02660-680">要求</span><span class="sxs-lookup"><span data-stu-id="02660-680">Requirement</span></span>| <span data-ttu-id="02660-681">值</span><span class="sxs-lookup"><span data-stu-id="02660-681">Value</span></span>|
|---|---|
|[<span data-ttu-id="02660-682">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="02660-682">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="02660-683">1.0</span><span class="sxs-lookup"><span data-stu-id="02660-683">1.0</span></span>|
|[<span data-ttu-id="02660-684">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="02660-684">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="02660-685">ReadItem</span><span class="sxs-lookup"><span data-stu-id="02660-685">ReadItem</span></span>|
|[<span data-ttu-id="02660-686">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="02660-686">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="02660-687">阅读</span><span class="sxs-lookup"><span data-stu-id="02660-687">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="02660-688">返回：</span><span class="sxs-lookup"><span data-stu-id="02660-688">Returns:</span></span>

<span data-ttu-id="02660-689">类型：[Entities](/javascript/api/outlook_1_1/office.entities)</span><span class="sxs-lookup"><span data-stu-id="02660-689">Type: [Entities](/javascript/api/outlook_1_1/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="02660-690">示例</span><span class="sxs-lookup"><span data-stu-id="02660-690">Example</span></span>

<span data-ttu-id="02660-691">以下示例访问当前项目的正文中的联系人实体。</span><span class="sxs-lookup"><span data-stu-id="02660-691">The following example accesses the contacts entities in the current item's body.</span></span>

```javascript
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlook11officecontactmeetingsuggestionjavascriptapioutlook11officemeetingsuggestionphonenumberjavascriptapioutlook11officephonenumbertasksuggestionjavascriptapioutlook11officetasksuggestion"></a><span data-ttu-id="02660-692">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_1/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_1/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_1/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_1/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="02660-692">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_1/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_1/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_1/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_1/office.tasksuggestion))>}</span></span>

<span data-ttu-id="02660-693">获取所选项目的正文中找到的指定实体类型的所有实体的数组。</span><span class="sxs-lookup"><span data-stu-id="02660-693">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="02660-694">IOS 或 Android 上的 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="02660-694">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="02660-695">参数</span><span class="sxs-lookup"><span data-stu-id="02660-695">Parameters</span></span>

|<span data-ttu-id="02660-696">名称</span><span class="sxs-lookup"><span data-stu-id="02660-696">Name</span></span>| <span data-ttu-id="02660-697">类型</span><span class="sxs-lookup"><span data-stu-id="02660-697">Type</span></span>| <span data-ttu-id="02660-698">说明</span><span class="sxs-lookup"><span data-stu-id="02660-698">Description</span></span>|
|---|---|---|
|`entityType`| [<span data-ttu-id="02660-699">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="02660-699">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook_1_1/office.MailboxEnums.entitytype)|<span data-ttu-id="02660-700">EntityType 枚举值之一。</span><span class="sxs-lookup"><span data-stu-id="02660-700">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="02660-701">要求</span><span class="sxs-lookup"><span data-stu-id="02660-701">Requirements</span></span>

|<span data-ttu-id="02660-702">要求</span><span class="sxs-lookup"><span data-stu-id="02660-702">Requirement</span></span>| <span data-ttu-id="02660-703">值</span><span class="sxs-lookup"><span data-stu-id="02660-703">Value</span></span>|
|---|---|
|[<span data-ttu-id="02660-704">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="02660-704">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="02660-705">1.0</span><span class="sxs-lookup"><span data-stu-id="02660-705">1.0</span></span>|
|[<span data-ttu-id="02660-706">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="02660-706">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="02660-707">受限</span><span class="sxs-lookup"><span data-stu-id="02660-707">Restricted</span></span>|
|[<span data-ttu-id="02660-708">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="02660-708">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="02660-709">阅读</span><span class="sxs-lookup"><span data-stu-id="02660-709">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="02660-710">返回：</span><span class="sxs-lookup"><span data-stu-id="02660-710">Returns:</span></span>

<span data-ttu-id="02660-711">如果在 `entityType` 中传递的值不是 `EntityType` 枚举的有效成员，该方法返回 null。</span><span class="sxs-lookup"><span data-stu-id="02660-711">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="02660-712">如果指定类型的任何实体都不存在于该项目的正文中，该方法将返回空数组。</span><span class="sxs-lookup"><span data-stu-id="02660-712">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="02660-713">否则，返回的数组中对象的类型取决于 `entityType` 参数中请求实体的类型。</span><span class="sxs-lookup"><span data-stu-id="02660-713">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="02660-714">当使用此方法的最低权限级别**受限**时，某些实体类型需要 **ReadItem** 才能进行访问，如下表中所指定。</span><span class="sxs-lookup"><span data-stu-id="02660-714">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

| <span data-ttu-id="02660-715">`entityType` 的值</span><span class="sxs-lookup"><span data-stu-id="02660-715">Value of `entityType`</span></span> | <span data-ttu-id="02660-716">返回的数组中对象的类型</span><span class="sxs-lookup"><span data-stu-id="02660-716">Type of objects in returned array</span></span> | <span data-ttu-id="02660-717">所需权限级别</span><span class="sxs-lookup"><span data-stu-id="02660-717">Required Permission Level</span></span> |
| --- | --- | --- |
| `Address` | <span data-ttu-id="02660-718">String</span><span class="sxs-lookup"><span data-stu-id="02660-718">String</span></span> | <span data-ttu-id="02660-719">**受限**</span><span class="sxs-lookup"><span data-stu-id="02660-719">**Restricted**</span></span> |
| `Contact` | <span data-ttu-id="02660-720">Contact</span><span class="sxs-lookup"><span data-stu-id="02660-720">Contact</span></span> | <span data-ttu-id="02660-721">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="02660-721">**ReadItem**</span></span> |
| `EmailAddress` | <span data-ttu-id="02660-722">String</span><span class="sxs-lookup"><span data-stu-id="02660-722">String</span></span> | <span data-ttu-id="02660-723">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="02660-723">**ReadItem**</span></span> |
| `MeetingSuggestion` | <span data-ttu-id="02660-724">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="02660-724">MeetingSuggestion</span></span> | <span data-ttu-id="02660-725">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="02660-725">**ReadItem**</span></span> |
| `PhoneNumber` | <span data-ttu-id="02660-726">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="02660-726">PhoneNumber</span></span> | <span data-ttu-id="02660-727">**受限**</span><span class="sxs-lookup"><span data-stu-id="02660-727">**Restricted**</span></span> |
| `TaskSuggestion` | <span data-ttu-id="02660-728">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="02660-728">TaskSuggestion</span></span> | <span data-ttu-id="02660-729">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="02660-729">**ReadItem**</span></span> |
| `URL` | <span data-ttu-id="02660-730">String</span><span class="sxs-lookup"><span data-stu-id="02660-730">String</span></span> | <span data-ttu-id="02660-731">**受限**</span><span class="sxs-lookup"><span data-stu-id="02660-731">**Restricted**</span></span> |

<span data-ttu-id="02660-732">类型：Array.<(String|[Contact](/javascript/api/outlook_1_1/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_1/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_1/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_1/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="02660-732">Type:  Array.<(String|[Contact](/javascript/api/outlook_1_1/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_1/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_1/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_1/office.tasksuggestion))></span></span>


##### <a name="example"></a><span data-ttu-id="02660-733">示例</span><span class="sxs-lookup"><span data-stu-id="02660-733">Example</span></span>

<span data-ttu-id="02660-734">以下示例显示了如何访问表示当前项目的正文中的邮政地址的字符串数组。</span><span class="sxs-lookup"><span data-stu-id="02660-734">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlook11officecontactmeetingsuggestionjavascriptapioutlook11officemeetingsuggestionphonenumberjavascriptapioutlook11officephonenumbertasksuggestionjavascriptapioutlook11officetasksuggestion"></a><span data-ttu-id="02660-735">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_1/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_1/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_1/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_1/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="02660-735">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_1/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_1/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_1/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_1/office.tasksuggestion))>}</span></span>

<span data-ttu-id="02660-736">返回传递清单 XML 文件中定义的命名筛选器的所选项目中的已知实体。</span><span class="sxs-lookup"><span data-stu-id="02660-736">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="02660-737">IOS 或 Android 上的 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="02660-737">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="02660-738">`getFilteredEntitiesByName` 方法返回匹配在具有指定 `FilterName` 元素值的清单 XML 文件中的 [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) 规则元素中定义的正则表达式的实体。</span><span class="sxs-lookup"><span data-stu-id="02660-738">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="02660-739">参数</span><span class="sxs-lookup"><span data-stu-id="02660-739">Parameters</span></span>

|<span data-ttu-id="02660-740">名称</span><span class="sxs-lookup"><span data-stu-id="02660-740">Name</span></span>| <span data-ttu-id="02660-741">类型</span><span class="sxs-lookup"><span data-stu-id="02660-741">Type</span></span>| <span data-ttu-id="02660-742">说明</span><span class="sxs-lookup"><span data-stu-id="02660-742">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="02660-743">字符串</span><span class="sxs-lookup"><span data-stu-id="02660-743">String</span></span>|<span data-ttu-id="02660-744">定义筛选器匹配的 `ItemHasKnownEntity` 规则元素的名称。</span><span class="sxs-lookup"><span data-stu-id="02660-744">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="02660-745">要求</span><span class="sxs-lookup"><span data-stu-id="02660-745">Requirements</span></span>

|<span data-ttu-id="02660-746">要求</span><span class="sxs-lookup"><span data-stu-id="02660-746">Requirement</span></span>| <span data-ttu-id="02660-747">值</span><span class="sxs-lookup"><span data-stu-id="02660-747">Value</span></span>|
|---|---|
|[<span data-ttu-id="02660-748">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="02660-748">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="02660-749">1.0</span><span class="sxs-lookup"><span data-stu-id="02660-749">1.0</span></span>|
|[<span data-ttu-id="02660-750">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="02660-750">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="02660-751">ReadItem</span><span class="sxs-lookup"><span data-stu-id="02660-751">ReadItem</span></span>|
|[<span data-ttu-id="02660-752">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="02660-752">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="02660-753">阅读</span><span class="sxs-lookup"><span data-stu-id="02660-753">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="02660-754">返回：</span><span class="sxs-lookup"><span data-stu-id="02660-754">Returns:</span></span>

<span data-ttu-id="02660-p146">如果具有匹配 `name` 参数的 `FilterName` 元素值的清单中没有任何 `ItemHasKnownEntity` 元素，则该方法返回 `null`。如果 `name` 参数匹配清单中的 `ItemHasKnownEntity` 元素，但在匹配的当前项目中没有实体，则该方法返回一个空数组。</span><span class="sxs-lookup"><span data-stu-id="02660-p146">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>


<span data-ttu-id="02660-757">类型：Array.<(String|[Contact](/javascript/api/outlook_1_1/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_1/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_1/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_1/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="02660-757">Type: Array.<(String|[Contact](/javascript/api/outlook_1_1/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_1/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_1/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_1/office.tasksuggestion))></span></span>


#### <a name="getregexmatches--object"></a><span data-ttu-id="02660-758">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="02660-758">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="02660-759">返回所选项目中匹配在清单 XML 文件中定义的正则表达式的字符串值。</span><span class="sxs-lookup"><span data-stu-id="02660-759">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="02660-760">IOS 或 Android 上的 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="02660-760">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="02660-p147">`getRegExMatches` 方法返回匹配在清单 XML 文件中的每个 `ItemHasRegularExpressionMatch` 或 `ItemHasKnownEntity` 规则元素中定义的正则表达式的字符串。对于 `ItemHasRegularExpressionMatch` 规则，匹配字符串必须发生在该规则指定的项目的属性中。`PropertyName` 简单类型定义支持的属性。</span><span class="sxs-lookup"><span data-stu-id="02660-p147">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="02660-764">例如，考虑一个外接程序清单具有以下 `Rule` 元素：</span><span class="sxs-lookup"><span data-stu-id="02660-764">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="02660-765">从 `getRegExMatches` 返回的对象应有两个属性：`fruits` 和 `veggies`。</span><span class="sxs-lookup"><span data-stu-id="02660-765">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

> [!NOTE]
> <span data-ttu-id="02660-p148">如果在项目的正文属性上指定 `ItemHasRegularExpressionMatch` 规则，则正则表达式应进一步筛选正文，不应尝试返回该项目的整个正文。使用正则表达式（如 `.*`）获取项目的整个正文并不总是返回预期的结果。</span><span class="sxs-lookup"><span data-stu-id="02660-p148">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="requirements"></a><span data-ttu-id="02660-768">要求</span><span class="sxs-lookup"><span data-stu-id="02660-768">Requirements</span></span>

|<span data-ttu-id="02660-769">要求</span><span class="sxs-lookup"><span data-stu-id="02660-769">Requirement</span></span>| <span data-ttu-id="02660-770">值</span><span class="sxs-lookup"><span data-stu-id="02660-770">Value</span></span>|
|---|---|
|[<span data-ttu-id="02660-771">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="02660-771">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="02660-772">1.0</span><span class="sxs-lookup"><span data-stu-id="02660-772">1.0</span></span>|
|[<span data-ttu-id="02660-773">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="02660-773">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="02660-774">ReadItem</span><span class="sxs-lookup"><span data-stu-id="02660-774">ReadItem</span></span>|
|[<span data-ttu-id="02660-775">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="02660-775">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="02660-776">阅读</span><span class="sxs-lookup"><span data-stu-id="02660-776">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="02660-777">返回：</span><span class="sxs-lookup"><span data-stu-id="02660-777">Returns:</span></span>

<span data-ttu-id="02660-p149">一个包含匹配在清单 XML 文件中定义的正则表达式的字符串数组的对象。每个数组的名称等于匹配 `ItemHasRegularExpressionMatch` 规则的 `RegExName` 属性或匹配 `ItemHasKnownEntity` 规则的 `FilterName` 属性的相应值。</span><span class="sxs-lookup"><span data-stu-id="02660-p149">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<dl class="param-type"><span data-ttu-id="02660-780">

<dt>
类型</dt>


</span><span class="sxs-lookup"><span data-stu-id="02660-780">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="02660-781">对象</span><span class="sxs-lookup"><span data-stu-id="02660-781">Object</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="02660-782">示例</span><span class="sxs-lookup"><span data-stu-id="02660-782">Example</span></span>

<span data-ttu-id="02660-783">以下示例显示了如何访问正则表达式 <rule> 元素 `fruits` 和 `veggies` 的匹配项的数组，这些元素在清单中指定。</rule></span><span class="sxs-lookup"><span data-stu-id="02660-783">The following example shows how to access the array of matches for the regular expression <rule>elements `fruits` and `veggies`, which are specified in the manifest.</rule></span></span>

```javascript
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veggies = allMatches.veggies;
```

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="02660-784">getRegExMatchesByName(name) → (nullable) {Array.<String>}</span><span class="sxs-lookup"><span data-stu-id="02660-784">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="02660-785">返回匹配在清单 XML 文件中定义的命名正则表达式的所选项目中的字符串值。</span><span class="sxs-lookup"><span data-stu-id="02660-785">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="02660-786">IOS 或 Android 上的 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="02660-786">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="02660-787">`getRegExMatchesByName` 方法返回匹配在具有指定 `RegExName` 元素值的清单 XML 文件中的 `ItemHasRegularExpressionMatch` 规则元素中定义的正则表达式的字符串。</span><span class="sxs-lookup"><span data-stu-id="02660-787">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="02660-p150">如果在项目的正文属性上指定 `ItemHasRegularExpressionMatch` 规则，则正则表达式应进一步筛选正文，不应尝试返回该项目的整个正文。使用正则表达式（如 `.*`）获取项目的整个正文并不总是返回预期的结果。</span><span class="sxs-lookup"><span data-stu-id="02660-p150">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="02660-790">参数</span><span class="sxs-lookup"><span data-stu-id="02660-790">Parameters</span></span>

|<span data-ttu-id="02660-791">名称</span><span class="sxs-lookup"><span data-stu-id="02660-791">Name</span></span>| <span data-ttu-id="02660-792">类型</span><span class="sxs-lookup"><span data-stu-id="02660-792">Type</span></span>| <span data-ttu-id="02660-793">说明</span><span class="sxs-lookup"><span data-stu-id="02660-793">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="02660-794">String</span><span class="sxs-lookup"><span data-stu-id="02660-794">String</span></span>|<span data-ttu-id="02660-795">定义筛选器匹配的 `ItemHasRegularExpressionMatch` 规则元素的名称。</span><span class="sxs-lookup"><span data-stu-id="02660-795">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="02660-796">要求</span><span class="sxs-lookup"><span data-stu-id="02660-796">Requirements</span></span>

|<span data-ttu-id="02660-797">要求</span><span class="sxs-lookup"><span data-stu-id="02660-797">Requirement</span></span>| <span data-ttu-id="02660-798">值</span><span class="sxs-lookup"><span data-stu-id="02660-798">Value</span></span>|
|---|---|
|[<span data-ttu-id="02660-799">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="02660-799">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="02660-800">1.0</span><span class="sxs-lookup"><span data-stu-id="02660-800">1.0</span></span>|
|[<span data-ttu-id="02660-801">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="02660-801">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="02660-802">ReadItem</span><span class="sxs-lookup"><span data-stu-id="02660-802">ReadItem</span></span>|
|[<span data-ttu-id="02660-803">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="02660-803">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="02660-804">阅读</span><span class="sxs-lookup"><span data-stu-id="02660-804">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="02660-805">返回：</span><span class="sxs-lookup"><span data-stu-id="02660-805">Returns:</span></span>

<span data-ttu-id="02660-806">一个包含匹配在清单 XML 文件中定义的正则表达式的字符串的数组。</span><span class="sxs-lookup"><span data-stu-id="02660-806">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<dl class="param-type"><span data-ttu-id="02660-807">

<dt>类型</dt>

</span><span class="sxs-lookup"><span data-stu-id="02660-807">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="02660-808">Array.< String ></span><span class="sxs-lookup"><span data-stu-id="02660-808">Array.< String ></span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="02660-809">示例</span><span class="sxs-lookup"><span data-stu-id="02660-809">Example</span></span>

```javascript
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

#### <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="02660-810">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="02660-810">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="02660-811">异步加载所选项目上此外接程序的自定义属性。</span><span class="sxs-lookup"><span data-stu-id="02660-811">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="02660-p151">自定义属性基于每个应用、每个项目存储为键/值对。此方法在回调中返回 `CustomProperties` 对象，该回调提供访问特定于当前项目和当前外接程序的自定义属性的方法。自定义属性未在项目上加密，因此这不应用作安全存储。</span><span class="sxs-lookup"><span data-stu-id="02660-p151">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="02660-815">参数</span><span class="sxs-lookup"><span data-stu-id="02660-815">Parameters</span></span>

|<span data-ttu-id="02660-816">名称</span><span class="sxs-lookup"><span data-stu-id="02660-816">Name</span></span>| <span data-ttu-id="02660-817">类型</span><span class="sxs-lookup"><span data-stu-id="02660-817">Type</span></span>| <span data-ttu-id="02660-818">属性</span><span class="sxs-lookup"><span data-stu-id="02660-818">Attributes</span></span>| <span data-ttu-id="02660-819">说明</span><span class="sxs-lookup"><span data-stu-id="02660-819">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="02660-820">函数</span><span class="sxs-lookup"><span data-stu-id="02660-820">function</span></span>||<span data-ttu-id="02660-821">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="02660-821">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="02660-822">自定义属性作为 `asyncResult.value` 属性中的 [`CustomProperties`](/javascript/api/outlook_1_1/office.customproperties) 对象提供。</span><span class="sxs-lookup"><span data-stu-id="02660-822">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook_1_1/office.customproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="02660-823">此对象可用于获取、设置以及从项目中删除自定义属性，并将自定义属性集的更改重新保存到服务器。</span><span class="sxs-lookup"><span data-stu-id="02660-823">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`| <span data-ttu-id="02660-824">Object</span><span class="sxs-lookup"><span data-stu-id="02660-824">Object</span></span>| <span data-ttu-id="02660-825">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="02660-825">&lt;optional&gt;</span></span>|<span data-ttu-id="02660-826">开发人员可以提供他们想要在回调函数中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="02660-826">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="02660-827">此对象可以通过回调函数中的 `asyncResult.asyncContext` 属性进行访问。</span><span class="sxs-lookup"><span data-stu-id="02660-827">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="02660-828">要求</span><span class="sxs-lookup"><span data-stu-id="02660-828">Requirements</span></span>

|<span data-ttu-id="02660-829">要求</span><span class="sxs-lookup"><span data-stu-id="02660-829">Requirement</span></span>| <span data-ttu-id="02660-830">值</span><span class="sxs-lookup"><span data-stu-id="02660-830">Value</span></span>|
|---|---|
|[<span data-ttu-id="02660-831">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="02660-831">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="02660-832">1.0</span><span class="sxs-lookup"><span data-stu-id="02660-832">1.0</span></span>|
|[<span data-ttu-id="02660-833">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="02660-833">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="02660-834">ReadItem</span><span class="sxs-lookup"><span data-stu-id="02660-834">ReadItem</span></span>|
|[<span data-ttu-id="02660-835">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="02660-835">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="02660-836">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="02660-836">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="02660-837">示例</span><span class="sxs-lookup"><span data-stu-id="02660-837">Example</span></span>

<span data-ttu-id="02660-p154">以下代码示例显示了如何使用 `loadCustomPropertiesAsync` 方法异步加载特定于当前项目的自定义属性。该示例还显示了如何使用 `CustomProperties.saveAsync` 方法将这些属性重新保存到服务器。加载自定义属性后，该代码示例将使用 `CustomProperties.get` 方法读取自定义属性 `myProp`，使用 `CustomProperties.set` 方法写入自定义属性 `otherProp`，最后调用 `saveAsync` 方法保存这些自定义属性。</span><span class="sxs-lookup"><span data-stu-id="02660-p154">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

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

#### <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="02660-841">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="02660-841">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="02660-842">将附件从邮件或约会中删除。</span><span class="sxs-lookup"><span data-stu-id="02660-842">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="02660-843">`removeAttachmentAsync` 方法删除项目中带指定标识符的附件。</span><span class="sxs-lookup"><span data-stu-id="02660-843">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item.</span></span> <span data-ttu-id="02660-844">最佳做法是，仅当同一个邮件应用程序在同一会话中添加了一个附件时，你才应使用该附件标识符来删除该附件。</span><span class="sxs-lookup"><span data-stu-id="02660-844">As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session.</span></span> <span data-ttu-id="02660-845">在 web 和移动设备上的 Outlook 中, 附件标识符仅在同一个会话中有效。</span><span class="sxs-lookup"><span data-stu-id="02660-845">In Outlook on the web and mobile devices, the attachment identifier is valid only within the same session.</span></span> <span data-ttu-id="02660-846">当用户关闭应用程序，或者如果用户开始在内嵌窗体中撰写，并在随后弹出的内嵌窗体中继续在单独的窗口撰写时，会话即结束。</span><span class="sxs-lookup"><span data-stu-id="02660-846">A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="02660-847">参数</span><span class="sxs-lookup"><span data-stu-id="02660-847">Parameters</span></span>

|<span data-ttu-id="02660-848">名称</span><span class="sxs-lookup"><span data-stu-id="02660-848">Name</span></span>| <span data-ttu-id="02660-849">类型</span><span class="sxs-lookup"><span data-stu-id="02660-849">Type</span></span>| <span data-ttu-id="02660-850">属性</span><span class="sxs-lookup"><span data-stu-id="02660-850">Attributes</span></span>| <span data-ttu-id="02660-851">说明</span><span class="sxs-lookup"><span data-stu-id="02660-851">Description</span></span>|
|---|---|---|---|
|`attachmentId`| <span data-ttu-id="02660-852">字符串</span><span class="sxs-lookup"><span data-stu-id="02660-852">String</span></span>||<span data-ttu-id="02660-853">要删除的附件的标识符。</span><span class="sxs-lookup"><span data-stu-id="02660-853">The identifier of the attachment to remove.</span></span>|
|`options`| <span data-ttu-id="02660-854">对象</span><span class="sxs-lookup"><span data-stu-id="02660-854">Object</span></span>| <span data-ttu-id="02660-855">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="02660-855">&lt;optional&gt;</span></span>|<span data-ttu-id="02660-856">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="02660-856">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="02660-857">Object</span><span class="sxs-lookup"><span data-stu-id="02660-857">Object</span></span>| <span data-ttu-id="02660-858">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="02660-858">&lt;optional&gt;</span></span>|<span data-ttu-id="02660-859">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="02660-859">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="02660-860">函数</span><span class="sxs-lookup"><span data-stu-id="02660-860">function</span></span>| <span data-ttu-id="02660-861">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="02660-861">&lt;optional&gt;</span></span>|<span data-ttu-id="02660-862">方法完成后，使用单个参数 `callback`（一个 [`asyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `AsyncResult` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="02660-862">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="02660-863">如果删除附件失败，`asyncResult.error` 属性将包含一个说明失败原因的错误代码。</span><span class="sxs-lookup"><span data-stu-id="02660-863">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="02660-864">错误</span><span class="sxs-lookup"><span data-stu-id="02660-864">Errors</span></span>

| <span data-ttu-id="02660-865">错误代码</span><span class="sxs-lookup"><span data-stu-id="02660-865">Error code</span></span> | <span data-ttu-id="02660-866">说明</span><span class="sxs-lookup"><span data-stu-id="02660-866">Description</span></span> |
|------------|-------------|
| `InvalidAttachmentId` | <span data-ttu-id="02660-867">附件标识符不存在。</span><span class="sxs-lookup"><span data-stu-id="02660-867">The attachment identifier does not exist.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="02660-868">要求</span><span class="sxs-lookup"><span data-stu-id="02660-868">Requirements</span></span>

|<span data-ttu-id="02660-869">要求</span><span class="sxs-lookup"><span data-stu-id="02660-869">Requirement</span></span>| <span data-ttu-id="02660-870">值</span><span class="sxs-lookup"><span data-stu-id="02660-870">Value</span></span>|
|---|---|
|[<span data-ttu-id="02660-871">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="02660-871">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="02660-872">1.1</span><span class="sxs-lookup"><span data-stu-id="02660-872">1.1</span></span>|
|[<span data-ttu-id="02660-873">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="02660-873">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="02660-874">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="02660-874">ReadWriteItem</span></span>|
|[<span data-ttu-id="02660-875">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="02660-875">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="02660-876">撰写</span><span class="sxs-lookup"><span data-stu-id="02660-876">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="02660-877">示例</span><span class="sxs-lookup"><span data-stu-id="02660-877">Example</span></span>

<span data-ttu-id="02660-878">以下代码删除包含标识符 '0' 的附件。</span><span class="sxs-lookup"><span data-stu-id="02660-878">The following code removes an attachment with an identifier of '0'.</span></span>

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
