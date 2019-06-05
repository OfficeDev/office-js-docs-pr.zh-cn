---
title: "\"Context\"-\"邮箱\"。项目-要求集1。4"
description: ''
ms.date: 05/30/2019
localization_priority: Normal
ms.openlocfilehash: e2b8f70317f364137ffd2c3980c5f92c708a901e
ms.sourcegitcommit: 567aa05d6ee6b3639f65c50188df2331b7685857
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 06/04/2019
ms.locfileid: "34706342"
---
# <a name="item"></a><span data-ttu-id="3d5f7-102">item</span><span class="sxs-lookup"><span data-stu-id="3d5f7-102">item</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmditem"></a><span data-ttu-id="3d5f7-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span><span class="sxs-lookup"><span data-stu-id="3d5f7-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span></span>

<span data-ttu-id="3d5f7-p101">`item` 命名空间用于访问当前选定的邮件、会议请求或约会。可以通过使用 [itemType](#itemtype-officemailboxenumsitemtype) 属性确定 `item` 的类型。</span><span class="sxs-lookup"><span data-stu-id="3d5f7-p101">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="3d5f7-106">Requirements</span><span class="sxs-lookup"><span data-stu-id="3d5f7-106">Requirements</span></span>

|<span data-ttu-id="3d5f7-107">要求</span><span class="sxs-lookup"><span data-stu-id="3d5f7-107">Requirement</span></span>| <span data-ttu-id="3d5f7-108">值</span><span class="sxs-lookup"><span data-stu-id="3d5f7-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="3d5f7-109">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="3d5f7-109">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3d5f7-110">1.0</span><span class="sxs-lookup"><span data-stu-id="3d5f7-110">1.0</span></span>|
|[<span data-ttu-id="3d5f7-111">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="3d5f7-111">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3d5f7-112">受限</span><span class="sxs-lookup"><span data-stu-id="3d5f7-112">Restricted</span></span>|
|[<span data-ttu-id="3d5f7-113">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="3d5f7-113">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="3d5f7-114">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="3d5f7-114">Compose or Read</span></span>|

### <a name="example"></a><span data-ttu-id="3d5f7-115">示例</span><span class="sxs-lookup"><span data-stu-id="3d5f7-115">Example</span></span>

<span data-ttu-id="3d5f7-116">以下 JavaScript 代码示例显示了如何访问 Outlook 中当前项目的 `subject` 属性。</span><span class="sxs-lookup"><span data-stu-id="3d5f7-116">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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

### <a name="members"></a><span data-ttu-id="3d5f7-117">成员</span><span class="sxs-lookup"><span data-stu-id="3d5f7-117">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlook14officeattachmentdetails"></a><span data-ttu-id="3d5f7-118">附件: Array. <[AttachmentDetails](/javascript/api/outlook_1_4/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="3d5f7-118">attachments: Array.<[AttachmentDetails](/javascript/api/outlook_1_4/office.attachmentdetails)></span></span>

<span data-ttu-id="3d5f7-p102">获取项目的附件数组。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="3d5f7-p102">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="3d5f7-121">由于潜在的安全问题，某些类型的文件会受到 Outlook 阻止，并且不会返回。</span><span class="sxs-lookup"><span data-stu-id="3d5f7-121">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="3d5f7-122">如需了解更多信息，请参阅 [Outlook 中阻止的附件](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519)。</span><span class="sxs-lookup"><span data-stu-id="3d5f7-122">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="3d5f7-123">类型</span><span class="sxs-lookup"><span data-stu-id="3d5f7-123">Type</span></span>

*   <span data-ttu-id="3d5f7-124">Array.<[AttachmentDetails](/javascript/api/outlook_1_4/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="3d5f7-124">Array.<[AttachmentDetails](/javascript/api/outlook_1_4/office.attachmentdetails)></span></span>

##### <a name="requirements"></a><span data-ttu-id="3d5f7-125">要求</span><span class="sxs-lookup"><span data-stu-id="3d5f7-125">Requirements</span></span>

|<span data-ttu-id="3d5f7-126">要求</span><span class="sxs-lookup"><span data-stu-id="3d5f7-126">Requirement</span></span>| <span data-ttu-id="3d5f7-127">值</span><span class="sxs-lookup"><span data-stu-id="3d5f7-127">Value</span></span>|
|---|---|
|[<span data-ttu-id="3d5f7-128">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="3d5f7-128">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3d5f7-129">1.0</span><span class="sxs-lookup"><span data-stu-id="3d5f7-129">1.0</span></span>|
|[<span data-ttu-id="3d5f7-130">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="3d5f7-130">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3d5f7-131">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3d5f7-131">ReadItem</span></span>|
|[<span data-ttu-id="3d5f7-132">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="3d5f7-132">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="3d5f7-133">阅读</span><span class="sxs-lookup"><span data-stu-id="3d5f7-133">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="3d5f7-134">示例</span><span class="sxs-lookup"><span data-stu-id="3d5f7-134">Example</span></span>

<span data-ttu-id="3d5f7-135">以下代码使用当前项目上所有附件的详细信息构成 HTML 字符串。</span><span class="sxs-lookup"><span data-stu-id="3d5f7-135">The following code builds an HTML string with details of all attachments on the current item.</span></span>

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

#### <a name="bcc-recipientsjavascriptapioutlook14officerecipients"></a><span data-ttu-id="3d5f7-136">密件抄送:[收件人](/javascript/api/outlook_1_4/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="3d5f7-136">bcc: [Recipients](/javascript/api/outlook_1_4/office.recipients)</span></span>

<span data-ttu-id="3d5f7-137">获取一个对象, 该对象提供用于获取或更新邮件的密件抄送 (密件抄送) 行的方法。</span><span class="sxs-lookup"><span data-stu-id="3d5f7-137">Gets an object that provides methods to get or update the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="3d5f7-138">仅限撰写模式。</span><span class="sxs-lookup"><span data-stu-id="3d5f7-138">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="3d5f7-139">类型</span><span class="sxs-lookup"><span data-stu-id="3d5f7-139">Type</span></span>

*   [<span data-ttu-id="3d5f7-140">收件人</span><span class="sxs-lookup"><span data-stu-id="3d5f7-140">Recipients</span></span>](/javascript/api/outlook_1_4/office.recipients)

##### <a name="requirements"></a><span data-ttu-id="3d5f7-141">要求</span><span class="sxs-lookup"><span data-stu-id="3d5f7-141">Requirements</span></span>

|<span data-ttu-id="3d5f7-142">要求</span><span class="sxs-lookup"><span data-stu-id="3d5f7-142">Requirement</span></span>| <span data-ttu-id="3d5f7-143">值</span><span class="sxs-lookup"><span data-stu-id="3d5f7-143">Value</span></span>|
|---|---|
|[<span data-ttu-id="3d5f7-144">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="3d5f7-144">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3d5f7-145">1.1</span><span class="sxs-lookup"><span data-stu-id="3d5f7-145">1.1</span></span>|
|[<span data-ttu-id="3d5f7-146">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="3d5f7-146">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3d5f7-147">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3d5f7-147">ReadItem</span></span>|
|[<span data-ttu-id="3d5f7-148">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="3d5f7-148">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="3d5f7-149">撰写</span><span class="sxs-lookup"><span data-stu-id="3d5f7-149">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="3d5f7-150">示例</span><span class="sxs-lookup"><span data-stu-id="3d5f7-150">Example</span></span>

```javascript
Office.context.mailbox.item.bcc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.bcc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.bcc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfBccRecipients = asyncResult.value;
}
```

#### <a name="body-bodyjavascriptapioutlook14officebody"></a><span data-ttu-id="3d5f7-151">正文:[正文](/javascript/api/outlook_1_4/office.body)</span><span class="sxs-lookup"><span data-stu-id="3d5f7-151">body: [Body](/javascript/api/outlook_1_4/office.body)</span></span>

<span data-ttu-id="3d5f7-152">获取一个提供用于处理项目正文的方法的对象。</span><span class="sxs-lookup"><span data-stu-id="3d5f7-152">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="3d5f7-153">类型</span><span class="sxs-lookup"><span data-stu-id="3d5f7-153">Type</span></span>

*   [<span data-ttu-id="3d5f7-154">Body</span><span class="sxs-lookup"><span data-stu-id="3d5f7-154">Body</span></span>](/javascript/api/outlook_1_4/office.body)

##### <a name="requirements"></a><span data-ttu-id="3d5f7-155">要求</span><span class="sxs-lookup"><span data-stu-id="3d5f7-155">Requirements</span></span>

|<span data-ttu-id="3d5f7-156">要求</span><span class="sxs-lookup"><span data-stu-id="3d5f7-156">Requirement</span></span>| <span data-ttu-id="3d5f7-157">值</span><span class="sxs-lookup"><span data-stu-id="3d5f7-157">Value</span></span>|
|---|---|
|[<span data-ttu-id="3d5f7-158">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="3d5f7-158">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3d5f7-159">1.1</span><span class="sxs-lookup"><span data-stu-id="3d5f7-159">1.1</span></span>|
|[<span data-ttu-id="3d5f7-160">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="3d5f7-160">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3d5f7-161">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3d5f7-161">ReadItem</span></span>|
|[<span data-ttu-id="3d5f7-162">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="3d5f7-162">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="3d5f7-163">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="3d5f7-163">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="3d5f7-164">示例</span><span class="sxs-lookup"><span data-stu-id="3d5f7-164">Example</span></span>

<span data-ttu-id="3d5f7-165">本示例获取纯文本格式的邮件正文。</span><span class="sxs-lookup"><span data-stu-id="3d5f7-165">This example gets the body of the message in plain text.</span></span>

```javascript
Office.context.mailbox.item.body.getAsync(
  "text",
  { asyncContext: "This is passed to the callback" },
  function callback(result) {
    // Do something with the result.
  });

```

<span data-ttu-id="3d5f7-166">以下是传递到回调函数的结果参数的示例。</span><span class="sxs-lookup"><span data-stu-id="3d5f7-166">The following is an example of the result parameter passed to the callback function.</span></span>

```json
{
  "value": "TEXT of whole body (including threads below)",
  "status": "succeeded",
  "asyncContext": "This is passed to the callback"
}
```

#### <a name="cc-arrayemailaddressdetailsjavascriptapioutlook14officeemailaddressdetailsrecipientsjavascriptapioutlook14officerecipients"></a><span data-ttu-id="3d5f7-167"><[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)>|[收件人](/javascript/api/outlook_1_4/office.recipients)的抄送: Array</span><span class="sxs-lookup"><span data-stu-id="3d5f7-167">cc: Array.<[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_4/office.recipients)</span></span>

<span data-ttu-id="3d5f7-168">提供对邮件的抄送 (Cc) 收件人的访问权限。</span><span class="sxs-lookup"><span data-stu-id="3d5f7-168">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="3d5f7-169">对象的类型和访问级别取决于当前项目的模式。</span><span class="sxs-lookup"><span data-stu-id="3d5f7-169">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="3d5f7-170">阅读模式</span><span class="sxs-lookup"><span data-stu-id="3d5f7-170">Read mode</span></span>

<span data-ttu-id="3d5f7-p106">`cc` 属性返回包含邮件的**抄送**行上所列的每个收件人的 `EmailAddressDetails` 对象的数组。集合上限为 100 个成员。</span><span class="sxs-lookup"><span data-stu-id="3d5f7-p106">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.</span></span>

```javascript
console.log(JSON.stringify(Office.context.mailbox.item.cc));
```

##### <a name="compose-mode"></a><span data-ttu-id="3d5f7-173">撰写模式</span><span class="sxs-lookup"><span data-stu-id="3d5f7-173">Compose mode</span></span>

<span data-ttu-id="3d5f7-174">`cc` 属性返回一个 `Recipients` 对象，该对象提供用于获取或更新邮件的“**抄送**”行上收件人的方法。</span><span class="sxs-lookup"><span data-stu-id="3d5f7-174">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span>

```javascript
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="3d5f7-175">类型</span><span class="sxs-lookup"><span data-stu-id="3d5f7-175">Type</span></span>

*   <span data-ttu-id="3d5f7-176">Array.<[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_4/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="3d5f7-176">Array.<[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_4/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="3d5f7-177">要求</span><span class="sxs-lookup"><span data-stu-id="3d5f7-177">Requirements</span></span>

|<span data-ttu-id="3d5f7-178">要求</span><span class="sxs-lookup"><span data-stu-id="3d5f7-178">Requirement</span></span>| <span data-ttu-id="3d5f7-179">值</span><span class="sxs-lookup"><span data-stu-id="3d5f7-179">Value</span></span>|
|---|---|
|[<span data-ttu-id="3d5f7-180">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="3d5f7-180">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3d5f7-181">1.0</span><span class="sxs-lookup"><span data-stu-id="3d5f7-181">1.0</span></span>|
|[<span data-ttu-id="3d5f7-182">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="3d5f7-182">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3d5f7-183">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3d5f7-183">ReadItem</span></span>|
|[<span data-ttu-id="3d5f7-184">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="3d5f7-184">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="3d5f7-185">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="3d5f7-185">Compose or Read</span></span>|

#### <a name="nullable-conversationid-string"></a><span data-ttu-id="3d5f7-186">(可以为 null) conversationId: String</span><span class="sxs-lookup"><span data-stu-id="3d5f7-186">(nullable) conversationId: String</span></span>

<span data-ttu-id="3d5f7-187">获取包含特定消息的电子邮件会话的标识符。</span><span class="sxs-lookup"><span data-stu-id="3d5f7-187">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="3d5f7-p107">如果在阅读窗体或撰写窗体的回复中激活邮件应用程序，则此属性可以获得一个整数值。如果用户随后更改了回复邮件的主题（若发送回复），则该邮件的对话 ID 将改变且之前获取的值将不适用。</span><span class="sxs-lookup"><span data-stu-id="3d5f7-p107">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="3d5f7-p108">对于撰写窗体的新项目，此属性获得一个 null 值。如果用户设置一个主题并保存该项目，`conversationId` 属性将返回一个值。</span><span class="sxs-lookup"><span data-stu-id="3d5f7-p108">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="3d5f7-192">Type</span><span class="sxs-lookup"><span data-stu-id="3d5f7-192">Type</span></span>

*   <span data-ttu-id="3d5f7-193">String</span><span class="sxs-lookup"><span data-stu-id="3d5f7-193">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="3d5f7-194">要求</span><span class="sxs-lookup"><span data-stu-id="3d5f7-194">Requirements</span></span>

|<span data-ttu-id="3d5f7-195">要求</span><span class="sxs-lookup"><span data-stu-id="3d5f7-195">Requirement</span></span>| <span data-ttu-id="3d5f7-196">值</span><span class="sxs-lookup"><span data-stu-id="3d5f7-196">Value</span></span>|
|---|---|
|[<span data-ttu-id="3d5f7-197">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="3d5f7-197">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3d5f7-198">1.0</span><span class="sxs-lookup"><span data-stu-id="3d5f7-198">1.0</span></span>|
|[<span data-ttu-id="3d5f7-199">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="3d5f7-199">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3d5f7-200">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3d5f7-200">ReadItem</span></span>|
|[<span data-ttu-id="3d5f7-201">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="3d5f7-201">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="3d5f7-202">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="3d5f7-202">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="3d5f7-203">示例</span><span class="sxs-lookup"><span data-stu-id="3d5f7-203">Example</span></span>

```javascript
var conversationId = Office.context.mailbox.item.conversationId;
console.log("conversationId: " + conversationId);
```

#### <a name="datetimecreated-date"></a><span data-ttu-id="3d5f7-204">dateTimeCreated: Date</span><span class="sxs-lookup"><span data-stu-id="3d5f7-204">dateTimeCreated: Date</span></span>

<span data-ttu-id="3d5f7-p109">获取项目创建的日期和时间。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="3d5f7-p109">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="3d5f7-207">类型</span><span class="sxs-lookup"><span data-stu-id="3d5f7-207">Type</span></span>

*   <span data-ttu-id="3d5f7-208">日期</span><span class="sxs-lookup"><span data-stu-id="3d5f7-208">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="3d5f7-209">要求</span><span class="sxs-lookup"><span data-stu-id="3d5f7-209">Requirements</span></span>

|<span data-ttu-id="3d5f7-210">要求</span><span class="sxs-lookup"><span data-stu-id="3d5f7-210">Requirement</span></span>| <span data-ttu-id="3d5f7-211">值</span><span class="sxs-lookup"><span data-stu-id="3d5f7-211">Value</span></span>|
|---|---|
|[<span data-ttu-id="3d5f7-212">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="3d5f7-212">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3d5f7-213">1.0</span><span class="sxs-lookup"><span data-stu-id="3d5f7-213">1.0</span></span>|
|[<span data-ttu-id="3d5f7-214">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="3d5f7-214">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3d5f7-215">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3d5f7-215">ReadItem</span></span>|
|[<span data-ttu-id="3d5f7-216">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="3d5f7-216">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="3d5f7-217">阅读</span><span class="sxs-lookup"><span data-stu-id="3d5f7-217">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="3d5f7-218">示例</span><span class="sxs-lookup"><span data-stu-id="3d5f7-218">Example</span></span>

```javascript
var dateTimeCreated = Office.context.mailbox.item.dateTimeCreated;
console.log("Date and time created: " + dateTimeCreated);
```

#### <a name="datetimemodified-date"></a><span data-ttu-id="3d5f7-219">dateTimeModified: Date</span><span class="sxs-lookup"><span data-stu-id="3d5f7-219">dateTimeModified: Date</span></span>

<span data-ttu-id="3d5f7-p110">获取项目最近一次修改的日期和时间。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="3d5f7-p110">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="3d5f7-222">在 Outlook for iOS 或 Outlook for Android 中不支持此成员。</span><span class="sxs-lookup"><span data-stu-id="3d5f7-222">This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="type"></a><span data-ttu-id="3d5f7-223">类型</span><span class="sxs-lookup"><span data-stu-id="3d5f7-223">Type</span></span>

*   <span data-ttu-id="3d5f7-224">日期</span><span class="sxs-lookup"><span data-stu-id="3d5f7-224">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="3d5f7-225">要求</span><span class="sxs-lookup"><span data-stu-id="3d5f7-225">Requirements</span></span>

|<span data-ttu-id="3d5f7-226">要求</span><span class="sxs-lookup"><span data-stu-id="3d5f7-226">Requirement</span></span>| <span data-ttu-id="3d5f7-227">值</span><span class="sxs-lookup"><span data-stu-id="3d5f7-227">Value</span></span>|
|---|---|
|[<span data-ttu-id="3d5f7-228">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="3d5f7-228">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3d5f7-229">1.0</span><span class="sxs-lookup"><span data-stu-id="3d5f7-229">1.0</span></span>|
|[<span data-ttu-id="3d5f7-230">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="3d5f7-230">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3d5f7-231">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3d5f7-231">ReadItem</span></span>|
|[<span data-ttu-id="3d5f7-232">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="3d5f7-232">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="3d5f7-233">阅读</span><span class="sxs-lookup"><span data-stu-id="3d5f7-233">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="3d5f7-234">示例</span><span class="sxs-lookup"><span data-stu-id="3d5f7-234">Example</span></span>

```javascript
var dateTimeModified = Office.context.mailbox.item.dateTimeModified;
console.log("Date and time modified: " + dateTimeModified);
```

#### <a name="end-datetimejavascriptapioutlook14officetime"></a><span data-ttu-id="3d5f7-235">结束: 日期 |[时间](/javascript/api/outlook_1_4/office.time)</span><span class="sxs-lookup"><span data-stu-id="3d5f7-235">end: Date|[Time](/javascript/api/outlook_1_4/office.time)</span></span>

<span data-ttu-id="3d5f7-236">获取或设置约会结束的日期和时间。</span><span class="sxs-lookup"><span data-stu-id="3d5f7-236">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="3d5f7-p111">将 `end` 属性表示为协调世界时 (UTC) 的日期和时间值。可使用 [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) 方法将 end 属性值转换为客户端的本地日期和时间。</span><span class="sxs-lookup"><span data-stu-id="3d5f7-p111">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="3d5f7-239">阅读模式</span><span class="sxs-lookup"><span data-stu-id="3d5f7-239">Read mode</span></span>

<span data-ttu-id="3d5f7-240">`end` 属性返回 `Date` 对象。</span><span class="sxs-lookup"><span data-stu-id="3d5f7-240">The `end` property returns a `Date` object.</span></span>

```javascript
var end = Office.context.mailbox.item.end;
console.log("Appointment end: " + end);
```

##### <a name="compose-mode"></a><span data-ttu-id="3d5f7-241">撰写模式</span><span class="sxs-lookup"><span data-stu-id="3d5f7-241">Compose mode</span></span>

<span data-ttu-id="3d5f7-242">`end` 属性返回 `Time` 对象。</span><span class="sxs-lookup"><span data-stu-id="3d5f7-242">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="3d5f7-243">使用 [`Time.setAsync`](/javascript/api/outlook_1_4/office.time#setasync-datetime--options--callback-) 方法设置结束时间时，应使用 [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) 方法将客户端的本地时间转换为服务器的 UTC。</span><span class="sxs-lookup"><span data-stu-id="3d5f7-243">When you use the [`Time.setAsync`](/javascript/api/outlook_1_4/office.time#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="3d5f7-244">以下示例使用 `Time` 对象的 [`setAsync`](/javascript/api/outlook_1_4/office.time#setasync-datetime--options--callback-) 方法设置约会的结束时间。</span><span class="sxs-lookup"><span data-stu-id="3d5f7-244">The following example sets the end time of an appointment by using the [`setAsync`](/javascript/api/outlook_1_4/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="3d5f7-245">类型</span><span class="sxs-lookup"><span data-stu-id="3d5f7-245">Type</span></span>

*   <span data-ttu-id="3d5f7-246">Date | [Time](/javascript/api/outlook_1_4/office.time)</span><span class="sxs-lookup"><span data-stu-id="3d5f7-246">Date | [Time](/javascript/api/outlook_1_4/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="3d5f7-247">要求</span><span class="sxs-lookup"><span data-stu-id="3d5f7-247">Requirements</span></span>

|<span data-ttu-id="3d5f7-248">要求</span><span class="sxs-lookup"><span data-stu-id="3d5f7-248">Requirement</span></span>| <span data-ttu-id="3d5f7-249">值</span><span class="sxs-lookup"><span data-stu-id="3d5f7-249">Value</span></span>|
|---|---|
|[<span data-ttu-id="3d5f7-250">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="3d5f7-250">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3d5f7-251">1.0</span><span class="sxs-lookup"><span data-stu-id="3d5f7-251">1.0</span></span>|
|[<span data-ttu-id="3d5f7-252">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="3d5f7-252">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3d5f7-253">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3d5f7-253">ReadItem</span></span>|
|[<span data-ttu-id="3d5f7-254">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="3d5f7-254">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="3d5f7-255">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="3d5f7-255">Compose or Read</span></span>|

#### <a name="from-emailaddressdetailsjavascriptapioutlook14officeemailaddressdetails"></a><span data-ttu-id="3d5f7-256">发件人: [EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="3d5f7-256">from: [EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)</span></span>

<span data-ttu-id="3d5f7-p112">获取邮件发件人的电子邮件地址。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="3d5f7-p112">Gets the email address of the sender of a message. Read mode only.</span></span>

<span data-ttu-id="3d5f7-p113">`from` 和 [`sender`](#sender-emailaddressdetails) 属性表示同一个人，邮件由代理人发送的除外。在此情况下，`from` 属性表示代理程序，而 sender 属性表示代理人。</span><span class="sxs-lookup"><span data-stu-id="3d5f7-p113">The `from` and [`sender`](#sender-emailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="3d5f7-261">`from` 属性中 `EmailAddressDetails` 对象的 `recipientType` 属性为 `undefined`。</span><span class="sxs-lookup"><span data-stu-id="3d5f7-261">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="3d5f7-262">类型</span><span class="sxs-lookup"><span data-stu-id="3d5f7-262">Type</span></span>

*   [<span data-ttu-id="3d5f7-263">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="3d5f7-263">EmailAddressDetails</span></span>](/javascript/api/outlook_1_4/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="3d5f7-264">要求</span><span class="sxs-lookup"><span data-stu-id="3d5f7-264">Requirements</span></span>

|<span data-ttu-id="3d5f7-265">要求</span><span class="sxs-lookup"><span data-stu-id="3d5f7-265">Requirement</span></span>| <span data-ttu-id="3d5f7-266">值</span><span class="sxs-lookup"><span data-stu-id="3d5f7-266">Value</span></span>|
|---|---|
|[<span data-ttu-id="3d5f7-267">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="3d5f7-267">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3d5f7-268">1.0</span><span class="sxs-lookup"><span data-stu-id="3d5f7-268">1.0</span></span>|
|[<span data-ttu-id="3d5f7-269">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="3d5f7-269">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3d5f7-270">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3d5f7-270">ReadItem</span></span>|
|[<span data-ttu-id="3d5f7-271">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="3d5f7-271">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="3d5f7-272">阅读</span><span class="sxs-lookup"><span data-stu-id="3d5f7-272">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="3d5f7-273">示例</span><span class="sxs-lookup"><span data-stu-id="3d5f7-273">Example</span></span>

```javascript
var from = Office.context.mailbox.item.from;
console.log("From " + from);
```

#### <a name="internetmessageid-string"></a><span data-ttu-id="3d5f7-274">internetMessageId: String</span><span class="sxs-lookup"><span data-stu-id="3d5f7-274">internetMessageId: String</span></span>

<span data-ttu-id="3d5f7-p114">获取电子邮件的 Internet 消息标识符。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="3d5f7-p114">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="3d5f7-277">Type</span><span class="sxs-lookup"><span data-stu-id="3d5f7-277">Type</span></span>

*   <span data-ttu-id="3d5f7-278">String</span><span class="sxs-lookup"><span data-stu-id="3d5f7-278">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="3d5f7-279">要求</span><span class="sxs-lookup"><span data-stu-id="3d5f7-279">Requirements</span></span>

|<span data-ttu-id="3d5f7-280">要求</span><span class="sxs-lookup"><span data-stu-id="3d5f7-280">Requirement</span></span>| <span data-ttu-id="3d5f7-281">值</span><span class="sxs-lookup"><span data-stu-id="3d5f7-281">Value</span></span>|
|---|---|
|[<span data-ttu-id="3d5f7-282">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="3d5f7-282">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3d5f7-283">1.0</span><span class="sxs-lookup"><span data-stu-id="3d5f7-283">1.0</span></span>|
|[<span data-ttu-id="3d5f7-284">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="3d5f7-284">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3d5f7-285">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3d5f7-285">ReadItem</span></span>|
|[<span data-ttu-id="3d5f7-286">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="3d5f7-286">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="3d5f7-287">阅读</span><span class="sxs-lookup"><span data-stu-id="3d5f7-287">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="3d5f7-288">示例</span><span class="sxs-lookup"><span data-stu-id="3d5f7-288">Example</span></span>

```javascript
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

#### <a name="itemclass-string"></a><span data-ttu-id="3d5f7-289">itemClass: String</span><span class="sxs-lookup"><span data-stu-id="3d5f7-289">itemClass: String</span></span>

<span data-ttu-id="3d5f7-p115">获取选定项目的 Exchange Web 服务项目类。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="3d5f7-p115">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="3d5f7-p116">`itemClass` 属性指定所选项目的邮件类别。以下是邮件或约会项目的默认邮件类别。</span><span class="sxs-lookup"><span data-stu-id="3d5f7-p116">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

| <span data-ttu-id="3d5f7-294">类型</span><span class="sxs-lookup"><span data-stu-id="3d5f7-294">Type</span></span> | <span data-ttu-id="3d5f7-295">说明</span><span class="sxs-lookup"><span data-stu-id="3d5f7-295">Description</span></span> | <span data-ttu-id="3d5f7-296">项目类</span><span class="sxs-lookup"><span data-stu-id="3d5f7-296">item class</span></span> |
| --- | --- | --- |
| <span data-ttu-id="3d5f7-297">约会项目</span><span class="sxs-lookup"><span data-stu-id="3d5f7-297">Appointment items</span></span> | <span data-ttu-id="3d5f7-298">这些是项目类 `IPM.Appointment` 或 `IPM.Appointment.Occurrence` 的日历项目。</span><span class="sxs-lookup"><span data-stu-id="3d5f7-298">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurrence`.</span></span> | `IPM.Appointment`<br />`IPM.Appointment.Occurrence` |
| <span data-ttu-id="3d5f7-299">邮件项目</span><span class="sxs-lookup"><span data-stu-id="3d5f7-299">Message items</span></span> | <span data-ttu-id="3d5f7-300">这些项目包括具有默认邮件类别 `IPM.Note` 的电子邮件，以及使用 `IPM.Schedule.Meeting` 作为基础邮件类别的会议请求、响应和取消。</span><span class="sxs-lookup"><span data-stu-id="3d5f7-300">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span> | `IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled` |

<span data-ttu-id="3d5f7-301">你可以创建用于扩展默认邮件类别的自定义邮件类别，例如，自定义约会邮件类别 `IPM.Appointment.Contoso`。</span><span class="sxs-lookup"><span data-stu-id="3d5f7-301">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="3d5f7-302">Type</span><span class="sxs-lookup"><span data-stu-id="3d5f7-302">Type</span></span>

*   <span data-ttu-id="3d5f7-303">String</span><span class="sxs-lookup"><span data-stu-id="3d5f7-303">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="3d5f7-304">要求</span><span class="sxs-lookup"><span data-stu-id="3d5f7-304">Requirements</span></span>

|<span data-ttu-id="3d5f7-305">要求</span><span class="sxs-lookup"><span data-stu-id="3d5f7-305">Requirement</span></span>| <span data-ttu-id="3d5f7-306">值</span><span class="sxs-lookup"><span data-stu-id="3d5f7-306">Value</span></span>|
|---|---|
|[<span data-ttu-id="3d5f7-307">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="3d5f7-307">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3d5f7-308">1.0</span><span class="sxs-lookup"><span data-stu-id="3d5f7-308">1.0</span></span>|
|[<span data-ttu-id="3d5f7-309">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="3d5f7-309">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3d5f7-310">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3d5f7-310">ReadItem</span></span>|
|[<span data-ttu-id="3d5f7-311">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="3d5f7-311">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="3d5f7-312">阅读</span><span class="sxs-lookup"><span data-stu-id="3d5f7-312">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="3d5f7-313">示例</span><span class="sxs-lookup"><span data-stu-id="3d5f7-313">Example</span></span>

```javascript
var itemClass = Office.context.mailbox.item.itemClass;
console.log("Item class: " + itemClass);
```

#### <a name="nullable-itemid-string"></a><span data-ttu-id="3d5f7-314">(可以为 null) itemId: String</span><span class="sxs-lookup"><span data-stu-id="3d5f7-314">(nullable) itemId: String</span></span>

<span data-ttu-id="3d5f7-p117">获取当前项目的 Exchange Web 服务项目标识符。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="3d5f7-p117">Gets the Exchange Web Services item identifier for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="3d5f7-317">`itemId` 属性返回的标识符与 Exchange Web 服务项目标识符相同。</span><span class="sxs-lookup"><span data-stu-id="3d5f7-317">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="3d5f7-318">`itemId` 属性与 Outlook 条目 ID 或 Outlook REST API 使用的 ID 不同。</span><span class="sxs-lookup"><span data-stu-id="3d5f7-318">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="3d5f7-319">使用此值进行 REST API 调用前，应使用 [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string) 对它进行转换。</span><span class="sxs-lookup"><span data-stu-id="3d5f7-319">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="3d5f7-320">有关详细信息，请参阅[从 Outlook 加载项使用 Outlook REST API](/outlook/add-ins/use-rest-api#get-the-item-id)。</span><span class="sxs-lookup"><span data-stu-id="3d5f7-320">For more details, see [Use the Outlook REST APIs from an Outlook add-in](/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

<span data-ttu-id="3d5f7-p119">`itemId` 属性在撰写模式下不可用。如果需要项目标识符，[`saveAsync`](#saveasyncoptions-callback) 方法可用于将项目保存到存储，这将在回调函数的 [`AsyncResult.value`](/javascript/api/office/office.asyncresult) 参数中返回项目标识符。</span><span class="sxs-lookup"><span data-stu-id="3d5f7-p119">The `itemId` property is not available in compose mode. If an item identifier is required, the [`saveAsync`](#saveasyncoptions-callback) method can be used to save the item to the store, which will return the item identifier in the [`AsyncResult.value`](/javascript/api/office/office.asyncresult) parameter in the callback function.</span></span>

##### <a name="type"></a><span data-ttu-id="3d5f7-323">Type</span><span class="sxs-lookup"><span data-stu-id="3d5f7-323">Type</span></span>

*   <span data-ttu-id="3d5f7-324">String</span><span class="sxs-lookup"><span data-stu-id="3d5f7-324">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="3d5f7-325">要求</span><span class="sxs-lookup"><span data-stu-id="3d5f7-325">Requirements</span></span>

|<span data-ttu-id="3d5f7-326">要求</span><span class="sxs-lookup"><span data-stu-id="3d5f7-326">Requirement</span></span>| <span data-ttu-id="3d5f7-327">值</span><span class="sxs-lookup"><span data-stu-id="3d5f7-327">Value</span></span>|
|---|---|
|[<span data-ttu-id="3d5f7-328">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="3d5f7-328">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3d5f7-329">1.0</span><span class="sxs-lookup"><span data-stu-id="3d5f7-329">1.0</span></span>|
|[<span data-ttu-id="3d5f7-330">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="3d5f7-330">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3d5f7-331">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3d5f7-331">ReadItem</span></span>|
|[<span data-ttu-id="3d5f7-332">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="3d5f7-332">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="3d5f7-333">阅读</span><span class="sxs-lookup"><span data-stu-id="3d5f7-333">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="3d5f7-334">示例</span><span class="sxs-lookup"><span data-stu-id="3d5f7-334">Example</span></span>

<span data-ttu-id="3d5f7-p120">以下代码检查项目标识符是否存在。如果 `itemId` 属性返回 `null` 或 `undefined`，则将项目保存到存储，并从异步结果中获取项目标识符。</span><span class="sxs-lookup"><span data-stu-id="3d5f7-p120">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

```javascript
var itemId = Office.context.mailbox.item.itemId;
if (itemId === null || itemId == undefined) {
  Office.context.mailbox.item.saveAsync(function(result) {
    itemId = result.value;
  });
}
```

#### <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlook14officemailboxenumsitemtype"></a><span data-ttu-id="3d5f7-337">itemType: [MailboxEnums](/javascript/api/outlook_1_4/office.mailboxenums.itemtype)</span><span class="sxs-lookup"><span data-stu-id="3d5f7-337">itemType: [Office.MailboxEnums.ItemType](/javascript/api/outlook_1_4/office.mailboxenums.itemtype)</span></span>

<span data-ttu-id="3d5f7-338">获取实例表示的项的类型。</span><span class="sxs-lookup"><span data-stu-id="3d5f7-338">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="3d5f7-339">`itemType` 属性返回其中一个 `ItemType` 枚举值，指示 `item` 对象实例是邮件还是约会。</span><span class="sxs-lookup"><span data-stu-id="3d5f7-339">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="3d5f7-340">类型</span><span class="sxs-lookup"><span data-stu-id="3d5f7-340">Type</span></span>

*   [<span data-ttu-id="3d5f7-341">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="3d5f7-341">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook_1_4/office.mailboxenums.itemtype)

##### <a name="requirements"></a><span data-ttu-id="3d5f7-342">要求</span><span class="sxs-lookup"><span data-stu-id="3d5f7-342">Requirements</span></span>

|<span data-ttu-id="3d5f7-343">要求</span><span class="sxs-lookup"><span data-stu-id="3d5f7-343">Requirement</span></span>| <span data-ttu-id="3d5f7-344">值</span><span class="sxs-lookup"><span data-stu-id="3d5f7-344">Value</span></span>|
|---|---|
|[<span data-ttu-id="3d5f7-345">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="3d5f7-345">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3d5f7-346">1.0</span><span class="sxs-lookup"><span data-stu-id="3d5f7-346">1.0</span></span>|
|[<span data-ttu-id="3d5f7-347">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="3d5f7-347">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3d5f7-348">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3d5f7-348">ReadItem</span></span>|
|[<span data-ttu-id="3d5f7-349">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="3d5f7-349">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="3d5f7-350">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="3d5f7-350">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="3d5f7-351">示例</span><span class="sxs-lookup"><span data-stu-id="3d5f7-351">Example</span></span>

```javascript
if (Office.context.mailbox.item.itemType === Office.MailboxEnums.ItemType.Message) {
  // Do something.
} else {
  // Do something else.
}
```

#### <a name="location-stringlocationjavascriptapioutlook14officelocation"></a><span data-ttu-id="3d5f7-352">位置: 字符串 |[位置](/javascript/api/outlook_1_4/office.location)</span><span class="sxs-lookup"><span data-stu-id="3d5f7-352">location: String|[Location](/javascript/api/outlook_1_4/office.location)</span></span>

<span data-ttu-id="3d5f7-353">获取或设置约会的位置。</span><span class="sxs-lookup"><span data-stu-id="3d5f7-353">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="3d5f7-354">阅读模式</span><span class="sxs-lookup"><span data-stu-id="3d5f7-354">Read mode</span></span>

<span data-ttu-id="3d5f7-355">`location` 属性返回一个包含约会位置的字符串。</span><span class="sxs-lookup"><span data-stu-id="3d5f7-355">The `location` property returns a string that contains the location of the appointment.</span></span>

```javascript
var location = Office.context.mailbox.item.location;
console.log("location: " + location);
```

##### <a name="compose-mode"></a><span data-ttu-id="3d5f7-356">撰写模式</span><span class="sxs-lookup"><span data-stu-id="3d5f7-356">Compose mode</span></span>

<span data-ttu-id="3d5f7-357">`location` 属性返回一个 `Location` 对象，该对象提供用于获取和设置约会位置的方法。</span><span class="sxs-lookup"><span data-stu-id="3d5f7-357">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

```javascript
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="3d5f7-358">类型</span><span class="sxs-lookup"><span data-stu-id="3d5f7-358">Type</span></span>

*   <span data-ttu-id="3d5f7-359">String | [Location](/javascript/api/outlook_1_4/office.location)</span><span class="sxs-lookup"><span data-stu-id="3d5f7-359">String | [Location](/javascript/api/outlook_1_4/office.location)</span></span>

##### <a name="requirements"></a><span data-ttu-id="3d5f7-360">要求</span><span class="sxs-lookup"><span data-stu-id="3d5f7-360">Requirements</span></span>

|<span data-ttu-id="3d5f7-361">要求</span><span class="sxs-lookup"><span data-stu-id="3d5f7-361">Requirement</span></span>| <span data-ttu-id="3d5f7-362">值</span><span class="sxs-lookup"><span data-stu-id="3d5f7-362">Value</span></span>|
|---|---|
|[<span data-ttu-id="3d5f7-363">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="3d5f7-363">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3d5f7-364">1.0</span><span class="sxs-lookup"><span data-stu-id="3d5f7-364">1.0</span></span>|
|[<span data-ttu-id="3d5f7-365">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="3d5f7-365">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3d5f7-366">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3d5f7-366">ReadItem</span></span>|
|[<span data-ttu-id="3d5f7-367">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="3d5f7-367">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="3d5f7-368">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="3d5f7-368">Compose or Read</span></span>|

#### <a name="normalizedsubject-string"></a><span data-ttu-id="3d5f7-369">normalizedSubject: String</span><span class="sxs-lookup"><span data-stu-id="3d5f7-369">normalizedSubject: String</span></span>

<span data-ttu-id="3d5f7-p121">获取删除了所有前缀（包括 `RE:` 和 `FWD:`）的项目主题。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="3d5f7-p121">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="3d5f7-p122">normalizedSubject 属性获取包含由电子邮件程序添加的任何标准前缀（如 `RE:` 和 `FW:`）的项目主题。若要获取包含完整前缀的项目主题，请使用 [`subject`](#subject-stringsubject) 属性。</span><span class="sxs-lookup"><span data-stu-id="3d5f7-p122">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="3d5f7-374">Type</span><span class="sxs-lookup"><span data-stu-id="3d5f7-374">Type</span></span>

*   <span data-ttu-id="3d5f7-375">String</span><span class="sxs-lookup"><span data-stu-id="3d5f7-375">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="3d5f7-376">要求</span><span class="sxs-lookup"><span data-stu-id="3d5f7-376">Requirements</span></span>

|<span data-ttu-id="3d5f7-377">要求</span><span class="sxs-lookup"><span data-stu-id="3d5f7-377">Requirement</span></span>| <span data-ttu-id="3d5f7-378">值</span><span class="sxs-lookup"><span data-stu-id="3d5f7-378">Value</span></span>|
|---|---|
|[<span data-ttu-id="3d5f7-379">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="3d5f7-379">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3d5f7-380">1.0</span><span class="sxs-lookup"><span data-stu-id="3d5f7-380">1.0</span></span>|
|[<span data-ttu-id="3d5f7-381">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="3d5f7-381">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3d5f7-382">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3d5f7-382">ReadItem</span></span>|
|[<span data-ttu-id="3d5f7-383">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="3d5f7-383">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="3d5f7-384">阅读</span><span class="sxs-lookup"><span data-stu-id="3d5f7-384">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="3d5f7-385">示例</span><span class="sxs-lookup"><span data-stu-id="3d5f7-385">Example</span></span>

```javascript
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
console.log("Normalized subject: " + normalizedSubject);
```

#### <a name="notificationmessages-notificationmessagesjavascriptapioutlook14officenotificationmessages"></a><span data-ttu-id="3d5f7-386">notificationMessages: [notificationMessages](/javascript/api/outlook_1_4/office.notificationmessages)</span><span class="sxs-lookup"><span data-stu-id="3d5f7-386">notificationMessages: [NotificationMessages](/javascript/api/outlook_1_4/office.notificationmessages)</span></span>

<span data-ttu-id="3d5f7-387">获取项目的通知邮件。</span><span class="sxs-lookup"><span data-stu-id="3d5f7-387">Gets the notification messages for an item.</span></span>

##### <a name="type"></a><span data-ttu-id="3d5f7-388">类型</span><span class="sxs-lookup"><span data-stu-id="3d5f7-388">Type</span></span>

*   [<span data-ttu-id="3d5f7-389">NotificationMessages</span><span class="sxs-lookup"><span data-stu-id="3d5f7-389">NotificationMessages</span></span>](/javascript/api/outlook_1_4/office.notificationmessages)

##### <a name="requirements"></a><span data-ttu-id="3d5f7-390">要求</span><span class="sxs-lookup"><span data-stu-id="3d5f7-390">Requirements</span></span>

|<span data-ttu-id="3d5f7-391">要求</span><span class="sxs-lookup"><span data-stu-id="3d5f7-391">Requirement</span></span>| <span data-ttu-id="3d5f7-392">值</span><span class="sxs-lookup"><span data-stu-id="3d5f7-392">Value</span></span>|
|---|---|
|[<span data-ttu-id="3d5f7-393">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="3d5f7-393">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3d5f7-394">1.3</span><span class="sxs-lookup"><span data-stu-id="3d5f7-394">1.3</span></span>|
|[<span data-ttu-id="3d5f7-395">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="3d5f7-395">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3d5f7-396">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3d5f7-396">ReadItem</span></span>|
|[<span data-ttu-id="3d5f7-397">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="3d5f7-397">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="3d5f7-398">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="3d5f7-398">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="3d5f7-399">示例</span><span class="sxs-lookup"><span data-stu-id="3d5f7-399">Example</span></span>

```javascript
// Get all notifications.
Office.context.mailbox.item.notificationMessages.getAllAsync(
  function (asyncResult) {
    console.log(JSON.stringify(asyncResult));
  }
);
```

#### <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlook14officeemailaddressdetailsrecipientsjavascriptapioutlook14officerecipients"></a><span data-ttu-id="3d5f7-400">optionalAttendees: <[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)>|[收件人](/javascript/api/outlook_1_4/office.recipients)的数组</span><span class="sxs-lookup"><span data-stu-id="3d5f7-400">optionalAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_4/office.recipients)</span></span>

<span data-ttu-id="3d5f7-401">提供对事件的可选与会者的访问权限。</span><span class="sxs-lookup"><span data-stu-id="3d5f7-401">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="3d5f7-402">对象的类型和访问级别取决于当前项目的模式。</span><span class="sxs-lookup"><span data-stu-id="3d5f7-402">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="3d5f7-403">阅读模式</span><span class="sxs-lookup"><span data-stu-id="3d5f7-403">Read mode</span></span>

<span data-ttu-id="3d5f7-404">`optionalAttendees` 属性返回一个数组，其中包含每个可选与会者的 `EmailAddressDetails` 对象。</span><span class="sxs-lookup"><span data-stu-id="3d5f7-404">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span>

```javascript
var optionalAttendees = Office.context.mailbox.item.optionalAttendees;
console.log("Optional attendees: " + JSON.stringify(optionalAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="3d5f7-405">撰写模式</span><span class="sxs-lookup"><span data-stu-id="3d5f7-405">Compose mode</span></span>

<span data-ttu-id="3d5f7-406">`optionalAttendees` 属性返回一个 `Recipients` 对象，该对象提供用于获取或更新可选与会者的方法。</span><span class="sxs-lookup"><span data-stu-id="3d5f7-406">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span>

```javascript
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="3d5f7-407">类型</span><span class="sxs-lookup"><span data-stu-id="3d5f7-407">Type</span></span>

*   <span data-ttu-id="3d5f7-408">Array.<[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_4/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="3d5f7-408">Array.<[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_4/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="3d5f7-409">要求</span><span class="sxs-lookup"><span data-stu-id="3d5f7-409">Requirements</span></span>

|<span data-ttu-id="3d5f7-410">要求</span><span class="sxs-lookup"><span data-stu-id="3d5f7-410">Requirement</span></span>| <span data-ttu-id="3d5f7-411">值</span><span class="sxs-lookup"><span data-stu-id="3d5f7-411">Value</span></span>|
|---|---|
|[<span data-ttu-id="3d5f7-412">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="3d5f7-412">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3d5f7-413">1.0</span><span class="sxs-lookup"><span data-stu-id="3d5f7-413">1.0</span></span>|
|[<span data-ttu-id="3d5f7-414">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="3d5f7-414">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3d5f7-415">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3d5f7-415">ReadItem</span></span>|
|[<span data-ttu-id="3d5f7-416">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="3d5f7-416">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="3d5f7-417">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="3d5f7-417">Compose or Read</span></span>|

#### <a name="organizer-emailaddressdetailsjavascriptapioutlook14officeemailaddressdetails"></a><span data-ttu-id="3d5f7-418">组织者: [EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="3d5f7-418">organizer: [EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)</span></span>

<span data-ttu-id="3d5f7-p124">获取指定会议的会议组织者的电子邮件地址。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="3d5f7-p124">Gets the email address of the meeting organizer for a specified meeting. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="3d5f7-421">类型</span><span class="sxs-lookup"><span data-stu-id="3d5f7-421">Type</span></span>

*   [<span data-ttu-id="3d5f7-422">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="3d5f7-422">EmailAddressDetails</span></span>](/javascript/api/outlook_1_4/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="3d5f7-423">要求</span><span class="sxs-lookup"><span data-stu-id="3d5f7-423">Requirements</span></span>

|<span data-ttu-id="3d5f7-424">要求</span><span class="sxs-lookup"><span data-stu-id="3d5f7-424">Requirement</span></span>| <span data-ttu-id="3d5f7-425">值</span><span class="sxs-lookup"><span data-stu-id="3d5f7-425">Value</span></span>|
|---|---|
|[<span data-ttu-id="3d5f7-426">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="3d5f7-426">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3d5f7-427">1.0</span><span class="sxs-lookup"><span data-stu-id="3d5f7-427">1.0</span></span>|
|[<span data-ttu-id="3d5f7-428">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="3d5f7-428">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3d5f7-429">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3d5f7-429">ReadItem</span></span>|
|[<span data-ttu-id="3d5f7-430">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="3d5f7-430">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="3d5f7-431">阅读</span><span class="sxs-lookup"><span data-stu-id="3d5f7-431">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="3d5f7-432">示例</span><span class="sxs-lookup"><span data-stu-id="3d5f7-432">Example</span></span>

```javascript
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
console.log("Organizer: " + organizerName + " (" + organizerAddress + ")");
```

#### <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlook14officeemailaddressdetailsrecipientsjavascriptapioutlook14officerecipients"></a><span data-ttu-id="3d5f7-433">requiredAttendees: <[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)>|[收件人](/javascript/api/outlook_1_4/office.recipients)的数组</span><span class="sxs-lookup"><span data-stu-id="3d5f7-433">requiredAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_4/office.recipients)</span></span>

<span data-ttu-id="3d5f7-434">提供对事件的必需与会者的访问权限。</span><span class="sxs-lookup"><span data-stu-id="3d5f7-434">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="3d5f7-435">对象的类型和访问级别取决于当前项目的模式。</span><span class="sxs-lookup"><span data-stu-id="3d5f7-435">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="3d5f7-436">阅读模式</span><span class="sxs-lookup"><span data-stu-id="3d5f7-436">Read mode</span></span>

<span data-ttu-id="3d5f7-437">`requiredAttendees` 属性返回一个数组，其中包含每个必需与会者的 `EmailAddressDetails` 对象。</span><span class="sxs-lookup"><span data-stu-id="3d5f7-437">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span>

```javascript
var requiredAttendees = Office.context.mailbox.item.requiredAttendees;
console.log("Required attendees: " + JSON.stringify(requiredAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="3d5f7-438">撰写模式</span><span class="sxs-lookup"><span data-stu-id="3d5f7-438">Compose mode</span></span>

<span data-ttu-id="3d5f7-439">`requiredAttendees` 属性返回一个 `Recipients` 对象，该对象提供用于获取或更新必需与会者的方法。</span><span class="sxs-lookup"><span data-stu-id="3d5f7-439">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span>

```javascript
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
  console.log(JSON.stringify(arrayOfRequiredAttendeesRecipients));
}
```

##### <a name="type"></a><span data-ttu-id="3d5f7-440">类型</span><span class="sxs-lookup"><span data-stu-id="3d5f7-440">Type</span></span>

*   <span data-ttu-id="3d5f7-441">Array.<[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_4/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="3d5f7-441">Array.<[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_4/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="3d5f7-442">要求</span><span class="sxs-lookup"><span data-stu-id="3d5f7-442">Requirements</span></span>

|<span data-ttu-id="3d5f7-443">要求</span><span class="sxs-lookup"><span data-stu-id="3d5f7-443">Requirement</span></span>| <span data-ttu-id="3d5f7-444">值</span><span class="sxs-lookup"><span data-stu-id="3d5f7-444">Value</span></span>|
|---|---|
|[<span data-ttu-id="3d5f7-445">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="3d5f7-445">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3d5f7-446">1.0</span><span class="sxs-lookup"><span data-stu-id="3d5f7-446">1.0</span></span>|
|[<span data-ttu-id="3d5f7-447">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="3d5f7-447">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3d5f7-448">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3d5f7-448">ReadItem</span></span>|
|[<span data-ttu-id="3d5f7-449">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="3d5f7-449">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="3d5f7-450">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="3d5f7-450">Compose or Read</span></span>|

#### <a name="sender-emailaddressdetailsjavascriptapioutlook14officeemailaddressdetails"></a><span data-ttu-id="3d5f7-451">发件人: [EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="3d5f7-451">sender: [EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)</span></span>

<span data-ttu-id="3d5f7-p126">获取电子邮件发件人的电子邮件地址。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="3d5f7-p126">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="3d5f7-p127">[`from`](#from-emailaddressdetails) 和 `sender` 属性表示同一个人，邮件由代理人发送的除外。在此情况下，`from` 属性表示代理程序，而 sender 属性表示代理人。</span><span class="sxs-lookup"><span data-stu-id="3d5f7-p127">The [`from`](#from-emailaddressdetails) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="3d5f7-456">`sender` 属性中 `EmailAddressDetails` 对象的 `recipientType` 属性为 `undefined`。</span><span class="sxs-lookup"><span data-stu-id="3d5f7-456">The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="3d5f7-457">类型</span><span class="sxs-lookup"><span data-stu-id="3d5f7-457">Type</span></span>

*   [<span data-ttu-id="3d5f7-458">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="3d5f7-458">EmailAddressDetails</span></span>](/javascript/api/outlook_1_4/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="3d5f7-459">要求</span><span class="sxs-lookup"><span data-stu-id="3d5f7-459">Requirements</span></span>

|<span data-ttu-id="3d5f7-460">要求</span><span class="sxs-lookup"><span data-stu-id="3d5f7-460">Requirement</span></span>| <span data-ttu-id="3d5f7-461">值</span><span class="sxs-lookup"><span data-stu-id="3d5f7-461">Value</span></span>|
|---|---|
|[<span data-ttu-id="3d5f7-462">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="3d5f7-462">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3d5f7-463">1.0</span><span class="sxs-lookup"><span data-stu-id="3d5f7-463">1.0</span></span>|
|[<span data-ttu-id="3d5f7-464">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="3d5f7-464">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3d5f7-465">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3d5f7-465">ReadItem</span></span>|
|[<span data-ttu-id="3d5f7-466">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="3d5f7-466">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="3d5f7-467">阅读</span><span class="sxs-lookup"><span data-stu-id="3d5f7-467">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="3d5f7-468">示例</span><span class="sxs-lookup"><span data-stu-id="3d5f7-468">Example</span></span>

```javascript
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
console.log("Sender: " + senderName + " (" + senderAddress + ")");
```

#### <a name="start-datetimejavascriptapioutlook14officetime"></a><span data-ttu-id="3d5f7-469">开始日期: 日期 |[时间](/javascript/api/outlook_1_4/office.time)</span><span class="sxs-lookup"><span data-stu-id="3d5f7-469">start: Date|[Time](/javascript/api/outlook_1_4/office.time)</span></span>

<span data-ttu-id="3d5f7-470">获取或设置约会开始的日期和时间。</span><span class="sxs-lookup"><span data-stu-id="3d5f7-470">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="3d5f7-p128">将 `start` 属性表示为协调世界时 (UTC) 的日期和时间值。可使用 [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) 方法将该值转换为客户端的本地日期和时间。</span><span class="sxs-lookup"><span data-stu-id="3d5f7-p128">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="3d5f7-473">阅读模式</span><span class="sxs-lookup"><span data-stu-id="3d5f7-473">Read mode</span></span>

<span data-ttu-id="3d5f7-474">`start` 属性返回 `Date` 对象。</span><span class="sxs-lookup"><span data-stu-id="3d5f7-474">The `start` property returns a `Date` object.</span></span>

```javascript
var start = Office.context.mailbox.item.start;
console.log("Appointment start: " + JSON.stringify(start));
```

##### <a name="compose-mode"></a><span data-ttu-id="3d5f7-475">撰写模式</span><span class="sxs-lookup"><span data-stu-id="3d5f7-475">Compose mode</span></span>

<span data-ttu-id="3d5f7-476">`start` 属性返回 `Time` 对象。</span><span class="sxs-lookup"><span data-stu-id="3d5f7-476">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="3d5f7-477">使用 [`Time.setAsync`](/javascript/api/outlook_1_4/office.time#setasync-datetime--options--callback-) 方法设置开始时间时，应使用 [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) 方法将客户端的本地时间转换为服务器的 UTC。</span><span class="sxs-lookup"><span data-stu-id="3d5f7-477">When you use the [`Time.setAsync`](/javascript/api/outlook_1_4/office.time#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="3d5f7-478">以下示例通过使用 `Time` 对象的 [`setAsync`](/javascript/api/outlook_1_4/office.time#setasync-datetime--options--callback-) 方法，设置撰写模式下约会的开始时间。</span><span class="sxs-lookup"><span data-stu-id="3d5f7-478">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_4/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="3d5f7-479">类型</span><span class="sxs-lookup"><span data-stu-id="3d5f7-479">Type</span></span>

*   <span data-ttu-id="3d5f7-480">Date | [Time](/javascript/api/outlook_1_4/office.time)</span><span class="sxs-lookup"><span data-stu-id="3d5f7-480">Date | [Time](/javascript/api/outlook_1_4/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="3d5f7-481">要求</span><span class="sxs-lookup"><span data-stu-id="3d5f7-481">Requirements</span></span>

|<span data-ttu-id="3d5f7-482">要求</span><span class="sxs-lookup"><span data-stu-id="3d5f7-482">Requirement</span></span>| <span data-ttu-id="3d5f7-483">值</span><span class="sxs-lookup"><span data-stu-id="3d5f7-483">Value</span></span>|
|---|---|
|[<span data-ttu-id="3d5f7-484">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="3d5f7-484">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3d5f7-485">1.0</span><span class="sxs-lookup"><span data-stu-id="3d5f7-485">1.0</span></span>|
|[<span data-ttu-id="3d5f7-486">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="3d5f7-486">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3d5f7-487">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3d5f7-487">ReadItem</span></span>|
|[<span data-ttu-id="3d5f7-488">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="3d5f7-488">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="3d5f7-489">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="3d5f7-489">Compose or Read</span></span>|

#### <a name="subject-stringsubjectjavascriptapioutlook14officesubject"></a><span data-ttu-id="3d5f7-490">subject: String |[主题](/javascript/api/outlook_1_4/office.subject)</span><span class="sxs-lookup"><span data-stu-id="3d5f7-490">subject: String|[Subject](/javascript/api/outlook_1_4/office.subject)</span></span>

<span data-ttu-id="3d5f7-491">获取或设置显示在项目的主题字段中的说明。</span><span class="sxs-lookup"><span data-stu-id="3d5f7-491">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="3d5f7-492">`subject` 属性获取或设置由电子邮件服务器发送项目时的整个主题。</span><span class="sxs-lookup"><span data-stu-id="3d5f7-492">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="3d5f7-493">阅读模式</span><span class="sxs-lookup"><span data-stu-id="3d5f7-493">Read mode</span></span>

<span data-ttu-id="3d5f7-p129">`subject` 属性返回一个字符串。使用 [`normalizedSubject`](#normalizedsubject-string) 属性获取不带任何前导前缀（如 `RE:` 和 `FW:`）的主题。</span><span class="sxs-lookup"><span data-stu-id="3d5f7-p129">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

```javascript
var subject = Office.context.mailbox.item.subject;
console.log(subject);
```

##### <a name="compose-mode"></a><span data-ttu-id="3d5f7-496">撰写模式</span><span class="sxs-lookup"><span data-stu-id="3d5f7-496">Compose mode</span></span>

<span data-ttu-id="3d5f7-497">`subject` 属性返回一个 `Subject` 对象，该对象提供用于获取和设置主题的方法。</span><span class="sxs-lookup"><span data-stu-id="3d5f7-497">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```javascript
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
  console.log(subject);
}
```

##### <a name="type"></a><span data-ttu-id="3d5f7-498">类型</span><span class="sxs-lookup"><span data-stu-id="3d5f7-498">Type</span></span>

*   <span data-ttu-id="3d5f7-499">String | [Subject](/javascript/api/outlook_1_4/office.subject)</span><span class="sxs-lookup"><span data-stu-id="3d5f7-499">String | [Subject](/javascript/api/outlook_1_4/office.subject)</span></span>

##### <a name="requirements"></a><span data-ttu-id="3d5f7-500">要求</span><span class="sxs-lookup"><span data-stu-id="3d5f7-500">Requirements</span></span>

|<span data-ttu-id="3d5f7-501">要求</span><span class="sxs-lookup"><span data-stu-id="3d5f7-501">Requirement</span></span>| <span data-ttu-id="3d5f7-502">值</span><span class="sxs-lookup"><span data-stu-id="3d5f7-502">Value</span></span>|
|---|---|
|[<span data-ttu-id="3d5f7-503">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="3d5f7-503">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3d5f7-504">1.0</span><span class="sxs-lookup"><span data-stu-id="3d5f7-504">1.0</span></span>|
|[<span data-ttu-id="3d5f7-505">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="3d5f7-505">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3d5f7-506">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3d5f7-506">ReadItem</span></span>|
|[<span data-ttu-id="3d5f7-507">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="3d5f7-507">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="3d5f7-508">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="3d5f7-508">Compose or Read</span></span>|

#### <a name="to-arrayemailaddressdetailsjavascriptapioutlook14officeemailaddressdetailsrecipientsjavascriptapioutlook14officerecipients"></a><span data-ttu-id="3d5f7-509">to: <[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)>|[收件人](/javascript/api/outlook_1_4/office.recipients)的数组</span><span class="sxs-lookup"><span data-stu-id="3d5f7-509">to: Array.<[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_4/office.recipients)</span></span>

<span data-ttu-id="3d5f7-510">提供对邮件的“**收件人**”行上的收件人的访问权限。</span><span class="sxs-lookup"><span data-stu-id="3d5f7-510">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="3d5f7-511">对象的类型和访问级别取决于当前项目的模式。</span><span class="sxs-lookup"><span data-stu-id="3d5f7-511">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="3d5f7-512">阅读模式</span><span class="sxs-lookup"><span data-stu-id="3d5f7-512">Read mode</span></span>

<span data-ttu-id="3d5f7-p131">`to` 属性返回包含邮件的**收件人**行上所列的每个收件人的 `EmailAddressDetails` 对象的数组。集合上限为 100 个成员。</span><span class="sxs-lookup"><span data-stu-id="3d5f7-p131">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.</span></span>

```javascript
console.log(JSON.stringify(Office.context.mailbox.item.to));
```

##### <a name="compose-mode"></a><span data-ttu-id="3d5f7-515">撰写模式</span><span class="sxs-lookup"><span data-stu-id="3d5f7-515">Compose mode</span></span>

<span data-ttu-id="3d5f7-516">`to` 属性返回一个 `Recipients` 对象，该对象提供用于获取或更新邮件的“**收件人**”行上收件人的方法。</span><span class="sxs-lookup"><span data-stu-id="3d5f7-516">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span>

```javascript
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="3d5f7-517">类型</span><span class="sxs-lookup"><span data-stu-id="3d5f7-517">Type</span></span>

*   <span data-ttu-id="3d5f7-518">Array.<[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_4/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="3d5f7-518">Array.<[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_4/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="3d5f7-519">要求</span><span class="sxs-lookup"><span data-stu-id="3d5f7-519">Requirements</span></span>

|<span data-ttu-id="3d5f7-520">要求</span><span class="sxs-lookup"><span data-stu-id="3d5f7-520">Requirement</span></span>| <span data-ttu-id="3d5f7-521">值</span><span class="sxs-lookup"><span data-stu-id="3d5f7-521">Value</span></span>|
|---|---|
|[<span data-ttu-id="3d5f7-522">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="3d5f7-522">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3d5f7-523">1.0</span><span class="sxs-lookup"><span data-stu-id="3d5f7-523">1.0</span></span>|
|[<span data-ttu-id="3d5f7-524">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="3d5f7-524">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3d5f7-525">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3d5f7-525">ReadItem</span></span>|
|[<span data-ttu-id="3d5f7-526">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="3d5f7-526">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="3d5f7-527">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="3d5f7-527">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="3d5f7-528">方法</span><span class="sxs-lookup"><span data-stu-id="3d5f7-528">Methods</span></span>

#### <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="3d5f7-529">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="3d5f7-529">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="3d5f7-530">将文件作为附件添加到邮件或约会。</span><span class="sxs-lookup"><span data-stu-id="3d5f7-530">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="3d5f7-531">`addFileAttachmentAsync` 方法在指定的 URI 上载文件并将其附加到撰写窗体中的项目。</span><span class="sxs-lookup"><span data-stu-id="3d5f7-531">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="3d5f7-532">随后可以将该标识符与 [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) 方法一同使用，以删除同一个会话中的附件。</span><span class="sxs-lookup"><span data-stu-id="3d5f7-532">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="3d5f7-533">参数</span><span class="sxs-lookup"><span data-stu-id="3d5f7-533">Parameters</span></span>

|<span data-ttu-id="3d5f7-534">名称</span><span class="sxs-lookup"><span data-stu-id="3d5f7-534">Name</span></span>| <span data-ttu-id="3d5f7-535">类型</span><span class="sxs-lookup"><span data-stu-id="3d5f7-535">Type</span></span>| <span data-ttu-id="3d5f7-536">属性</span><span class="sxs-lookup"><span data-stu-id="3d5f7-536">Attributes</span></span>| <span data-ttu-id="3d5f7-537">说明</span><span class="sxs-lookup"><span data-stu-id="3d5f7-537">Description</span></span>|
|---|---|---|---|
|`uri`| <span data-ttu-id="3d5f7-538">String</span><span class="sxs-lookup"><span data-stu-id="3d5f7-538">String</span></span>||<span data-ttu-id="3d5f7-p132">提供附加到邮件或约会的文件的位置的 URI。最大长度为 2048 个字符。</span><span class="sxs-lookup"><span data-stu-id="3d5f7-p132">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="3d5f7-541">字符串</span><span class="sxs-lookup"><span data-stu-id="3d5f7-541">String</span></span>||<span data-ttu-id="3d5f7-p133">在附件上载过程中显示的附件名称。最大长度为 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="3d5f7-p133">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="3d5f7-544">Object</span><span class="sxs-lookup"><span data-stu-id="3d5f7-544">Object</span></span>| <span data-ttu-id="3d5f7-545">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="3d5f7-545">&lt;optional&gt;</span></span>|<span data-ttu-id="3d5f7-546">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="3d5f7-546">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="3d5f7-547">对象</span><span class="sxs-lookup"><span data-stu-id="3d5f7-547">Object</span></span>| <span data-ttu-id="3d5f7-548">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="3d5f7-548">&lt;optional&gt;</span></span>|<span data-ttu-id="3d5f7-549">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="3d5f7-549">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="3d5f7-550">函数</span><span class="sxs-lookup"><span data-stu-id="3d5f7-550">function</span></span>| <span data-ttu-id="3d5f7-551">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="3d5f7-551">&lt;optional&gt;</span></span>|<span data-ttu-id="3d5f7-552">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="3d5f7-552">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="3d5f7-553">如果成功，附件标识符将在 `asyncResult.value` 属性中提供。</span><span class="sxs-lookup"><span data-stu-id="3d5f7-553">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="3d5f7-554">如果上传附件失败，`asyncResult` 对象将包含一个提供错误说明的 `Error` 对象。</span><span class="sxs-lookup"><span data-stu-id="3d5f7-554">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="3d5f7-555">错误</span><span class="sxs-lookup"><span data-stu-id="3d5f7-555">Errors</span></span>

| <span data-ttu-id="3d5f7-556">错误代码</span><span class="sxs-lookup"><span data-stu-id="3d5f7-556">Error code</span></span> | <span data-ttu-id="3d5f7-557">说明</span><span class="sxs-lookup"><span data-stu-id="3d5f7-557">Description</span></span> |
|------------|-------------|
| `AttachmentSizeExceeded` | <span data-ttu-id="3d5f7-558">附件大小超过了允许的大小。</span><span class="sxs-lookup"><span data-stu-id="3d5f7-558">The attachment is larger than allowed.</span></span> |
| `FileTypeNotSupported` | <span data-ttu-id="3d5f7-559">该附件的扩展名不是允许的扩展名。</span><span class="sxs-lookup"><span data-stu-id="3d5f7-559">The attachment has an extension that is not allowed.</span></span> |
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="3d5f7-560">邮件或约会具有的附件过多。</span><span class="sxs-lookup"><span data-stu-id="3d5f7-560">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="3d5f7-561">要求</span><span class="sxs-lookup"><span data-stu-id="3d5f7-561">Requirements</span></span>

|<span data-ttu-id="3d5f7-562">要求</span><span class="sxs-lookup"><span data-stu-id="3d5f7-562">Requirement</span></span>| <span data-ttu-id="3d5f7-563">值</span><span class="sxs-lookup"><span data-stu-id="3d5f7-563">Value</span></span>|
|---|---|
|[<span data-ttu-id="3d5f7-564">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="3d5f7-564">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3d5f7-565">1.1</span><span class="sxs-lookup"><span data-stu-id="3d5f7-565">1.1</span></span>|
|[<span data-ttu-id="3d5f7-566">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="3d5f7-566">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3d5f7-567">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="3d5f7-567">ReadWriteItem</span></span>|
|[<span data-ttu-id="3d5f7-568">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="3d5f7-568">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="3d5f7-569">撰写</span><span class="sxs-lookup"><span data-stu-id="3d5f7-569">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="3d5f7-570">示例</span><span class="sxs-lookup"><span data-stu-id="3d5f7-570">Example</span></span>

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

#### <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="3d5f7-571">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="3d5f7-571">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="3d5f7-572">将 Exchange 项目（如邮件）作为附件添加到邮件或约会。</span><span class="sxs-lookup"><span data-stu-id="3d5f7-572">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="3d5f7-p134">`addItemAttachmentAsync` 方法将包含指定 Exchange 标识符的项目附加到撰写窗体中的项目。如果指定一个回调方法，此方法使用 `asyncResult` 参数调用，该参数包含一个附件标识符或代码，指示附加项目过程中出现的任何错误。可以使用 `options` 参数将状态信息传递给回调方法（如果需要）。</span><span class="sxs-lookup"><span data-stu-id="3d5f7-p134">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="3d5f7-576">随后可以将该标识符与 [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) 方法一同使用，以删除同一个会话中的附件。</span><span class="sxs-lookup"><span data-stu-id="3d5f7-576">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="3d5f7-577">如果 Office 加载项在 Outlook Web App 中运行，则 `addItemAttachmentAsync` 方法可以将项目附加到项目（正在编辑的项目除外）中；然而，不支持也不建议这样做。</span><span class="sxs-lookup"><span data-stu-id="3d5f7-577">If your Office Add-in is running in Outlook Web App, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="3d5f7-578">参数</span><span class="sxs-lookup"><span data-stu-id="3d5f7-578">Parameters</span></span>

|<span data-ttu-id="3d5f7-579">名称</span><span class="sxs-lookup"><span data-stu-id="3d5f7-579">Name</span></span>| <span data-ttu-id="3d5f7-580">类型</span><span class="sxs-lookup"><span data-stu-id="3d5f7-580">Type</span></span>| <span data-ttu-id="3d5f7-581">属性</span><span class="sxs-lookup"><span data-stu-id="3d5f7-581">Attributes</span></span>| <span data-ttu-id="3d5f7-582">说明</span><span class="sxs-lookup"><span data-stu-id="3d5f7-582">Description</span></span>|
|---|---|---|---|
|`itemId`| <span data-ttu-id="3d5f7-583">String</span><span class="sxs-lookup"><span data-stu-id="3d5f7-583">String</span></span>||<span data-ttu-id="3d5f7-p135">要附加的项目的 Exchange 标识符。最大长度为 100 个字符。</span><span class="sxs-lookup"><span data-stu-id="3d5f7-p135">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="3d5f7-586">String</span><span class="sxs-lookup"><span data-stu-id="3d5f7-586">String</span></span>||<span data-ttu-id="3d5f7-587">要附加的项目的主题。</span><span class="sxs-lookup"><span data-stu-id="3d5f7-587">The subject of the item to be attached.</span></span> <span data-ttu-id="3d5f7-588">最大长度为 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="3d5f7-588">The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="3d5f7-589">对象</span><span class="sxs-lookup"><span data-stu-id="3d5f7-589">Object</span></span>| <span data-ttu-id="3d5f7-590">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="3d5f7-590">&lt;optional&gt;</span></span>|<span data-ttu-id="3d5f7-591">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="3d5f7-591">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="3d5f7-592">对象</span><span class="sxs-lookup"><span data-stu-id="3d5f7-592">Object</span></span>| <span data-ttu-id="3d5f7-593">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="3d5f7-593">&lt;optional&gt;</span></span>|<span data-ttu-id="3d5f7-594">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="3d5f7-594">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="3d5f7-595">函数</span><span class="sxs-lookup"><span data-stu-id="3d5f7-595">function</span></span>| <span data-ttu-id="3d5f7-596">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="3d5f7-596">&lt;optional&gt;</span></span>|<span data-ttu-id="3d5f7-597">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="3d5f7-597">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="3d5f7-598">如果成功，附件标识符将在 `asyncResult.value` 属性中提供。</span><span class="sxs-lookup"><span data-stu-id="3d5f7-598">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="3d5f7-599">如果添加附件失败，`asyncResult` 对象将包含一个提供错误说明的 `Error` 对象。</span><span class="sxs-lookup"><span data-stu-id="3d5f7-599">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="3d5f7-600">错误</span><span class="sxs-lookup"><span data-stu-id="3d5f7-600">Errors</span></span>

| <span data-ttu-id="3d5f7-601">错误代码</span><span class="sxs-lookup"><span data-stu-id="3d5f7-601">Error code</span></span> | <span data-ttu-id="3d5f7-602">说明</span><span class="sxs-lookup"><span data-stu-id="3d5f7-602">Description</span></span> |
|------------|-------------|
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="3d5f7-603">邮件或约会具有的附件过多。</span><span class="sxs-lookup"><span data-stu-id="3d5f7-603">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="3d5f7-604">要求</span><span class="sxs-lookup"><span data-stu-id="3d5f7-604">Requirements</span></span>

|<span data-ttu-id="3d5f7-605">要求</span><span class="sxs-lookup"><span data-stu-id="3d5f7-605">Requirement</span></span>| <span data-ttu-id="3d5f7-606">值</span><span class="sxs-lookup"><span data-stu-id="3d5f7-606">Value</span></span>|
|---|---|
|[<span data-ttu-id="3d5f7-607">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="3d5f7-607">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3d5f7-608">1.1</span><span class="sxs-lookup"><span data-stu-id="3d5f7-608">1.1</span></span>|
|[<span data-ttu-id="3d5f7-609">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="3d5f7-609">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3d5f7-610">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="3d5f7-610">ReadWriteItem</span></span>|
|[<span data-ttu-id="3d5f7-611">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="3d5f7-611">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="3d5f7-612">撰写</span><span class="sxs-lookup"><span data-stu-id="3d5f7-612">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="3d5f7-613">示例</span><span class="sxs-lookup"><span data-stu-id="3d5f7-613">Example</span></span>

<span data-ttu-id="3d5f7-614">以下示例将现有的 Outlook 项目添加为名为 `My Attachment` 的附件。</span><span class="sxs-lookup"><span data-stu-id="3d5f7-614">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

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

#### <a name="close"></a><span data-ttu-id="3d5f7-615">close()</span><span class="sxs-lookup"><span data-stu-id="3d5f7-615">close()</span></span>

<span data-ttu-id="3d5f7-616">关闭当前正在撰写的项目。</span><span class="sxs-lookup"><span data-stu-id="3d5f7-616">Closes the current item that is being composed.</span></span>

<span data-ttu-id="3d5f7-p137">
            \`close\` 方法的行为取决于要撰写的项目的当前状态。如果项目具有未保存的更改，客户端将提示用户保存、放弃或取消关闭操作。</span><span class="sxs-lookup"><span data-stu-id="3d5f7-p137">The behavior of the `close` method depends on the current state of the item being composed. If the item has unsaved changes, the client prompts the user to save, discard, or cancel the close action.</span></span>

> [!NOTE]
> <span data-ttu-id="3d5f7-619">在 Outlook 网页版中，如果该项目是约会并且之前已使用 `saveAsync` 保存，则即使自上次保存项目后未发生任何更改，也会提示用户保存、放弃或取消。</span><span class="sxs-lookup"><span data-stu-id="3d5f7-619">In Outlook on the web, if the item is an appointment and it has previously been saved using `saveAsync`, the user is prompted to save, discard, or cancel even if no changes have occurred since the item was last saved.</span></span>

<span data-ttu-id="3d5f7-620">在 Outlook 桌面客户端中，如果邮件是内联答复，`close` 方法不起作用。</span><span class="sxs-lookup"><span data-stu-id="3d5f7-620">In the Outlook desktop client, if the message is an inline reply, the `close` method has no effect.</span></span>

##### <a name="requirements"></a><span data-ttu-id="3d5f7-621">要求</span><span class="sxs-lookup"><span data-stu-id="3d5f7-621">Requirements</span></span>

|<span data-ttu-id="3d5f7-622">要求</span><span class="sxs-lookup"><span data-stu-id="3d5f7-622">Requirement</span></span>| <span data-ttu-id="3d5f7-623">值</span><span class="sxs-lookup"><span data-stu-id="3d5f7-623">Value</span></span>|
|---|---|
|[<span data-ttu-id="3d5f7-624">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="3d5f7-624">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3d5f7-625">1.3</span><span class="sxs-lookup"><span data-stu-id="3d5f7-625">1.3</span></span>|
|[<span data-ttu-id="3d5f7-626">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="3d5f7-626">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3d5f7-627">受限</span><span class="sxs-lookup"><span data-stu-id="3d5f7-627">Restricted</span></span>|
|[<span data-ttu-id="3d5f7-628">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="3d5f7-628">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="3d5f7-629">撰写</span><span class="sxs-lookup"><span data-stu-id="3d5f7-629">Compose</span></span>|

#### <a name="displayreplyallformformdata-callback"></a><span data-ttu-id="3d5f7-630">displayReplyAllForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="3d5f7-630">displayReplyAllForm(formData, [callback])</span></span>

<span data-ttu-id="3d5f7-631">显示答复窗体，其中包括所选邮件的发件人和所有收件人或所选约会的组织者和所有与会者。</span><span class="sxs-lookup"><span data-stu-id="3d5f7-631">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="3d5f7-632">在 Outlook for iOS 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="3d5f7-632">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="3d5f7-633">在 Outlook Web App 中，答复窗体显示为包含 3 列视图的弹出式窗体以及包含 2 列或 1 列视图的弹出式窗体。</span><span class="sxs-lookup"><span data-stu-id="3d5f7-633">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="3d5f7-634">如果任意字符串参数超出其限制，`displayReplyAllForm` 将引发异常。</span><span class="sxs-lookup"><span data-stu-id="3d5f7-634">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="3d5f7-p138">当在 `formData.attachments` 参数中指定附件时，Outlook 和 Outlook Web App 尝试下载所有附件并将其附加到答复窗体。如果无法添加任何附件，则在窗体 UI 中显示错误。如果这不可能，则不引发错误消息。</span><span class="sxs-lookup"><span data-stu-id="3d5f7-p138">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="3d5f7-638">参数</span><span class="sxs-lookup"><span data-stu-id="3d5f7-638">Parameters</span></span>

|<span data-ttu-id="3d5f7-639">名称</span><span class="sxs-lookup"><span data-stu-id="3d5f7-639">Name</span></span>| <span data-ttu-id="3d5f7-640">类型</span><span class="sxs-lookup"><span data-stu-id="3d5f7-640">Type</span></span>| <span data-ttu-id="3d5f7-641">说明</span><span class="sxs-lookup"><span data-stu-id="3d5f7-641">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="3d5f7-642">字符串 &#124; 对象</span><span class="sxs-lookup"><span data-stu-id="3d5f7-642">String &#124; Object</span></span>| |<span data-ttu-id="3d5f7-p139">一个包含文本和 HTML 且表示答复窗体的正文的字符串。字符串限制为 32 KB。</span><span class="sxs-lookup"><span data-stu-id="3d5f7-p139">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="3d5f7-645">**或**</span><span class="sxs-lookup"><span data-stu-id="3d5f7-645">**OR**</span></span><br/><span data-ttu-id="3d5f7-p140">包含正文或附件数据和回调函数的对象。对象定义如下。</span><span class="sxs-lookup"><span data-stu-id="3d5f7-p140">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="3d5f7-648">String</span><span class="sxs-lookup"><span data-stu-id="3d5f7-648">String</span></span> | <span data-ttu-id="3d5f7-649">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="3d5f7-649">&lt;optional&gt;</span></span> | <span data-ttu-id="3d5f7-p141">一个包含文本和 HTML 且表示答复窗体的正文的字符串。字符串限制为 32 KB。</span><span class="sxs-lookup"><span data-stu-id="3d5f7-p141">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="3d5f7-652">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="3d5f7-652">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="3d5f7-653">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="3d5f7-653">&lt;optional&gt;</span></span> | <span data-ttu-id="3d5f7-654">JSON 对象（是文件或项目附件）的数组。</span><span class="sxs-lookup"><span data-stu-id="3d5f7-654">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="3d5f7-655">String</span><span class="sxs-lookup"><span data-stu-id="3d5f7-655">String</span></span> | | <span data-ttu-id="3d5f7-p142">指示附件的类型。必须是文件附件的 `file` 或项目附件的 `item`。</span><span class="sxs-lookup"><span data-stu-id="3d5f7-p142">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="3d5f7-658">字符串</span><span class="sxs-lookup"><span data-stu-id="3d5f7-658">String</span></span> | | <span data-ttu-id="3d5f7-659">一个包含附件的名称的字符串，最多包含 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="3d5f7-659">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="3d5f7-660">String</span><span class="sxs-lookup"><span data-stu-id="3d5f7-660">String</span></span> | | <span data-ttu-id="3d5f7-p143">仅在将 `type` 设置为 `file` 时使用。文件的位置的 URI。</span><span class="sxs-lookup"><span data-stu-id="3d5f7-p143">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="3d5f7-663">String</span><span class="sxs-lookup"><span data-stu-id="3d5f7-663">String</span></span> | | <span data-ttu-id="3d5f7-p144">仅在将 `type` 设置为 `item` 时使用。附件的 EWS 项目 ID。字符串最长为 100 个字符。</span><span class="sxs-lookup"><span data-stu-id="3d5f7-p144">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="3d5f7-667">函数</span><span class="sxs-lookup"><span data-stu-id="3d5f7-667">function</span></span> | <span data-ttu-id="3d5f7-668">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="3d5f7-668">&lt;optional&gt;</span></span> | <span data-ttu-id="3d5f7-669">方法完成后，使用单个参数 `asyncResult`（一个 [AsyncResult](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="3d5f7-669">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="3d5f7-670">要求</span><span class="sxs-lookup"><span data-stu-id="3d5f7-670">Requirements</span></span>

|<span data-ttu-id="3d5f7-671">要求</span><span class="sxs-lookup"><span data-stu-id="3d5f7-671">Requirement</span></span>| <span data-ttu-id="3d5f7-672">值</span><span class="sxs-lookup"><span data-stu-id="3d5f7-672">Value</span></span>|
|---|---|
|[<span data-ttu-id="3d5f7-673">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="3d5f7-673">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3d5f7-674">1.0</span><span class="sxs-lookup"><span data-stu-id="3d5f7-674">1.0</span></span>|
|[<span data-ttu-id="3d5f7-675">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="3d5f7-675">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3d5f7-676">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3d5f7-676">ReadItem</span></span>|
|[<span data-ttu-id="3d5f7-677">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="3d5f7-677">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="3d5f7-678">阅读</span><span class="sxs-lookup"><span data-stu-id="3d5f7-678">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="3d5f7-679">示例</span><span class="sxs-lookup"><span data-stu-id="3d5f7-679">Examples</span></span>

<span data-ttu-id="3d5f7-680">以下代码将一个字符串传递到 `displayReplyAllForm` 函数。</span><span class="sxs-lookup"><span data-stu-id="3d5f7-680">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="3d5f7-681">使用空白正文答复。</span><span class="sxs-lookup"><span data-stu-id="3d5f7-681">Reply with an empty body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="3d5f7-682">仅使用正文答复。</span><span class="sxs-lookup"><span data-stu-id="3d5f7-682">Reply with just a body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="3d5f7-683">使用正文和文件附件答复。</span><span class="sxs-lookup"><span data-stu-id="3d5f7-683">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="3d5f7-684">使用正文和项目附件答复。</span><span class="sxs-lookup"><span data-stu-id="3d5f7-684">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="3d5f7-685">使用正文、文件附件、项目附件和回调答复。</span><span class="sxs-lookup"><span data-stu-id="3d5f7-685">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="displayreplyformformdata-callback"></a><span data-ttu-id="3d5f7-686">displayReplyForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="3d5f7-686">displayReplyForm(formData, [callback])</span></span>

<span data-ttu-id="3d5f7-687">显示答复窗体，其中仅包括所选邮件的发件人或所选约会的组织者。</span><span class="sxs-lookup"><span data-stu-id="3d5f7-687">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="3d5f7-688">在 Outlook for iOS 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="3d5f7-688">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="3d5f7-689">在 Outlook Web App 中，答复窗体显示为包含 3 列视图的弹出式窗体以及包含 2 列或 1 列视图的弹出式窗体。</span><span class="sxs-lookup"><span data-stu-id="3d5f7-689">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="3d5f7-690">如果任意字符串参数超出其限制，`displayReplyForm` 将引发异常。</span><span class="sxs-lookup"><span data-stu-id="3d5f7-690">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="3d5f7-p145">当在 `formData.attachments` 参数中指定附件时，Outlook 和 Outlook Web App 尝试下载所有附件并将其附加到答复窗体。如果无法添加任何附件，则在窗体 UI 中显示错误。如果这不可能，则不引发错误消息。</span><span class="sxs-lookup"><span data-stu-id="3d5f7-p145">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="3d5f7-694">参数</span><span class="sxs-lookup"><span data-stu-id="3d5f7-694">Parameters</span></span>

|<span data-ttu-id="3d5f7-695">名称</span><span class="sxs-lookup"><span data-stu-id="3d5f7-695">Name</span></span>| <span data-ttu-id="3d5f7-696">类型</span><span class="sxs-lookup"><span data-stu-id="3d5f7-696">Type</span></span>| <span data-ttu-id="3d5f7-697">说明</span><span class="sxs-lookup"><span data-stu-id="3d5f7-697">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="3d5f7-698">字符串 &#124; 对象</span><span class="sxs-lookup"><span data-stu-id="3d5f7-698">String &#124; Object</span></span>| | <span data-ttu-id="3d5f7-p146">一个包含文本和 HTML 且表示答复窗体的正文的字符串。字符串限制为 32 KB。</span><span class="sxs-lookup"><span data-stu-id="3d5f7-p146">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="3d5f7-701">**或**</span><span class="sxs-lookup"><span data-stu-id="3d5f7-701">**OR**</span></span><br/><span data-ttu-id="3d5f7-p147">包含正文或附件数据和回调函数的对象。对象定义如下。</span><span class="sxs-lookup"><span data-stu-id="3d5f7-p147">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="3d5f7-704">String</span><span class="sxs-lookup"><span data-stu-id="3d5f7-704">String</span></span> | <span data-ttu-id="3d5f7-705">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="3d5f7-705">&lt;optional&gt;</span></span> | <span data-ttu-id="3d5f7-p148">一个包含文本和 HTML 且表示答复窗体的正文的字符串。字符串限制为 32 KB。</span><span class="sxs-lookup"><span data-stu-id="3d5f7-p148">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="3d5f7-708">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="3d5f7-708">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="3d5f7-709">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="3d5f7-709">&lt;optional&gt;</span></span> | <span data-ttu-id="3d5f7-710">JSON 对象（是文件或项目附件）的数组。</span><span class="sxs-lookup"><span data-stu-id="3d5f7-710">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="3d5f7-711">String</span><span class="sxs-lookup"><span data-stu-id="3d5f7-711">String</span></span> | | <span data-ttu-id="3d5f7-p149">指示附件的类型。必须是文件附件的 `file` 或项目附件的 `item`。</span><span class="sxs-lookup"><span data-stu-id="3d5f7-p149">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="3d5f7-714">字符串</span><span class="sxs-lookup"><span data-stu-id="3d5f7-714">String</span></span> | | <span data-ttu-id="3d5f7-715">一个包含附件的名称的字符串，最多包含 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="3d5f7-715">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="3d5f7-716">String</span><span class="sxs-lookup"><span data-stu-id="3d5f7-716">String</span></span> | | <span data-ttu-id="3d5f7-p150">仅在将 `type` 设置为 `file` 时使用。文件的位置的 URI。</span><span class="sxs-lookup"><span data-stu-id="3d5f7-p150">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="3d5f7-719">字符串</span><span class="sxs-lookup"><span data-stu-id="3d5f7-719">String</span></span> | | <span data-ttu-id="3d5f7-p151">仅在将 `type` 设置为 `item` 时使用。附件的 EWS 项目 ID。字符串最长为 100 个字符。</span><span class="sxs-lookup"><span data-stu-id="3d5f7-p151">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="3d5f7-723">函数</span><span class="sxs-lookup"><span data-stu-id="3d5f7-723">function</span></span> | <span data-ttu-id="3d5f7-724">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="3d5f7-724">&lt;optional&gt;</span></span> | <span data-ttu-id="3d5f7-725">方法完成后，使用单个参数 `asyncResult`（一个 [AsyncResult](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="3d5f7-725">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="3d5f7-726">要求</span><span class="sxs-lookup"><span data-stu-id="3d5f7-726">Requirements</span></span>

|<span data-ttu-id="3d5f7-727">要求</span><span class="sxs-lookup"><span data-stu-id="3d5f7-727">Requirement</span></span>| <span data-ttu-id="3d5f7-728">值</span><span class="sxs-lookup"><span data-stu-id="3d5f7-728">Value</span></span>|
|---|---|
|[<span data-ttu-id="3d5f7-729">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="3d5f7-729">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3d5f7-730">1.0</span><span class="sxs-lookup"><span data-stu-id="3d5f7-730">1.0</span></span>|
|[<span data-ttu-id="3d5f7-731">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="3d5f7-731">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3d5f7-732">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3d5f7-732">ReadItem</span></span>|
|[<span data-ttu-id="3d5f7-733">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="3d5f7-733">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="3d5f7-734">阅读</span><span class="sxs-lookup"><span data-stu-id="3d5f7-734">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="3d5f7-735">示例</span><span class="sxs-lookup"><span data-stu-id="3d5f7-735">Examples</span></span>

<span data-ttu-id="3d5f7-736">以下代码将一个字符串传递到 `displayReplyForm` 函数。</span><span class="sxs-lookup"><span data-stu-id="3d5f7-736">The following code passes a string to the `displayReplyForm` function.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="3d5f7-737">使用空白正文答复。</span><span class="sxs-lookup"><span data-stu-id="3d5f7-737">Reply with an empty body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="3d5f7-738">仅使用正文答复。</span><span class="sxs-lookup"><span data-stu-id="3d5f7-738">Reply with just a body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="3d5f7-739">使用正文和文件附件答复。</span><span class="sxs-lookup"><span data-stu-id="3d5f7-739">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="3d5f7-740">使用正文和项目附件答复。</span><span class="sxs-lookup"><span data-stu-id="3d5f7-740">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="3d5f7-741">使用正文、文件附件、项目附件和回调答复。</span><span class="sxs-lookup"><span data-stu-id="3d5f7-741">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="getentities--entitiesjavascriptapioutlook14officeentities"></a><span data-ttu-id="3d5f7-742">getEntities() → {[Entities](/javascript/api/outlook_1_4/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="3d5f7-742">getEntities() → {[Entities](/javascript/api/outlook_1_4/office.entities)}</span></span>

<span data-ttu-id="3d5f7-743">获取在所选项目的正文中找到的实体。</span><span class="sxs-lookup"><span data-stu-id="3d5f7-743">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="3d5f7-744">在 Outlook for iOS 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="3d5f7-744">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="3d5f7-745">要求</span><span class="sxs-lookup"><span data-stu-id="3d5f7-745">Requirements</span></span>

|<span data-ttu-id="3d5f7-746">要求</span><span class="sxs-lookup"><span data-stu-id="3d5f7-746">Requirement</span></span>| <span data-ttu-id="3d5f7-747">值</span><span class="sxs-lookup"><span data-stu-id="3d5f7-747">Value</span></span>|
|---|---|
|[<span data-ttu-id="3d5f7-748">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="3d5f7-748">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3d5f7-749">1.0</span><span class="sxs-lookup"><span data-stu-id="3d5f7-749">1.0</span></span>|
|[<span data-ttu-id="3d5f7-750">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="3d5f7-750">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3d5f7-751">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3d5f7-751">ReadItem</span></span>|
|[<span data-ttu-id="3d5f7-752">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="3d5f7-752">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="3d5f7-753">阅读</span><span class="sxs-lookup"><span data-stu-id="3d5f7-753">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="3d5f7-754">返回：</span><span class="sxs-lookup"><span data-stu-id="3d5f7-754">Returns:</span></span>

<span data-ttu-id="3d5f7-755">类型：[Entities](/javascript/api/outlook_1_4/office.entities)</span><span class="sxs-lookup"><span data-stu-id="3d5f7-755">Type: [Entities](/javascript/api/outlook_1_4/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="3d5f7-756">示例</span><span class="sxs-lookup"><span data-stu-id="3d5f7-756">Example</span></span>

<span data-ttu-id="3d5f7-757">以下示例访问当前项目的正文中的联系人实体。</span><span class="sxs-lookup"><span data-stu-id="3d5f7-757">The following example accesses the contacts entities in the current item's body.</span></span>

```javascript
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlook14officecontactmeetingsuggestionjavascriptapioutlook14officemeetingsuggestionphonenumberjavascriptapioutlook14officephonenumbertasksuggestionjavascriptapioutlook14officetasksuggestion"></a><span data-ttu-id="3d5f7-758">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_4/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_4/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_4/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_4/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="3d5f7-758">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_4/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_4/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_4/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_4/office.tasksuggestion))>}</span></span>

<span data-ttu-id="3d5f7-759">获取所选项目的正文中找到的指定实体类型的所有实体的数组。</span><span class="sxs-lookup"><span data-stu-id="3d5f7-759">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="3d5f7-760">在 Outlook for iOS 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="3d5f7-760">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="3d5f7-761">参数</span><span class="sxs-lookup"><span data-stu-id="3d5f7-761">Parameters</span></span>

|<span data-ttu-id="3d5f7-762">名称</span><span class="sxs-lookup"><span data-stu-id="3d5f7-762">Name</span></span>| <span data-ttu-id="3d5f7-763">类型</span><span class="sxs-lookup"><span data-stu-id="3d5f7-763">Type</span></span>| <span data-ttu-id="3d5f7-764">说明</span><span class="sxs-lookup"><span data-stu-id="3d5f7-764">Description</span></span>|
|---|---|---|
|`entityType`| [<span data-ttu-id="3d5f7-765">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="3d5f7-765">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook_1_4/office.mailboxenums.entitytype)|<span data-ttu-id="3d5f7-766">EntityType 枚举值之一。</span><span class="sxs-lookup"><span data-stu-id="3d5f7-766">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="3d5f7-767">要求</span><span class="sxs-lookup"><span data-stu-id="3d5f7-767">Requirements</span></span>

|<span data-ttu-id="3d5f7-768">要求</span><span class="sxs-lookup"><span data-stu-id="3d5f7-768">Requirement</span></span>| <span data-ttu-id="3d5f7-769">值</span><span class="sxs-lookup"><span data-stu-id="3d5f7-769">Value</span></span>|
|---|---|
|[<span data-ttu-id="3d5f7-770">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="3d5f7-770">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3d5f7-771">1.0</span><span class="sxs-lookup"><span data-stu-id="3d5f7-771">1.0</span></span>|
|[<span data-ttu-id="3d5f7-772">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="3d5f7-772">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3d5f7-773">受限</span><span class="sxs-lookup"><span data-stu-id="3d5f7-773">Restricted</span></span>|
|[<span data-ttu-id="3d5f7-774">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="3d5f7-774">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="3d5f7-775">阅读</span><span class="sxs-lookup"><span data-stu-id="3d5f7-775">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="3d5f7-776">返回：</span><span class="sxs-lookup"><span data-stu-id="3d5f7-776">Returns:</span></span>

<span data-ttu-id="3d5f7-777">如果在 `entityType` 中传递的值不是 `EntityType` 枚举的有效成员，该方法返回 null。</span><span class="sxs-lookup"><span data-stu-id="3d5f7-777">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="3d5f7-778">如果指定类型的任何实体都不存在于该项目的正文中，该方法将返回空数组。</span><span class="sxs-lookup"><span data-stu-id="3d5f7-778">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="3d5f7-779">否则，返回的数组中对象的类型取决于 `entityType` 参数中请求实体的类型。</span><span class="sxs-lookup"><span data-stu-id="3d5f7-779">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="3d5f7-780">当使用此方法的最低权限级别**受限**时，某些实体类型需要 **ReadItem** 才能进行访问，如下表中所指定。</span><span class="sxs-lookup"><span data-stu-id="3d5f7-780">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

| <span data-ttu-id="3d5f7-781">`entityType` 的值</span><span class="sxs-lookup"><span data-stu-id="3d5f7-781">Value of `entityType`</span></span> | <span data-ttu-id="3d5f7-782">返回的数组中对象的类型</span><span class="sxs-lookup"><span data-stu-id="3d5f7-782">Type of objects in returned array</span></span> | <span data-ttu-id="3d5f7-783">所需权限级别</span><span class="sxs-lookup"><span data-stu-id="3d5f7-783">Required Permission Level</span></span> |
| --- | --- | --- |
| `Address` | <span data-ttu-id="3d5f7-784">字符串</span><span class="sxs-lookup"><span data-stu-id="3d5f7-784">String</span></span> | <span data-ttu-id="3d5f7-785">**受限**</span><span class="sxs-lookup"><span data-stu-id="3d5f7-785">**Restricted**</span></span> |
| `Contact` | <span data-ttu-id="3d5f7-786">Contact</span><span class="sxs-lookup"><span data-stu-id="3d5f7-786">Contact</span></span> | <span data-ttu-id="3d5f7-787">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="3d5f7-787">**ReadItem**</span></span> |
| `EmailAddress` | <span data-ttu-id="3d5f7-788">String</span><span class="sxs-lookup"><span data-stu-id="3d5f7-788">String</span></span> | <span data-ttu-id="3d5f7-789">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="3d5f7-789">**ReadItem**</span></span> |
| `MeetingSuggestion` | <span data-ttu-id="3d5f7-790">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="3d5f7-790">MeetingSuggestion</span></span> | <span data-ttu-id="3d5f7-791">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="3d5f7-791">**ReadItem**</span></span> |
| `PhoneNumber` | <span data-ttu-id="3d5f7-792">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="3d5f7-792">PhoneNumber</span></span> | <span data-ttu-id="3d5f7-793">**受限**</span><span class="sxs-lookup"><span data-stu-id="3d5f7-793">**Restricted**</span></span> |
| `TaskSuggestion` | <span data-ttu-id="3d5f7-794">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="3d5f7-794">TaskSuggestion</span></span> | <span data-ttu-id="3d5f7-795">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="3d5f7-795">**ReadItem**</span></span> |
| `URL` | <span data-ttu-id="3d5f7-796">String</span><span class="sxs-lookup"><span data-stu-id="3d5f7-796">String</span></span> | <span data-ttu-id="3d5f7-797">**受限**</span><span class="sxs-lookup"><span data-stu-id="3d5f7-797">**Restricted**</span></span> |

<span data-ttu-id="3d5f7-798">类型：Array.<(String|[Contact](/javascript/api/outlook_1_4/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_4/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_4/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_4/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="3d5f7-798">Type: Array.<(String|[Contact](/javascript/api/outlook_1_4/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_4/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_4/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_4/office.tasksuggestion))></span></span>

##### <a name="example"></a><span data-ttu-id="3d5f7-799">示例</span><span class="sxs-lookup"><span data-stu-id="3d5f7-799">Example</span></span>

<span data-ttu-id="3d5f7-800">以下示例显示了如何访问表示当前项目的正文中的邮政地址的字符串数组。</span><span class="sxs-lookup"><span data-stu-id="3d5f7-800">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlook14officecontactmeetingsuggestionjavascriptapioutlook14officemeetingsuggestionphonenumberjavascriptapioutlook14officephonenumbertasksuggestionjavascriptapioutlook14officetasksuggestion"></a><span data-ttu-id="3d5f7-801">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_4/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_4/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_4/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_4/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="3d5f7-801">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_4/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_4/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_4/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_4/office.tasksuggestion))>}</span></span>

<span data-ttu-id="3d5f7-802">返回传递清单 XML 文件中定义的命名筛选器的所选项目中的已知实体。</span><span class="sxs-lookup"><span data-stu-id="3d5f7-802">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="3d5f7-803">在 Outlook for iOS 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="3d5f7-803">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="3d5f7-804">`getFilteredEntitiesByName` 方法返回匹配在具有指定 `FilterName` 元素值的清单 XML 文件中的 [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) 规则元素中定义的正则表达式的实体。</span><span class="sxs-lookup"><span data-stu-id="3d5f7-804">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="3d5f7-805">参数</span><span class="sxs-lookup"><span data-stu-id="3d5f7-805">Parameters</span></span>

|<span data-ttu-id="3d5f7-806">名称</span><span class="sxs-lookup"><span data-stu-id="3d5f7-806">Name</span></span>| <span data-ttu-id="3d5f7-807">类型</span><span class="sxs-lookup"><span data-stu-id="3d5f7-807">Type</span></span>| <span data-ttu-id="3d5f7-808">说明</span><span class="sxs-lookup"><span data-stu-id="3d5f7-808">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="3d5f7-809">String</span><span class="sxs-lookup"><span data-stu-id="3d5f7-809">String</span></span>|<span data-ttu-id="3d5f7-810">定义筛选器匹配的 `ItemHasKnownEntity` 规则元素的名称。</span><span class="sxs-lookup"><span data-stu-id="3d5f7-810">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="3d5f7-811">要求</span><span class="sxs-lookup"><span data-stu-id="3d5f7-811">Requirements</span></span>

|<span data-ttu-id="3d5f7-812">要求</span><span class="sxs-lookup"><span data-stu-id="3d5f7-812">Requirement</span></span>| <span data-ttu-id="3d5f7-813">值</span><span class="sxs-lookup"><span data-stu-id="3d5f7-813">Value</span></span>|
|---|---|
|[<span data-ttu-id="3d5f7-814">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="3d5f7-814">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3d5f7-815">1.0</span><span class="sxs-lookup"><span data-stu-id="3d5f7-815">1.0</span></span>|
|[<span data-ttu-id="3d5f7-816">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="3d5f7-816">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3d5f7-817">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3d5f7-817">ReadItem</span></span>|
|[<span data-ttu-id="3d5f7-818">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="3d5f7-818">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="3d5f7-819">阅读</span><span class="sxs-lookup"><span data-stu-id="3d5f7-819">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="3d5f7-820">返回：</span><span class="sxs-lookup"><span data-stu-id="3d5f7-820">Returns:</span></span>

<span data-ttu-id="3d5f7-p153">如果具有匹配 `name` 参数的 `FilterName` 元素值的清单中没有任何 `ItemHasKnownEntity` 元素，则该方法返回 `null`。如果 `name` 参数匹配清单中的 `ItemHasKnownEntity` 元素，但在匹配的当前项目中没有实体，则该方法返回一个空数组。</span><span class="sxs-lookup"><span data-stu-id="3d5f7-p153">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="3d5f7-823">类型：Array.<(String|[Contact](/javascript/api/outlook_1_4/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_4/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_4/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_4/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="3d5f7-823">Type: Array.<(String|[Contact](/javascript/api/outlook_1_4/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_4/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_4/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_4/office.tasksuggestion))></span></span>

#### <a name="getregexmatches--object"></a><span data-ttu-id="3d5f7-824">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="3d5f7-824">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="3d5f7-825">返回所选项目中匹配在清单 XML 文件中定义的正则表达式的字符串值。</span><span class="sxs-lookup"><span data-stu-id="3d5f7-825">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="3d5f7-826">在 Outlook for iOS 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="3d5f7-826">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="3d5f7-p154">`getRegExMatches` 方法返回匹配在清单 XML 文件中的每个 `ItemHasRegularExpressionMatch` 或 `ItemHasKnownEntity` 规则元素中定义的正则表达式的字符串。对于 `ItemHasRegularExpressionMatch` 规则，匹配字符串必须发生在该规则指定的项目的属性中。`PropertyName` 简单类型定义支持的属性。</span><span class="sxs-lookup"><span data-stu-id="3d5f7-p154">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="3d5f7-830">例如，考虑一个外接程序清单具有以下 `Rule` 元素：</span><span class="sxs-lookup"><span data-stu-id="3d5f7-830">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="3d5f7-831">从 `getRegExMatches` 返回的对象应有两个属性：`fruits` 和 `veggies`。</span><span class="sxs-lookup"><span data-stu-id="3d5f7-831">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="3d5f7-p155">如果在项目的正文属性上指定 `ItemHasRegularExpressionMatch` 规则，则正则表达式应进一步筛选正文，不应尝试返回该项目的整个正文。使用正则表达式（如 `.*`）获取项目的整个正文并不总是返回预期的结果。而是使用 [`Body.getAsync`](/javascript/api/outlook_1_4/office.body#getasync-coerciontype--options--callback-) 方法检索整个正文。</span><span class="sxs-lookup"><span data-stu-id="3d5f7-p155">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook_1_4/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="3d5f7-835">Requirements</span><span class="sxs-lookup"><span data-stu-id="3d5f7-835">Requirements</span></span>

|<span data-ttu-id="3d5f7-836">要求</span><span class="sxs-lookup"><span data-stu-id="3d5f7-836">Requirement</span></span>| <span data-ttu-id="3d5f7-837">值</span><span class="sxs-lookup"><span data-stu-id="3d5f7-837">Value</span></span>|
|---|---|
|[<span data-ttu-id="3d5f7-838">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="3d5f7-838">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3d5f7-839">1.0</span><span class="sxs-lookup"><span data-stu-id="3d5f7-839">1.0</span></span>|
|[<span data-ttu-id="3d5f7-840">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="3d5f7-840">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3d5f7-841">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3d5f7-841">ReadItem</span></span>|
|[<span data-ttu-id="3d5f7-842">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="3d5f7-842">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="3d5f7-843">阅读</span><span class="sxs-lookup"><span data-stu-id="3d5f7-843">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="3d5f7-844">返回：</span><span class="sxs-lookup"><span data-stu-id="3d5f7-844">Returns:</span></span>

<span data-ttu-id="3d5f7-p156">一个包含匹配在清单 XML 文件中定义的正则表达式的字符串数组的对象。每个数组的名称等于匹配 `ItemHasRegularExpressionMatch` 规则的 `RegExName` 属性或匹配 `ItemHasKnownEntity` 规则的 `FilterName` 属性的相应值。</span><span class="sxs-lookup"><span data-stu-id="3d5f7-p156">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<dl class="param-type"><span data-ttu-id="3d5f7-847">

<dt>
类型</dt>


</span><span class="sxs-lookup"><span data-stu-id="3d5f7-847">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="3d5f7-848">对象</span><span class="sxs-lookup"><span data-stu-id="3d5f7-848">Object</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="3d5f7-849">示例</span><span class="sxs-lookup"><span data-stu-id="3d5f7-849">Example</span></span>

<span data-ttu-id="3d5f7-850">以下示例显示了如何访问正则表达式 <rule> 元素 `fruits` 和 `veggies` 的匹配项的数组，这些元素在清单中指定。</rule></span><span class="sxs-lookup"><span data-stu-id="3d5f7-850">The following example shows how to access the array of matches for the regular expression <rule>elements `fruits` and `veggies`, which are specified in the manifest.</rule></span></span>

```javascript
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veggies = allMatches.veggies;
```

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="3d5f7-851">getRegExMatchesByName(name) → (nullable) {Array.<String>}</span><span class="sxs-lookup"><span data-stu-id="3d5f7-851">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="3d5f7-852">返回匹配在清单 XML 文件中定义的命名正则表达式的所选项目中的字符串值。</span><span class="sxs-lookup"><span data-stu-id="3d5f7-852">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="3d5f7-853">在 Outlook for iOS 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="3d5f7-853">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="3d5f7-854">`getRegExMatchesByName` 方法返回匹配在具有指定 `RegExName` 元素值的清单 XML 文件中的 `ItemHasRegularExpressionMatch` 规则元素中定义的正则表达式的字符串。</span><span class="sxs-lookup"><span data-stu-id="3d5f7-854">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="3d5f7-p157">如果在项目的正文属性上指定 `ItemHasRegularExpressionMatch` 规则，则正则表达式应进一步筛选正文，不应尝试返回该项目的整个正文。使用正则表达式（如 `.*`）获取项目的整个正文并不总是返回预期的结果。</span><span class="sxs-lookup"><span data-stu-id="3d5f7-p157">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="3d5f7-857">参数</span><span class="sxs-lookup"><span data-stu-id="3d5f7-857">Parameters</span></span>

|<span data-ttu-id="3d5f7-858">名称</span><span class="sxs-lookup"><span data-stu-id="3d5f7-858">Name</span></span>| <span data-ttu-id="3d5f7-859">类型</span><span class="sxs-lookup"><span data-stu-id="3d5f7-859">Type</span></span>| <span data-ttu-id="3d5f7-860">说明</span><span class="sxs-lookup"><span data-stu-id="3d5f7-860">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="3d5f7-861">String</span><span class="sxs-lookup"><span data-stu-id="3d5f7-861">String</span></span>|<span data-ttu-id="3d5f7-862">定义筛选器匹配的 `ItemHasRegularExpressionMatch` 规则元素的名称。</span><span class="sxs-lookup"><span data-stu-id="3d5f7-862">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="3d5f7-863">要求</span><span class="sxs-lookup"><span data-stu-id="3d5f7-863">Requirements</span></span>

|<span data-ttu-id="3d5f7-864">要求</span><span class="sxs-lookup"><span data-stu-id="3d5f7-864">Requirement</span></span>| <span data-ttu-id="3d5f7-865">值</span><span class="sxs-lookup"><span data-stu-id="3d5f7-865">Value</span></span>|
|---|---|
|[<span data-ttu-id="3d5f7-866">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="3d5f7-866">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3d5f7-867">1.0</span><span class="sxs-lookup"><span data-stu-id="3d5f7-867">1.0</span></span>|
|[<span data-ttu-id="3d5f7-868">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="3d5f7-868">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3d5f7-869">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3d5f7-869">ReadItem</span></span>|
|[<span data-ttu-id="3d5f7-870">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="3d5f7-870">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="3d5f7-871">阅读</span><span class="sxs-lookup"><span data-stu-id="3d5f7-871">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="3d5f7-872">返回：</span><span class="sxs-lookup"><span data-stu-id="3d5f7-872">Returns:</span></span>

<span data-ttu-id="3d5f7-873">一个包含匹配在清单 XML 文件中定义的正则表达式的字符串的数组。</span><span class="sxs-lookup"><span data-stu-id="3d5f7-873">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<dl class="param-type"><span data-ttu-id="3d5f7-874">

<dt>类型</dt>

</span><span class="sxs-lookup"><span data-stu-id="3d5f7-874">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="3d5f7-875">Array.< String ></span><span class="sxs-lookup"><span data-stu-id="3d5f7-875">Array.< String ></span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="3d5f7-876">示例</span><span class="sxs-lookup"><span data-stu-id="3d5f7-876">Example</span></span>

```javascript
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

#### <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="3d5f7-877">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="3d5f7-877">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="3d5f7-878">以异步方式返回邮件的主题或正文中选定的数据。</span><span class="sxs-lookup"><span data-stu-id="3d5f7-878">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="3d5f7-p158">如果没有选定内容，但光标位于正文或主题中，此方法将会为所选数据返回 null。如果选定的是字段，而不是正文或主题，则此方法返回 `InvalidSelection` 错误。</span><span class="sxs-lookup"><span data-stu-id="3d5f7-p158">If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="3d5f7-881">参数</span><span class="sxs-lookup"><span data-stu-id="3d5f7-881">Parameters</span></span>

|<span data-ttu-id="3d5f7-882">名称</span><span class="sxs-lookup"><span data-stu-id="3d5f7-882">Name</span></span>| <span data-ttu-id="3d5f7-883">类型</span><span class="sxs-lookup"><span data-stu-id="3d5f7-883">Type</span></span>| <span data-ttu-id="3d5f7-884">属性</span><span class="sxs-lookup"><span data-stu-id="3d5f7-884">Attributes</span></span>| <span data-ttu-id="3d5f7-885">说明</span><span class="sxs-lookup"><span data-stu-id="3d5f7-885">Description</span></span>|
|---|---|---|---|
|`coercionType`| [<span data-ttu-id="3d5f7-886">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="3d5f7-886">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="3d5f7-p159">请求数据的格式。如果为文本，则此方法返回纯文本作为字符串，删除任何显示的 HTML 标记。如果为 HTML，则此方法返回所选文本，不论是纯文本还是 HTML。</span><span class="sxs-lookup"><span data-stu-id="3d5f7-p159">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`| <span data-ttu-id="3d5f7-890">对象</span><span class="sxs-lookup"><span data-stu-id="3d5f7-890">Object</span></span>| <span data-ttu-id="3d5f7-891">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="3d5f7-891">&lt;optional&gt;</span></span>|<span data-ttu-id="3d5f7-892">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="3d5f7-892">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="3d5f7-893">Object</span><span class="sxs-lookup"><span data-stu-id="3d5f7-893">Object</span></span>| <span data-ttu-id="3d5f7-894">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="3d5f7-894">&lt;optional&gt;</span></span>|<span data-ttu-id="3d5f7-895">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="3d5f7-895">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="3d5f7-896">函数</span><span class="sxs-lookup"><span data-stu-id="3d5f7-896">function</span></span>||<span data-ttu-id="3d5f7-897">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="3d5f7-897">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="3d5f7-898">若要从回调方法访问所选数据，请调用 `asyncResult.value.data`。</span><span class="sxs-lookup"><span data-stu-id="3d5f7-898">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="3d5f7-899">若要访问选定内容的源属性，请调用 `asyncResult.value.sourceProperty`，这将为 `body` 或 `subject`。</span><span class="sxs-lookup"><span data-stu-id="3d5f7-899">To access the source property that the selection comes from, call `asyncResult.value.sourceProperty`, which will be either `body` or `subject`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="3d5f7-900">要求</span><span class="sxs-lookup"><span data-stu-id="3d5f7-900">Requirements</span></span>

|<span data-ttu-id="3d5f7-901">要求</span><span class="sxs-lookup"><span data-stu-id="3d5f7-901">Requirement</span></span>| <span data-ttu-id="3d5f7-902">值</span><span class="sxs-lookup"><span data-stu-id="3d5f7-902">Value</span></span>|
|---|---|
|[<span data-ttu-id="3d5f7-903">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="3d5f7-903">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3d5f7-904">1.2</span><span class="sxs-lookup"><span data-stu-id="3d5f7-904">1.2</span></span>|
|[<span data-ttu-id="3d5f7-905">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="3d5f7-905">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3d5f7-906">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="3d5f7-906">ReadWriteItem</span></span>|
|[<span data-ttu-id="3d5f7-907">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="3d5f7-907">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="3d5f7-908">撰写</span><span class="sxs-lookup"><span data-stu-id="3d5f7-908">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="3d5f7-909">返回：</span><span class="sxs-lookup"><span data-stu-id="3d5f7-909">Returns:</span></span>

<span data-ttu-id="3d5f7-910">作为字符串的所选数据的格式由 `coercionType` 确定。</span><span class="sxs-lookup"><span data-stu-id="3d5f7-910">The selected data as a string with format determined by `coercionType`.</span></span>

<dl class="param-type"><span data-ttu-id="3d5f7-911">

<dt>
类型</dt>


</span><span class="sxs-lookup"><span data-stu-id="3d5f7-911">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="3d5f7-912">字符串</span><span class="sxs-lookup"><span data-stu-id="3d5f7-912">String</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="3d5f7-913">示例</span><span class="sxs-lookup"><span data-stu-id="3d5f7-913">Example</span></span>

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

#### <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="3d5f7-914">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="3d5f7-914">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="3d5f7-915">异步加载所选项目上此外接程序的自定义属性。</span><span class="sxs-lookup"><span data-stu-id="3d5f7-915">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="3d5f7-p161">自定义属性基于每个应用、每个项目存储为键/值对。此方法在回调中返回 `CustomProperties` 对象，该回调提供访问特定于当前项目和当前外接程序的自定义属性的方法。自定义属性未在项目上加密，因此这不应用作安全存储。</span><span class="sxs-lookup"><span data-stu-id="3d5f7-p161">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="3d5f7-919">参数</span><span class="sxs-lookup"><span data-stu-id="3d5f7-919">Parameters</span></span>

|<span data-ttu-id="3d5f7-920">名称</span><span class="sxs-lookup"><span data-stu-id="3d5f7-920">Name</span></span>| <span data-ttu-id="3d5f7-921">类型</span><span class="sxs-lookup"><span data-stu-id="3d5f7-921">Type</span></span>| <span data-ttu-id="3d5f7-922">属性</span><span class="sxs-lookup"><span data-stu-id="3d5f7-922">Attributes</span></span>| <span data-ttu-id="3d5f7-923">说明</span><span class="sxs-lookup"><span data-stu-id="3d5f7-923">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="3d5f7-924">函数</span><span class="sxs-lookup"><span data-stu-id="3d5f7-924">function</span></span>||<span data-ttu-id="3d5f7-925">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="3d5f7-925">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="3d5f7-926">自定义属性作为 `asyncResult.value` 属性中的 [`CustomProperties`](/javascript/api/outlook_1_4/office.customproperties) 对象提供。</span><span class="sxs-lookup"><span data-stu-id="3d5f7-926">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook_1_4/office.customproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="3d5f7-927">此对象可用于获取、设置以及从项目中删除自定义属性，并将自定义属性集的更改重新保存到服务器。</span><span class="sxs-lookup"><span data-stu-id="3d5f7-927">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`| <span data-ttu-id="3d5f7-928">对象</span><span class="sxs-lookup"><span data-stu-id="3d5f7-928">Object</span></span>| <span data-ttu-id="3d5f7-929">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="3d5f7-929">&lt;optional&gt;</span></span>|<span data-ttu-id="3d5f7-930">开发人员可以提供他们想要在回调函数中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="3d5f7-930">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="3d5f7-931">此对象可以通过回调函数中的 `asyncResult.asyncContext` 属性进行访问。</span><span class="sxs-lookup"><span data-stu-id="3d5f7-931">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="3d5f7-932">要求</span><span class="sxs-lookup"><span data-stu-id="3d5f7-932">Requirements</span></span>

|<span data-ttu-id="3d5f7-933">要求</span><span class="sxs-lookup"><span data-stu-id="3d5f7-933">Requirement</span></span>| <span data-ttu-id="3d5f7-934">值</span><span class="sxs-lookup"><span data-stu-id="3d5f7-934">Value</span></span>|
|---|---|
|[<span data-ttu-id="3d5f7-935">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="3d5f7-935">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3d5f7-936">1.0</span><span class="sxs-lookup"><span data-stu-id="3d5f7-936">1.0</span></span>|
|[<span data-ttu-id="3d5f7-937">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="3d5f7-937">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3d5f7-938">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3d5f7-938">ReadItem</span></span>|
|[<span data-ttu-id="3d5f7-939">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="3d5f7-939">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="3d5f7-940">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="3d5f7-940">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="3d5f7-941">示例</span><span class="sxs-lookup"><span data-stu-id="3d5f7-941">Example</span></span>

<span data-ttu-id="3d5f7-p164">以下代码示例显示了如何使用 `loadCustomPropertiesAsync` 方法异步加载特定于当前项目的自定义属性。该示例还显示了如何使用 `CustomProperties.saveAsync` 方法将这些属性重新保存到服务器。加载自定义属性后，该代码示例将使用 `CustomProperties.get` 方法读取自定义属性 `myProp`，使用 `CustomProperties.set` 方法写入自定义属性 `otherProp`，最后调用 `saveAsync` 方法保存这些自定义属性。</span><span class="sxs-lookup"><span data-stu-id="3d5f7-p164">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

```javascript
// The initialize function is required for all add-ins.
Office.initialize = function () {
  // Checks for the DOM to load using the jQuery ready function.
  $(document).ready(function () {
    // After the DOM is loaded, add-in-specific code can run.
    var item = Office.context.mailbox.item;
    item.loadCustomPropertiesAsync(customPropsCallback);
  });
};

function customPropsCallback(asyncResult) {
  var customProps = asyncResult.value;
  var myProp = customProps.get("myProp");

  customProps.set("otherProp", "value");
  customProps.saveAsync(saveCallback);
}

function saveCallback(asyncResult) {
}
```

#### <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="3d5f7-945">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="3d5f7-945">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="3d5f7-946">将附件从邮件或约会中删除。</span><span class="sxs-lookup"><span data-stu-id="3d5f7-946">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="3d5f7-p165">`removeAttachmentAsync` 方法删除项目中带指定标识符的附件。最佳做法是，仅当同一个邮件应用程序在同一会话中添加了一个附件时，你才应使用该附件标识符来删除该附件。在 Outlook Web App 和适用于设备的 OWA 中，附件标识符只在同一个会话中才有效。当用户关闭应用程序，或者如果用户开始在内嵌窗体中撰写，并在随后弹出的内嵌窗体中继续在单独的窗口撰写时，会话即结束。</span><span class="sxs-lookup"><span data-stu-id="3d5f7-p165">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item. As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session. In Outlook Web App and OWA for Devices, the attachment identifier is valid only within the same session. A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="3d5f7-951">参数</span><span class="sxs-lookup"><span data-stu-id="3d5f7-951">Parameters</span></span>

|<span data-ttu-id="3d5f7-952">名称</span><span class="sxs-lookup"><span data-stu-id="3d5f7-952">Name</span></span>| <span data-ttu-id="3d5f7-953">类型</span><span class="sxs-lookup"><span data-stu-id="3d5f7-953">Type</span></span>| <span data-ttu-id="3d5f7-954">属性</span><span class="sxs-lookup"><span data-stu-id="3d5f7-954">Attributes</span></span>| <span data-ttu-id="3d5f7-955">说明</span><span class="sxs-lookup"><span data-stu-id="3d5f7-955">Description</span></span>|
|---|---|---|---|
|`attachmentId`| <span data-ttu-id="3d5f7-956">String</span><span class="sxs-lookup"><span data-stu-id="3d5f7-956">String</span></span>||<span data-ttu-id="3d5f7-957">要删除的附件的标识符。</span><span class="sxs-lookup"><span data-stu-id="3d5f7-957">The identifier of the attachment to remove.</span></span>|
|`options`| <span data-ttu-id="3d5f7-958">对象</span><span class="sxs-lookup"><span data-stu-id="3d5f7-958">Object</span></span>| <span data-ttu-id="3d5f7-959">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="3d5f7-959">&lt;optional&gt;</span></span>|<span data-ttu-id="3d5f7-960">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="3d5f7-960">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="3d5f7-961">对象</span><span class="sxs-lookup"><span data-stu-id="3d5f7-961">Object</span></span>| <span data-ttu-id="3d5f7-962">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="3d5f7-962">&lt;optional&gt;</span></span>|<span data-ttu-id="3d5f7-963">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="3d5f7-963">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="3d5f7-964">函数</span><span class="sxs-lookup"><span data-stu-id="3d5f7-964">function</span></span>| <span data-ttu-id="3d5f7-965">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="3d5f7-965">&lt;optional&gt;</span></span>|<span data-ttu-id="3d5f7-966">方法完成后，使用单个参数 `callback`（一个 [`asyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `AsyncResult` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="3d5f7-966">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="3d5f7-967">如果删除附件失败，`asyncResult.error` 属性将包含一个说明失败原因的错误代码。</span><span class="sxs-lookup"><span data-stu-id="3d5f7-967">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="3d5f7-968">错误</span><span class="sxs-lookup"><span data-stu-id="3d5f7-968">Errors</span></span>

| <span data-ttu-id="3d5f7-969">错误代码</span><span class="sxs-lookup"><span data-stu-id="3d5f7-969">Error code</span></span> | <span data-ttu-id="3d5f7-970">说明</span><span class="sxs-lookup"><span data-stu-id="3d5f7-970">Description</span></span> |
|------------|-------------|
| `InvalidAttachmentId` | <span data-ttu-id="3d5f7-971">附件标识符不存在。</span><span class="sxs-lookup"><span data-stu-id="3d5f7-971">The attachment identifier does not exist.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="3d5f7-972">要求</span><span class="sxs-lookup"><span data-stu-id="3d5f7-972">Requirements</span></span>

|<span data-ttu-id="3d5f7-973">要求</span><span class="sxs-lookup"><span data-stu-id="3d5f7-973">Requirement</span></span>| <span data-ttu-id="3d5f7-974">值</span><span class="sxs-lookup"><span data-stu-id="3d5f7-974">Value</span></span>|
|---|---|
|[<span data-ttu-id="3d5f7-975">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="3d5f7-975">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3d5f7-976">1.1</span><span class="sxs-lookup"><span data-stu-id="3d5f7-976">1.1</span></span>|
|[<span data-ttu-id="3d5f7-977">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="3d5f7-977">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3d5f7-978">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="3d5f7-978">ReadWriteItem</span></span>|
|[<span data-ttu-id="3d5f7-979">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="3d5f7-979">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="3d5f7-980">撰写</span><span class="sxs-lookup"><span data-stu-id="3d5f7-980">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="3d5f7-981">示例</span><span class="sxs-lookup"><span data-stu-id="3d5f7-981">Example</span></span>

<span data-ttu-id="3d5f7-982">以下代码删除包含标识符 '0' 的附件。</span><span class="sxs-lookup"><span data-stu-id="3d5f7-982">The following code removes an attachment with an identifier of '0'.</span></span>

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

#### <a name="saveasyncoptions-callback"></a><span data-ttu-id="3d5f7-983">saveAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="3d5f7-983">saveAsync([options], callback)</span></span>

<span data-ttu-id="3d5f7-984">异步保存项目。</span><span class="sxs-lookup"><span data-stu-id="3d5f7-984">Asynchronously saves an item.</span></span>

<span data-ttu-id="3d5f7-p166">调用时，此方法将当前邮件保存为草稿，并通过回调方法返回项目 ID。在 Outlook Web App 或 Outlook 联机模式下，该项目被保存到服务器中。在 Outlook 缓存模式下，该项目被保存到本地缓存中。</span><span class="sxs-lookup"><span data-stu-id="3d5f7-p166">When invoked, this method saves the current message as a draft and returns the item id via the callback method. In Outlook Web App or Outlook in online mode, the item is saved to the server. In Outlook in cached mode, the item is saved to the local cache.</span></span>

> [!NOTE]
> <span data-ttu-id="3d5f7-988">如果加载项在撰写模式下对某个项目调用 `saveAsync` 来获得 `itemId`，以便与 EWS 或 REST API 一同使用，请注意，当 Outlook 处于高速缓存模式时，可能需要一段时间项目才能真正同步到服务器。</span><span class="sxs-lookup"><span data-stu-id="3d5f7-988">If your add-in calls `saveAsync` on an item in compose mode in order to get an `itemId` to use with EWS or the REST API, be aware that when Outlook is in cached mode, it may take some time before the item is actually synced to the server.</span></span> <span data-ttu-id="3d5f7-989">在项目同步前，使用 `itemId` 将返回一个错误。</span><span class="sxs-lookup"><span data-stu-id="3d5f7-989">Until the item is synced, using the `itemId` will return an error.</span></span>

<span data-ttu-id="3d5f7-p168">由于约会没有草稿状态，如果以撰写模式在约会中调用 `saveAsync`，则该项将被保存为用户日历中的正常约会。对于之前未保存过的新约会，则不会发送邀请。保存现有约会将向添加或删除的与会者发送更新。</span><span class="sxs-lookup"><span data-stu-id="3d5f7-p168">Since appointments have no draft state, if `saveAsync` is called on an appointment in compose mode, the item will be saved as a normal appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent. Saving an existing appointment will send an update to added or removed attendees.</span></span>

> [!NOTE]
> <span data-ttu-id="3d5f7-993">以下客户端在撰写模式下对约会上的 `saveAsync` 具有不同的行为：</span><span class="sxs-lookup"><span data-stu-id="3d5f7-993">The following clients have different behavior for `saveAsync` on appointments in compose mode:</span></span>
>
> - <span data-ttu-id="3d5f7-994">Outlook for Mac 不支持保存会议。</span><span class="sxs-lookup"><span data-stu-id="3d5f7-994">Outlook for Mac does not support saving a meeting.</span></span> <span data-ttu-id="3d5f7-995">在`saveAsync`撰写模式下从会议中调用时, 此方法将失败。</span><span class="sxs-lookup"><span data-stu-id="3d5f7-995">The `saveAsync` method fails when called from a meeting in compose mode.</span></span> <span data-ttu-id="3d5f7-996">若要解决此问题, 请参阅[使用 OFFICE JS API 将会议保存为 Outlook For Mac 中的草稿](https://support.microsoft.com/help/4505745)。</span><span class="sxs-lookup"><span data-stu-id="3d5f7-996">See [Cannot save a meeting as a draft in Outlook for Mac by using Office JS API](https://support.microsoft.com/help/4505745) for a workaround.</span></span>
> - <span data-ttu-id="3d5f7-997">在撰写模式下的约会上调用 `saveAsync` 时，Outlook 网页版始终发送邀请或更新。</span><span class="sxs-lookup"><span data-stu-id="3d5f7-997">Outlook on the web always sends an invitation or update when `saveAsync` is called on an appointment in compose mode.</span></span>

##### <a name="parameters"></a><span data-ttu-id="3d5f7-998">参数</span><span class="sxs-lookup"><span data-stu-id="3d5f7-998">Parameters</span></span>

|<span data-ttu-id="3d5f7-999">名称</span><span class="sxs-lookup"><span data-stu-id="3d5f7-999">Name</span></span>| <span data-ttu-id="3d5f7-1000">类型</span><span class="sxs-lookup"><span data-stu-id="3d5f7-1000">Type</span></span>| <span data-ttu-id="3d5f7-1001">属性</span><span class="sxs-lookup"><span data-stu-id="3d5f7-1001">Attributes</span></span>| <span data-ttu-id="3d5f7-1002">说明</span><span class="sxs-lookup"><span data-stu-id="3d5f7-1002">Description</span></span>|
|---|---|---|---|
|`options`| <span data-ttu-id="3d5f7-1003">对象</span><span class="sxs-lookup"><span data-stu-id="3d5f7-1003">Object</span></span>| <span data-ttu-id="3d5f7-1004">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="3d5f7-1004">&lt;optional&gt;</span></span>|<span data-ttu-id="3d5f7-1005">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="3d5f7-1005">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="3d5f7-1006">对象</span><span class="sxs-lookup"><span data-stu-id="3d5f7-1006">Object</span></span>| <span data-ttu-id="3d5f7-1007">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="3d5f7-1007">&lt;optional&gt;</span></span>|<span data-ttu-id="3d5f7-1008">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="3d5f7-1008">Developers can provide any object they wish to access in the callback method.</span></span>||
|`callback`| <span data-ttu-id="3d5f7-1009">function</span><span class="sxs-lookup"><span data-stu-id="3d5f7-1009">function</span></span>||<span data-ttu-id="3d5f7-1010">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="3d5f7-1010">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="3d5f7-1011">如果成功，该项目标识符将在 `asyncResult.value` 属性中提供。</span><span class="sxs-lookup"><span data-stu-id="3d5f7-1011">On success, the item identifier is provided in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="3d5f7-1012">要求</span><span class="sxs-lookup"><span data-stu-id="3d5f7-1012">Requirements</span></span>

|<span data-ttu-id="3d5f7-1013">要求</span><span class="sxs-lookup"><span data-stu-id="3d5f7-1013">Requirement</span></span>| <span data-ttu-id="3d5f7-1014">值</span><span class="sxs-lookup"><span data-stu-id="3d5f7-1014">Value</span></span>|
|---|---|
|[<span data-ttu-id="3d5f7-1015">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="3d5f7-1015">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3d5f7-1016">1.3</span><span class="sxs-lookup"><span data-stu-id="3d5f7-1016">1.3</span></span>|
|[<span data-ttu-id="3d5f7-1017">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="3d5f7-1017">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3d5f7-1018">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="3d5f7-1018">ReadWriteItem</span></span>|
|[<span data-ttu-id="3d5f7-1019">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="3d5f7-1019">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="3d5f7-1020">撰写</span><span class="sxs-lookup"><span data-stu-id="3d5f7-1020">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="3d5f7-1021">示例</span><span class="sxs-lookup"><span data-stu-id="3d5f7-1021">Examples</span></span>

```javascript
Office.context.mailbox.item.saveAsync(
  function callback(result) {
    // Process the result.
  });
```

<span data-ttu-id="3d5f7-p170">下面是传递给回调函数的 `result` 参数的示例。`value` 属性包含的项目的项目 ID。</span><span class="sxs-lookup"><span data-stu-id="3d5f7-p170">The following is an example of the `result` parameter passed to the callback function. The `value` property contains the item ID of the item.</span></span>

```json
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

#### <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="3d5f7-1024">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="3d5f7-1024">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="3d5f7-1025">以异步方式将数据插入到邮件的正文或主题中。</span><span class="sxs-lookup"><span data-stu-id="3d5f7-1025">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="3d5f7-p171">`setSelectedDataAsync` 方法将指定字符串插入到项目主题或正文的光标位置，或者，如果在编辑器中已选择文本，则该方法将替换选择的文本。如果光标不在正文或主题字段中，则返回错误。插入之后，光标会位于插入内容的末尾。</span><span class="sxs-lookup"><span data-stu-id="3d5f7-p171">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="3d5f7-1029">参数</span><span class="sxs-lookup"><span data-stu-id="3d5f7-1029">Parameters</span></span>

|<span data-ttu-id="3d5f7-1030">名称</span><span class="sxs-lookup"><span data-stu-id="3d5f7-1030">Name</span></span>| <span data-ttu-id="3d5f7-1031">类型</span><span class="sxs-lookup"><span data-stu-id="3d5f7-1031">Type</span></span>| <span data-ttu-id="3d5f7-1032">属性</span><span class="sxs-lookup"><span data-stu-id="3d5f7-1032">Attributes</span></span>| <span data-ttu-id="3d5f7-1033">说明</span><span class="sxs-lookup"><span data-stu-id="3d5f7-1033">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="3d5f7-1034">字符串</span><span class="sxs-lookup"><span data-stu-id="3d5f7-1034">String</span></span>||<span data-ttu-id="3d5f7-p172">要插入的数据。数据不得超过 1,000,000 个字符。如果传入的数据超过 1,000,000 个字符，则会引发 `ArgumentOutOfRange` 异常。</span><span class="sxs-lookup"><span data-stu-id="3d5f7-p172">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`| <span data-ttu-id="3d5f7-1038">Object</span><span class="sxs-lookup"><span data-stu-id="3d5f7-1038">Object</span></span>| <span data-ttu-id="3d5f7-1039">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="3d5f7-1039">&lt;optional&gt;</span></span>|<span data-ttu-id="3d5f7-1040">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="3d5f7-1040">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="3d5f7-1041">对象</span><span class="sxs-lookup"><span data-stu-id="3d5f7-1041">Object</span></span>| <span data-ttu-id="3d5f7-1042">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="3d5f7-1042">&lt;optional&gt;</span></span>|<span data-ttu-id="3d5f7-1043">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="3d5f7-1043">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`| [<span data-ttu-id="3d5f7-1044">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="3d5f7-1044">Office.CoercionType</span></span>](office.md#coerciontype-string)| <span data-ttu-id="3d5f7-1045">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="3d5f7-1045">&lt;optional&gt;</span></span>|<span data-ttu-id="3d5f7-p173">如果为 `text`，则在 Outlook Web App 和 Outlook 中应用当前样式。如果该字段是 HTML 编辑器，则仅插入文本数据，即使数据为 HTML。</span><span class="sxs-lookup"><span data-stu-id="3d5f7-p173">If `text`, the current style is applied in Outlook Web App and Outlook. If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="3d5f7-p174">如果 `html` 和该字段支持 HTML（主题不支持），则在 Outlook Web App 中应用当前样式，而在 Outlook 中应用默认样式。如果该字段是文本字段，则返回 `InvalidDataFormat` 错误。</span><span class="sxs-lookup"><span data-stu-id="3d5f7-p174">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook Web App and the default style is applied in Outlook. If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="3d5f7-1050">如果未设置 `coercionType`，则结果取决于该字段：如果该字段是 HTML，则使用 HTML；如果该字段是文本，则使用纯文本。</span><span class="sxs-lookup"><span data-stu-id="3d5f7-1050">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`| <span data-ttu-id="3d5f7-1051">function</span><span class="sxs-lookup"><span data-stu-id="3d5f7-1051">function</span></span>||<span data-ttu-id="3d5f7-1052">方法完成后，使用单个参数 `callback`（一个 [`asyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `AsyncResult` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="3d5f7-1052">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="3d5f7-1053">Requirements</span><span class="sxs-lookup"><span data-stu-id="3d5f7-1053">Requirements</span></span>

|<span data-ttu-id="3d5f7-1054">要求</span><span class="sxs-lookup"><span data-stu-id="3d5f7-1054">Requirement</span></span>| <span data-ttu-id="3d5f7-1055">值</span><span class="sxs-lookup"><span data-stu-id="3d5f7-1055">Value</span></span>|
|---|---|
|[<span data-ttu-id="3d5f7-1056">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="3d5f7-1056">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3d5f7-1057">1.2</span><span class="sxs-lookup"><span data-stu-id="3d5f7-1057">1.2</span></span>|
|[<span data-ttu-id="3d5f7-1058">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="3d5f7-1058">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3d5f7-1059">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="3d5f7-1059">ReadWriteItem</span></span>|
|[<span data-ttu-id="3d5f7-1060">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="3d5f7-1060">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="3d5f7-1061">撰写</span><span class="sxs-lookup"><span data-stu-id="3d5f7-1061">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="3d5f7-1062">示例</span><span class="sxs-lookup"><span data-stu-id="3d5f7-1062">Example</span></span>

```javascript
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```
