---
title: "\"Context\"-\"邮箱\"。项目-要求集1。2"
description: ''
ms.date: 06/20/2019
localization_priority: Normal
ms.openlocfilehash: dabc3717738539f92d211acf6d407ff06a796a89
ms.sourcegitcommit: 3f5d7f4794e3d3c8bc3a79fa05c54157613b9376
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 08/02/2019
ms.locfileid: "36064408"
---
# <a name="item"></a><span data-ttu-id="9e14c-102">item</span><span class="sxs-lookup"><span data-stu-id="9e14c-102">item</span></span>

### <span data-ttu-id="9e14c-p101">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md). 项目</span><span class="sxs-lookup"><span data-stu-id="9e14c-p101">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md). item</span></span>

<span data-ttu-id="9e14c-p102">`item` 命名空间用于访问当前选定的邮件、会议请求或约会。可以通过使用 [itemType](#itemtype-officemailboxenumsitemtype) 属性确定 `item` 的类型。</span><span class="sxs-lookup"><span data-stu-id="9e14c-p102">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="9e14c-107">Requirements</span><span class="sxs-lookup"><span data-stu-id="9e14c-107">Requirements</span></span>

|<span data-ttu-id="9e14c-108">要求</span><span class="sxs-lookup"><span data-stu-id="9e14c-108">Requirement</span></span>| <span data-ttu-id="9e14c-109">值</span><span class="sxs-lookup"><span data-stu-id="9e14c-109">Value</span></span>|
|---|---|
|[<span data-ttu-id="9e14c-110">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="9e14c-110">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9e14c-111">1.0</span><span class="sxs-lookup"><span data-stu-id="9e14c-111">1.0</span></span>|
|[<span data-ttu-id="9e14c-112">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="9e14c-112">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9e14c-113">受限</span><span class="sxs-lookup"><span data-stu-id="9e14c-113">Restricted</span></span>|
|[<span data-ttu-id="9e14c-114">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="9e14c-114">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="9e14c-115">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="9e14c-115">Compose or Read</span></span>|

### <a name="example"></a><span data-ttu-id="9e14c-116">示例</span><span class="sxs-lookup"><span data-stu-id="9e14c-116">Example</span></span>

<span data-ttu-id="9e14c-117">以下 JavaScript 代码示例显示了如何访问 Outlook 中当前项目的 `subject` 属性。</span><span class="sxs-lookup"><span data-stu-id="9e14c-117">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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

### <a name="members"></a><span data-ttu-id="9e14c-118">成员</span><span class="sxs-lookup"><span data-stu-id="9e14c-118">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetailsviewoutlook-js-12"></a><span data-ttu-id="9e14c-119">附件: Array. <[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.2)></span><span class="sxs-lookup"><span data-stu-id="9e14c-119">attachments: Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.2)></span></span>

<span data-ttu-id="9e14c-p103">获取项目的附件数组。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="9e14c-p103">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="9e14c-122">由于潜在的安全问题，某些类型的文件会受到 Outlook 阻止，并且不会返回。</span><span class="sxs-lookup"><span data-stu-id="9e14c-122">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="9e14c-123">如需了解更多信息，请参阅 [Outlook 中阻止的附件](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519)。</span><span class="sxs-lookup"><span data-stu-id="9e14c-123">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="9e14c-124">类型</span><span class="sxs-lookup"><span data-stu-id="9e14c-124">Type</span></span>

*   <span data-ttu-id="9e14c-125">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.2)></span><span class="sxs-lookup"><span data-stu-id="9e14c-125">Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.2)></span></span>

##### <a name="requirements"></a><span data-ttu-id="9e14c-126">要求</span><span class="sxs-lookup"><span data-stu-id="9e14c-126">Requirements</span></span>

|<span data-ttu-id="9e14c-127">要求</span><span class="sxs-lookup"><span data-stu-id="9e14c-127">Requirement</span></span>| <span data-ttu-id="9e14c-128">值</span><span class="sxs-lookup"><span data-stu-id="9e14c-128">Value</span></span>|
|---|---|
|[<span data-ttu-id="9e14c-129">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="9e14c-129">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9e14c-130">1.0</span><span class="sxs-lookup"><span data-stu-id="9e14c-130">1.0</span></span>|
|[<span data-ttu-id="9e14c-131">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="9e14c-131">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9e14c-132">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9e14c-132">ReadItem</span></span>|
|[<span data-ttu-id="9e14c-133">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="9e14c-133">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="9e14c-134">阅读</span><span class="sxs-lookup"><span data-stu-id="9e14c-134">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="9e14c-135">示例</span><span class="sxs-lookup"><span data-stu-id="9e14c-135">Example</span></span>

<span data-ttu-id="9e14c-136">以下代码使用当前项目上所有附件的详细信息构成 HTML 字符串。</span><span class="sxs-lookup"><span data-stu-id="9e14c-136">The following code builds an HTML string with details of all attachments on the current item.</span></span>

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

#### <a name="bcc-recipientsjavascriptapioutlookofficerecipientsviewoutlook-js-12"></a><span data-ttu-id="9e14c-137">密件抄送:[收件人](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="9e14c-137">bcc: [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span></span>

<span data-ttu-id="9e14c-138">获取对象，该对象提供用于获取或更新邮件的密件抄送 (Bcc) 行上的收件人的方法。</span><span class="sxs-lookup"><span data-stu-id="9e14c-138">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="9e14c-139">仅限撰写模式。</span><span class="sxs-lookup"><span data-stu-id="9e14c-139">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="9e14c-140">类型</span><span class="sxs-lookup"><span data-stu-id="9e14c-140">Type</span></span>

*   [<span data-ttu-id="9e14c-141">收件人</span><span class="sxs-lookup"><span data-stu-id="9e14c-141">Recipients</span></span>](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)

##### <a name="requirements"></a><span data-ttu-id="9e14c-142">要求</span><span class="sxs-lookup"><span data-stu-id="9e14c-142">Requirements</span></span>

|<span data-ttu-id="9e14c-143">要求</span><span class="sxs-lookup"><span data-stu-id="9e14c-143">Requirement</span></span>| <span data-ttu-id="9e14c-144">值</span><span class="sxs-lookup"><span data-stu-id="9e14c-144">Value</span></span>|
|---|---|
|[<span data-ttu-id="9e14c-145">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="9e14c-145">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9e14c-146">1.1</span><span class="sxs-lookup"><span data-stu-id="9e14c-146">1.1</span></span>|
|[<span data-ttu-id="9e14c-147">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="9e14c-147">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9e14c-148">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9e14c-148">ReadItem</span></span>|
|[<span data-ttu-id="9e14c-149">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="9e14c-149">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="9e14c-150">撰写</span><span class="sxs-lookup"><span data-stu-id="9e14c-150">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="9e14c-151">示例</span><span class="sxs-lookup"><span data-stu-id="9e14c-151">Example</span></span>

```javascript
Office.context.mailbox.item.bcc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.bcc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.bcc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfBccRecipients = asyncResult.value;
}
```

#### <a name="body-bodyjavascriptapioutlookofficebodyviewoutlook-js-12"></a><span data-ttu-id="9e14c-152">正文:[正文](/javascript/api/outlook/office.body?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="9e14c-152">body: [Body](/javascript/api/outlook/office.body?view=outlook-js-1.2)</span></span>

<span data-ttu-id="9e14c-153">获取一个提供用于处理项目正文的方法的对象。</span><span class="sxs-lookup"><span data-stu-id="9e14c-153">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="9e14c-154">类型</span><span class="sxs-lookup"><span data-stu-id="9e14c-154">Type</span></span>

*   [<span data-ttu-id="9e14c-155">Body</span><span class="sxs-lookup"><span data-stu-id="9e14c-155">Body</span></span>](/javascript/api/outlook/office.body?view=outlook-js-1.2)

##### <a name="requirements"></a><span data-ttu-id="9e14c-156">要求</span><span class="sxs-lookup"><span data-stu-id="9e14c-156">Requirements</span></span>

|<span data-ttu-id="9e14c-157">要求</span><span class="sxs-lookup"><span data-stu-id="9e14c-157">Requirement</span></span>| <span data-ttu-id="9e14c-158">值</span><span class="sxs-lookup"><span data-stu-id="9e14c-158">Value</span></span>|
|---|---|
|[<span data-ttu-id="9e14c-159">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="9e14c-159">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9e14c-160">1.1</span><span class="sxs-lookup"><span data-stu-id="9e14c-160">1.1</span></span>|
|[<span data-ttu-id="9e14c-161">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="9e14c-161">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9e14c-162">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9e14c-162">ReadItem</span></span>|
|[<span data-ttu-id="9e14c-163">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="9e14c-163">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="9e14c-164">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="9e14c-164">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="9e14c-165">示例</span><span class="sxs-lookup"><span data-stu-id="9e14c-165">Example</span></span>

<span data-ttu-id="9e14c-166">本示例获取纯文本格式的邮件正文。</span><span class="sxs-lookup"><span data-stu-id="9e14c-166">This example gets the body of the message in plain text.</span></span>

```javascript
Office.context.mailbox.item.body.getAsync(
  "text",
  { asyncContext: "This is passed to the callback" },
  function callback(result) {
    // Do something with the result.
  });

```

<span data-ttu-id="9e14c-167">以下是传递到回调函数的结果参数的示例。</span><span class="sxs-lookup"><span data-stu-id="9e14c-167">The following is an example of the result parameter passed to the callback function.</span></span>

```json
{
  "value": "TEXT of whole body (including threads below)",
  "status": "succeeded",
  "asyncContext": "This is passed to the callback"
}
```

#### <a name="cc-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipientsviewoutlook-js-12"></a><span data-ttu-id="9e14c-168"><[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[收件人](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)的抄送: Array</span><span class="sxs-lookup"><span data-stu-id="9e14c-168">cc: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span></span>

<span data-ttu-id="9e14c-169">提供对邮件的抄送 (Cc) 收件人的访问权限。</span><span class="sxs-lookup"><span data-stu-id="9e14c-169">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="9e14c-170">对象的类型和访问级别取决于当前项目的模式。</span><span class="sxs-lookup"><span data-stu-id="9e14c-170">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="9e14c-171">阅读模式</span><span class="sxs-lookup"><span data-stu-id="9e14c-171">Read mode</span></span>

<span data-ttu-id="9e14c-p107">`cc` 属性返回包含邮件的**抄送**行上所列的每个收件人的 `EmailAddressDetails` 对象的数组。集合上限为 100 个成员。</span><span class="sxs-lookup"><span data-stu-id="9e14c-p107">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.</span></span>

```javascript
console.log(JSON.stringify(Office.context.mailbox.item.cc));
```

##### <a name="compose-mode"></a><span data-ttu-id="9e14c-174">撰写模式</span><span class="sxs-lookup"><span data-stu-id="9e14c-174">Compose mode</span></span>

<span data-ttu-id="9e14c-175">`cc` 属性返回一个 `Recipients` 对象，该对象提供用于获取或更新邮件的“**抄送**”行上收件人的方法。</span><span class="sxs-lookup"><span data-stu-id="9e14c-175">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span>

```javascript
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="9e14c-176">类型</span><span class="sxs-lookup"><span data-stu-id="9e14c-176">Type</span></span>

*   <span data-ttu-id="9e14c-177">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="9e14c-177">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span></span>

##### <a name="requirements"></a><span data-ttu-id="9e14c-178">要求</span><span class="sxs-lookup"><span data-stu-id="9e14c-178">Requirements</span></span>

|<span data-ttu-id="9e14c-179">要求</span><span class="sxs-lookup"><span data-stu-id="9e14c-179">Requirement</span></span>| <span data-ttu-id="9e14c-180">值</span><span class="sxs-lookup"><span data-stu-id="9e14c-180">Value</span></span>|
|---|---|
|[<span data-ttu-id="9e14c-181">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="9e14c-181">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9e14c-182">1.0</span><span class="sxs-lookup"><span data-stu-id="9e14c-182">1.0</span></span>|
|[<span data-ttu-id="9e14c-183">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="9e14c-183">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9e14c-184">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9e14c-184">ReadItem</span></span>|
|[<span data-ttu-id="9e14c-185">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="9e14c-185">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="9e14c-186">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="9e14c-186">Compose or Read</span></span>|

#### <a name="nullable-conversationid-string"></a><span data-ttu-id="9e14c-187">(可以为 null) conversationId: String</span><span class="sxs-lookup"><span data-stu-id="9e14c-187">(nullable) conversationId: String</span></span>

<span data-ttu-id="9e14c-188">获取包含特定消息的电子邮件会话的标识符。</span><span class="sxs-lookup"><span data-stu-id="9e14c-188">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="9e14c-p108">如果在阅读窗体或撰写窗体的回复中激活邮件应用程序，则此属性可以获得一个整数值。如果用户随后更改了回复邮件的主题（若发送回复），则该邮件的对话 ID 将改变且之前获取的值将不适用。</span><span class="sxs-lookup"><span data-stu-id="9e14c-p108">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="9e14c-p109">对于撰写窗体的新项目，此属性获得一个 null 值。如果用户设置一个主题并保存该项目，`conversationId` 属性将返回一个值。</span><span class="sxs-lookup"><span data-stu-id="9e14c-p109">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="9e14c-193">Type</span><span class="sxs-lookup"><span data-stu-id="9e14c-193">Type</span></span>

*   <span data-ttu-id="9e14c-194">String</span><span class="sxs-lookup"><span data-stu-id="9e14c-194">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="9e14c-195">要求</span><span class="sxs-lookup"><span data-stu-id="9e14c-195">Requirements</span></span>

|<span data-ttu-id="9e14c-196">要求</span><span class="sxs-lookup"><span data-stu-id="9e14c-196">Requirement</span></span>| <span data-ttu-id="9e14c-197">值</span><span class="sxs-lookup"><span data-stu-id="9e14c-197">Value</span></span>|
|---|---|
|[<span data-ttu-id="9e14c-198">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="9e14c-198">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9e14c-199">1.0</span><span class="sxs-lookup"><span data-stu-id="9e14c-199">1.0</span></span>|
|[<span data-ttu-id="9e14c-200">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="9e14c-200">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9e14c-201">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9e14c-201">ReadItem</span></span>|
|[<span data-ttu-id="9e14c-202">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="9e14c-202">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="9e14c-203">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="9e14c-203">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="9e14c-204">示例</span><span class="sxs-lookup"><span data-stu-id="9e14c-204">Example</span></span>

```javascript
var conversationId = Office.context.mailbox.item.conversationId;
console.log("conversationId: " + conversationId);
```

#### <a name="datetimecreated-date"></a><span data-ttu-id="9e14c-205">dateTimeCreated: Date</span><span class="sxs-lookup"><span data-stu-id="9e14c-205">dateTimeCreated: Date</span></span>

<span data-ttu-id="9e14c-p110">获取项目创建的日期和时间。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="9e14c-p110">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="9e14c-208">类型</span><span class="sxs-lookup"><span data-stu-id="9e14c-208">Type</span></span>

*   <span data-ttu-id="9e14c-209">日期</span><span class="sxs-lookup"><span data-stu-id="9e14c-209">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="9e14c-210">要求</span><span class="sxs-lookup"><span data-stu-id="9e14c-210">Requirements</span></span>

|<span data-ttu-id="9e14c-211">要求</span><span class="sxs-lookup"><span data-stu-id="9e14c-211">Requirement</span></span>| <span data-ttu-id="9e14c-212">值</span><span class="sxs-lookup"><span data-stu-id="9e14c-212">Value</span></span>|
|---|---|
|[<span data-ttu-id="9e14c-213">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="9e14c-213">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9e14c-214">1.0</span><span class="sxs-lookup"><span data-stu-id="9e14c-214">1.0</span></span>|
|[<span data-ttu-id="9e14c-215">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="9e14c-215">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9e14c-216">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9e14c-216">ReadItem</span></span>|
|[<span data-ttu-id="9e14c-217">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="9e14c-217">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="9e14c-218">阅读</span><span class="sxs-lookup"><span data-stu-id="9e14c-218">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="9e14c-219">示例</span><span class="sxs-lookup"><span data-stu-id="9e14c-219">Example</span></span>

```javascript
var dateTimeCreated = Office.context.mailbox.item.dateTimeCreated;
console.log("Date and time created: " + dateTimeCreated);
```

#### <a name="datetimemodified-date"></a><span data-ttu-id="9e14c-220">dateTimeModified: Date</span><span class="sxs-lookup"><span data-stu-id="9e14c-220">dateTimeModified: Date</span></span>

<span data-ttu-id="9e14c-221">获取项目最近一次修改的日期和时间。</span><span class="sxs-lookup"><span data-stu-id="9e14c-221">Gets the date and time that an item was last modified.</span></span> <span data-ttu-id="9e14c-222">仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="9e14c-222">Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="9e14c-223">IOS 或 Android 上的 Outlook 不支持此成员。</span><span class="sxs-lookup"><span data-stu-id="9e14c-223">This member is not supported in Outlook on iOS or Android.</span></span>

##### <a name="type"></a><span data-ttu-id="9e14c-224">类型</span><span class="sxs-lookup"><span data-stu-id="9e14c-224">Type</span></span>

*   <span data-ttu-id="9e14c-225">日期</span><span class="sxs-lookup"><span data-stu-id="9e14c-225">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="9e14c-226">要求</span><span class="sxs-lookup"><span data-stu-id="9e14c-226">Requirements</span></span>

|<span data-ttu-id="9e14c-227">要求</span><span class="sxs-lookup"><span data-stu-id="9e14c-227">Requirement</span></span>| <span data-ttu-id="9e14c-228">值</span><span class="sxs-lookup"><span data-stu-id="9e14c-228">Value</span></span>|
|---|---|
|[<span data-ttu-id="9e14c-229">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="9e14c-229">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9e14c-230">1.0</span><span class="sxs-lookup"><span data-stu-id="9e14c-230">1.0</span></span>|
|[<span data-ttu-id="9e14c-231">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="9e14c-231">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9e14c-232">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9e14c-232">ReadItem</span></span>|
|[<span data-ttu-id="9e14c-233">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="9e14c-233">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="9e14c-234">阅读</span><span class="sxs-lookup"><span data-stu-id="9e14c-234">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="9e14c-235">示例</span><span class="sxs-lookup"><span data-stu-id="9e14c-235">Example</span></span>

```javascript
var dateTimeModified = Office.context.mailbox.item.dateTimeModified;
console.log("Date and time modified: " + dateTimeModified);
```

#### <a name="end-datetimejavascriptapioutlookofficetimeviewoutlook-js-12"></a><span data-ttu-id="9e14c-236">结束: 日期 |[时间](/javascript/api/outlook/office.time?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="9e14c-236">end: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.2)</span></span>

<span data-ttu-id="9e14c-237">获取或设置约会结束的日期和时间。</span><span class="sxs-lookup"><span data-stu-id="9e14c-237">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="9e14c-p112">将 `end` 属性表示为协调世界时 (UTC) 的日期和时间值。可使用 [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) 方法将 end 属性值转换为客户端的本地日期和时间。</span><span class="sxs-lookup"><span data-stu-id="9e14c-p112">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="9e14c-240">阅读模式</span><span class="sxs-lookup"><span data-stu-id="9e14c-240">Read mode</span></span>

<span data-ttu-id="9e14c-241">`end` 属性返回 `Date` 对象。</span><span class="sxs-lookup"><span data-stu-id="9e14c-241">The `end` property returns a `Date` object.</span></span>

```javascript
var end = Office.context.mailbox.item.end;
console.log("Appointment end: " + end);
```

##### <a name="compose-mode"></a><span data-ttu-id="9e14c-242">撰写模式</span><span class="sxs-lookup"><span data-stu-id="9e14c-242">Compose mode</span></span>

<span data-ttu-id="9e14c-243">`end` 属性返回 `Time` 对象。</span><span class="sxs-lookup"><span data-stu-id="9e14c-243">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="9e14c-244">使用 [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.2#setasync-datetime--options--callback-) 方法设置结束时间时，应使用 [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) 方法将客户端的本地时间转换为服务器的 UTC。</span><span class="sxs-lookup"><span data-stu-id="9e14c-244">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.2#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

<span data-ttu-id="9e14c-245">以下示例使用 `Time` 对象的 [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.2#setasync-datetime--options--callback-) 方法设置约会的结束时间。</span><span class="sxs-lookup"><span data-stu-id="9e14c-245">The following example sets the end time of an appointment by using the [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.2#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="9e14c-246">类型</span><span class="sxs-lookup"><span data-stu-id="9e14c-246">Type</span></span>

*   <span data-ttu-id="9e14c-247">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="9e14c-247">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.2)</span></span>

##### <a name="requirements"></a><span data-ttu-id="9e14c-248">要求</span><span class="sxs-lookup"><span data-stu-id="9e14c-248">Requirements</span></span>

|<span data-ttu-id="9e14c-249">要求</span><span class="sxs-lookup"><span data-stu-id="9e14c-249">Requirement</span></span>| <span data-ttu-id="9e14c-250">值</span><span class="sxs-lookup"><span data-stu-id="9e14c-250">Value</span></span>|
|---|---|
|[<span data-ttu-id="9e14c-251">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="9e14c-251">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9e14c-252">1.0</span><span class="sxs-lookup"><span data-stu-id="9e14c-252">1.0</span></span>|
|[<span data-ttu-id="9e14c-253">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="9e14c-253">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9e14c-254">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9e14c-254">ReadItem</span></span>|
|[<span data-ttu-id="9e14c-255">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="9e14c-255">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="9e14c-256">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="9e14c-256">Compose or Read</span></span>|

#### <a name="from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-12"></a><span data-ttu-id="9e14c-257">发件人: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="9e14c-257">from: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)</span></span>

<span data-ttu-id="9e14c-p113">获取邮件发件人的电子邮件地址。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="9e14c-p113">Gets the email address of the sender of a message. Read mode only.</span></span>

<span data-ttu-id="9e14c-p114">`from` 和 [`sender`](#sender-emailaddressdetails) 属性表示同一个人，邮件由代理人发送的除外。在此情况下，`from` 属性表示代理程序，而 sender 属性表示代理人。</span><span class="sxs-lookup"><span data-stu-id="9e14c-p114">The `from` and [`sender`](#sender-emailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="9e14c-262">`from` 属性中 `EmailAddressDetails` 对象的 `recipientType` 属性为 `undefined`。</span><span class="sxs-lookup"><span data-stu-id="9e14c-262">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="9e14c-263">类型</span><span class="sxs-lookup"><span data-stu-id="9e14c-263">Type</span></span>

*   [<span data-ttu-id="9e14c-264">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="9e14c-264">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)

##### <a name="requirements"></a><span data-ttu-id="9e14c-265">要求</span><span class="sxs-lookup"><span data-stu-id="9e14c-265">Requirements</span></span>

|<span data-ttu-id="9e14c-266">要求</span><span class="sxs-lookup"><span data-stu-id="9e14c-266">Requirement</span></span>| <span data-ttu-id="9e14c-267">值</span><span class="sxs-lookup"><span data-stu-id="9e14c-267">Value</span></span>|
|---|---|
|[<span data-ttu-id="9e14c-268">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="9e14c-268">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9e14c-269">1.0</span><span class="sxs-lookup"><span data-stu-id="9e14c-269">1.0</span></span>|
|[<span data-ttu-id="9e14c-270">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="9e14c-270">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9e14c-271">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9e14c-271">ReadItem</span></span>|
|[<span data-ttu-id="9e14c-272">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="9e14c-272">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="9e14c-273">阅读</span><span class="sxs-lookup"><span data-stu-id="9e14c-273">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="9e14c-274">示例</span><span class="sxs-lookup"><span data-stu-id="9e14c-274">Example</span></span>

```javascript
var from = Office.context.mailbox.item.from;
console.log("From " + from);
```

#### <a name="internetmessageid-string"></a><span data-ttu-id="9e14c-275">internetMessageId: String</span><span class="sxs-lookup"><span data-stu-id="9e14c-275">internetMessageId: String</span></span>

<span data-ttu-id="9e14c-p115">获取电子邮件的 Internet 消息标识符。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="9e14c-p115">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="9e14c-278">Type</span><span class="sxs-lookup"><span data-stu-id="9e14c-278">Type</span></span>

*   <span data-ttu-id="9e14c-279">String</span><span class="sxs-lookup"><span data-stu-id="9e14c-279">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="9e14c-280">要求</span><span class="sxs-lookup"><span data-stu-id="9e14c-280">Requirements</span></span>

|<span data-ttu-id="9e14c-281">要求</span><span class="sxs-lookup"><span data-stu-id="9e14c-281">Requirement</span></span>| <span data-ttu-id="9e14c-282">值</span><span class="sxs-lookup"><span data-stu-id="9e14c-282">Value</span></span>|
|---|---|
|[<span data-ttu-id="9e14c-283">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="9e14c-283">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9e14c-284">1.0</span><span class="sxs-lookup"><span data-stu-id="9e14c-284">1.0</span></span>|
|[<span data-ttu-id="9e14c-285">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="9e14c-285">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9e14c-286">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9e14c-286">ReadItem</span></span>|
|[<span data-ttu-id="9e14c-287">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="9e14c-287">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="9e14c-288">阅读</span><span class="sxs-lookup"><span data-stu-id="9e14c-288">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="9e14c-289">示例</span><span class="sxs-lookup"><span data-stu-id="9e14c-289">Example</span></span>

```javascript
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

#### <a name="itemclass-string"></a><span data-ttu-id="9e14c-290">itemClass: String</span><span class="sxs-lookup"><span data-stu-id="9e14c-290">itemClass: String</span></span>

<span data-ttu-id="9e14c-p116">获取选定项目的 Exchange Web 服务项目类。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="9e14c-p116">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="9e14c-p117">`itemClass` 属性指定所选项目的邮件类别。以下是邮件或约会项目的默认邮件类别。</span><span class="sxs-lookup"><span data-stu-id="9e14c-p117">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

| <span data-ttu-id="9e14c-295">类型</span><span class="sxs-lookup"><span data-stu-id="9e14c-295">Type</span></span> | <span data-ttu-id="9e14c-296">说明</span><span class="sxs-lookup"><span data-stu-id="9e14c-296">Description</span></span> | <span data-ttu-id="9e14c-297">项目类</span><span class="sxs-lookup"><span data-stu-id="9e14c-297">item class</span></span> |
| --- | --- | --- |
| <span data-ttu-id="9e14c-298">约会项目</span><span class="sxs-lookup"><span data-stu-id="9e14c-298">Appointment items</span></span> | <span data-ttu-id="9e14c-299">这些是项目类 `IPM.Appointment` 或 `IPM.Appointment.Occurrence` 的日历项目。</span><span class="sxs-lookup"><span data-stu-id="9e14c-299">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurrence`.</span></span> | `IPM.Appointment`<br />`IPM.Appointment.Occurrence` |
| <span data-ttu-id="9e14c-300">邮件项目</span><span class="sxs-lookup"><span data-stu-id="9e14c-300">Message items</span></span> | <span data-ttu-id="9e14c-301">这些项目包括具有默认邮件类别 `IPM.Note` 的电子邮件，以及使用 `IPM.Schedule.Meeting` 作为基础邮件类别的会议请求、响应和取消。</span><span class="sxs-lookup"><span data-stu-id="9e14c-301">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span> | `IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled` |

<span data-ttu-id="9e14c-302">你可以创建用于扩展默认邮件类别的自定义邮件类别，例如，自定义约会邮件类别 `IPM.Appointment.Contoso`。</span><span class="sxs-lookup"><span data-stu-id="9e14c-302">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="9e14c-303">类型</span><span class="sxs-lookup"><span data-stu-id="9e14c-303">Type</span></span>

*   <span data-ttu-id="9e14c-304">String</span><span class="sxs-lookup"><span data-stu-id="9e14c-304">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="9e14c-305">要求</span><span class="sxs-lookup"><span data-stu-id="9e14c-305">Requirements</span></span>

|<span data-ttu-id="9e14c-306">要求</span><span class="sxs-lookup"><span data-stu-id="9e14c-306">Requirement</span></span>| <span data-ttu-id="9e14c-307">值</span><span class="sxs-lookup"><span data-stu-id="9e14c-307">Value</span></span>|
|---|---|
|[<span data-ttu-id="9e14c-308">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="9e14c-308">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9e14c-309">1.0</span><span class="sxs-lookup"><span data-stu-id="9e14c-309">1.0</span></span>|
|[<span data-ttu-id="9e14c-310">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="9e14c-310">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9e14c-311">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9e14c-311">ReadItem</span></span>|
|[<span data-ttu-id="9e14c-312">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="9e14c-312">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="9e14c-313">阅读</span><span class="sxs-lookup"><span data-stu-id="9e14c-313">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="9e14c-314">示例</span><span class="sxs-lookup"><span data-stu-id="9e14c-314">Example</span></span>

```javascript
var itemClass = Office.context.mailbox.item.itemClass;
console.log("Item class: " + itemClass);
```

#### <a name="nullable-itemid-string"></a><span data-ttu-id="9e14c-315">(可以为 null) itemId: String</span><span class="sxs-lookup"><span data-stu-id="9e14c-315">(nullable) itemId: String</span></span>

<span data-ttu-id="9e14c-316">获取当前项目的 Exchange Web 服务项目标识符。</span><span class="sxs-lookup"><span data-stu-id="9e14c-316">Gets the Exchange Web Services item identifier for the current item.</span></span> <span data-ttu-id="9e14c-317">仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="9e14c-317">Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="9e14c-318">`itemId` 属性返回的标识符与 Exchange Web 服务项目标识符相同。</span><span class="sxs-lookup"><span data-stu-id="9e14c-318">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="9e14c-319">`itemId` 属性与 Outlook 条目 ID 或 Outlook REST API 使用的 ID 不同。</span><span class="sxs-lookup"><span data-stu-id="9e14c-319">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="9e14c-320">在使用此值进行 REST API 调用之前, 应使用`Office.context.mailbox.convertToRestId`转换它, 这可从要求集1.3 中开始。</span><span class="sxs-lookup"><span data-stu-id="9e14c-320">Before making REST API calls using this value, it should be converted using `Office.context.mailbox.convertToRestId`, which is available starting in requirement set 1.3.</span></span> <span data-ttu-id="9e14c-321">有关详细信息，请参阅[从 Outlook 加载项使用 Outlook REST API](/outlook/add-ins/use-rest-api#get-the-item-id)。</span><span class="sxs-lookup"><span data-stu-id="9e14c-321">For more details, see [Use the Outlook REST APIs from an Outlook add-in](/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

##### <a name="type"></a><span data-ttu-id="9e14c-322">Type</span><span class="sxs-lookup"><span data-stu-id="9e14c-322">Type</span></span>

*   <span data-ttu-id="9e14c-323">String</span><span class="sxs-lookup"><span data-stu-id="9e14c-323">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="9e14c-324">要求</span><span class="sxs-lookup"><span data-stu-id="9e14c-324">Requirements</span></span>

|<span data-ttu-id="9e14c-325">要求</span><span class="sxs-lookup"><span data-stu-id="9e14c-325">Requirement</span></span>| <span data-ttu-id="9e14c-326">值</span><span class="sxs-lookup"><span data-stu-id="9e14c-326">Value</span></span>|
|---|---|
|[<span data-ttu-id="9e14c-327">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="9e14c-327">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9e14c-328">1.0</span><span class="sxs-lookup"><span data-stu-id="9e14c-328">1.0</span></span>|
|[<span data-ttu-id="9e14c-329">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="9e14c-329">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9e14c-330">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9e14c-330">ReadItem</span></span>|
|[<span data-ttu-id="9e14c-331">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="9e14c-331">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="9e14c-332">阅读</span><span class="sxs-lookup"><span data-stu-id="9e14c-332">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="9e14c-333">示例</span><span class="sxs-lookup"><span data-stu-id="9e14c-333">Example</span></span>

<span data-ttu-id="9e14c-p120">以下代码检查项目标识符是否存在。如果 `itemId` 属性返回 `null` 或 `undefined`，则将项目保存到存储，并从异步结果中获取项目标识符。</span><span class="sxs-lookup"><span data-stu-id="9e14c-p120">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

```javascript
var itemId = Office.context.mailbox.item.itemId;
if (itemId === null || itemId == undefined) {
  Office.context.mailbox.item.saveAsync(function(result) {
    itemId = result.value;
  });
}
```

#### <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtypeviewoutlook-js-12"></a><span data-ttu-id="9e14c-336">itemType: [MailboxEnums](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="9e14c-336">itemType: [Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.2)</span></span>

<span data-ttu-id="9e14c-337">获取实例表示的项的类型。</span><span class="sxs-lookup"><span data-stu-id="9e14c-337">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="9e14c-338">`itemType` 属性返回其中一个 `ItemType` 枚举值，指示 `item` 对象实例是邮件还是约会。</span><span class="sxs-lookup"><span data-stu-id="9e14c-338">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="9e14c-339">类型</span><span class="sxs-lookup"><span data-stu-id="9e14c-339">Type</span></span>

*   [<span data-ttu-id="9e14c-340">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="9e14c-340">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.2)

##### <a name="requirements"></a><span data-ttu-id="9e14c-341">要求</span><span class="sxs-lookup"><span data-stu-id="9e14c-341">Requirements</span></span>

|<span data-ttu-id="9e14c-342">要求</span><span class="sxs-lookup"><span data-stu-id="9e14c-342">Requirement</span></span>| <span data-ttu-id="9e14c-343">值</span><span class="sxs-lookup"><span data-stu-id="9e14c-343">Value</span></span>|
|---|---|
|[<span data-ttu-id="9e14c-344">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="9e14c-344">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9e14c-345">1.0</span><span class="sxs-lookup"><span data-stu-id="9e14c-345">1.0</span></span>|
|[<span data-ttu-id="9e14c-346">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="9e14c-346">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9e14c-347">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9e14c-347">ReadItem</span></span>|
|[<span data-ttu-id="9e14c-348">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="9e14c-348">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="9e14c-349">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="9e14c-349">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="9e14c-350">示例</span><span class="sxs-lookup"><span data-stu-id="9e14c-350">Example</span></span>

```javascript
if (Office.context.mailbox.item.itemType === Office.MailboxEnums.ItemType.Message) {
  // Do something.
} else {
  // Do something else.
}
```

#### <a name="location-stringlocationjavascriptapioutlookofficelocationviewoutlook-js-12"></a><span data-ttu-id="9e14c-351">位置: 字符串 |[位置](/javascript/api/outlook/office.location?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="9e14c-351">location: String|[Location](/javascript/api/outlook/office.location?view=outlook-js-1.2)</span></span>

<span data-ttu-id="9e14c-352">获取或设置约会的位置。</span><span class="sxs-lookup"><span data-stu-id="9e14c-352">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="9e14c-353">阅读模式</span><span class="sxs-lookup"><span data-stu-id="9e14c-353">Read mode</span></span>

<span data-ttu-id="9e14c-354">`location` 属性返回一个包含约会位置的字符串。</span><span class="sxs-lookup"><span data-stu-id="9e14c-354">The `location` property returns a string that contains the location of the appointment.</span></span>

```javascript
var location = Office.context.mailbox.item.location;
console.log("location: " + location);
```

##### <a name="compose-mode"></a><span data-ttu-id="9e14c-355">撰写模式</span><span class="sxs-lookup"><span data-stu-id="9e14c-355">Compose mode</span></span>

<span data-ttu-id="9e14c-356">`location` 属性返回一个 `Location` 对象，该对象提供用于获取和设置约会位置的方法。</span><span class="sxs-lookup"><span data-stu-id="9e14c-356">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

```javascript
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="9e14c-357">类型</span><span class="sxs-lookup"><span data-stu-id="9e14c-357">Type</span></span>

*   <span data-ttu-id="9e14c-358">String | [Location](/javascript/api/outlook/office.location?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="9e14c-358">String | [Location](/javascript/api/outlook/office.location?view=outlook-js-1.2)</span></span>

##### <a name="requirements"></a><span data-ttu-id="9e14c-359">要求</span><span class="sxs-lookup"><span data-stu-id="9e14c-359">Requirements</span></span>

|<span data-ttu-id="9e14c-360">要求</span><span class="sxs-lookup"><span data-stu-id="9e14c-360">Requirement</span></span>| <span data-ttu-id="9e14c-361">值</span><span class="sxs-lookup"><span data-stu-id="9e14c-361">Value</span></span>|
|---|---|
|[<span data-ttu-id="9e14c-362">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="9e14c-362">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9e14c-363">1.0</span><span class="sxs-lookup"><span data-stu-id="9e14c-363">1.0</span></span>|
|[<span data-ttu-id="9e14c-364">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="9e14c-364">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9e14c-365">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9e14c-365">ReadItem</span></span>|
|[<span data-ttu-id="9e14c-366">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="9e14c-366">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="9e14c-367">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="9e14c-367">Compose or Read</span></span>|

#### <a name="normalizedsubject-string"></a><span data-ttu-id="9e14c-368">normalizedSubject: String</span><span class="sxs-lookup"><span data-stu-id="9e14c-368">normalizedSubject: String</span></span>

<span data-ttu-id="9e14c-p121">获取删除了所有前缀（包括 `RE:` 和 `FWD:`）的项目主题。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="9e14c-p121">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="9e14c-p122">normalizedSubject 属性获取包含由电子邮件程序添加的任何标准前缀（如 `RE:` 和 `FW:`）的项目主题。若要获取包含完整前缀的项目主题，请使用 [`subject`](#subject-stringsubject) 属性。</span><span class="sxs-lookup"><span data-stu-id="9e14c-p122">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="9e14c-373">Type</span><span class="sxs-lookup"><span data-stu-id="9e14c-373">Type</span></span>

*   <span data-ttu-id="9e14c-374">String</span><span class="sxs-lookup"><span data-stu-id="9e14c-374">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="9e14c-375">要求</span><span class="sxs-lookup"><span data-stu-id="9e14c-375">Requirements</span></span>

|<span data-ttu-id="9e14c-376">要求</span><span class="sxs-lookup"><span data-stu-id="9e14c-376">Requirement</span></span>| <span data-ttu-id="9e14c-377">值</span><span class="sxs-lookup"><span data-stu-id="9e14c-377">Value</span></span>|
|---|---|
|[<span data-ttu-id="9e14c-378">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="9e14c-378">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9e14c-379">1.0</span><span class="sxs-lookup"><span data-stu-id="9e14c-379">1.0</span></span>|
|[<span data-ttu-id="9e14c-380">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="9e14c-380">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9e14c-381">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9e14c-381">ReadItem</span></span>|
|[<span data-ttu-id="9e14c-382">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="9e14c-382">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="9e14c-383">阅读</span><span class="sxs-lookup"><span data-stu-id="9e14c-383">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="9e14c-384">示例</span><span class="sxs-lookup"><span data-stu-id="9e14c-384">Example</span></span>

```javascript
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
console.log("Normalized subject: " + normalizedSubject);
```

#### <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipientsviewoutlook-js-12"></a><span data-ttu-id="9e14c-385">optionalAttendees: <[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[收件人](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)的数组</span><span class="sxs-lookup"><span data-stu-id="9e14c-385">optionalAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span></span>

<span data-ttu-id="9e14c-386">提供对事件的可选与会者的访问权限。</span><span class="sxs-lookup"><span data-stu-id="9e14c-386">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="9e14c-387">对象的类型和访问级别取决于当前项目的模式。</span><span class="sxs-lookup"><span data-stu-id="9e14c-387">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="9e14c-388">阅读模式</span><span class="sxs-lookup"><span data-stu-id="9e14c-388">Read mode</span></span>

<span data-ttu-id="9e14c-389">`optionalAttendees` 属性返回一个数组，其中包含每个可选与会者的 `EmailAddressDetails` 对象。</span><span class="sxs-lookup"><span data-stu-id="9e14c-389">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span>

```javascript
var optionalAttendees = Office.context.mailbox.item.optionalAttendees;
console.log("Optional attendees: " + JSON.stringify(optionalAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="9e14c-390">撰写模式</span><span class="sxs-lookup"><span data-stu-id="9e14c-390">Compose mode</span></span>

<span data-ttu-id="9e14c-391">`optionalAttendees` 属性返回一个 `Recipients` 对象，该对象提供用于获取或更新可选与会者的方法。</span><span class="sxs-lookup"><span data-stu-id="9e14c-391">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span>

```javascript
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="9e14c-392">类型</span><span class="sxs-lookup"><span data-stu-id="9e14c-392">Type</span></span>

*   <span data-ttu-id="9e14c-393">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="9e14c-393">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span></span>

##### <a name="requirements"></a><span data-ttu-id="9e14c-394">要求</span><span class="sxs-lookup"><span data-stu-id="9e14c-394">Requirements</span></span>

|<span data-ttu-id="9e14c-395">要求</span><span class="sxs-lookup"><span data-stu-id="9e14c-395">Requirement</span></span>| <span data-ttu-id="9e14c-396">值</span><span class="sxs-lookup"><span data-stu-id="9e14c-396">Value</span></span>|
|---|---|
|[<span data-ttu-id="9e14c-397">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="9e14c-397">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9e14c-398">1.0</span><span class="sxs-lookup"><span data-stu-id="9e14c-398">1.0</span></span>|
|[<span data-ttu-id="9e14c-399">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="9e14c-399">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9e14c-400">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9e14c-400">ReadItem</span></span>|
|[<span data-ttu-id="9e14c-401">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="9e14c-401">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="9e14c-402">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="9e14c-402">Compose or Read</span></span>|

#### <a name="organizer-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-12"></a><span data-ttu-id="9e14c-403">组织者: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="9e14c-403">organizer: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)</span></span>

<span data-ttu-id="9e14c-p124">获取指定会议的会议组织者的电子邮件地址。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="9e14c-p124">Gets the email address of the meeting organizer for a specified meeting. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="9e14c-406">类型</span><span class="sxs-lookup"><span data-stu-id="9e14c-406">Type</span></span>

*   [<span data-ttu-id="9e14c-407">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="9e14c-407">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)

##### <a name="requirements"></a><span data-ttu-id="9e14c-408">要求</span><span class="sxs-lookup"><span data-stu-id="9e14c-408">Requirements</span></span>

|<span data-ttu-id="9e14c-409">要求</span><span class="sxs-lookup"><span data-stu-id="9e14c-409">Requirement</span></span>| <span data-ttu-id="9e14c-410">值</span><span class="sxs-lookup"><span data-stu-id="9e14c-410">Value</span></span>|
|---|---|
|[<span data-ttu-id="9e14c-411">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="9e14c-411">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9e14c-412">1.0</span><span class="sxs-lookup"><span data-stu-id="9e14c-412">1.0</span></span>|
|[<span data-ttu-id="9e14c-413">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="9e14c-413">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9e14c-414">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9e14c-414">ReadItem</span></span>|
|[<span data-ttu-id="9e14c-415">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="9e14c-415">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="9e14c-416">阅读</span><span class="sxs-lookup"><span data-stu-id="9e14c-416">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="9e14c-417">示例</span><span class="sxs-lookup"><span data-stu-id="9e14c-417">Example</span></span>

```javascript
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
console.log("Organizer: " + organizerName + " (" + organizerAddress + ")");
```

#### <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipientsviewoutlook-js-12"></a><span data-ttu-id="9e14c-418">requiredAttendees: <[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[收件人](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)的数组</span><span class="sxs-lookup"><span data-stu-id="9e14c-418">requiredAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span></span>

<span data-ttu-id="9e14c-419">提供对事件的必需与会者的访问权限。</span><span class="sxs-lookup"><span data-stu-id="9e14c-419">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="9e14c-420">对象的类型和访问级别取决于当前项目的模式。</span><span class="sxs-lookup"><span data-stu-id="9e14c-420">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="9e14c-421">阅读模式</span><span class="sxs-lookup"><span data-stu-id="9e14c-421">Read mode</span></span>

<span data-ttu-id="9e14c-422">`requiredAttendees` 属性返回一个数组，其中包含每个必需与会者的 `EmailAddressDetails` 对象。</span><span class="sxs-lookup"><span data-stu-id="9e14c-422">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span>

```javascript
var requiredAttendees = Office.context.mailbox.item.requiredAttendees;
console.log("Required attendees: " + JSON.stringify(requiredAttendees));
```

##### <a name="compose-mode"></a><span data-ttu-id="9e14c-423">撰写模式</span><span class="sxs-lookup"><span data-stu-id="9e14c-423">Compose mode</span></span>

<span data-ttu-id="9e14c-424">`requiredAttendees` 属性返回一个 `Recipients` 对象，该对象提供用于获取或更新必需与会者的方法。</span><span class="sxs-lookup"><span data-stu-id="9e14c-424">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span>

```javascript
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
  console.log(JSON.stringify(arrayOfRequiredAttendeesRecipients));
}
```

##### <a name="type"></a><span data-ttu-id="9e14c-425">类型</span><span class="sxs-lookup"><span data-stu-id="9e14c-425">Type</span></span>

*   <span data-ttu-id="9e14c-426">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="9e14c-426">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span></span>

##### <a name="requirements"></a><span data-ttu-id="9e14c-427">要求</span><span class="sxs-lookup"><span data-stu-id="9e14c-427">Requirements</span></span>

|<span data-ttu-id="9e14c-428">要求</span><span class="sxs-lookup"><span data-stu-id="9e14c-428">Requirement</span></span>| <span data-ttu-id="9e14c-429">值</span><span class="sxs-lookup"><span data-stu-id="9e14c-429">Value</span></span>|
|---|---|
|[<span data-ttu-id="9e14c-430">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="9e14c-430">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9e14c-431">1.0</span><span class="sxs-lookup"><span data-stu-id="9e14c-431">1.0</span></span>|
|[<span data-ttu-id="9e14c-432">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="9e14c-432">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9e14c-433">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9e14c-433">ReadItem</span></span>|
|[<span data-ttu-id="9e14c-434">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="9e14c-434">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="9e14c-435">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="9e14c-435">Compose or Read</span></span>|

#### <a name="sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsviewoutlook-js-12"></a><span data-ttu-id="9e14c-436">发件人: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="9e14c-436">sender: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)</span></span>

<span data-ttu-id="9e14c-p126">获取电子邮件发件人的电子邮件地址。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="9e14c-p126">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="9e14c-p127">[`from`](#from-emailaddressdetails) 和 `sender` 属性表示同一个人，邮件由代理人发送的除外。在此情况下，`from` 属性表示代理程序，而 sender 属性表示代理人。</span><span class="sxs-lookup"><span data-stu-id="9e14c-p127">The [`from`](#from-emailaddressdetails) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="9e14c-441">`sender` 属性中 `EmailAddressDetails` 对象的 `recipientType` 属性为 `undefined`。</span><span class="sxs-lookup"><span data-stu-id="9e14c-441">The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="9e14c-442">类型</span><span class="sxs-lookup"><span data-stu-id="9e14c-442">Type</span></span>

*   [<span data-ttu-id="9e14c-443">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="9e14c-443">EmailAddressDetails</span></span>](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)

##### <a name="requirements"></a><span data-ttu-id="9e14c-444">要求</span><span class="sxs-lookup"><span data-stu-id="9e14c-444">Requirements</span></span>

|<span data-ttu-id="9e14c-445">要求</span><span class="sxs-lookup"><span data-stu-id="9e14c-445">Requirement</span></span>| <span data-ttu-id="9e14c-446">值</span><span class="sxs-lookup"><span data-stu-id="9e14c-446">Value</span></span>|
|---|---|
|[<span data-ttu-id="9e14c-447">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="9e14c-447">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9e14c-448">1.0</span><span class="sxs-lookup"><span data-stu-id="9e14c-448">1.0</span></span>|
|[<span data-ttu-id="9e14c-449">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="9e14c-449">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9e14c-450">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9e14c-450">ReadItem</span></span>|
|[<span data-ttu-id="9e14c-451">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="9e14c-451">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="9e14c-452">阅读</span><span class="sxs-lookup"><span data-stu-id="9e14c-452">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="9e14c-453">示例</span><span class="sxs-lookup"><span data-stu-id="9e14c-453">Example</span></span>

```javascript
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
console.log("Sender: " + senderName + " (" + senderAddress + ")");
```

#### <a name="start-datetimejavascriptapioutlookofficetimeviewoutlook-js-12"></a><span data-ttu-id="9e14c-454">开始日期: 日期 |[时间](/javascript/api/outlook/office.time?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="9e14c-454">start: Date|[Time](/javascript/api/outlook/office.time?view=outlook-js-1.2)</span></span>

<span data-ttu-id="9e14c-455">获取或设置约会开始的日期和时间。</span><span class="sxs-lookup"><span data-stu-id="9e14c-455">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="9e14c-p128">将 `start` 属性表示为协调世界时 (UTC) 的日期和时间值。可使用 [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) 方法将该值转换为客户端的本地日期和时间。</span><span class="sxs-lookup"><span data-stu-id="9e14c-p128">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="9e14c-458">阅读模式</span><span class="sxs-lookup"><span data-stu-id="9e14c-458">Read mode</span></span>

<span data-ttu-id="9e14c-459">`start` 属性返回 `Date` 对象。</span><span class="sxs-lookup"><span data-stu-id="9e14c-459">The `start` property returns a `Date` object.</span></span>

```javascript
var start = Office.context.mailbox.item.start;
console.log("Appointment start: " + JSON.stringify(start));
```

##### <a name="compose-mode"></a><span data-ttu-id="9e14c-460">撰写模式</span><span class="sxs-lookup"><span data-stu-id="9e14c-460">Compose mode</span></span>

<span data-ttu-id="9e14c-461">`start` 属性返回 `Time` 对象。</span><span class="sxs-lookup"><span data-stu-id="9e14c-461">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="9e14c-462">使用 [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.2#setasync-datetime--options--callback-) 方法设置开始时间时，应使用 [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) 方法将客户端的本地时间转换为服务器的 UTC。</span><span class="sxs-lookup"><span data-stu-id="9e14c-462">When you use the [`Time.setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.2#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>
<span data-ttu-id="9e14c-463">以下示例通过使用 `Time` 对象的 [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.2#setasync-datetime--options--callback-) 方法，设置撰写模式下约会的开始时间。</span><span class="sxs-lookup"><span data-stu-id="9e14c-463">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook/office.time?view=outlook-js-1.2#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

##### <a name="type"></a><span data-ttu-id="9e14c-464">类型</span><span class="sxs-lookup"><span data-stu-id="9e14c-464">Type</span></span>

*   <span data-ttu-id="9e14c-465">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="9e14c-465">Date | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.2)</span></span>

##### <a name="requirements"></a><span data-ttu-id="9e14c-466">要求</span><span class="sxs-lookup"><span data-stu-id="9e14c-466">Requirements</span></span>

|<span data-ttu-id="9e14c-467">要求</span><span class="sxs-lookup"><span data-stu-id="9e14c-467">Requirement</span></span>| <span data-ttu-id="9e14c-468">值</span><span class="sxs-lookup"><span data-stu-id="9e14c-468">Value</span></span>|
|---|---|
|[<span data-ttu-id="9e14c-469">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="9e14c-469">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9e14c-470">1.0</span><span class="sxs-lookup"><span data-stu-id="9e14c-470">1.0</span></span>|
|[<span data-ttu-id="9e14c-471">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="9e14c-471">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9e14c-472">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9e14c-472">ReadItem</span></span>|
|[<span data-ttu-id="9e14c-473">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="9e14c-473">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="9e14c-474">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="9e14c-474">Compose or Read</span></span>|

#### <a name="subject-stringsubjectjavascriptapioutlookofficesubjectviewoutlook-js-12"></a><span data-ttu-id="9e14c-475">subject: String |[主题](/javascript/api/outlook/office.subject?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="9e14c-475">subject: String|[Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.2)</span></span>

<span data-ttu-id="9e14c-476">获取或设置显示在项目的主题字段中的说明。</span><span class="sxs-lookup"><span data-stu-id="9e14c-476">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="9e14c-477">`subject` 属性获取或设置由电子邮件服务器发送项目时的整个主题。</span><span class="sxs-lookup"><span data-stu-id="9e14c-477">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="9e14c-478">阅读模式</span><span class="sxs-lookup"><span data-stu-id="9e14c-478">Read mode</span></span>

<span data-ttu-id="9e14c-p130">`subject` 属性返回一个字符串。使用 [`normalizedSubject`](#normalizedsubject-string) 属性获取不带任何前导前缀（如 `RE:` 和 `FW:`）的主题。</span><span class="sxs-lookup"><span data-stu-id="9e14c-p130">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

```javascript
var subject = Office.context.mailbox.item.subject;
console.log(subject);
```

##### <a name="compose-mode"></a><span data-ttu-id="9e14c-481">撰写模式</span><span class="sxs-lookup"><span data-stu-id="9e14c-481">Compose mode</span></span>

<span data-ttu-id="9e14c-482">`subject` 属性返回一个 `Subject` 对象，该对象提供用于获取和设置主题的方法。</span><span class="sxs-lookup"><span data-stu-id="9e14c-482">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```javascript
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
  console.log(subject);
}
```

##### <a name="type"></a><span data-ttu-id="9e14c-483">类型</span><span class="sxs-lookup"><span data-stu-id="9e14c-483">Type</span></span>

*   <span data-ttu-id="9e14c-484">String | [Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="9e14c-484">String | [Subject](/javascript/api/outlook/office.subject?view=outlook-js-1.2)</span></span>

##### <a name="requirements"></a><span data-ttu-id="9e14c-485">要求</span><span class="sxs-lookup"><span data-stu-id="9e14c-485">Requirements</span></span>

|<span data-ttu-id="9e14c-486">要求</span><span class="sxs-lookup"><span data-stu-id="9e14c-486">Requirement</span></span>| <span data-ttu-id="9e14c-487">值</span><span class="sxs-lookup"><span data-stu-id="9e14c-487">Value</span></span>|
|---|---|
|[<span data-ttu-id="9e14c-488">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="9e14c-488">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9e14c-489">1.0</span><span class="sxs-lookup"><span data-stu-id="9e14c-489">1.0</span></span>|
|[<span data-ttu-id="9e14c-490">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="9e14c-490">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9e14c-491">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9e14c-491">ReadItem</span></span>|
|[<span data-ttu-id="9e14c-492">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="9e14c-492">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="9e14c-493">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="9e14c-493">Compose or Read</span></span>|

#### <a name="to-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipientsviewoutlook-js-12"></a><span data-ttu-id="9e14c-494">to: <[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[收件人](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)的数组</span><span class="sxs-lookup"><span data-stu-id="9e14c-494">to: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span></span>

<span data-ttu-id="9e14c-495">提供对邮件的“**收件人**”行上的收件人的访问权限。</span><span class="sxs-lookup"><span data-stu-id="9e14c-495">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="9e14c-496">对象的类型和访问级别取决于当前项目的模式。</span><span class="sxs-lookup"><span data-stu-id="9e14c-496">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="9e14c-497">阅读模式</span><span class="sxs-lookup"><span data-stu-id="9e14c-497">Read mode</span></span>

<span data-ttu-id="9e14c-p132">`to` 属性返回包含邮件的**收件人**行上所列的每个收件人的 `EmailAddressDetails` 对象的数组。集合上限为 100 个成员。</span><span class="sxs-lookup"><span data-stu-id="9e14c-p132">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.</span></span>

```javascript
console.log(JSON.stringify(Office.context.mailbox.item.to));
```

##### <a name="compose-mode"></a><span data-ttu-id="9e14c-500">撰写模式</span><span class="sxs-lookup"><span data-stu-id="9e14c-500">Compose mode</span></span>

<span data-ttu-id="9e14c-501">`to` 属性返回一个 `Recipients` 对象，该对象提供用于获取或更新邮件的“**收件人**”行上收件人的方法。</span><span class="sxs-lookup"><span data-stu-id="9e14c-501">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span>

```javascript
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="9e14c-502">类型</span><span class="sxs-lookup"><span data-stu-id="9e14c-502">Type</span></span>

*   <span data-ttu-id="9e14c-503">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="9e14c-503">Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-1.2)</span></span>

##### <a name="requirements"></a><span data-ttu-id="9e14c-504">要求</span><span class="sxs-lookup"><span data-stu-id="9e14c-504">Requirements</span></span>

|<span data-ttu-id="9e14c-505">要求</span><span class="sxs-lookup"><span data-stu-id="9e14c-505">Requirement</span></span>| <span data-ttu-id="9e14c-506">值</span><span class="sxs-lookup"><span data-stu-id="9e14c-506">Value</span></span>|
|---|---|
|[<span data-ttu-id="9e14c-507">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="9e14c-507">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9e14c-508">1.0</span><span class="sxs-lookup"><span data-stu-id="9e14c-508">1.0</span></span>|
|[<span data-ttu-id="9e14c-509">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="9e14c-509">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9e14c-510">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9e14c-510">ReadItem</span></span>|
|[<span data-ttu-id="9e14c-511">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="9e14c-511">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="9e14c-512">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="9e14c-512">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="9e14c-513">方法</span><span class="sxs-lookup"><span data-stu-id="9e14c-513">Methods</span></span>

#### <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="9e14c-514">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="9e14c-514">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="9e14c-515">将文件作为附件添加到邮件或约会。</span><span class="sxs-lookup"><span data-stu-id="9e14c-515">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="9e14c-516">`addFileAttachmentAsync` 方法在指定的 URI 上载文件并将其附加到撰写窗体中的项目。</span><span class="sxs-lookup"><span data-stu-id="9e14c-516">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="9e14c-517">随后可以将该标识符与 [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) 方法一同使用，以删除同一个会话中的附件。</span><span class="sxs-lookup"><span data-stu-id="9e14c-517">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="9e14c-518">参数</span><span class="sxs-lookup"><span data-stu-id="9e14c-518">Parameters</span></span>

|<span data-ttu-id="9e14c-519">名称</span><span class="sxs-lookup"><span data-stu-id="9e14c-519">Name</span></span>| <span data-ttu-id="9e14c-520">类型</span><span class="sxs-lookup"><span data-stu-id="9e14c-520">Type</span></span>| <span data-ttu-id="9e14c-521">属性</span><span class="sxs-lookup"><span data-stu-id="9e14c-521">Attributes</span></span>| <span data-ttu-id="9e14c-522">说明</span><span class="sxs-lookup"><span data-stu-id="9e14c-522">Description</span></span>|
|---|---|---|---|
|`uri`| <span data-ttu-id="9e14c-523">String</span><span class="sxs-lookup"><span data-stu-id="9e14c-523">String</span></span>||<span data-ttu-id="9e14c-p133">提供附加到邮件或约会的文件的位置的 URI。最大长度为 2048 个字符。</span><span class="sxs-lookup"><span data-stu-id="9e14c-p133">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="9e14c-526">字符串</span><span class="sxs-lookup"><span data-stu-id="9e14c-526">String</span></span>||<span data-ttu-id="9e14c-p134">在附件上载过程中显示的附件名称。最大长度为 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="9e14c-p134">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="9e14c-529">Object</span><span class="sxs-lookup"><span data-stu-id="9e14c-529">Object</span></span>| <span data-ttu-id="9e14c-530">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="9e14c-530">&lt;optional&gt;</span></span>|<span data-ttu-id="9e14c-531">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="9e14c-531">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="9e14c-532">对象</span><span class="sxs-lookup"><span data-stu-id="9e14c-532">Object</span></span>| <span data-ttu-id="9e14c-533">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="9e14c-533">&lt;optional&gt;</span></span>|<span data-ttu-id="9e14c-534">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="9e14c-534">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="9e14c-535">函数</span><span class="sxs-lookup"><span data-stu-id="9e14c-535">function</span></span>| <span data-ttu-id="9e14c-536">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="9e14c-536">&lt;optional&gt;</span></span>|<span data-ttu-id="9e14c-537">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="9e14c-537">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="9e14c-538">如果成功，附件标识符将在 `asyncResult.value` 属性中提供。</span><span class="sxs-lookup"><span data-stu-id="9e14c-538">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="9e14c-539">如果上传附件失败，`asyncResult` 对象将包含一个提供错误说明的 `Error` 对象。</span><span class="sxs-lookup"><span data-stu-id="9e14c-539">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="9e14c-540">错误</span><span class="sxs-lookup"><span data-stu-id="9e14c-540">Errors</span></span>

| <span data-ttu-id="9e14c-541">错误代码</span><span class="sxs-lookup"><span data-stu-id="9e14c-541">Error code</span></span> | <span data-ttu-id="9e14c-542">说明</span><span class="sxs-lookup"><span data-stu-id="9e14c-542">Description</span></span> |
|------------|-------------|
| `AttachmentSizeExceeded` | <span data-ttu-id="9e14c-543">附件大小超过了允许的大小。</span><span class="sxs-lookup"><span data-stu-id="9e14c-543">The attachment is larger than allowed.</span></span> |
| `FileTypeNotSupported` | <span data-ttu-id="9e14c-544">该附件的扩展名不是允许的扩展名。</span><span class="sxs-lookup"><span data-stu-id="9e14c-544">The attachment has an extension that is not allowed.</span></span> |
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="9e14c-545">邮件或约会具有的附件过多。</span><span class="sxs-lookup"><span data-stu-id="9e14c-545">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="9e14c-546">要求</span><span class="sxs-lookup"><span data-stu-id="9e14c-546">Requirements</span></span>

|<span data-ttu-id="9e14c-547">要求</span><span class="sxs-lookup"><span data-stu-id="9e14c-547">Requirement</span></span>| <span data-ttu-id="9e14c-548">值</span><span class="sxs-lookup"><span data-stu-id="9e14c-548">Value</span></span>|
|---|---|
|[<span data-ttu-id="9e14c-549">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="9e14c-549">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9e14c-550">1.1</span><span class="sxs-lookup"><span data-stu-id="9e14c-550">1.1</span></span>|
|[<span data-ttu-id="9e14c-551">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="9e14c-551">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9e14c-552">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="9e14c-552">ReadWriteItem</span></span>|
|[<span data-ttu-id="9e14c-553">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="9e14c-553">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="9e14c-554">撰写</span><span class="sxs-lookup"><span data-stu-id="9e14c-554">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="9e14c-555">示例</span><span class="sxs-lookup"><span data-stu-id="9e14c-555">Example</span></span>

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

#### <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="9e14c-556">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="9e14c-556">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="9e14c-557">将 Exchange 项目（如邮件）作为附件添加到邮件或约会。</span><span class="sxs-lookup"><span data-stu-id="9e14c-557">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="9e14c-p135">`addItemAttachmentAsync` 方法将包含指定 Exchange 标识符的项目附加到撰写窗体中的项目。如果指定一个回调方法，此方法使用 `asyncResult` 参数调用，该参数包含一个附件标识符或代码，指示附加项目过程中出现的任何错误。可以使用 `options` 参数将状态信息传递给回调方法（如果需要）。</span><span class="sxs-lookup"><span data-stu-id="9e14c-p135">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="9e14c-561">随后可以将该标识符与 [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) 方法一同使用，以删除同一个会话中的附件。</span><span class="sxs-lookup"><span data-stu-id="9e14c-561">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="9e14c-562">如果 Office 外接程序在 web 上的 Outlook 中运行, 则该`addItemAttachmentAsync`方法可以将项目附加到您正在编辑的项目之外的项目中;但是, 不支持这种情况, 建议不要这样做。</span><span class="sxs-lookup"><span data-stu-id="9e14c-562">If your Office Add-in is running in Outlook on the web, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="9e14c-563">参数</span><span class="sxs-lookup"><span data-stu-id="9e14c-563">Parameters</span></span>

|<span data-ttu-id="9e14c-564">名称</span><span class="sxs-lookup"><span data-stu-id="9e14c-564">Name</span></span>| <span data-ttu-id="9e14c-565">类型</span><span class="sxs-lookup"><span data-stu-id="9e14c-565">Type</span></span>| <span data-ttu-id="9e14c-566">属性</span><span class="sxs-lookup"><span data-stu-id="9e14c-566">Attributes</span></span>| <span data-ttu-id="9e14c-567">说明</span><span class="sxs-lookup"><span data-stu-id="9e14c-567">Description</span></span>|
|---|---|---|---|
|`itemId`| <span data-ttu-id="9e14c-568">String</span><span class="sxs-lookup"><span data-stu-id="9e14c-568">String</span></span>||<span data-ttu-id="9e14c-p136">要附加的项目的 Exchange 标识符。最大长度为 100 个字符。</span><span class="sxs-lookup"><span data-stu-id="9e14c-p136">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="9e14c-571">String</span><span class="sxs-lookup"><span data-stu-id="9e14c-571">String</span></span>||<span data-ttu-id="9e14c-572">要附加的项目的主题。</span><span class="sxs-lookup"><span data-stu-id="9e14c-572">The subject of the item to be attached.</span></span> <span data-ttu-id="9e14c-573">最大长度为 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="9e14c-573">The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="9e14c-574">对象</span><span class="sxs-lookup"><span data-stu-id="9e14c-574">Object</span></span>| <span data-ttu-id="9e14c-575">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="9e14c-575">&lt;optional&gt;</span></span>|<span data-ttu-id="9e14c-576">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="9e14c-576">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="9e14c-577">对象</span><span class="sxs-lookup"><span data-stu-id="9e14c-577">Object</span></span>| <span data-ttu-id="9e14c-578">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="9e14c-578">&lt;optional&gt;</span></span>|<span data-ttu-id="9e14c-579">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="9e14c-579">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="9e14c-580">函数</span><span class="sxs-lookup"><span data-stu-id="9e14c-580">function</span></span>| <span data-ttu-id="9e14c-581">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="9e14c-581">&lt;optional&gt;</span></span>|<span data-ttu-id="9e14c-582">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="9e14c-582">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="9e14c-583">如果成功，附件标识符将在 `asyncResult.value` 属性中提供。</span><span class="sxs-lookup"><span data-stu-id="9e14c-583">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="9e14c-584">如果添加附件失败，`asyncResult` 对象将包含一个提供错误说明的 `Error` 对象。</span><span class="sxs-lookup"><span data-stu-id="9e14c-584">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="9e14c-585">错误</span><span class="sxs-lookup"><span data-stu-id="9e14c-585">Errors</span></span>

| <span data-ttu-id="9e14c-586">错误代码</span><span class="sxs-lookup"><span data-stu-id="9e14c-586">Error code</span></span> | <span data-ttu-id="9e14c-587">说明</span><span class="sxs-lookup"><span data-stu-id="9e14c-587">Description</span></span> |
|------------|-------------|
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="9e14c-588">邮件或约会具有的附件过多。</span><span class="sxs-lookup"><span data-stu-id="9e14c-588">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="9e14c-589">要求</span><span class="sxs-lookup"><span data-stu-id="9e14c-589">Requirements</span></span>

|<span data-ttu-id="9e14c-590">要求</span><span class="sxs-lookup"><span data-stu-id="9e14c-590">Requirement</span></span>| <span data-ttu-id="9e14c-591">值</span><span class="sxs-lookup"><span data-stu-id="9e14c-591">Value</span></span>|
|---|---|
|[<span data-ttu-id="9e14c-592">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="9e14c-592">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9e14c-593">1.1</span><span class="sxs-lookup"><span data-stu-id="9e14c-593">1.1</span></span>|
|[<span data-ttu-id="9e14c-594">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="9e14c-594">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9e14c-595">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="9e14c-595">ReadWriteItem</span></span>|
|[<span data-ttu-id="9e14c-596">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="9e14c-596">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="9e14c-597">撰写</span><span class="sxs-lookup"><span data-stu-id="9e14c-597">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="9e14c-598">示例</span><span class="sxs-lookup"><span data-stu-id="9e14c-598">Example</span></span>

<span data-ttu-id="9e14c-599">以下示例将现有的 Outlook 项目添加为名为 `My Attachment` 的附件。</span><span class="sxs-lookup"><span data-stu-id="9e14c-599">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

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

#### <a name="displayreplyallformformdata-callback"></a><span data-ttu-id="9e14c-600">displayReplyAllForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="9e14c-600">displayReplyAllForm(formData, [callback])</span></span>

<span data-ttu-id="9e14c-601">显示答复窗体，其中包括所选邮件的发件人和所有收件人或所选约会的组织者和所有与会者。</span><span class="sxs-lookup"><span data-stu-id="9e14c-601">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="9e14c-602">IOS 或 Android 上的 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="9e14c-602">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="9e14c-603">在 web 上的 Outlook 中, 答复窗体显示为3列视图中的弹出窗体和2列或1列视图中的弹出式窗体。</span><span class="sxs-lookup"><span data-stu-id="9e14c-603">In Outlook on the web, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="9e14c-604">如果任意字符串参数超出其限制，`displayReplyAllForm` 将引发异常。</span><span class="sxs-lookup"><span data-stu-id="9e14c-604">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="9e14c-605">如果在`formData.attachments`参数中指定了附件, 则 web 上的 Outlook 和桌面客户端将尝试下载所有附件并将其附加到答复窗体。</span><span class="sxs-lookup"><span data-stu-id="9e14c-605">When attachments are specified in the `formData.attachments` parameter, Outlook on the web and desktop clients attempt to download all attachments and attach them to the reply form.</span></span> <span data-ttu-id="9e14c-606">如果无法添加任何附件，则在窗体 UI 中显示错误。</span><span class="sxs-lookup"><span data-stu-id="9e14c-606">If any attachments fail to be added, an error is shown in the form UI.</span></span> <span data-ttu-id="9e14c-607">如果这不可能，则不引发错误消息。</span><span class="sxs-lookup"><span data-stu-id="9e14c-607">If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="9e14c-608">参数</span><span class="sxs-lookup"><span data-stu-id="9e14c-608">Parameters</span></span>

|<span data-ttu-id="9e14c-609">名称</span><span class="sxs-lookup"><span data-stu-id="9e14c-609">Name</span></span>| <span data-ttu-id="9e14c-610">类型</span><span class="sxs-lookup"><span data-stu-id="9e14c-610">Type</span></span>| <span data-ttu-id="9e14c-611">说明</span><span class="sxs-lookup"><span data-stu-id="9e14c-611">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="9e14c-612">字符串 &#124; 对象</span><span class="sxs-lookup"><span data-stu-id="9e14c-612">String &#124; Object</span></span>| |<span data-ttu-id="9e14c-p139">一个包含文本和 HTML 且表示答复窗体的正文的字符串。字符串限制为 32 KB。</span><span class="sxs-lookup"><span data-stu-id="9e14c-p139">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="9e14c-615">**或**</span><span class="sxs-lookup"><span data-stu-id="9e14c-615">**OR**</span></span><br/><span data-ttu-id="9e14c-p140">包含正文或附件数据和回调函数的对象。对象定义如下。</span><span class="sxs-lookup"><span data-stu-id="9e14c-p140">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="9e14c-618">String</span><span class="sxs-lookup"><span data-stu-id="9e14c-618">String</span></span> | <span data-ttu-id="9e14c-619">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="9e14c-619">&lt;optional&gt;</span></span> | <span data-ttu-id="9e14c-p141">一个包含文本和 HTML 且表示答复窗体的正文的字符串。字符串限制为 32 KB。</span><span class="sxs-lookup"><span data-stu-id="9e14c-p141">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="9e14c-622">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="9e14c-622">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="9e14c-623">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="9e14c-623">&lt;optional&gt;</span></span> | <span data-ttu-id="9e14c-624">JSON 对象（是文件或项目附件）的数组。</span><span class="sxs-lookup"><span data-stu-id="9e14c-624">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="9e14c-625">String</span><span class="sxs-lookup"><span data-stu-id="9e14c-625">String</span></span> | | <span data-ttu-id="9e14c-p142">指示附件的类型。必须是文件附件的 `file` 或项目附件的 `item`。</span><span class="sxs-lookup"><span data-stu-id="9e14c-p142">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="9e14c-628">字符串</span><span class="sxs-lookup"><span data-stu-id="9e14c-628">String</span></span> | | <span data-ttu-id="9e14c-629">一个包含附件的名称的字符串，最多包含 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="9e14c-629">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="9e14c-630">String</span><span class="sxs-lookup"><span data-stu-id="9e14c-630">String</span></span> | | <span data-ttu-id="9e14c-p143">仅在将 `type` 设置为 `file` 时使用。文件的位置的 URI。</span><span class="sxs-lookup"><span data-stu-id="9e14c-p143">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="9e14c-633">String</span><span class="sxs-lookup"><span data-stu-id="9e14c-633">String</span></span> | | <span data-ttu-id="9e14c-p144">仅在将 `type` 设置为 `item` 时使用。附件的 EWS 项目 ID。字符串最长为 100 个字符。</span><span class="sxs-lookup"><span data-stu-id="9e14c-p144">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="9e14c-637">函数</span><span class="sxs-lookup"><span data-stu-id="9e14c-637">function</span></span> | <span data-ttu-id="9e14c-638">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="9e14c-638">&lt;optional&gt;</span></span> | <span data-ttu-id="9e14c-639">方法完成后，使用单个参数 `asyncResult`（一个 [AsyncResult](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="9e14c-639">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="9e14c-640">要求</span><span class="sxs-lookup"><span data-stu-id="9e14c-640">Requirements</span></span>

|<span data-ttu-id="9e14c-641">要求</span><span class="sxs-lookup"><span data-stu-id="9e14c-641">Requirement</span></span>| <span data-ttu-id="9e14c-642">值</span><span class="sxs-lookup"><span data-stu-id="9e14c-642">Value</span></span>|
|---|---|
|[<span data-ttu-id="9e14c-643">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="9e14c-643">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9e14c-644">1.0</span><span class="sxs-lookup"><span data-stu-id="9e14c-644">1.0</span></span>|
|[<span data-ttu-id="9e14c-645">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="9e14c-645">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9e14c-646">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9e14c-646">ReadItem</span></span>|
|[<span data-ttu-id="9e14c-647">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="9e14c-647">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="9e14c-648">阅读</span><span class="sxs-lookup"><span data-stu-id="9e14c-648">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="9e14c-649">示例</span><span class="sxs-lookup"><span data-stu-id="9e14c-649">Examples</span></span>

<span data-ttu-id="9e14c-650">以下代码将一个字符串传递到 `displayReplyAllForm` 函数。</span><span class="sxs-lookup"><span data-stu-id="9e14c-650">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="9e14c-651">使用空白正文答复。</span><span class="sxs-lookup"><span data-stu-id="9e14c-651">Reply with an empty body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="9e14c-652">仅使用正文答复。</span><span class="sxs-lookup"><span data-stu-id="9e14c-652">Reply with just a body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="9e14c-653">使用正文和文件附件答复。</span><span class="sxs-lookup"><span data-stu-id="9e14c-653">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="9e14c-654">使用正文和项目附件答复。</span><span class="sxs-lookup"><span data-stu-id="9e14c-654">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="9e14c-655">使用正文、文件附件、项目附件和回调答复。</span><span class="sxs-lookup"><span data-stu-id="9e14c-655">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="displayreplyformformdata-callback"></a><span data-ttu-id="9e14c-656">displayReplyForm(formData, [callback])</span><span class="sxs-lookup"><span data-stu-id="9e14c-656">displayReplyForm(formData, [callback])</span></span>

<span data-ttu-id="9e14c-657">显示答复窗体，其中仅包括所选邮件的发件人或所选约会的组织者。</span><span class="sxs-lookup"><span data-stu-id="9e14c-657">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="9e14c-658">IOS 或 Android 上的 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="9e14c-658">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="9e14c-659">在 web 上的 Outlook 中, 答复窗体显示为3列视图中的弹出窗体和2列或1列视图中的弹出式窗体。</span><span class="sxs-lookup"><span data-stu-id="9e14c-659">In Outlook on the web, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="9e14c-660">如果任意字符串参数超出其限制，`displayReplyForm` 将引发异常。</span><span class="sxs-lookup"><span data-stu-id="9e14c-660">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="9e14c-661">如果在`formData.attachments`参数中指定了附件, 则 web 上的 Outlook 和桌面客户端将尝试下载所有附件并将其附加到答复窗体。</span><span class="sxs-lookup"><span data-stu-id="9e14c-661">When attachments are specified in the `formData.attachments` parameter, Outlook on the web and desktop clients attempt to download all attachments and attach them to the reply form.</span></span> <span data-ttu-id="9e14c-662">如果无法添加任何附件，则在窗体 UI 中显示错误。</span><span class="sxs-lookup"><span data-stu-id="9e14c-662">If any attachments fail to be added, an error is shown in the form UI.</span></span> <span data-ttu-id="9e14c-663">如果这不可能，则不引发错误消息。</span><span class="sxs-lookup"><span data-stu-id="9e14c-663">If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="9e14c-664">参数</span><span class="sxs-lookup"><span data-stu-id="9e14c-664">Parameters</span></span>

|<span data-ttu-id="9e14c-665">名称</span><span class="sxs-lookup"><span data-stu-id="9e14c-665">Name</span></span>| <span data-ttu-id="9e14c-666">类型</span><span class="sxs-lookup"><span data-stu-id="9e14c-666">Type</span></span>| <span data-ttu-id="9e14c-667">说明</span><span class="sxs-lookup"><span data-stu-id="9e14c-667">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="9e14c-668">字符串 &#124; 对象</span><span class="sxs-lookup"><span data-stu-id="9e14c-668">String &#124; Object</span></span>| | <span data-ttu-id="9e14c-p146">一个包含文本和 HTML 且表示答复窗体的正文的字符串。字符串限制为 32 KB。</span><span class="sxs-lookup"><span data-stu-id="9e14c-p146">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="9e14c-671">**或**</span><span class="sxs-lookup"><span data-stu-id="9e14c-671">**OR**</span></span><br/><span data-ttu-id="9e14c-p147">包含正文或附件数据和回调函数的对象。对象定义如下。</span><span class="sxs-lookup"><span data-stu-id="9e14c-p147">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="9e14c-674">String</span><span class="sxs-lookup"><span data-stu-id="9e14c-674">String</span></span> | <span data-ttu-id="9e14c-675">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="9e14c-675">&lt;optional&gt;</span></span> | <span data-ttu-id="9e14c-p148">一个包含文本和 HTML 且表示答复窗体的正文的字符串。字符串限制为 32 KB。</span><span class="sxs-lookup"><span data-stu-id="9e14c-p148">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="9e14c-678">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="9e14c-678">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="9e14c-679">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="9e14c-679">&lt;optional&gt;</span></span> | <span data-ttu-id="9e14c-680">JSON 对象（是文件或项目附件）的数组。</span><span class="sxs-lookup"><span data-stu-id="9e14c-680">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="9e14c-681">String</span><span class="sxs-lookup"><span data-stu-id="9e14c-681">String</span></span> | | <span data-ttu-id="9e14c-p149">指示附件的类型。必须是文件附件的 `file` 或项目附件的 `item`。</span><span class="sxs-lookup"><span data-stu-id="9e14c-p149">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="9e14c-684">字符串</span><span class="sxs-lookup"><span data-stu-id="9e14c-684">String</span></span> | | <span data-ttu-id="9e14c-685">一个包含附件的名称的字符串，最多包含 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="9e14c-685">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="9e14c-686">String</span><span class="sxs-lookup"><span data-stu-id="9e14c-686">String</span></span> | | <span data-ttu-id="9e14c-p150">仅在将 `type` 设置为 `file` 时使用。文件的位置的 URI。</span><span class="sxs-lookup"><span data-stu-id="9e14c-p150">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="9e14c-689">字符串</span><span class="sxs-lookup"><span data-stu-id="9e14c-689">String</span></span> | | <span data-ttu-id="9e14c-p151">仅在将 `type` 设置为 `item` 时使用。附件的 EWS 项目 ID。字符串最长为 100 个字符。</span><span class="sxs-lookup"><span data-stu-id="9e14c-p151">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="9e14c-693">函数</span><span class="sxs-lookup"><span data-stu-id="9e14c-693">function</span></span> | <span data-ttu-id="9e14c-694">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="9e14c-694">&lt;optional&gt;</span></span> | <span data-ttu-id="9e14c-695">方法完成后，使用单个参数 `asyncResult`（一个 [AsyncResult](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="9e14c-695">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="9e14c-696">要求</span><span class="sxs-lookup"><span data-stu-id="9e14c-696">Requirements</span></span>

|<span data-ttu-id="9e14c-697">要求</span><span class="sxs-lookup"><span data-stu-id="9e14c-697">Requirement</span></span>| <span data-ttu-id="9e14c-698">值</span><span class="sxs-lookup"><span data-stu-id="9e14c-698">Value</span></span>|
|---|---|
|[<span data-ttu-id="9e14c-699">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="9e14c-699">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9e14c-700">1.0</span><span class="sxs-lookup"><span data-stu-id="9e14c-700">1.0</span></span>|
|[<span data-ttu-id="9e14c-701">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="9e14c-701">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9e14c-702">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9e14c-702">ReadItem</span></span>|
|[<span data-ttu-id="9e14c-703">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="9e14c-703">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="9e14c-704">阅读</span><span class="sxs-lookup"><span data-stu-id="9e14c-704">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="9e14c-705">示例</span><span class="sxs-lookup"><span data-stu-id="9e14c-705">Examples</span></span>

<span data-ttu-id="9e14c-706">以下代码将一个字符串传递到 `displayReplyForm` 函数。</span><span class="sxs-lookup"><span data-stu-id="9e14c-706">The following code passes a string to the `displayReplyForm` function.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="9e14c-707">使用空白正文答复。</span><span class="sxs-lookup"><span data-stu-id="9e14c-707">Reply with an empty body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="9e14c-708">仅使用正文答复。</span><span class="sxs-lookup"><span data-stu-id="9e14c-708">Reply with just a body.</span></span>

```javascript
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="9e14c-709">使用正文和文件附件答复。</span><span class="sxs-lookup"><span data-stu-id="9e14c-709">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="9e14c-710">使用正文和项目附件答复。</span><span class="sxs-lookup"><span data-stu-id="9e14c-710">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="9e14c-711">使用正文、文件附件、项目附件和回调答复。</span><span class="sxs-lookup"><span data-stu-id="9e14c-711">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="getentities--entitiesjavascriptapioutlookofficeentitiesviewoutlook-js-12"></a><span data-ttu-id="9e14c-712">getEntities() → {[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.2)}</span><span class="sxs-lookup"><span data-stu-id="9e14c-712">getEntities() → {[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.2)}</span></span>

<span data-ttu-id="9e14c-713">获取在所选项目的正文中找到的实体。</span><span class="sxs-lookup"><span data-stu-id="9e14c-713">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="9e14c-714">IOS 或 Android 上的 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="9e14c-714">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="9e14c-715">要求</span><span class="sxs-lookup"><span data-stu-id="9e14c-715">Requirements</span></span>

|<span data-ttu-id="9e14c-716">要求</span><span class="sxs-lookup"><span data-stu-id="9e14c-716">Requirement</span></span>| <span data-ttu-id="9e14c-717">值</span><span class="sxs-lookup"><span data-stu-id="9e14c-717">Value</span></span>|
|---|---|
|[<span data-ttu-id="9e14c-718">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="9e14c-718">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9e14c-719">1.0</span><span class="sxs-lookup"><span data-stu-id="9e14c-719">1.0</span></span>|
|[<span data-ttu-id="9e14c-720">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="9e14c-720">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9e14c-721">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9e14c-721">ReadItem</span></span>|
|[<span data-ttu-id="9e14c-722">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="9e14c-722">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="9e14c-723">阅读</span><span class="sxs-lookup"><span data-stu-id="9e14c-723">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="9e14c-724">返回：</span><span class="sxs-lookup"><span data-stu-id="9e14c-724">Returns:</span></span>

<span data-ttu-id="9e14c-725">类型：[Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="9e14c-725">Type: [Entities](/javascript/api/outlook/office.entities?view=outlook-js-1.2)</span></span>

##### <a name="example"></a><span data-ttu-id="9e14c-726">示例</span><span class="sxs-lookup"><span data-stu-id="9e14c-726">Example</span></span>

<span data-ttu-id="9e14c-727">以下示例访问当前项目的正文中的联系人实体。</span><span class="sxs-lookup"><span data-stu-id="9e14c-727">The following example accesses the contacts entities in the current item's body.</span></span>

```javascript
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlookofficecontactmeetingsuggestionjavascriptapioutlookofficemeetingsuggestionphonenumberjavascriptapioutlookofficephonenumbertasksuggestionjavascriptapioutlookofficetasksuggestionviewoutlook-js-12"></a><span data-ttu-id="9e14c-728">getEntitiesByType (entityType) → (nullable) {Array. < (String |[联系人](/javascript/api/outlook/office.contact)|[](/javascript/api/outlook/office.meetingsuggestion)MeetingSuggestion|[](/javascript/api/outlook/office.phonenumber)PhoneNumber|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion)？ view = outlook-js-1.2) >}</span><span class="sxs-lookup"><span data-stu-id="9e14c-728">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion)?view=outlook-js-1.2)>}</span></span>

<span data-ttu-id="9e14c-729">获取所选项目的正文中找到的指定实体类型的所有实体的数组。</span><span class="sxs-lookup"><span data-stu-id="9e14c-729">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="9e14c-730">IOS 或 Android 上的 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="9e14c-730">This method is not supported in Outlook on iOS or Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="9e14c-731">参数</span><span class="sxs-lookup"><span data-stu-id="9e14c-731">Parameters</span></span>

|<span data-ttu-id="9e14c-732">名称</span><span class="sxs-lookup"><span data-stu-id="9e14c-732">Name</span></span>| <span data-ttu-id="9e14c-733">类型</span><span class="sxs-lookup"><span data-stu-id="9e14c-733">Type</span></span>| <span data-ttu-id="9e14c-734">说明</span><span class="sxs-lookup"><span data-stu-id="9e14c-734">Description</span></span>|
|---|---|---|
|`entityType`| [<span data-ttu-id="9e14c-735">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="9e14c-735">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook/office.mailboxenums.entitytype?view=outlook-js-1.2)|<span data-ttu-id="9e14c-736">EntityType 枚举值之一。</span><span class="sxs-lookup"><span data-stu-id="9e14c-736">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="9e14c-737">要求</span><span class="sxs-lookup"><span data-stu-id="9e14c-737">Requirements</span></span>

|<span data-ttu-id="9e14c-738">要求</span><span class="sxs-lookup"><span data-stu-id="9e14c-738">Requirement</span></span>| <span data-ttu-id="9e14c-739">值</span><span class="sxs-lookup"><span data-stu-id="9e14c-739">Value</span></span>|
|---|---|
|[<span data-ttu-id="9e14c-740">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="9e14c-740">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9e14c-741">1.0</span><span class="sxs-lookup"><span data-stu-id="9e14c-741">1.0</span></span>|
|[<span data-ttu-id="9e14c-742">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="9e14c-742">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9e14c-743">受限</span><span class="sxs-lookup"><span data-stu-id="9e14c-743">Restricted</span></span>|
|[<span data-ttu-id="9e14c-744">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="9e14c-744">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="9e14c-745">阅读</span><span class="sxs-lookup"><span data-stu-id="9e14c-745">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="9e14c-746">返回：</span><span class="sxs-lookup"><span data-stu-id="9e14c-746">Returns:</span></span>

<span data-ttu-id="9e14c-747">如果在 `entityType` 中传递的值不是 `EntityType` 枚举的有效成员，该方法返回 null。</span><span class="sxs-lookup"><span data-stu-id="9e14c-747">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="9e14c-748">如果指定类型的任何实体都不存在于该项目的正文中，该方法将返回空数组。</span><span class="sxs-lookup"><span data-stu-id="9e14c-748">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="9e14c-749">否则，返回的数组中对象的类型取决于 `entityType` 参数中请求实体的类型。</span><span class="sxs-lookup"><span data-stu-id="9e14c-749">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="9e14c-750">当使用此方法的最低权限级别**受限**时，某些实体类型需要 **ReadItem** 才能进行访问，如下表中所指定。</span><span class="sxs-lookup"><span data-stu-id="9e14c-750">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

| <span data-ttu-id="9e14c-751">`entityType` 的值</span><span class="sxs-lookup"><span data-stu-id="9e14c-751">Value of `entityType`</span></span> | <span data-ttu-id="9e14c-752">返回的数组中对象的类型</span><span class="sxs-lookup"><span data-stu-id="9e14c-752">Type of objects in returned array</span></span> | <span data-ttu-id="9e14c-753">所需权限级别</span><span class="sxs-lookup"><span data-stu-id="9e14c-753">Required Permission Level</span></span> |
| --- | --- | --- |
| `Address` | <span data-ttu-id="9e14c-754">字符串</span><span class="sxs-lookup"><span data-stu-id="9e14c-754">String</span></span> | <span data-ttu-id="9e14c-755">**受限**</span><span class="sxs-lookup"><span data-stu-id="9e14c-755">**Restricted**</span></span> |
| `Contact` | <span data-ttu-id="9e14c-756">Contact</span><span class="sxs-lookup"><span data-stu-id="9e14c-756">Contact</span></span> | <span data-ttu-id="9e14c-757">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="9e14c-757">**ReadItem**</span></span> |
| `EmailAddress` | <span data-ttu-id="9e14c-758">String</span><span class="sxs-lookup"><span data-stu-id="9e14c-758">String</span></span> | <span data-ttu-id="9e14c-759">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="9e14c-759">**ReadItem**</span></span> |
| `MeetingSuggestion` | <span data-ttu-id="9e14c-760">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="9e14c-760">MeetingSuggestion</span></span> | <span data-ttu-id="9e14c-761">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="9e14c-761">**ReadItem**</span></span> |
| `PhoneNumber` | <span data-ttu-id="9e14c-762">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="9e14c-762">PhoneNumber</span></span> | <span data-ttu-id="9e14c-763">**受限**</span><span class="sxs-lookup"><span data-stu-id="9e14c-763">**Restricted**</span></span> |
| `TaskSuggestion` | <span data-ttu-id="9e14c-764">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="9e14c-764">TaskSuggestion</span></span> | <span data-ttu-id="9e14c-765">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="9e14c-765">**ReadItem**</span></span> |
| `URL` | <span data-ttu-id="9e14c-766">字符串</span><span class="sxs-lookup"><span data-stu-id="9e14c-766">String</span></span> | <span data-ttu-id="9e14c-767">**受限**</span><span class="sxs-lookup"><span data-stu-id="9e14c-767">**Restricted**</span></span> |

<span data-ttu-id="9e14c-768">类型: Array. < (String |[联系人](/javascript/api/outlook/office.contact)|[](/javascript/api/outlook/office.meetingsuggestion)MeetingSuggestion|[](/javascript/api/outlook/office.phonenumber)PhoneNumber|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion)？ view = outlook-js-1.2) ></span><span class="sxs-lookup"><span data-stu-id="9e14c-768">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion)?view=outlook-js-1.2)></span></span>

##### <a name="example"></a><span data-ttu-id="9e14c-769">示例</span><span class="sxs-lookup"><span data-stu-id="9e14c-769">Example</span></span>

<span data-ttu-id="9e14c-770">以下示例显示了如何访问表示当前项目的正文中的邮政地址的字符串数组。</span><span class="sxs-lookup"><span data-stu-id="9e14c-770">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlookofficecontactmeetingsuggestionjavascriptapioutlookofficemeetingsuggestionphonenumberjavascriptapioutlookofficephonenumbertasksuggestionjavascriptapioutlookofficetasksuggestionviewoutlook-js-12"></a><span data-ttu-id="9e14c-771">getFilteredEntitiesByName (name) → (nullable) {Array. < (String |[联系人](/javascript/api/outlook/office.contact)|[](/javascript/api/outlook/office.meetingsuggestion)MeetingSuggestion|[](/javascript/api/outlook/office.phonenumber)PhoneNumber|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion)？ view = outlook-js-1.2) >}</span><span class="sxs-lookup"><span data-stu-id="9e14c-771">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion)?view=outlook-js-1.2)>}</span></span>

<span data-ttu-id="9e14c-772">返回传递清单 XML 文件中定义的命名筛选器的所选项目中的已知实体。</span><span class="sxs-lookup"><span data-stu-id="9e14c-772">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="9e14c-773">IOS 或 Android 上的 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="9e14c-773">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="9e14c-774">`getFilteredEntitiesByName` 方法返回匹配在具有指定 `FilterName` 元素值的清单 XML 文件中的 [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) 规则元素中定义的正则表达式的实体。</span><span class="sxs-lookup"><span data-stu-id="9e14c-774">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="9e14c-775">参数</span><span class="sxs-lookup"><span data-stu-id="9e14c-775">Parameters</span></span>

|<span data-ttu-id="9e14c-776">名称</span><span class="sxs-lookup"><span data-stu-id="9e14c-776">Name</span></span>| <span data-ttu-id="9e14c-777">类型</span><span class="sxs-lookup"><span data-stu-id="9e14c-777">Type</span></span>| <span data-ttu-id="9e14c-778">说明</span><span class="sxs-lookup"><span data-stu-id="9e14c-778">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="9e14c-779">String</span><span class="sxs-lookup"><span data-stu-id="9e14c-779">String</span></span>|<span data-ttu-id="9e14c-780">定义筛选器匹配的 `ItemHasKnownEntity` 规则元素的名称。</span><span class="sxs-lookup"><span data-stu-id="9e14c-780">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="9e14c-781">要求</span><span class="sxs-lookup"><span data-stu-id="9e14c-781">Requirements</span></span>

|<span data-ttu-id="9e14c-782">要求</span><span class="sxs-lookup"><span data-stu-id="9e14c-782">Requirement</span></span>| <span data-ttu-id="9e14c-783">值</span><span class="sxs-lookup"><span data-stu-id="9e14c-783">Value</span></span>|
|---|---|
|[<span data-ttu-id="9e14c-784">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="9e14c-784">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9e14c-785">1.0</span><span class="sxs-lookup"><span data-stu-id="9e14c-785">1.0</span></span>|
|[<span data-ttu-id="9e14c-786">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="9e14c-786">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9e14c-787">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9e14c-787">ReadItem</span></span>|
|[<span data-ttu-id="9e14c-788">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="9e14c-788">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="9e14c-789">阅读</span><span class="sxs-lookup"><span data-stu-id="9e14c-789">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="9e14c-790">返回：</span><span class="sxs-lookup"><span data-stu-id="9e14c-790">Returns:</span></span>

<span data-ttu-id="9e14c-p153">如果具有匹配 `name` 参数的 `FilterName` 元素值的清单中没有任何 `ItemHasKnownEntity` 元素，则该方法返回 `null`。如果 `name` 参数匹配清单中的 `ItemHasKnownEntity` 元素，但在匹配的当前项目中没有实体，则该方法返回一个空数组。</span><span class="sxs-lookup"><span data-stu-id="9e14c-p153">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="9e14c-793">类型: Array. < (String |[联系人](/javascript/api/outlook/office.contact)|[](/javascript/api/outlook/office.meetingsuggestion)MeetingSuggestion|[](/javascript/api/outlook/office.phonenumber)PhoneNumber|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion)？ view = outlook-js-1.2) ></span><span class="sxs-lookup"><span data-stu-id="9e14c-793">Type: Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion)?view=outlook-js-1.2)></span></span>

#### <a name="getregexmatches--object"></a><span data-ttu-id="9e14c-794">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="9e14c-794">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="9e14c-795">返回所选项目中匹配在清单 XML 文件中定义的正则表达式的字符串值。</span><span class="sxs-lookup"><span data-stu-id="9e14c-795">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="9e14c-796">IOS 或 Android 上的 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="9e14c-796">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="9e14c-p154">`getRegExMatches` 方法返回匹配在清单 XML 文件中的每个 `ItemHasRegularExpressionMatch` 或 `ItemHasKnownEntity` 规则元素中定义的正则表达式的字符串。对于 `ItemHasRegularExpressionMatch` 规则，匹配字符串必须发生在该规则指定的项目的属性中。`PropertyName` 简单类型定义支持的属性。</span><span class="sxs-lookup"><span data-stu-id="9e14c-p154">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="9e14c-800">例如，考虑一个外接程序清单具有以下 `Rule` 元素：</span><span class="sxs-lookup"><span data-stu-id="9e14c-800">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="9e14c-801">从 `getRegExMatches` 返回的对象应有两个属性：`fruits` 和 `veggies`。</span><span class="sxs-lookup"><span data-stu-id="9e14c-801">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

> [!NOTE]
> <span data-ttu-id="9e14c-p155">如果在项目的正文属性上指定 `ItemHasRegularExpressionMatch` 规则，则正则表达式应进一步筛选正文，不应尝试返回该项目的整个正文。使用正则表达式（如 `.*`）获取项目的整个正文并不总是返回预期的结果。</span><span class="sxs-lookup"><span data-stu-id="9e14c-p155">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="requirements"></a><span data-ttu-id="9e14c-804">要求</span><span class="sxs-lookup"><span data-stu-id="9e14c-804">Requirements</span></span>

|<span data-ttu-id="9e14c-805">要求</span><span class="sxs-lookup"><span data-stu-id="9e14c-805">Requirement</span></span>| <span data-ttu-id="9e14c-806">值</span><span class="sxs-lookup"><span data-stu-id="9e14c-806">Value</span></span>|
|---|---|
|[<span data-ttu-id="9e14c-807">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="9e14c-807">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9e14c-808">1.0</span><span class="sxs-lookup"><span data-stu-id="9e14c-808">1.0</span></span>|
|[<span data-ttu-id="9e14c-809">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="9e14c-809">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9e14c-810">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9e14c-810">ReadItem</span></span>|
|[<span data-ttu-id="9e14c-811">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="9e14c-811">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="9e14c-812">阅读</span><span class="sxs-lookup"><span data-stu-id="9e14c-812">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="9e14c-813">返回：</span><span class="sxs-lookup"><span data-stu-id="9e14c-813">Returns:</span></span>

<span data-ttu-id="9e14c-p156">一个包含匹配在清单 XML 文件中定义的正则表达式的字符串数组的对象。每个数组的名称等于匹配 `ItemHasRegularExpressionMatch` 规则的 `RegExName` 属性或匹配 `ItemHasKnownEntity` 规则的 `FilterName` 属性的相应值。</span><span class="sxs-lookup"><span data-stu-id="9e14c-p156">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<dl class="param-type"><span data-ttu-id="9e14c-816">

<dt>
类型</dt>


</span><span class="sxs-lookup"><span data-stu-id="9e14c-816">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="9e14c-817">对象</span><span class="sxs-lookup"><span data-stu-id="9e14c-817">Object</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="9e14c-818">示例</span><span class="sxs-lookup"><span data-stu-id="9e14c-818">Example</span></span>

<span data-ttu-id="9e14c-819">以下示例显示了如何访问正则表达式 <rule> 元素 `fruits` 和 `veggies` 的匹配项的数组，这些元素在清单中指定。</rule></span><span class="sxs-lookup"><span data-stu-id="9e14c-819">The following example shows how to access the array of matches for the regular expression <rule>elements `fruits` and `veggies`, which are specified in the manifest.</rule></span></span>

```javascript
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veggies = allMatches.veggies;
```

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="9e14c-820">getRegExMatchesByName(name) → (nullable) {Array.<String>}</span><span class="sxs-lookup"><span data-stu-id="9e14c-820">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="9e14c-821">返回匹配在清单 XML 文件中定义的命名正则表达式的所选项目中的字符串值。</span><span class="sxs-lookup"><span data-stu-id="9e14c-821">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="9e14c-822">IOS 或 Android 上的 Outlook 不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="9e14c-822">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="9e14c-823">`getRegExMatchesByName` 方法返回匹配在具有指定 `RegExName` 元素值的清单 XML 文件中的 `ItemHasRegularExpressionMatch` 规则元素中定义的正则表达式的字符串。</span><span class="sxs-lookup"><span data-stu-id="9e14c-823">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="9e14c-p157">如果在项目的正文属性上指定 `ItemHasRegularExpressionMatch` 规则，则正则表达式应进一步筛选正文，不应尝试返回该项目的整个正文。使用正则表达式（如 `.*`）获取项目的整个正文并不总是返回预期的结果。</span><span class="sxs-lookup"><span data-stu-id="9e14c-p157">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="9e14c-826">参数</span><span class="sxs-lookup"><span data-stu-id="9e14c-826">Parameters</span></span>

|<span data-ttu-id="9e14c-827">名称</span><span class="sxs-lookup"><span data-stu-id="9e14c-827">Name</span></span>| <span data-ttu-id="9e14c-828">类型</span><span class="sxs-lookup"><span data-stu-id="9e14c-828">Type</span></span>| <span data-ttu-id="9e14c-829">说明</span><span class="sxs-lookup"><span data-stu-id="9e14c-829">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="9e14c-830">String</span><span class="sxs-lookup"><span data-stu-id="9e14c-830">String</span></span>|<span data-ttu-id="9e14c-831">定义筛选器匹配的 `ItemHasRegularExpressionMatch` 规则元素的名称。</span><span class="sxs-lookup"><span data-stu-id="9e14c-831">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="9e14c-832">要求</span><span class="sxs-lookup"><span data-stu-id="9e14c-832">Requirements</span></span>

|<span data-ttu-id="9e14c-833">要求</span><span class="sxs-lookup"><span data-stu-id="9e14c-833">Requirement</span></span>| <span data-ttu-id="9e14c-834">值</span><span class="sxs-lookup"><span data-stu-id="9e14c-834">Value</span></span>|
|---|---|
|[<span data-ttu-id="9e14c-835">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="9e14c-835">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9e14c-836">1.0</span><span class="sxs-lookup"><span data-stu-id="9e14c-836">1.0</span></span>|
|[<span data-ttu-id="9e14c-837">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="9e14c-837">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9e14c-838">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9e14c-838">ReadItem</span></span>|
|[<span data-ttu-id="9e14c-839">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="9e14c-839">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="9e14c-840">阅读</span><span class="sxs-lookup"><span data-stu-id="9e14c-840">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="9e14c-841">返回：</span><span class="sxs-lookup"><span data-stu-id="9e14c-841">Returns:</span></span>

<span data-ttu-id="9e14c-842">一个包含匹配在清单 XML 文件中定义的正则表达式的字符串的数组。</span><span class="sxs-lookup"><span data-stu-id="9e14c-842">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<dl class="param-type"><span data-ttu-id="9e14c-843">

<dt>类型</dt>

</span><span class="sxs-lookup"><span data-stu-id="9e14c-843">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="9e14c-844">Array.< String ></span><span class="sxs-lookup"><span data-stu-id="9e14c-844">Array.< String ></span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="9e14c-845">示例</span><span class="sxs-lookup"><span data-stu-id="9e14c-845">Example</span></span>

```javascript
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

#### <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="9e14c-846">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="9e14c-846">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="9e14c-847">以异步方式返回邮件的主题或正文中选定的数据。</span><span class="sxs-lookup"><span data-stu-id="9e14c-847">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="9e14c-p158">如果没有选定内容，但光标位于正文或主题中，此方法将会为所选数据返回 null。如果选定的是字段，而不是正文或主题，则此方法返回 `InvalidSelection` 错误。</span><span class="sxs-lookup"><span data-stu-id="9e14c-p158">If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="9e14c-850">参数</span><span class="sxs-lookup"><span data-stu-id="9e14c-850">Parameters</span></span>

|<span data-ttu-id="9e14c-851">名称</span><span class="sxs-lookup"><span data-stu-id="9e14c-851">Name</span></span>| <span data-ttu-id="9e14c-852">类型</span><span class="sxs-lookup"><span data-stu-id="9e14c-852">Type</span></span>| <span data-ttu-id="9e14c-853">属性</span><span class="sxs-lookup"><span data-stu-id="9e14c-853">Attributes</span></span>| <span data-ttu-id="9e14c-854">说明</span><span class="sxs-lookup"><span data-stu-id="9e14c-854">Description</span></span>|
|---|---|---|---|
|`coercionType`| [<span data-ttu-id="9e14c-855">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="9e14c-855">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="9e14c-p159">请求数据的格式。如果为文本，则此方法返回纯文本作为字符串，删除任何显示的 HTML 标记。如果为 HTML，则此方法返回所选文本，不论是纯文本还是 HTML。</span><span class="sxs-lookup"><span data-stu-id="9e14c-p159">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`| <span data-ttu-id="9e14c-859">Object</span><span class="sxs-lookup"><span data-stu-id="9e14c-859">Object</span></span>| <span data-ttu-id="9e14c-860">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="9e14c-860">&lt;optional&gt;</span></span>|<span data-ttu-id="9e14c-861">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="9e14c-861">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="9e14c-862">Object</span><span class="sxs-lookup"><span data-stu-id="9e14c-862">Object</span></span>| <span data-ttu-id="9e14c-863">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="9e14c-863">&lt;optional&gt;</span></span>|<span data-ttu-id="9e14c-864">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="9e14c-864">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="9e14c-865">function</span><span class="sxs-lookup"><span data-stu-id="9e14c-865">function</span></span>||<span data-ttu-id="9e14c-866">方法完成后，使用单个参数 `callback`（一个 [`asyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `AsyncResult` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="9e14c-866">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="9e14c-867">若要从回调方法访问所选数据，请调用 `asyncResult.value.data`。</span><span class="sxs-lookup"><span data-stu-id="9e14c-867">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="9e14c-868">若要访问选定内容的源属性，请调用 `asyncResult.value.sourceProperty`，这将为 `body` 或 `subject`。</span><span class="sxs-lookup"><span data-stu-id="9e14c-868">To access the source property that the selection comes from, call `asyncResult.value.sourceProperty`, which will be either `body` or `subject`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="9e14c-869">要求</span><span class="sxs-lookup"><span data-stu-id="9e14c-869">Requirements</span></span>

|<span data-ttu-id="9e14c-870">要求</span><span class="sxs-lookup"><span data-stu-id="9e14c-870">Requirement</span></span>| <span data-ttu-id="9e14c-871">值</span><span class="sxs-lookup"><span data-stu-id="9e14c-871">Value</span></span>|
|---|---|
|[<span data-ttu-id="9e14c-872">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="9e14c-872">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9e14c-873">1.2</span><span class="sxs-lookup"><span data-stu-id="9e14c-873">1.2</span></span>|
|[<span data-ttu-id="9e14c-874">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="9e14c-874">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9e14c-875">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="9e14c-875">ReadWriteItem</span></span>|
|[<span data-ttu-id="9e14c-876">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="9e14c-876">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="9e14c-877">撰写</span><span class="sxs-lookup"><span data-stu-id="9e14c-877">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="9e14c-878">返回：</span><span class="sxs-lookup"><span data-stu-id="9e14c-878">Returns:</span></span>

<span data-ttu-id="9e14c-879">作为字符串的所选数据的格式由 `coercionType` 确定。</span><span class="sxs-lookup"><span data-stu-id="9e14c-879">The selected data as a string with format determined by `coercionType`.</span></span>

<dl class="param-type"><span data-ttu-id="9e14c-880">

<dt>
类型</dt>


</span><span class="sxs-lookup"><span data-stu-id="9e14c-880">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="9e14c-881">字符串</span><span class="sxs-lookup"><span data-stu-id="9e14c-881">String</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="9e14c-882">示例</span><span class="sxs-lookup"><span data-stu-id="9e14c-882">Example</span></span>

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

#### <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="9e14c-883">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="9e14c-883">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="9e14c-884">异步加载所选项目上此外接程序的自定义属性。</span><span class="sxs-lookup"><span data-stu-id="9e14c-884">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="9e14c-p161">自定义属性基于每个应用、每个项目存储为键/值对。此方法在回调中返回 `CustomProperties` 对象，该回调提供访问特定于当前项目和当前外接程序的自定义属性的方法。自定义属性未在项目上加密，因此这不应用作安全存储。</span><span class="sxs-lookup"><span data-stu-id="9e14c-p161">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="9e14c-888">参数</span><span class="sxs-lookup"><span data-stu-id="9e14c-888">Parameters</span></span>

|<span data-ttu-id="9e14c-889">名称</span><span class="sxs-lookup"><span data-stu-id="9e14c-889">Name</span></span>| <span data-ttu-id="9e14c-890">类型</span><span class="sxs-lookup"><span data-stu-id="9e14c-890">Type</span></span>| <span data-ttu-id="9e14c-891">属性</span><span class="sxs-lookup"><span data-stu-id="9e14c-891">Attributes</span></span>| <span data-ttu-id="9e14c-892">说明</span><span class="sxs-lookup"><span data-stu-id="9e14c-892">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="9e14c-893">函数</span><span class="sxs-lookup"><span data-stu-id="9e14c-893">function</span></span>||<span data-ttu-id="9e14c-894">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="9e14c-894">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="9e14c-895">自定义属性作为 `asyncResult.value` 属性中的 [`CustomProperties`](/javascript/api/outlook/office.customproperties?view=outlook-js-1.2) 对象提供。</span><span class="sxs-lookup"><span data-stu-id="9e14c-895">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook/office.customproperties?view=outlook-js-1.2) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="9e14c-896">此对象可用于获取、设置以及从项目中删除自定义属性，并将自定义属性集的更改重新保存到服务器。</span><span class="sxs-lookup"><span data-stu-id="9e14c-896">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`| <span data-ttu-id="9e14c-897">对象</span><span class="sxs-lookup"><span data-stu-id="9e14c-897">Object</span></span>| <span data-ttu-id="9e14c-898">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="9e14c-898">&lt;optional&gt;</span></span>|<span data-ttu-id="9e14c-899">开发人员可以提供他们想要在回调函数中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="9e14c-899">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="9e14c-900">此对象可以通过回调函数中的 `asyncResult.asyncContext` 属性进行访问。</span><span class="sxs-lookup"><span data-stu-id="9e14c-900">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="9e14c-901">要求</span><span class="sxs-lookup"><span data-stu-id="9e14c-901">Requirements</span></span>

|<span data-ttu-id="9e14c-902">要求</span><span class="sxs-lookup"><span data-stu-id="9e14c-902">Requirement</span></span>| <span data-ttu-id="9e14c-903">值</span><span class="sxs-lookup"><span data-stu-id="9e14c-903">Value</span></span>|
|---|---|
|[<span data-ttu-id="9e14c-904">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="9e14c-904">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9e14c-905">1.0</span><span class="sxs-lookup"><span data-stu-id="9e14c-905">1.0</span></span>|
|[<span data-ttu-id="9e14c-906">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="9e14c-906">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9e14c-907">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9e14c-907">ReadItem</span></span>|
|[<span data-ttu-id="9e14c-908">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="9e14c-908">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="9e14c-909">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="9e14c-909">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="9e14c-910">示例</span><span class="sxs-lookup"><span data-stu-id="9e14c-910">Example</span></span>

<span data-ttu-id="9e14c-p164">以下代码示例显示了如何使用 `loadCustomPropertiesAsync` 方法异步加载特定于当前项目的自定义属性。该示例还显示了如何使用 `CustomProperties.saveAsync` 方法将这些属性重新保存到服务器。加载自定义属性后，该代码示例将使用 `CustomProperties.get` 方法读取自定义属性 `myProp`，使用 `CustomProperties.set` 方法写入自定义属性 `otherProp`，最后调用 `saveAsync` 方法保存这些自定义属性。</span><span class="sxs-lookup"><span data-stu-id="9e14c-p164">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

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

#### <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="9e14c-914">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="9e14c-914">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="9e14c-915">将附件从邮件或约会中删除。</span><span class="sxs-lookup"><span data-stu-id="9e14c-915">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="9e14c-916">`removeAttachmentAsync` 方法删除项目中带指定标识符的附件。</span><span class="sxs-lookup"><span data-stu-id="9e14c-916">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item.</span></span> <span data-ttu-id="9e14c-917">最佳做法是，仅当同一个邮件应用程序在同一会话中添加了一个附件时，你才应使用该附件标识符来删除该附件。</span><span class="sxs-lookup"><span data-stu-id="9e14c-917">As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session.</span></span> <span data-ttu-id="9e14c-918">在 web 和移动设备上的 Outlook 中, 附件标识符仅在同一个会话中有效。</span><span class="sxs-lookup"><span data-stu-id="9e14c-918">In Outlook on the web and mobile devices, the attachment identifier is valid only within the same session.</span></span> <span data-ttu-id="9e14c-919">当用户关闭应用程序，或者如果用户开始在内嵌窗体中撰写，并在随后弹出的内嵌窗体中继续在单独的窗口撰写时，会话即结束。</span><span class="sxs-lookup"><span data-stu-id="9e14c-919">A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="9e14c-920">参数</span><span class="sxs-lookup"><span data-stu-id="9e14c-920">Parameters</span></span>

|<span data-ttu-id="9e14c-921">名称</span><span class="sxs-lookup"><span data-stu-id="9e14c-921">Name</span></span>| <span data-ttu-id="9e14c-922">类型</span><span class="sxs-lookup"><span data-stu-id="9e14c-922">Type</span></span>| <span data-ttu-id="9e14c-923">属性</span><span class="sxs-lookup"><span data-stu-id="9e14c-923">Attributes</span></span>| <span data-ttu-id="9e14c-924">说明</span><span class="sxs-lookup"><span data-stu-id="9e14c-924">Description</span></span>|
|---|---|---|---|
|`attachmentId`| <span data-ttu-id="9e14c-925">String</span><span class="sxs-lookup"><span data-stu-id="9e14c-925">String</span></span>||<span data-ttu-id="9e14c-926">要删除的附件的标识符。</span><span class="sxs-lookup"><span data-stu-id="9e14c-926">The identifier of the attachment to remove.</span></span>|
|`options`| <span data-ttu-id="9e14c-927">对象</span><span class="sxs-lookup"><span data-stu-id="9e14c-927">Object</span></span>| <span data-ttu-id="9e14c-928">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="9e14c-928">&lt;optional&gt;</span></span>|<span data-ttu-id="9e14c-929">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="9e14c-929">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="9e14c-930">对象</span><span class="sxs-lookup"><span data-stu-id="9e14c-930">Object</span></span>| <span data-ttu-id="9e14c-931">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="9e14c-931">&lt;optional&gt;</span></span>|<span data-ttu-id="9e14c-932">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="9e14c-932">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="9e14c-933">函数</span><span class="sxs-lookup"><span data-stu-id="9e14c-933">function</span></span>| <span data-ttu-id="9e14c-934">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="9e14c-934">&lt;optional&gt;</span></span>|<span data-ttu-id="9e14c-935">方法完成后，使用单个参数 `callback`（一个 [`asyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `AsyncResult` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="9e14c-935">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="9e14c-936">如果删除附件失败，`asyncResult.error` 属性将包含一个说明失败原因的错误代码。</span><span class="sxs-lookup"><span data-stu-id="9e14c-936">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="9e14c-937">错误</span><span class="sxs-lookup"><span data-stu-id="9e14c-937">Errors</span></span>

| <span data-ttu-id="9e14c-938">错误代码</span><span class="sxs-lookup"><span data-stu-id="9e14c-938">Error code</span></span> | <span data-ttu-id="9e14c-939">说明</span><span class="sxs-lookup"><span data-stu-id="9e14c-939">Description</span></span> |
|------------|-------------|
| `InvalidAttachmentId` | <span data-ttu-id="9e14c-940">附件标识符不存在。</span><span class="sxs-lookup"><span data-stu-id="9e14c-940">The attachment identifier does not exist.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="9e14c-941">要求</span><span class="sxs-lookup"><span data-stu-id="9e14c-941">Requirements</span></span>

|<span data-ttu-id="9e14c-942">要求</span><span class="sxs-lookup"><span data-stu-id="9e14c-942">Requirement</span></span>| <span data-ttu-id="9e14c-943">值</span><span class="sxs-lookup"><span data-stu-id="9e14c-943">Value</span></span>|
|---|---|
|[<span data-ttu-id="9e14c-944">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="9e14c-944">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9e14c-945">1.1</span><span class="sxs-lookup"><span data-stu-id="9e14c-945">1.1</span></span>|
|[<span data-ttu-id="9e14c-946">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="9e14c-946">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9e14c-947">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="9e14c-947">ReadWriteItem</span></span>|
|[<span data-ttu-id="9e14c-948">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="9e14c-948">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="9e14c-949">撰写</span><span class="sxs-lookup"><span data-stu-id="9e14c-949">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="9e14c-950">示例</span><span class="sxs-lookup"><span data-stu-id="9e14c-950">Example</span></span>

<span data-ttu-id="9e14c-951">以下代码删除包含标识符 '0' 的附件。</span><span class="sxs-lookup"><span data-stu-id="9e14c-951">The following code removes an attachment with an identifier of '0'.</span></span>

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

#### <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="9e14c-952">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="9e14c-952">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="9e14c-953">以异步方式将数据插入到邮件的正文或主题中。</span><span class="sxs-lookup"><span data-stu-id="9e14c-953">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="9e14c-p166">`setSelectedDataAsync` 方法将指定字符串插入到项目主题或正文的光标位置，或者，如果在编辑器中已选择文本，则该方法将替换选择的文本。如果光标不在正文或主题字段中，则返回错误。插入之后，光标会位于插入内容的末尾。</span><span class="sxs-lookup"><span data-stu-id="9e14c-p166">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="9e14c-957">参数</span><span class="sxs-lookup"><span data-stu-id="9e14c-957">Parameters</span></span>

|<span data-ttu-id="9e14c-958">名称</span><span class="sxs-lookup"><span data-stu-id="9e14c-958">Name</span></span>| <span data-ttu-id="9e14c-959">类型</span><span class="sxs-lookup"><span data-stu-id="9e14c-959">Type</span></span>| <span data-ttu-id="9e14c-960">属性</span><span class="sxs-lookup"><span data-stu-id="9e14c-960">Attributes</span></span>| <span data-ttu-id="9e14c-961">说明</span><span class="sxs-lookup"><span data-stu-id="9e14c-961">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="9e14c-962">字符串</span><span class="sxs-lookup"><span data-stu-id="9e14c-962">String</span></span>||<span data-ttu-id="9e14c-p167">要插入的数据。数据不得超过 1,000,000 个字符。如果传入的数据超过 1,000,000 个字符，则会引发 `ArgumentOutOfRange` 异常。</span><span class="sxs-lookup"><span data-stu-id="9e14c-p167">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`| <span data-ttu-id="9e14c-966">Object</span><span class="sxs-lookup"><span data-stu-id="9e14c-966">Object</span></span>| <span data-ttu-id="9e14c-967">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="9e14c-967">&lt;optional&gt;</span></span>|<span data-ttu-id="9e14c-968">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="9e14c-968">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="9e14c-969">对象</span><span class="sxs-lookup"><span data-stu-id="9e14c-969">Object</span></span>| <span data-ttu-id="9e14c-970">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="9e14c-970">&lt;optional&gt;</span></span>|<span data-ttu-id="9e14c-971">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="9e14c-971">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`|[<span data-ttu-id="9e14c-972">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="9e14c-972">Office.CoercionType</span></span>](office.md#coerciontype-string)|<span data-ttu-id="9e14c-973">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="9e14c-973">&lt;optional&gt;</span></span>|<span data-ttu-id="9e14c-974">如果`text`为, 则当前样式应用于 web 上的 Outlook 和桌面客户端。</span><span class="sxs-lookup"><span data-stu-id="9e14c-974">If `text`, the current style is applied in Outlook on the web and desktop clients.</span></span> <span data-ttu-id="9e14c-975">如果该字段是 HTML 编辑器，则仅插入文本数据，即使数据为 HTML。</span><span class="sxs-lookup"><span data-stu-id="9e14c-975">If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="9e14c-976">如果`html`和字段支持 HTML (主题不), 则当前样式应用于 web 上的 outlook, 并且在 outlook 桌面客户端中应用了默认样式。</span><span class="sxs-lookup"><span data-stu-id="9e14c-976">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook on the web and the default style is applied in Outlook desktop clients.</span></span> <span data-ttu-id="9e14c-977">如果该字段是文本字段，则返回 `InvalidDataFormat` 错误。</span><span class="sxs-lookup"><span data-stu-id="9e14c-977">If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="9e14c-978">如果未设置 `coercionType`，则结果取决于该字段：如果该字段是 HTML，则使用 HTML；如果该字段是文本，则使用纯文本。</span><span class="sxs-lookup"><span data-stu-id="9e14c-978">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`| <span data-ttu-id="9e14c-979">function</span><span class="sxs-lookup"><span data-stu-id="9e14c-979">function</span></span>||<span data-ttu-id="9e14c-980">方法完成后，使用单个参数 `callback`（一个 [`asyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `AsyncResult` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="9e14c-980">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="9e14c-981">Requirements</span><span class="sxs-lookup"><span data-stu-id="9e14c-981">Requirements</span></span>

|<span data-ttu-id="9e14c-982">要求</span><span class="sxs-lookup"><span data-stu-id="9e14c-982">Requirement</span></span>| <span data-ttu-id="9e14c-983">值</span><span class="sxs-lookup"><span data-stu-id="9e14c-983">Value</span></span>|
|---|---|
|[<span data-ttu-id="9e14c-984">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="9e14c-984">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9e14c-985">1.2</span><span class="sxs-lookup"><span data-stu-id="9e14c-985">1.2</span></span>|
|[<span data-ttu-id="9e14c-986">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="9e14c-986">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9e14c-987">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="9e14c-987">ReadWriteItem</span></span>|
|[<span data-ttu-id="9e14c-988">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="9e14c-988">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="9e14c-989">撰写</span><span class="sxs-lookup"><span data-stu-id="9e14c-989">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="9e14c-990">示例</span><span class="sxs-lookup"><span data-stu-id="9e14c-990">Example</span></span>

```javascript
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```
