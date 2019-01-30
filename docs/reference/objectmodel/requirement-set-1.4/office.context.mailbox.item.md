---
title: Office.context.mailbox.item-要求设置 1.4
description: ''
ms.date: 12/18/2018
localization_priority: Normal
ms.openlocfilehash: 3a559f71dc4dd5b4cbea901b117e2615acaf196e
ms.sourcegitcommit: d1aa7201820176ed986b9f00bb9c88e055906c77
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 01/23/2019
ms.locfileid: "29388624"
---
# <a name="item"></a><span data-ttu-id="c93e0-102">item</span><span class="sxs-lookup"><span data-stu-id="c93e0-102">item</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmditem"></a><span data-ttu-id="c93e0-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span><span class="sxs-lookup"><span data-stu-id="c93e0-103">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item</span></span>

<span data-ttu-id="c93e0-p101">`item` 命名空间用于访问当前选定的邮件、会议请求或约会。可以通过使用 [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlook14officemailboxenumsitemtype) 属性确定 `item` 的类型。</span><span class="sxs-lookup"><span data-stu-id="c93e0-p101">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlook14officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="c93e0-106">Requirements</span><span class="sxs-lookup"><span data-stu-id="c93e0-106">Requirements</span></span>

|<span data-ttu-id="c93e0-107">要求</span><span class="sxs-lookup"><span data-stu-id="c93e0-107">Requirement</span></span>| <span data-ttu-id="c93e0-108">值</span><span class="sxs-lookup"><span data-stu-id="c93e0-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="c93e0-109">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="c93e0-109">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c93e0-110">1.0</span><span class="sxs-lookup"><span data-stu-id="c93e0-110">1.0</span></span>|
|[<span data-ttu-id="c93e0-111">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="c93e0-111">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c93e0-112">受限</span><span class="sxs-lookup"><span data-stu-id="c93e0-112">Restricted</span></span>|
|[<span data-ttu-id="c93e0-113">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c93e0-113">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c93e0-114">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="c93e0-114">Compose or read</span></span>|

### <a name="example"></a><span data-ttu-id="c93e0-115">示例</span><span class="sxs-lookup"><span data-stu-id="c93e0-115">Example</span></span>

<span data-ttu-id="c93e0-116">以下 JavaScript 代码示例显示了如何访问 Outlook 中当前项目的 `subject` 属性。</span><span class="sxs-lookup"><span data-stu-id="c93e0-116">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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

### <a name="members"></a><span data-ttu-id="c93e0-117">成员</span><span class="sxs-lookup"><span data-stu-id="c93e0-117">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlook14officeattachmentdetails"></a><span data-ttu-id="c93e0-118">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_4/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="c93e0-118">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_4/office.attachmentdetails)></span></span>

<span data-ttu-id="c93e0-p102">获取项目的附件数组。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="c93e0-p102">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="c93e0-121">由于潜在的安全问题，某些类型的文件会受到 Outlook 阻止，并且不会返回。</span><span class="sxs-lookup"><span data-stu-id="c93e0-121">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="c93e0-122">如需了解更多信息，请参阅 [Outlook 中阻止的附件](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519)。</span><span class="sxs-lookup"><span data-stu-id="c93e0-122">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="c93e0-123">类型：</span><span class="sxs-lookup"><span data-stu-id="c93e0-123">Type:</span></span>

*   <span data-ttu-id="c93e0-124">Array.<[AttachmentDetails](/javascript/api/outlook_1_4/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="c93e0-124">Array.<[AttachmentDetails](/javascript/api/outlook_1_4/office.attachmentdetails)></span></span>

##### <a name="requirements"></a><span data-ttu-id="c93e0-125">要求</span><span class="sxs-lookup"><span data-stu-id="c93e0-125">Requirements</span></span>

|<span data-ttu-id="c93e0-126">要求</span><span class="sxs-lookup"><span data-stu-id="c93e0-126">Requirement</span></span>| <span data-ttu-id="c93e0-127">值</span><span class="sxs-lookup"><span data-stu-id="c93e0-127">Value</span></span>|
|---|---|
|[<span data-ttu-id="c93e0-128">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="c93e0-128">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c93e0-129">1.0</span><span class="sxs-lookup"><span data-stu-id="c93e0-129">1.0</span></span>|
|[<span data-ttu-id="c93e0-130">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="c93e0-130">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c93e0-131">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c93e0-131">ReadItem</span></span>|
|[<span data-ttu-id="c93e0-132">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c93e0-132">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c93e0-133">阅读</span><span class="sxs-lookup"><span data-stu-id="c93e0-133">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c93e0-134">示例</span><span class="sxs-lookup"><span data-stu-id="c93e0-134">Example</span></span>

<span data-ttu-id="c93e0-135">以下代码使用当前项目上所有附件的详细信息构成 HTML 字符串。</span><span class="sxs-lookup"><span data-stu-id="c93e0-135">The following code builds an HTML string with details of all attachments on the current item.</span></span>

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

####  <a name="bcc-recipientsjavascriptapioutlook14officerecipients"></a><span data-ttu-id="c93e0-136">bcc :[Recipients](/javascript/api/outlook_1_4/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="c93e0-136">bcc :[Recipients](/javascript/api/outlook_1_4/office.recipients)</span></span>

<span data-ttu-id="c93e0-137">获取对象，该对象提供用于获取或更新邮件的密件抄送 (Bcc) 行的方法。</span><span class="sxs-lookup"><span data-stu-id="c93e0-137">Gets an object that provides methods to get or update the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="c93e0-138">仅限撰写模式。</span><span class="sxs-lookup"><span data-stu-id="c93e0-138">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="c93e0-139">类型：</span><span class="sxs-lookup"><span data-stu-id="c93e0-139">Type:</span></span>

*   [<span data-ttu-id="c93e0-140">收件人</span><span class="sxs-lookup"><span data-stu-id="c93e0-140">Recipients</span></span>](/javascript/api/outlook_1_4/office.recipients)

##### <a name="requirements"></a><span data-ttu-id="c93e0-141">要求</span><span class="sxs-lookup"><span data-stu-id="c93e0-141">Requirements</span></span>

|<span data-ttu-id="c93e0-142">要求</span><span class="sxs-lookup"><span data-stu-id="c93e0-142">Requirement</span></span>| <span data-ttu-id="c93e0-143">值</span><span class="sxs-lookup"><span data-stu-id="c93e0-143">Value</span></span>|
|---|---|
|[<span data-ttu-id="c93e0-144">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="c93e0-144">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c93e0-145">1.1</span><span class="sxs-lookup"><span data-stu-id="c93e0-145">1.1</span></span>|
|[<span data-ttu-id="c93e0-146">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="c93e0-146">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c93e0-147">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c93e0-147">ReadItem</span></span>|
|[<span data-ttu-id="c93e0-148">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c93e0-148">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c93e0-149">撰写</span><span class="sxs-lookup"><span data-stu-id="c93e0-149">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="c93e0-150">示例</span><span class="sxs-lookup"><span data-stu-id="c93e0-150">Example</span></span>

```js
Office.context.mailbox.item.bcc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.bcc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.bcc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfBccRecipients = asyncResult.value;
}
```

####  <a name="body-bodyjavascriptapioutlook14officebody"></a><span data-ttu-id="c93e0-151">body :[Body](/javascript/api/outlook_1_4/office.body)</span><span class="sxs-lookup"><span data-stu-id="c93e0-151">body :[Body](/javascript/api/outlook_1_4/office.body)</span></span>

<span data-ttu-id="c93e0-152">获取一个提供用于处理项目正文的方法的对象。</span><span class="sxs-lookup"><span data-stu-id="c93e0-152">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="c93e0-153">类型：</span><span class="sxs-lookup"><span data-stu-id="c93e0-153">Type:</span></span>

*   [<span data-ttu-id="c93e0-154">Body</span><span class="sxs-lookup"><span data-stu-id="c93e0-154">Body</span></span>](/javascript/api/outlook_1_4/office.body)

##### <a name="requirements"></a><span data-ttu-id="c93e0-155">要求</span><span class="sxs-lookup"><span data-stu-id="c93e0-155">Requirements</span></span>

|<span data-ttu-id="c93e0-156">要求</span><span class="sxs-lookup"><span data-stu-id="c93e0-156">Requirement</span></span>| <span data-ttu-id="c93e0-157">值</span><span class="sxs-lookup"><span data-stu-id="c93e0-157">Value</span></span>|
|---|---|
|[<span data-ttu-id="c93e0-158">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="c93e0-158">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c93e0-159">1.1</span><span class="sxs-lookup"><span data-stu-id="c93e0-159">1.1</span></span>|
|[<span data-ttu-id="c93e0-160">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="c93e0-160">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c93e0-161">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c93e0-161">ReadItem</span></span>|
|[<span data-ttu-id="c93e0-162">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c93e0-162">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c93e0-163">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="c93e0-163">Compose or read</span></span>|

####  <a name="cc-arrayemailaddressdetailsjavascriptapioutlook14officeemailaddressdetailsrecipientsjavascriptapioutlook14officerecipients"></a><span data-ttu-id="c93e0-164">cc :Array.<[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_4/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="c93e0-164">cc :Array.<[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_4/office.recipients)</span></span>

<span data-ttu-id="c93e0-165">提供对邮件的抄送 (Cc) 收件人的访问权限。</span><span class="sxs-lookup"><span data-stu-id="c93e0-165">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="c93e0-166">对象的类型和访问级别取决于当前项目的模式。</span><span class="sxs-lookup"><span data-stu-id="c93e0-166">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="c93e0-167">阅读模式</span><span class="sxs-lookup"><span data-stu-id="c93e0-167">Read mode</span></span>

<span data-ttu-id="c93e0-p106">`cc` 属性返回包含邮件的**抄送**行上所列的每个收件人的 `EmailAddressDetails` 对象的数组。集合上限为 100 个成员。</span><span class="sxs-lookup"><span data-stu-id="c93e0-p106">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="c93e0-170">撰写模式</span><span class="sxs-lookup"><span data-stu-id="c93e0-170">Compose mode</span></span>

<span data-ttu-id="c93e0-171">`cc` 属性返回一个 `Recipients` 对象，该对象提供用于获取或更新邮件的“**抄送**”行上收件人的方法。</span><span class="sxs-lookup"><span data-stu-id="c93e0-171">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="c93e0-172">类型：</span><span class="sxs-lookup"><span data-stu-id="c93e0-172">Type:</span></span>

*   <span data-ttu-id="c93e0-173">Array.<[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_4/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="c93e0-173">Array.<[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_4/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="c93e0-174">要求</span><span class="sxs-lookup"><span data-stu-id="c93e0-174">Requirements</span></span>

|<span data-ttu-id="c93e0-175">要求</span><span class="sxs-lookup"><span data-stu-id="c93e0-175">Requirement</span></span>| <span data-ttu-id="c93e0-176">值</span><span class="sxs-lookup"><span data-stu-id="c93e0-176">Value</span></span>|
|---|---|
|[<span data-ttu-id="c93e0-177">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="c93e0-177">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c93e0-178">1.0</span><span class="sxs-lookup"><span data-stu-id="c93e0-178">1.0</span></span>|
|[<span data-ttu-id="c93e0-179">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="c93e0-179">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c93e0-180">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c93e0-180">ReadItem</span></span>|
|[<span data-ttu-id="c93e0-181">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c93e0-181">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c93e0-182">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="c93e0-182">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="c93e0-183">示例</span><span class="sxs-lookup"><span data-stu-id="c93e0-183">Example</span></span>

```js
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

####  <a name="nullable-conversationid-string"></a><span data-ttu-id="c93e0-184">(nullable) conversationId :String</span><span class="sxs-lookup"><span data-stu-id="c93e0-184">(nullable) conversationId :String</span></span>

<span data-ttu-id="c93e0-185">获取包含特定消息的电子邮件会话的标识符。</span><span class="sxs-lookup"><span data-stu-id="c93e0-185">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="c93e0-p107">如果在阅读窗体或撰写窗体的回复中激活邮件应用程序，则此属性可以获得一个整数值。如果用户随后更改了回复邮件的主题（若发送回复），则该邮件的对话 ID 将改变且之前获取的值将不适用。</span><span class="sxs-lookup"><span data-stu-id="c93e0-p107">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="c93e0-p108">对于撰写窗体的新项目，此属性获得一个 null 值。如果用户设置一个主题并保存该项目，`conversationId` 属性将返回一个值。</span><span class="sxs-lookup"><span data-stu-id="c93e0-p108">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="c93e0-190">类型:</span><span class="sxs-lookup"><span data-stu-id="c93e0-190">Type:</span></span>

*   <span data-ttu-id="c93e0-191">String</span><span class="sxs-lookup"><span data-stu-id="c93e0-191">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="c93e0-192">要求</span><span class="sxs-lookup"><span data-stu-id="c93e0-192">Requirements</span></span>

|<span data-ttu-id="c93e0-193">要求</span><span class="sxs-lookup"><span data-stu-id="c93e0-193">Requirement</span></span>| <span data-ttu-id="c93e0-194">值</span><span class="sxs-lookup"><span data-stu-id="c93e0-194">Value</span></span>|
|---|---|
|[<span data-ttu-id="c93e0-195">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="c93e0-195">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c93e0-196">1.0</span><span class="sxs-lookup"><span data-stu-id="c93e0-196">1.0</span></span>|
|[<span data-ttu-id="c93e0-197">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="c93e0-197">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c93e0-198">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c93e0-198">ReadItem</span></span>|
|[<span data-ttu-id="c93e0-199">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c93e0-199">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c93e0-200">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="c93e0-200">Compose or read</span></span>|

#### <a name="datetimecreated-date"></a><span data-ttu-id="c93e0-201">dateTimeCreated :Date</span><span class="sxs-lookup"><span data-stu-id="c93e0-201">dateTimeCreated :Date</span></span>

<span data-ttu-id="c93e0-p109">获取项目创建的日期和时间。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="c93e0-p109">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="c93e0-204">类型：</span><span class="sxs-lookup"><span data-stu-id="c93e0-204">Type:</span></span>

*   <span data-ttu-id="c93e0-205">日期</span><span class="sxs-lookup"><span data-stu-id="c93e0-205">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="c93e0-206">要求</span><span class="sxs-lookup"><span data-stu-id="c93e0-206">Requirements</span></span>

|<span data-ttu-id="c93e0-207">要求</span><span class="sxs-lookup"><span data-stu-id="c93e0-207">Requirement</span></span>| <span data-ttu-id="c93e0-208">值</span><span class="sxs-lookup"><span data-stu-id="c93e0-208">Value</span></span>|
|---|---|
|[<span data-ttu-id="c93e0-209">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="c93e0-209">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c93e0-210">1.0</span><span class="sxs-lookup"><span data-stu-id="c93e0-210">1.0</span></span>|
|[<span data-ttu-id="c93e0-211">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="c93e0-211">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c93e0-212">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c93e0-212">ReadItem</span></span>|
|[<span data-ttu-id="c93e0-213">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c93e0-213">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c93e0-214">阅读</span><span class="sxs-lookup"><span data-stu-id="c93e0-214">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c93e0-215">示例</span><span class="sxs-lookup"><span data-stu-id="c93e0-215">Example</span></span>

```js
var created = Office.context.mailbox.item.dateTimeCreated;
```

#### <a name="datetimemodified-date"></a><span data-ttu-id="c93e0-216">dateTimeModified :Date</span><span class="sxs-lookup"><span data-stu-id="c93e0-216">dateTimeModified :Date</span></span>

<span data-ttu-id="c93e0-p110">获取项目最近一次修改的日期和时间。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="c93e0-p110">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="c93e0-219">在 Outlook for iOS 或 Outlook for Android 中不支持此成员。</span><span class="sxs-lookup"><span data-stu-id="c93e0-219">This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="type"></a><span data-ttu-id="c93e0-220">类型:</span><span class="sxs-lookup"><span data-stu-id="c93e0-220">Type:</span></span>

*   <span data-ttu-id="c93e0-221">日期</span><span class="sxs-lookup"><span data-stu-id="c93e0-221">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="c93e0-222">要求</span><span class="sxs-lookup"><span data-stu-id="c93e0-222">Requirements</span></span>

|<span data-ttu-id="c93e0-223">要求</span><span class="sxs-lookup"><span data-stu-id="c93e0-223">Requirement</span></span>| <span data-ttu-id="c93e0-224">值</span><span class="sxs-lookup"><span data-stu-id="c93e0-224">Value</span></span>|
|---|---|
|[<span data-ttu-id="c93e0-225">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="c93e0-225">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c93e0-226">1.0</span><span class="sxs-lookup"><span data-stu-id="c93e0-226">1.0</span></span>|
|[<span data-ttu-id="c93e0-227">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="c93e0-227">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c93e0-228">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c93e0-228">ReadItem</span></span>|
|[<span data-ttu-id="c93e0-229">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c93e0-229">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c93e0-230">阅读</span><span class="sxs-lookup"><span data-stu-id="c93e0-230">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c93e0-231">示例</span><span class="sxs-lookup"><span data-stu-id="c93e0-231">Example</span></span>

```js
var modified = Office.context.mailbox.item.dateTimeModified;
```

####  <a name="end-datetimejavascriptapioutlook14officetime"></a><span data-ttu-id="c93e0-232">end :Date|[Time](/javascript/api/outlook_1_4/office.time)</span><span class="sxs-lookup"><span data-stu-id="c93e0-232">end :Date|[Time](/javascript/api/outlook_1_4/office.time)</span></span>

<span data-ttu-id="c93e0-233">获取或设置约会结束的日期和时间。</span><span class="sxs-lookup"><span data-stu-id="c93e0-233">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="c93e0-p111">将 `end` 属性表示为协调世界时 (UTC) 的日期和时间值。可使用 [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook14officelocalclienttime) 方法将 end 属性值转换为客户端的本地日期和时间。</span><span class="sxs-lookup"><span data-stu-id="c93e0-p111">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook14officelocalclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="c93e0-236">阅读模式</span><span class="sxs-lookup"><span data-stu-id="c93e0-236">Read mode</span></span>

<span data-ttu-id="c93e0-237">`end` 属性返回 `Date` 对象。</span><span class="sxs-lookup"><span data-stu-id="c93e0-237">The `end` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="c93e0-238">撰写模式</span><span class="sxs-lookup"><span data-stu-id="c93e0-238">Compose mode</span></span>

<span data-ttu-id="c93e0-239">`end` 属性返回 `Time` 对象。</span><span class="sxs-lookup"><span data-stu-id="c93e0-239">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="c93e0-240">使用 [`Time.setAsync`](/javascript/api/outlook_1_4/office.time#setasync-datetime--options--callback-) 方法设置结束时间时，应使用 [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) 方法将客户端的本地时间转换为服务器的 UTC。</span><span class="sxs-lookup"><span data-stu-id="c93e0-240">When you use the [`Time.setAsync`](/javascript/api/outlook_1_4/office.time#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="c93e0-241">类型：</span><span class="sxs-lookup"><span data-stu-id="c93e0-241">Type:</span></span>

*   <span data-ttu-id="c93e0-242">Date | [Time](/javascript/api/outlook_1_4/office.time)</span><span class="sxs-lookup"><span data-stu-id="c93e0-242">Date | [Time](/javascript/api/outlook_1_4/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="c93e0-243">要求</span><span class="sxs-lookup"><span data-stu-id="c93e0-243">Requirements</span></span>

|<span data-ttu-id="c93e0-244">要求</span><span class="sxs-lookup"><span data-stu-id="c93e0-244">Requirement</span></span>| <span data-ttu-id="c93e0-245">值</span><span class="sxs-lookup"><span data-stu-id="c93e0-245">Value</span></span>|
|---|---|
|[<span data-ttu-id="c93e0-246">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="c93e0-246">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c93e0-247">1.0</span><span class="sxs-lookup"><span data-stu-id="c93e0-247">1.0</span></span>|
|[<span data-ttu-id="c93e0-248">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="c93e0-248">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c93e0-249">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c93e0-249">ReadItem</span></span>|
|[<span data-ttu-id="c93e0-250">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c93e0-250">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c93e0-251">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="c93e0-251">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="c93e0-252">示例</span><span class="sxs-lookup"><span data-stu-id="c93e0-252">Example</span></span>

<span data-ttu-id="c93e0-253">以下示例通过使用 `Time` 对象的 [`setAsync`](/javascript/api/outlook_1_4/office.time#setasync-datetime--options--callback-) 方法，设置撰写模式下约会的结束时间。</span><span class="sxs-lookup"><span data-stu-id="c93e0-253">The following example sets the end time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_4/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

#### <a name="from-emailaddressdetailsjavascriptapioutlook14officeemailaddressdetails"></a><span data-ttu-id="c93e0-254">from :[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="c93e0-254">from :[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)</span></span>

<span data-ttu-id="c93e0-p112">获取邮件发件人的电子邮件地址。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="c93e0-p112">Gets the email address of the sender of a message. Read mode only.</span></span>

<span data-ttu-id="c93e0-p113">`from` 和 [`sender`](#sender-emailaddressdetailsjavascriptapioutlook14officeemailaddressdetails) 属性表示同一个人，邮件由代理人发送的除外。在此情况下，`from` 属性表示代理程序，而 sender 属性表示代理人。</span><span class="sxs-lookup"><span data-stu-id="c93e0-p113">The `from` and [`sender`](#sender-emailaddressdetailsjavascriptapioutlook14officeemailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="c93e0-259">`from` 属性中 `EmailAddressDetails` 对象的 `recipientType` 属性为 `undefined`。</span><span class="sxs-lookup"><span data-stu-id="c93e0-259">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="c93e0-260">类型：</span><span class="sxs-lookup"><span data-stu-id="c93e0-260">Type:</span></span>

*   [<span data-ttu-id="c93e0-261">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="c93e0-261">EmailAddressDetails</span></span>](/javascript/api/outlook_1_4/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="c93e0-262">要求</span><span class="sxs-lookup"><span data-stu-id="c93e0-262">Requirements</span></span>

|<span data-ttu-id="c93e0-263">要求</span><span class="sxs-lookup"><span data-stu-id="c93e0-263">Requirement</span></span>| <span data-ttu-id="c93e0-264">值</span><span class="sxs-lookup"><span data-stu-id="c93e0-264">Value</span></span>|
|---|---|
|[<span data-ttu-id="c93e0-265">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="c93e0-265">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c93e0-266">1.0</span><span class="sxs-lookup"><span data-stu-id="c93e0-266">1.0</span></span>|
|[<span data-ttu-id="c93e0-267">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="c93e0-267">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c93e0-268">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c93e0-268">ReadItem</span></span>|
|[<span data-ttu-id="c93e0-269">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c93e0-269">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c93e0-270">阅读</span><span class="sxs-lookup"><span data-stu-id="c93e0-270">Read</span></span>|

#### <a name="internetmessageid-string"></a><span data-ttu-id="c93e0-271">internetMessageId :String</span><span class="sxs-lookup"><span data-stu-id="c93e0-271">internetMessageId :String</span></span>

<span data-ttu-id="c93e0-p114">获取电子邮件的 Internet 消息标识符。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="c93e0-p114">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="c93e0-274">类型：</span><span class="sxs-lookup"><span data-stu-id="c93e0-274">Type:</span></span>

*   <span data-ttu-id="c93e0-275">String</span><span class="sxs-lookup"><span data-stu-id="c93e0-275">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="c93e0-276">要求</span><span class="sxs-lookup"><span data-stu-id="c93e0-276">Requirements</span></span>

|<span data-ttu-id="c93e0-277">要求</span><span class="sxs-lookup"><span data-stu-id="c93e0-277">Requirement</span></span>| <span data-ttu-id="c93e0-278">值</span><span class="sxs-lookup"><span data-stu-id="c93e0-278">Value</span></span>|
|---|---|
|[<span data-ttu-id="c93e0-279">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="c93e0-279">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c93e0-280">1.0</span><span class="sxs-lookup"><span data-stu-id="c93e0-280">1.0</span></span>|
|[<span data-ttu-id="c93e0-281">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="c93e0-281">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c93e0-282">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c93e0-282">ReadItem</span></span>|
|[<span data-ttu-id="c93e0-283">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c93e0-283">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c93e0-284">阅读</span><span class="sxs-lookup"><span data-stu-id="c93e0-284">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c93e0-285">示例</span><span class="sxs-lookup"><span data-stu-id="c93e0-285">Example</span></span>

```js
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

#### <a name="itemclass-string"></a><span data-ttu-id="c93e0-286">itemClass :String</span><span class="sxs-lookup"><span data-stu-id="c93e0-286">itemClass :String</span></span>

<span data-ttu-id="c93e0-p115">获取选定项目的 Exchange Web 服务项目类。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="c93e0-p115">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="c93e0-p116">`itemClass` 属性指定所选项目的邮件类别。以下是邮件或约会项目的默认邮件类别。</span><span class="sxs-lookup"><span data-stu-id="c93e0-p116">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

| <span data-ttu-id="c93e0-291">类型</span><span class="sxs-lookup"><span data-stu-id="c93e0-291">Type</span></span> | <span data-ttu-id="c93e0-292">说明</span><span class="sxs-lookup"><span data-stu-id="c93e0-292">Description</span></span> | <span data-ttu-id="c93e0-293">项目类</span><span class="sxs-lookup"><span data-stu-id="c93e0-293">item class</span></span> |
| --- | --- | --- |
| <span data-ttu-id="c93e0-294">约会项目</span><span class="sxs-lookup"><span data-stu-id="c93e0-294">Appointment items</span></span> | <span data-ttu-id="c93e0-295">这些是项目类 `IPM.Appointment` 或 `IPM.Appointment.Occurence` 的日历项目。</span><span class="sxs-lookup"><span data-stu-id="c93e0-295">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurence`.</span></span> | `IPM.Appointment`<br />`IPM.Appointment.Occurence` |
| <span data-ttu-id="c93e0-296">邮件项目</span><span class="sxs-lookup"><span data-stu-id="c93e0-296">Message items</span></span> | <span data-ttu-id="c93e0-297">这些项目包括具有默认邮件类别 `IPM.Note` 的电子邮件，以及使用 `IPM.Schedule.Meeting` 作为基础邮件类别的会议请求、响应和取消。</span><span class="sxs-lookup"><span data-stu-id="c93e0-297">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span> | `IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled` |

<span data-ttu-id="c93e0-298">你可以创建用于扩展默认邮件类别的自定义邮件类别，例如，自定义约会邮件类别 `IPM.Appointment.Contoso`。</span><span class="sxs-lookup"><span data-stu-id="c93e0-298">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="c93e0-299">类型：</span><span class="sxs-lookup"><span data-stu-id="c93e0-299">Type:</span></span>

*   <span data-ttu-id="c93e0-300">String</span><span class="sxs-lookup"><span data-stu-id="c93e0-300">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="c93e0-301">要求</span><span class="sxs-lookup"><span data-stu-id="c93e0-301">Requirements</span></span>

|<span data-ttu-id="c93e0-302">要求</span><span class="sxs-lookup"><span data-stu-id="c93e0-302">Requirement</span></span>| <span data-ttu-id="c93e0-303">值</span><span class="sxs-lookup"><span data-stu-id="c93e0-303">Value</span></span>|
|---|---|
|[<span data-ttu-id="c93e0-304">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="c93e0-304">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c93e0-305">1.0</span><span class="sxs-lookup"><span data-stu-id="c93e0-305">1.0</span></span>|
|[<span data-ttu-id="c93e0-306">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="c93e0-306">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c93e0-307">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c93e0-307">ReadItem</span></span>|
|[<span data-ttu-id="c93e0-308">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c93e0-308">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c93e0-309">阅读</span><span class="sxs-lookup"><span data-stu-id="c93e0-309">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c93e0-310">示例</span><span class="sxs-lookup"><span data-stu-id="c93e0-310">Example</span></span>

```js
var itemClass = Office.context.mailbox.item.itemClass;
```

#### <a name="nullable-itemid-string"></a><span data-ttu-id="c93e0-311">(nullable) itemId :String</span><span class="sxs-lookup"><span data-stu-id="c93e0-311">(nullable) itemId :String</span></span>

<span data-ttu-id="c93e0-p117">获取当前项目的 Exchange Web 服务项目标识符。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="c93e0-p117">Gets the Exchange Web Services item identifier for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="c93e0-314">`itemId` 属性返回的标识符与 Exchange Web 服务项目标识符相同。</span><span class="sxs-lookup"><span data-stu-id="c93e0-314">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="c93e0-315">`itemId` 属性与 Outlook 条目 ID 或 Outlook REST API 使用的 ID 不同。</span><span class="sxs-lookup"><span data-stu-id="c93e0-315">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="c93e0-316">使用此值进行 REST API 调用前，应使用 [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string) 对它进行转换。</span><span class="sxs-lookup"><span data-stu-id="c93e0-316">Before making REST API calls using this value, it should be converted using [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string).</span></span> <span data-ttu-id="c93e0-317">有关详细信息，请参阅[从 Outlook 加载项使用 Outlook REST API](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id)。</span><span class="sxs-lookup"><span data-stu-id="c93e0-317">For more details, see [Use the Outlook REST APIs from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

<span data-ttu-id="c93e0-p119">`itemId` 属性在撰写模式下不可用。如果需要项目标识符，[`saveAsync`](#saveasyncoptions-callback) 方法可用于将项目保存到存储，这将在回调函数的 [`AsyncResult.value`](/javascript/api/office/office.asyncresult) 参数中返回项目标识符。</span><span class="sxs-lookup"><span data-stu-id="c93e0-p119">The `itemId` property is not available in compose mode. If an item identifier is required, the [`saveAsync`](#saveasyncoptions-callback) method can be used to save the item to the store, which will return the item identifier in the [`AsyncResult.value`](/javascript/api/office/office.asyncresult) parameter in the callback function.</span></span>

##### <a name="type"></a><span data-ttu-id="c93e0-320">类型:</span><span class="sxs-lookup"><span data-stu-id="c93e0-320">Type:</span></span>

*   <span data-ttu-id="c93e0-321">String</span><span class="sxs-lookup"><span data-stu-id="c93e0-321">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="c93e0-322">要求</span><span class="sxs-lookup"><span data-stu-id="c93e0-322">Requirements</span></span>

|<span data-ttu-id="c93e0-323">要求</span><span class="sxs-lookup"><span data-stu-id="c93e0-323">Requirement</span></span>| <span data-ttu-id="c93e0-324">值</span><span class="sxs-lookup"><span data-stu-id="c93e0-324">Value</span></span>|
|---|---|
|[<span data-ttu-id="c93e0-325">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="c93e0-325">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c93e0-326">1.0</span><span class="sxs-lookup"><span data-stu-id="c93e0-326">1.0</span></span>|
|[<span data-ttu-id="c93e0-327">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="c93e0-327">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c93e0-328">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c93e0-328">ReadItem</span></span>|
|[<span data-ttu-id="c93e0-329">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c93e0-329">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c93e0-330">阅读</span><span class="sxs-lookup"><span data-stu-id="c93e0-330">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c93e0-331">示例</span><span class="sxs-lookup"><span data-stu-id="c93e0-331">Example</span></span>

<span data-ttu-id="c93e0-p120">以下代码检查项目标识符是否存在。如果 `itemId` 属性返回 `null` 或 `undefined`，则将项目保存到存储，并从异步结果中获取项目标识符。</span><span class="sxs-lookup"><span data-stu-id="c93e0-p120">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

```js
var itemId = Office.context.mailbox.item.itemId;
if (itemId === null || itemId == undefined) {
  Office.context.mailbox.item.saveAsync(function(result){
    itemId = result.value;
  });
}
```

####  <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlook14officemailboxenumsitemtype"></a><span data-ttu-id="c93e0-334">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_4/office.mailboxenums.itemtype)</span><span class="sxs-lookup"><span data-stu-id="c93e0-334">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_4/office.mailboxenums.itemtype)</span></span>

<span data-ttu-id="c93e0-335">获取实例表示的项的类型。</span><span class="sxs-lookup"><span data-stu-id="c93e0-335">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="c93e0-336">`itemType` 属性返回其中一个 `ItemType` 枚举值，指示 `item` 对象实例是邮件还是约会。</span><span class="sxs-lookup"><span data-stu-id="c93e0-336">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="c93e0-337">类型：</span><span class="sxs-lookup"><span data-stu-id="c93e0-337">Type:</span></span>

*   [<span data-ttu-id="c93e0-338">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="c93e0-338">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook_1_4/office.mailboxenums.itemtype)

##### <a name="requirements"></a><span data-ttu-id="c93e0-339">要求</span><span class="sxs-lookup"><span data-stu-id="c93e0-339">Requirements</span></span>

|<span data-ttu-id="c93e0-340">要求</span><span class="sxs-lookup"><span data-stu-id="c93e0-340">Requirement</span></span>| <span data-ttu-id="c93e0-341">值</span><span class="sxs-lookup"><span data-stu-id="c93e0-341">Value</span></span>|
|---|---|
|[<span data-ttu-id="c93e0-342">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="c93e0-342">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c93e0-343">1.0</span><span class="sxs-lookup"><span data-stu-id="c93e0-343">1.0</span></span>|
|[<span data-ttu-id="c93e0-344">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="c93e0-344">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c93e0-345">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c93e0-345">ReadItem</span></span>|
|[<span data-ttu-id="c93e0-346">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c93e0-346">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c93e0-347">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="c93e0-347">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="c93e0-348">示例</span><span class="sxs-lookup"><span data-stu-id="c93e0-348">Example</span></span>

```js
if (Office.context.mailbox.item.itemType == Office.MailboxEnums.ItemType.Message)
  // do something
else
  // do something else
```

####  <a name="location-stringlocationjavascriptapioutlook14officelocation"></a><span data-ttu-id="c93e0-349">location :String|[Location](/javascript/api/outlook_1_4/office.location)</span><span class="sxs-lookup"><span data-stu-id="c93e0-349">location :String|[Location](/javascript/api/outlook_1_4/office.location)</span></span>

<span data-ttu-id="c93e0-350">获取或设置约会的位置。</span><span class="sxs-lookup"><span data-stu-id="c93e0-350">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="c93e0-351">阅读模式</span><span class="sxs-lookup"><span data-stu-id="c93e0-351">Read mode</span></span>

<span data-ttu-id="c93e0-352">`location` 属性返回一个包含约会位置的字符串。</span><span class="sxs-lookup"><span data-stu-id="c93e0-352">The `location` property returns a string that contains the location of the appointment.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="c93e0-353">撰写模式</span><span class="sxs-lookup"><span data-stu-id="c93e0-353">Compose mode</span></span>

<span data-ttu-id="c93e0-354">`location` 属性返回一个 `Location` 对象，该对象提供用于获取和设置约会位置的方法。</span><span class="sxs-lookup"><span data-stu-id="c93e0-354">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="c93e0-355">类型：</span><span class="sxs-lookup"><span data-stu-id="c93e0-355">Type:</span></span>

*   <span data-ttu-id="c93e0-356">String | [Location](/javascript/api/outlook_1_4/office.location)</span><span class="sxs-lookup"><span data-stu-id="c93e0-356">String | [Location](/javascript/api/outlook_1_4/office.location)</span></span>

##### <a name="requirements"></a><span data-ttu-id="c93e0-357">要求</span><span class="sxs-lookup"><span data-stu-id="c93e0-357">Requirements</span></span>

|<span data-ttu-id="c93e0-358">要求</span><span class="sxs-lookup"><span data-stu-id="c93e0-358">Requirement</span></span>| <span data-ttu-id="c93e0-359">值</span><span class="sxs-lookup"><span data-stu-id="c93e0-359">Value</span></span>|
|---|---|
|[<span data-ttu-id="c93e0-360">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="c93e0-360">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c93e0-361">1.0</span><span class="sxs-lookup"><span data-stu-id="c93e0-361">1.0</span></span>|
|[<span data-ttu-id="c93e0-362">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="c93e0-362">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c93e0-363">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c93e0-363">ReadItem</span></span>|
|[<span data-ttu-id="c93e0-364">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c93e0-364">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c93e0-365">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="c93e0-365">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="c93e0-366">示例</span><span class="sxs-lookup"><span data-stu-id="c93e0-366">Example</span></span>

```js
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

#### <a name="normalizedsubject-string"></a><span data-ttu-id="c93e0-367">normalizedSubject :String</span><span class="sxs-lookup"><span data-stu-id="c93e0-367">normalizedSubject :String</span></span>

<span data-ttu-id="c93e0-p121">获取删除了所有前缀（包括 `RE:` 和 `FWD:`）的项目主题。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="c93e0-p121">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="c93e0-p122">normalizedSubject 属性获取包含由电子邮件程序添加的任何标准前缀（如 `RE:` 和 `FW:`）的项目主题。若要获取包含完整前缀的项目主题，请使用 [`subject`](#subject-stringsubjectjavascriptapioutlook14officesubject) 属性。</span><span class="sxs-lookup"><span data-stu-id="c93e0-p122">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubjectjavascriptapioutlook14officesubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="c93e0-372">类型:</span><span class="sxs-lookup"><span data-stu-id="c93e0-372">Type:</span></span>

*   <span data-ttu-id="c93e0-373">String</span><span class="sxs-lookup"><span data-stu-id="c93e0-373">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="c93e0-374">要求</span><span class="sxs-lookup"><span data-stu-id="c93e0-374">Requirements</span></span>

|<span data-ttu-id="c93e0-375">要求</span><span class="sxs-lookup"><span data-stu-id="c93e0-375">Requirement</span></span>| <span data-ttu-id="c93e0-376">值</span><span class="sxs-lookup"><span data-stu-id="c93e0-376">Value</span></span>|
|---|---|
|[<span data-ttu-id="c93e0-377">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="c93e0-377">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c93e0-378">1.0</span><span class="sxs-lookup"><span data-stu-id="c93e0-378">1.0</span></span>|
|[<span data-ttu-id="c93e0-379">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="c93e0-379">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c93e0-380">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c93e0-380">ReadItem</span></span>|
|[<span data-ttu-id="c93e0-381">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c93e0-381">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c93e0-382">阅读</span><span class="sxs-lookup"><span data-stu-id="c93e0-382">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c93e0-383">示例</span><span class="sxs-lookup"><span data-stu-id="c93e0-383">Example</span></span>

```js
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
```

####  <a name="notificationmessages-notificationmessagesjavascriptapioutlook14officenotificationmessages"></a><span data-ttu-id="c93e0-384">notificationMessages :[NotificationMessages](/javascript/api/outlook_1_4/office.notificationmessages)</span><span class="sxs-lookup"><span data-stu-id="c93e0-384">notificationMessages :[NotificationMessages](/javascript/api/outlook_1_4/office.notificationmessages)</span></span>

<span data-ttu-id="c93e0-385">获取项目的通知邮件。</span><span class="sxs-lookup"><span data-stu-id="c93e0-385">Gets the notification messages for an item.</span></span>

##### <a name="type"></a><span data-ttu-id="c93e0-386">类型：</span><span class="sxs-lookup"><span data-stu-id="c93e0-386">Type:</span></span>

*   [<span data-ttu-id="c93e0-387">NotificationMessages</span><span class="sxs-lookup"><span data-stu-id="c93e0-387">NotificationMessages</span></span>](/javascript/api/outlook_1_4/office.notificationmessages)

##### <a name="requirements"></a><span data-ttu-id="c93e0-388">要求</span><span class="sxs-lookup"><span data-stu-id="c93e0-388">Requirements</span></span>

|<span data-ttu-id="c93e0-389">要求</span><span class="sxs-lookup"><span data-stu-id="c93e0-389">Requirement</span></span>| <span data-ttu-id="c93e0-390">值</span><span class="sxs-lookup"><span data-stu-id="c93e0-390">Value</span></span>|
|---|---|
|[<span data-ttu-id="c93e0-391">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="c93e0-391">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c93e0-392">1.3</span><span class="sxs-lookup"><span data-stu-id="c93e0-392">1.3</span></span>|
|[<span data-ttu-id="c93e0-393">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="c93e0-393">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c93e0-394">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c93e0-394">ReadItem</span></span>|
|[<span data-ttu-id="c93e0-395">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c93e0-395">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c93e0-396">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="c93e0-396">Compose or read</span></span>|

####  <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlook14officeemailaddressdetailsrecipientsjavascriptapioutlook14officerecipients"></a><span data-ttu-id="c93e0-397">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_4/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="c93e0-397">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_4/office.recipients)</span></span>

<span data-ttu-id="c93e0-398">提供对事件的可选与会者的访问权限。</span><span class="sxs-lookup"><span data-stu-id="c93e0-398">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="c93e0-399">对象的类型和访问级别取决于当前项目的模式。</span><span class="sxs-lookup"><span data-stu-id="c93e0-399">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="c93e0-400">阅读模式</span><span class="sxs-lookup"><span data-stu-id="c93e0-400">Read mode</span></span>

<span data-ttu-id="c93e0-401">`optionalAttendees` 属性返回一个数组，其中包含每个可选与会者的 `EmailAddressDetails` 对象。</span><span class="sxs-lookup"><span data-stu-id="c93e0-401">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="c93e0-402">撰写模式</span><span class="sxs-lookup"><span data-stu-id="c93e0-402">Compose mode</span></span>

<span data-ttu-id="c93e0-403">`optionalAttendees` 属性返回一个 `Recipients` 对象，该对象提供用于获取或更新可选与会者的方法。</span><span class="sxs-lookup"><span data-stu-id="c93e0-403">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="c93e0-404">类型：</span><span class="sxs-lookup"><span data-stu-id="c93e0-404">Type:</span></span>

*   <span data-ttu-id="c93e0-405">Array.<[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_4/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="c93e0-405">Array.<[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_4/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="c93e0-406">要求</span><span class="sxs-lookup"><span data-stu-id="c93e0-406">Requirements</span></span>

|<span data-ttu-id="c93e0-407">要求</span><span class="sxs-lookup"><span data-stu-id="c93e0-407">Requirement</span></span>| <span data-ttu-id="c93e0-408">值</span><span class="sxs-lookup"><span data-stu-id="c93e0-408">Value</span></span>|
|---|---|
|[<span data-ttu-id="c93e0-409">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="c93e0-409">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c93e0-410">1.0</span><span class="sxs-lookup"><span data-stu-id="c93e0-410">1.0</span></span>|
|[<span data-ttu-id="c93e0-411">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="c93e0-411">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c93e0-412">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c93e0-412">ReadItem</span></span>|
|[<span data-ttu-id="c93e0-413">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c93e0-413">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c93e0-414">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="c93e0-414">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="c93e0-415">示例</span><span class="sxs-lookup"><span data-stu-id="c93e0-415">Example</span></span>

```js
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

#### <a name="organizer-emailaddressdetailsjavascriptapioutlook14officeemailaddressdetails"></a><span data-ttu-id="c93e0-416">organizer :[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="c93e0-416">organizer :[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)</span></span>

<span data-ttu-id="c93e0-p124">获取指定会议的会议组织者的电子邮件地址。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="c93e0-p124">Gets the email address of the meeting organizer for a specified meeting. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="c93e0-419">类型:</span><span class="sxs-lookup"><span data-stu-id="c93e0-419">Type:</span></span>

*   [<span data-ttu-id="c93e0-420">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="c93e0-420">EmailAddressDetails</span></span>](/javascript/api/outlook_1_4/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="c93e0-421">要求</span><span class="sxs-lookup"><span data-stu-id="c93e0-421">Requirements</span></span>

|<span data-ttu-id="c93e0-422">要求</span><span class="sxs-lookup"><span data-stu-id="c93e0-422">Requirement</span></span>| <span data-ttu-id="c93e0-423">值</span><span class="sxs-lookup"><span data-stu-id="c93e0-423">Value</span></span>|
|---|---|
|[<span data-ttu-id="c93e0-424">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="c93e0-424">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c93e0-425">1.0</span><span class="sxs-lookup"><span data-stu-id="c93e0-425">1.0</span></span>|
|[<span data-ttu-id="c93e0-426">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="c93e0-426">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c93e0-427">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c93e0-427">ReadItem</span></span>|
|[<span data-ttu-id="c93e0-428">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c93e0-428">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c93e0-429">阅读</span><span class="sxs-lookup"><span data-stu-id="c93e0-429">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c93e0-430">示例</span><span class="sxs-lookup"><span data-stu-id="c93e0-430">Example</span></span>

```js
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
```

####  <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlook14officeemailaddressdetailsrecipientsjavascriptapioutlook14officerecipients"></a><span data-ttu-id="c93e0-431">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_4/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="c93e0-431">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_4/office.recipients)</span></span>

<span data-ttu-id="c93e0-432">提供对事件的必需与会者的访问权限。</span><span class="sxs-lookup"><span data-stu-id="c93e0-432">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="c93e0-433">对象的类型和访问级别取决于当前项目的模式。</span><span class="sxs-lookup"><span data-stu-id="c93e0-433">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="c93e0-434">阅读模式</span><span class="sxs-lookup"><span data-stu-id="c93e0-434">Read mode</span></span>

<span data-ttu-id="c93e0-435">`requiredAttendees` 属性返回一个数组，其中包含每个必需与会者的 `EmailAddressDetails` 对象。</span><span class="sxs-lookup"><span data-stu-id="c93e0-435">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="c93e0-436">撰写模式</span><span class="sxs-lookup"><span data-stu-id="c93e0-436">Compose mode</span></span>

<span data-ttu-id="c93e0-437">`requiredAttendees` 属性返回一个 `Recipients` 对象，该对象提供用于获取或更新必需与会者的方法。</span><span class="sxs-lookup"><span data-stu-id="c93e0-437">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="c93e0-438">类型：</span><span class="sxs-lookup"><span data-stu-id="c93e0-438">Type:</span></span>

*   <span data-ttu-id="c93e0-439">Array.<[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_4/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="c93e0-439">Array.<[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_4/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="c93e0-440">要求</span><span class="sxs-lookup"><span data-stu-id="c93e0-440">Requirements</span></span>

|<span data-ttu-id="c93e0-441">要求</span><span class="sxs-lookup"><span data-stu-id="c93e0-441">Requirement</span></span>| <span data-ttu-id="c93e0-442">值</span><span class="sxs-lookup"><span data-stu-id="c93e0-442">Value</span></span>|
|---|---|
|[<span data-ttu-id="c93e0-443">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="c93e0-443">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c93e0-444">1.0</span><span class="sxs-lookup"><span data-stu-id="c93e0-444">1.0</span></span>|
|[<span data-ttu-id="c93e0-445">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="c93e0-445">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c93e0-446">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c93e0-446">ReadItem</span></span>|
|[<span data-ttu-id="c93e0-447">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c93e0-447">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c93e0-448">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="c93e0-448">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="c93e0-449">示例</span><span class="sxs-lookup"><span data-stu-id="c93e0-449">Example</span></span>

```js
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
}
```

#### <a name="sender-emailaddressdetailsjavascriptapioutlook14officeemailaddressdetails"></a><span data-ttu-id="c93e0-450">sender :[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="c93e0-450">sender :[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)</span></span>

<span data-ttu-id="c93e0-p126">获取电子邮件发件人的电子邮件地址。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="c93e0-p126">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="c93e0-p127">[`from`](#from-emailaddressdetailsjavascriptapioutlook14officeemailaddressdetails) 和 `sender` 属性表示同一个人，邮件由代理人发送的除外。在此情况下，`from` 属性表示代理程序，而 sender 属性表示代理人。</span><span class="sxs-lookup"><span data-stu-id="c93e0-p127">The [`from`](#from-emailaddressdetailsjavascriptapioutlook14officeemailaddressdetails) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="c93e0-455">`sender` 属性中 `EmailAddressDetails` 对象的 `recipientType` 属性为 `undefined`。</span><span class="sxs-lookup"><span data-stu-id="c93e0-455">The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="c93e0-456">类型：</span><span class="sxs-lookup"><span data-stu-id="c93e0-456">Type:</span></span>

*   [<span data-ttu-id="c93e0-457">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="c93e0-457">EmailAddressDetails</span></span>](/javascript/api/outlook_1_4/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="c93e0-458">要求</span><span class="sxs-lookup"><span data-stu-id="c93e0-458">Requirements</span></span>

|<span data-ttu-id="c93e0-459">要求</span><span class="sxs-lookup"><span data-stu-id="c93e0-459">Requirement</span></span>| <span data-ttu-id="c93e0-460">值</span><span class="sxs-lookup"><span data-stu-id="c93e0-460">Value</span></span>|
|---|---|
|[<span data-ttu-id="c93e0-461">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="c93e0-461">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c93e0-462">1.0</span><span class="sxs-lookup"><span data-stu-id="c93e0-462">1.0</span></span>|
|[<span data-ttu-id="c93e0-463">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="c93e0-463">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c93e0-464">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c93e0-464">ReadItem</span></span>|
|[<span data-ttu-id="c93e0-465">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c93e0-465">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c93e0-466">阅读</span><span class="sxs-lookup"><span data-stu-id="c93e0-466">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c93e0-467">示例</span><span class="sxs-lookup"><span data-stu-id="c93e0-467">Example</span></span>

```js
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
```

####  <a name="start-datetimejavascriptapioutlook14officetime"></a><span data-ttu-id="c93e0-468">start :Date|[Time](/javascript/api/outlook_1_4/office.time)</span><span class="sxs-lookup"><span data-stu-id="c93e0-468">start :Date|[Time](/javascript/api/outlook_1_4/office.time)</span></span>

<span data-ttu-id="c93e0-469">获取或设置约会开始的日期和时间。</span><span class="sxs-lookup"><span data-stu-id="c93e0-469">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="c93e0-p128">将 `start` 属性表示为协调世界时 (UTC) 的日期和时间值。可使用 [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook14officelocalclienttime) 方法将该值转换为客户端的本地日期和时间。</span><span class="sxs-lookup"><span data-stu-id="c93e0-p128">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook14officelocalclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="c93e0-472">阅读模式</span><span class="sxs-lookup"><span data-stu-id="c93e0-472">Read mode</span></span>

<span data-ttu-id="c93e0-473">`start` 属性返回 `Date` 对象。</span><span class="sxs-lookup"><span data-stu-id="c93e0-473">The `start` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="c93e0-474">撰写模式</span><span class="sxs-lookup"><span data-stu-id="c93e0-474">Compose mode</span></span>

<span data-ttu-id="c93e0-475">`start` 属性返回 `Time` 对象。</span><span class="sxs-lookup"><span data-stu-id="c93e0-475">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="c93e0-476">使用 [`Time.setAsync`](/javascript/api/outlook_1_4/office.time#setasync-datetime--options--callback-) 方法设置开始时间时，应使用 [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) 方法将客户端的本地时间转换为服务器的 UTC。</span><span class="sxs-lookup"><span data-stu-id="c93e0-476">When you use the [`Time.setAsync`](/javascript/api/outlook_1_4/office.time#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="c93e0-477">类型：</span><span class="sxs-lookup"><span data-stu-id="c93e0-477">Type:</span></span>

*   <span data-ttu-id="c93e0-478">Date | [Time](/javascript/api/outlook_1_4/office.time)</span><span class="sxs-lookup"><span data-stu-id="c93e0-478">Date | [Time](/javascript/api/outlook_1_4/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="c93e0-479">要求</span><span class="sxs-lookup"><span data-stu-id="c93e0-479">Requirements</span></span>

|<span data-ttu-id="c93e0-480">要求</span><span class="sxs-lookup"><span data-stu-id="c93e0-480">Requirement</span></span>| <span data-ttu-id="c93e0-481">值</span><span class="sxs-lookup"><span data-stu-id="c93e0-481">Value</span></span>|
|---|---|
|[<span data-ttu-id="c93e0-482">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="c93e0-482">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c93e0-483">1.0</span><span class="sxs-lookup"><span data-stu-id="c93e0-483">1.0</span></span>|
|[<span data-ttu-id="c93e0-484">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="c93e0-484">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c93e0-485">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c93e0-485">ReadItem</span></span>|
|[<span data-ttu-id="c93e0-486">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c93e0-486">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c93e0-487">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="c93e0-487">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="c93e0-488">示例</span><span class="sxs-lookup"><span data-stu-id="c93e0-488">Example</span></span>

<span data-ttu-id="c93e0-489">以下示例通过使用 `Time` 对象的 [`setAsync`](/javascript/api/outlook_1_4/office.time#setasync-datetime--options--callback-) 方法，设置撰写模式下约会的开始时间。</span><span class="sxs-lookup"><span data-stu-id="c93e0-489">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_4/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

####  <a name="subject-stringsubjectjavascriptapioutlook14officesubject"></a><span data-ttu-id="c93e0-490">subject :String|[Subject](/javascript/api/outlook_1_4/office.subject)</span><span class="sxs-lookup"><span data-stu-id="c93e0-490">subject :String|[Subject](/javascript/api/outlook_1_4/office.subject)</span></span>

<span data-ttu-id="c93e0-491">获取或设置显示在项目的主题字段中的说明。</span><span class="sxs-lookup"><span data-stu-id="c93e0-491">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="c93e0-492">`subject` 属性获取或设置由电子邮件服务器发送项目时的整个主题。</span><span class="sxs-lookup"><span data-stu-id="c93e0-492">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="c93e0-493">阅读模式</span><span class="sxs-lookup"><span data-stu-id="c93e0-493">Read mode</span></span>

<span data-ttu-id="c93e0-p129">`subject` 属性返回一个字符串。使用 [`normalizedSubject`](#normalizedsubject-string) 属性获取不带任何前导前缀（如 `RE:` 和 `FW:`）的主题。</span><span class="sxs-lookup"><span data-stu-id="c93e0-p129">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

```js
var subject = Office.context.mailbox.item.subject;
```

##### <a name="compose-mode"></a><span data-ttu-id="c93e0-496">撰写模式</span><span class="sxs-lookup"><span data-stu-id="c93e0-496">Compose mode</span></span>

<span data-ttu-id="c93e0-497">`subject` 属性返回一个 `Subject` 对象，该对象提供用于获取和设置主题的方法。</span><span class="sxs-lookup"><span data-stu-id="c93e0-497">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```js
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="c93e0-498">类型：</span><span class="sxs-lookup"><span data-stu-id="c93e0-498">Type:</span></span>

*   <span data-ttu-id="c93e0-499">String | [Subject](/javascript/api/outlook_1_4/office.subject)</span><span class="sxs-lookup"><span data-stu-id="c93e0-499">String | [Subject](/javascript/api/outlook_1_4/office.subject)</span></span>

##### <a name="requirements"></a><span data-ttu-id="c93e0-500">要求</span><span class="sxs-lookup"><span data-stu-id="c93e0-500">Requirements</span></span>

|<span data-ttu-id="c93e0-501">要求</span><span class="sxs-lookup"><span data-stu-id="c93e0-501">Requirement</span></span>| <span data-ttu-id="c93e0-502">值</span><span class="sxs-lookup"><span data-stu-id="c93e0-502">Value</span></span>|
|---|---|
|[<span data-ttu-id="c93e0-503">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="c93e0-503">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c93e0-504">1.0</span><span class="sxs-lookup"><span data-stu-id="c93e0-504">1.0</span></span>|
|[<span data-ttu-id="c93e0-505">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="c93e0-505">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c93e0-506">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c93e0-506">ReadItem</span></span>|
|[<span data-ttu-id="c93e0-507">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c93e0-507">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c93e0-508">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="c93e0-508">Compose or read</span></span>|

####  <a name="to-arrayemailaddressdetailsjavascriptapioutlook14officeemailaddressdetailsrecipientsjavascriptapioutlook14officerecipients"></a><span data-ttu-id="c93e0-509">to :Array.<[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_4/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="c93e0-509">to :Array.<[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_4/office.recipients)</span></span>

<span data-ttu-id="c93e0-510">提供对邮件的“**收件人**”行上的收件人的访问权限。</span><span class="sxs-lookup"><span data-stu-id="c93e0-510">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="c93e0-511">对象的类型和访问级别取决于当前项目的模式。</span><span class="sxs-lookup"><span data-stu-id="c93e0-511">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="c93e0-512">阅读模式</span><span class="sxs-lookup"><span data-stu-id="c93e0-512">Read mode</span></span>

<span data-ttu-id="c93e0-p131">`to` 属性返回包含邮件的**收件人**行上所列的每个收件人的 `EmailAddressDetails` 对象的数组。集合上限为 100 个成员。</span><span class="sxs-lookup"><span data-stu-id="c93e0-p131">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="c93e0-515">撰写模式</span><span class="sxs-lookup"><span data-stu-id="c93e0-515">Compose mode</span></span>

<span data-ttu-id="c93e0-516">`to` 属性返回一个 `Recipients` 对象，该对象提供用于获取或更新邮件的“**收件人**”行上收件人的方法。</span><span class="sxs-lookup"><span data-stu-id="c93e0-516">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="c93e0-517">类型：</span><span class="sxs-lookup"><span data-stu-id="c93e0-517">Type:</span></span>

*   <span data-ttu-id="c93e0-518">Array.<[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_4/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="c93e0-518">Array.<[EmailAddressDetails](/javascript/api/outlook_1_4/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_4/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="c93e0-519">要求</span><span class="sxs-lookup"><span data-stu-id="c93e0-519">Requirements</span></span>

|<span data-ttu-id="c93e0-520">要求</span><span class="sxs-lookup"><span data-stu-id="c93e0-520">Requirement</span></span>| <span data-ttu-id="c93e0-521">值</span><span class="sxs-lookup"><span data-stu-id="c93e0-521">Value</span></span>|
|---|---|
|[<span data-ttu-id="c93e0-522">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="c93e0-522">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c93e0-523">1.0</span><span class="sxs-lookup"><span data-stu-id="c93e0-523">1.0</span></span>|
|[<span data-ttu-id="c93e0-524">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="c93e0-524">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c93e0-525">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c93e0-525">ReadItem</span></span>|
|[<span data-ttu-id="c93e0-526">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c93e0-526">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c93e0-527">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="c93e0-527">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="c93e0-528">示例</span><span class="sxs-lookup"><span data-stu-id="c93e0-528">Example</span></span>

```js
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

### <a name="methods"></a><span data-ttu-id="c93e0-529">方法</span><span class="sxs-lookup"><span data-stu-id="c93e0-529">Methods</span></span>

####  <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="c93e0-530">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="c93e0-530">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="c93e0-531">将文件作为附件添加到邮件或约会。</span><span class="sxs-lookup"><span data-stu-id="c93e0-531">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="c93e0-532">`addFileAttachmentAsync` 方法在指定的 URI 上载文件并将其附加到撰写窗体中的项目。</span><span class="sxs-lookup"><span data-stu-id="c93e0-532">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="c93e0-533">你随后可以将该标识符与 [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) 方法一同使用，以删除同一个会话中的附件。</span><span class="sxs-lookup"><span data-stu-id="c93e0-533">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c93e0-534">参数：</span><span class="sxs-lookup"><span data-stu-id="c93e0-534">Parameters:</span></span>

|<span data-ttu-id="c93e0-535">名称</span><span class="sxs-lookup"><span data-stu-id="c93e0-535">Name</span></span>| <span data-ttu-id="c93e0-536">类型</span><span class="sxs-lookup"><span data-stu-id="c93e0-536">Type</span></span>| <span data-ttu-id="c93e0-537">属性</span><span class="sxs-lookup"><span data-stu-id="c93e0-537">Attributes</span></span>| <span data-ttu-id="c93e0-538">说明</span><span class="sxs-lookup"><span data-stu-id="c93e0-538">Description</span></span>|
|---|---|---|---|
|`uri`| <span data-ttu-id="c93e0-539">字符串</span><span class="sxs-lookup"><span data-stu-id="c93e0-539">String</span></span>||<span data-ttu-id="c93e0-p132">提供附加到邮件或约会的文件的位置的 URI。最大长度为 2048 个字符。</span><span class="sxs-lookup"><span data-stu-id="c93e0-p132">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="c93e0-542">字符串</span><span class="sxs-lookup"><span data-stu-id="c93e0-542">String</span></span>||<span data-ttu-id="c93e0-p133">在附件上载过程中显示的附件名称。最大长度为 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="c93e0-p133">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="c93e0-545">对象</span><span class="sxs-lookup"><span data-stu-id="c93e0-545">Object</span></span>| <span data-ttu-id="c93e0-546">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="c93e0-546">&lt;optional&gt;</span></span>|<span data-ttu-id="c93e0-547">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="c93e0-547">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="c93e0-548">对象</span><span class="sxs-lookup"><span data-stu-id="c93e0-548">Object</span></span>| <span data-ttu-id="c93e0-549">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="c93e0-549">&lt;optional&gt;</span></span>|<span data-ttu-id="c93e0-550">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="c93e0-550">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="c93e0-551">函数</span><span class="sxs-lookup"><span data-stu-id="c93e0-551">function</span></span>| <span data-ttu-id="c93e0-552">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="c93e0-552">&lt;optional&gt;</span></span>|<span data-ttu-id="c93e0-553">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="c93e0-553">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="c93e0-554">如果成功，附件标识符将在 `asyncResult.value` 属性中提供。</span><span class="sxs-lookup"><span data-stu-id="c93e0-554">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="c93e0-555">如果上传附件失败，`asyncResult` 对象将包含一个提供错误说明的 `Error` 对象。</span><span class="sxs-lookup"><span data-stu-id="c93e0-555">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="c93e0-556">错误</span><span class="sxs-lookup"><span data-stu-id="c93e0-556">Errors</span></span>

| <span data-ttu-id="c93e0-557">错误代码</span><span class="sxs-lookup"><span data-stu-id="c93e0-557">Error code</span></span> | <span data-ttu-id="c93e0-558">说明</span><span class="sxs-lookup"><span data-stu-id="c93e0-558">Description</span></span> |
|------------|-------------|
| `AttachmentSizeExceeded` | <span data-ttu-id="c93e0-559">附件大小超过了允许的大小。</span><span class="sxs-lookup"><span data-stu-id="c93e0-559">The attachment is larger than allowed.</span></span> |
| `FileTypeNotSupported` | <span data-ttu-id="c93e0-560">该附件的扩展名不是允许的扩展名。</span><span class="sxs-lookup"><span data-stu-id="c93e0-560">The attachment has an extension that is not allowed.</span></span> |
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="c93e0-561">邮件或约会具有的附件过多。</span><span class="sxs-lookup"><span data-stu-id="c93e0-561">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="c93e0-562">要求</span><span class="sxs-lookup"><span data-stu-id="c93e0-562">Requirements</span></span>

|<span data-ttu-id="c93e0-563">要求</span><span class="sxs-lookup"><span data-stu-id="c93e0-563">Requirement</span></span>| <span data-ttu-id="c93e0-564">值</span><span class="sxs-lookup"><span data-stu-id="c93e0-564">Value</span></span>|
|---|---|
|[<span data-ttu-id="c93e0-565">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="c93e0-565">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c93e0-566">1.1</span><span class="sxs-lookup"><span data-stu-id="c93e0-566">1.1</span></span>|
|[<span data-ttu-id="c93e0-567">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="c93e0-567">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c93e0-568">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="c93e0-568">ReadWriteItem</span></span>|
|[<span data-ttu-id="c93e0-569">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c93e0-569">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c93e0-570">撰写</span><span class="sxs-lookup"><span data-stu-id="c93e0-570">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="c93e0-571">示例</span><span class="sxs-lookup"><span data-stu-id="c93e0-571">Example</span></span>

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

####  <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="c93e0-572">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="c93e0-572">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="c93e0-573">将 Exchange 项目（如邮件）作为附件添加到邮件或约会。</span><span class="sxs-lookup"><span data-stu-id="c93e0-573">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="c93e0-p134">`addItemAttachmentAsync` 方法将包含指定 Exchange 标识符的项目附加到撰写窗体中的项目。如果指定一个回调方法，此方法使用 `asyncResult` 参数调用，该参数包含一个附件标识符或代码，指示附加项目过程中出现的任何错误。可以使用 `options` 参数将状态信息传递给回调方法（如果需要）。</span><span class="sxs-lookup"><span data-stu-id="c93e0-p134">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="c93e0-577">随后可以将该标识符与 [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) 方法一同使用，以删除同一个会话中的附件。</span><span class="sxs-lookup"><span data-stu-id="c93e0-577">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="c93e0-578">如果 Office 加载项在 Outlook Web App 中运行，则 `addItemAttachmentAsync` 方法可以将项目附加到项目（正在编辑的项目除外）中；然而，不支持也不建议这样做。</span><span class="sxs-lookup"><span data-stu-id="c93e0-578">If your Office Add-in is running in Outlook Web App, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c93e0-579">参数：</span><span class="sxs-lookup"><span data-stu-id="c93e0-579">Parameters:</span></span>

|<span data-ttu-id="c93e0-580">名称</span><span class="sxs-lookup"><span data-stu-id="c93e0-580">Name</span></span>| <span data-ttu-id="c93e0-581">类型</span><span class="sxs-lookup"><span data-stu-id="c93e0-581">Type</span></span>| <span data-ttu-id="c93e0-582">属性</span><span class="sxs-lookup"><span data-stu-id="c93e0-582">Attributes</span></span>| <span data-ttu-id="c93e0-583">说明</span><span class="sxs-lookup"><span data-stu-id="c93e0-583">Description</span></span>|
|---|---|---|---|
|`itemId`| <span data-ttu-id="c93e0-584">字符串</span><span class="sxs-lookup"><span data-stu-id="c93e0-584">String</span></span>||<span data-ttu-id="c93e0-p135">要附加的项目的 Exchange 标识符。最大长度为 100 个字符。</span><span class="sxs-lookup"><span data-stu-id="c93e0-p135">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="c93e0-587">字符串</span><span class="sxs-lookup"><span data-stu-id="c93e0-587">String</span></span>||<span data-ttu-id="c93e0-p136">要附加的项目的主题。最大长度为 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="c93e0-p136">The sujbect of the item to be attached. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="c93e0-590">对象</span><span class="sxs-lookup"><span data-stu-id="c93e0-590">Object</span></span>| <span data-ttu-id="c93e0-591">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="c93e0-591">&lt;optional&gt;</span></span>|<span data-ttu-id="c93e0-592">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="c93e0-592">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="c93e0-593">对象</span><span class="sxs-lookup"><span data-stu-id="c93e0-593">Object</span></span>| <span data-ttu-id="c93e0-594">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="c93e0-594">&lt;optional&gt;</span></span>|<span data-ttu-id="c93e0-595">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="c93e0-595">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="c93e0-596">函数</span><span class="sxs-lookup"><span data-stu-id="c93e0-596">function</span></span>| <span data-ttu-id="c93e0-597">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="c93e0-597">&lt;optional&gt;</span></span>|<span data-ttu-id="c93e0-598">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="c93e0-598">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="c93e0-599">如果成功，附件标识符将在 `asyncResult.value` 属性中提供。</span><span class="sxs-lookup"><span data-stu-id="c93e0-599">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="c93e0-600">如果添加附件失败，`asyncResult` 对象将包含一个提供错误说明的 `Error` 对象。</span><span class="sxs-lookup"><span data-stu-id="c93e0-600">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="c93e0-601">错误</span><span class="sxs-lookup"><span data-stu-id="c93e0-601">Errors</span></span>

| <span data-ttu-id="c93e0-602">错误代码</span><span class="sxs-lookup"><span data-stu-id="c93e0-602">Error code</span></span> | <span data-ttu-id="c93e0-603">说明</span><span class="sxs-lookup"><span data-stu-id="c93e0-603">Description</span></span> |
|------------|-------------|
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="c93e0-604">邮件或约会具有的附件过多。</span><span class="sxs-lookup"><span data-stu-id="c93e0-604">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="c93e0-605">要求</span><span class="sxs-lookup"><span data-stu-id="c93e0-605">Requirements</span></span>

|<span data-ttu-id="c93e0-606">要求</span><span class="sxs-lookup"><span data-stu-id="c93e0-606">Requirement</span></span>| <span data-ttu-id="c93e0-607">值</span><span class="sxs-lookup"><span data-stu-id="c93e0-607">Value</span></span>|
|---|---|
|[<span data-ttu-id="c93e0-608">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="c93e0-608">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c93e0-609">1.1</span><span class="sxs-lookup"><span data-stu-id="c93e0-609">1.1</span></span>|
|[<span data-ttu-id="c93e0-610">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="c93e0-610">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c93e0-611">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="c93e0-611">ReadWriteItem</span></span>|
|[<span data-ttu-id="c93e0-612">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c93e0-612">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c93e0-613">撰写</span><span class="sxs-lookup"><span data-stu-id="c93e0-613">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="c93e0-614">示例</span><span class="sxs-lookup"><span data-stu-id="c93e0-614">Example</span></span>

<span data-ttu-id="c93e0-615">以下示例将现有的 Outlook 项目添加为名为 `My Attachment` 的附件。</span><span class="sxs-lookup"><span data-stu-id="c93e0-615">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

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

####  <a name="close"></a><span data-ttu-id="c93e0-616">close()</span><span class="sxs-lookup"><span data-stu-id="c93e0-616">close()</span></span>

<span data-ttu-id="c93e0-617">关闭当前正在撰写的项目。</span><span class="sxs-lookup"><span data-stu-id="c93e0-617">Closes the current item that is being composed.</span></span>

<span data-ttu-id="c93e0-p137">
            \`close\` 方法的行为取决于要撰写的项目的当前状态。如果项目具有未保存的更改，客户端将提示用户保存、放弃或取消关闭操作。</span><span class="sxs-lookup"><span data-stu-id="c93e0-p137">The behavior of the `close` method depends on the current state of the item being composed. If the item has unsaved changes, the client prompts the user to save, discard, or cancel the close action.</span></span>

> [!NOTE]
> <span data-ttu-id="c93e0-620">在 Outlook 网页版中，如果该项目是约会并且之前已使用 `saveAsync` 保存，则即使自上次保存项目后未发生任何更改，也会提示用户保存、放弃或取消。</span><span class="sxs-lookup"><span data-stu-id="c93e0-620">In Outlook on the web, if the item is an appointment and it has previously been saved using `saveAsync`, the user is prompted to save, discard, or cancel even if no changes have occurred since the item was last saved.</span></span>

<span data-ttu-id="c93e0-621">在 Outlook 桌面客户端中，如果邮件是内联答复，`close` 方法不起作用。</span><span class="sxs-lookup"><span data-stu-id="c93e0-621">In the Outlook desktop client, if the message is an inline reply, the `close` method has no effect.</span></span>

##### <a name="requirements"></a><span data-ttu-id="c93e0-622">要求</span><span class="sxs-lookup"><span data-stu-id="c93e0-622">Requirements</span></span>

|<span data-ttu-id="c93e0-623">要求</span><span class="sxs-lookup"><span data-stu-id="c93e0-623">Requirement</span></span>| <span data-ttu-id="c93e0-624">值</span><span class="sxs-lookup"><span data-stu-id="c93e0-624">Value</span></span>|
|---|---|
|[<span data-ttu-id="c93e0-625">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="c93e0-625">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c93e0-626">1.3</span><span class="sxs-lookup"><span data-stu-id="c93e0-626">1.3</span></span>|
|[<span data-ttu-id="c93e0-627">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="c93e0-627">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c93e0-628">受限</span><span class="sxs-lookup"><span data-stu-id="c93e0-628">Restricted</span></span>|
|[<span data-ttu-id="c93e0-629">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c93e0-629">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c93e0-630">撰写</span><span class="sxs-lookup"><span data-stu-id="c93e0-630">Compose</span></span>|

#### <a name="displayreplyallformformdata"></a><span data-ttu-id="c93e0-631">displayReplyAllForm(formData)</span><span class="sxs-lookup"><span data-stu-id="c93e0-631">displayReplyAllForm(formData)</span></span>

<span data-ttu-id="c93e0-632">显示答复窗体，其中包括所选邮件的发件人和所有收件人或所选约会的组织者和所有与会者。</span><span class="sxs-lookup"><span data-stu-id="c93e0-632">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="c93e0-633">在 Outlook for iOS 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="c93e0-633">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="c93e0-634">在 Outlook Web App 中，答复窗体显示为包含 3 列视图的弹出式窗体以及包含 2 列或 1 列视图的弹出式窗体。</span><span class="sxs-lookup"><span data-stu-id="c93e0-634">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="c93e0-635">如果任意字符串参数超出其限制，`displayReplyAllForm` 将引发异常。</span><span class="sxs-lookup"><span data-stu-id="c93e0-635">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="c93e0-p138">当在 `formData.attachments` 参数中指定附件时，Outlook 和 Outlook Web App 尝试下载所有附件并将其附加到答复窗体。如果无法添加任何附件，则在窗体 UI 中显示错误。如果这不可能，则不引发错误消息。</span><span class="sxs-lookup"><span data-stu-id="c93e0-p138">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c93e0-639">参数：</span><span class="sxs-lookup"><span data-stu-id="c93e0-639">Parameters:</span></span>

|<span data-ttu-id="c93e0-640">名称</span><span class="sxs-lookup"><span data-stu-id="c93e0-640">Name</span></span>| <span data-ttu-id="c93e0-641">类型</span><span class="sxs-lookup"><span data-stu-id="c93e0-641">Type</span></span>| <span data-ttu-id="c93e0-642">说明</span><span class="sxs-lookup"><span data-stu-id="c93e0-642">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="c93e0-643">字符串 &#124; 对象</span><span class="sxs-lookup"><span data-stu-id="c93e0-643">String &#124; Object</span></span>| |<span data-ttu-id="c93e0-p139">一个包含文本和 HTML 且表示答复窗体的正文的字符串。字符串限制为 32 KB。</span><span class="sxs-lookup"><span data-stu-id="c93e0-p139">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="c93e0-646">**OR**</span><span class="sxs-lookup"><span data-stu-id="c93e0-646">**OR**</span></span><br/><span data-ttu-id="c93e0-p140">包含正文或附件数据和回调函数的对象。对象定义如下。</span><span class="sxs-lookup"><span data-stu-id="c93e0-p140">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="c93e0-649">字符串</span><span class="sxs-lookup"><span data-stu-id="c93e0-649">String</span></span> | <span data-ttu-id="c93e0-650">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="c93e0-650">&lt;optional&gt;</span></span> | <span data-ttu-id="c93e0-p141">一个包含文本和 HTML 且表示答复窗体的正文的字符串。字符串限制为 32 KB。</span><span class="sxs-lookup"><span data-stu-id="c93e0-p141">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="c93e0-653">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="c93e0-653">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="c93e0-654">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="c93e0-654">&lt;optional&gt;</span></span> | <span data-ttu-id="c93e0-655">JSON 对象（是文件或项目附件）的数组。</span><span class="sxs-lookup"><span data-stu-id="c93e0-655">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="c93e0-656">字符串</span><span class="sxs-lookup"><span data-stu-id="c93e0-656">String</span></span> | | <span data-ttu-id="c93e0-p142">指示附件的类型。必须是文件附件的 `file` 或项目附件的 `item`。</span><span class="sxs-lookup"><span data-stu-id="c93e0-p142">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="c93e0-659">字符串</span><span class="sxs-lookup"><span data-stu-id="c93e0-659">String</span></span> | | <span data-ttu-id="c93e0-660">一个包含附件的名称的字符串，最多包含 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="c93e0-660">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="c93e0-661">String</span><span class="sxs-lookup"><span data-stu-id="c93e0-661">String</span></span> | | <span data-ttu-id="c93e0-p143">仅在将 `type` 设置为 `file` 时使用。文件的位置的 URI。</span><span class="sxs-lookup"><span data-stu-id="c93e0-p143">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="c93e0-664">String</span><span class="sxs-lookup"><span data-stu-id="c93e0-664">String</span></span> | | <span data-ttu-id="c93e0-p144">仅在将 `type` 设置为 `item` 时使用。附件的 EWS 项目 ID。字符串最长为 100 个字符。</span><span class="sxs-lookup"><span data-stu-id="c93e0-p144">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="c93e0-668">函数</span><span class="sxs-lookup"><span data-stu-id="c93e0-668">function</span></span> | <span data-ttu-id="c93e0-669">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="c93e0-669">&lt;optional&gt;</span></span> | <span data-ttu-id="c93e0-670">方法完成后，使用单个参数 `asyncResult`（一个 [AsyncResult](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="c93e0-670">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="c93e0-671">要求</span><span class="sxs-lookup"><span data-stu-id="c93e0-671">Requirements</span></span>

|<span data-ttu-id="c93e0-672">要求</span><span class="sxs-lookup"><span data-stu-id="c93e0-672">Requirement</span></span>| <span data-ttu-id="c93e0-673">值</span><span class="sxs-lookup"><span data-stu-id="c93e0-673">Value</span></span>|
|---|---|
|[<span data-ttu-id="c93e0-674">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="c93e0-674">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c93e0-675">1.0</span><span class="sxs-lookup"><span data-stu-id="c93e0-675">1.0</span></span>|
|[<span data-ttu-id="c93e0-676">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="c93e0-676">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c93e0-677">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c93e0-677">ReadItem</span></span>|
|[<span data-ttu-id="c93e0-678">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c93e0-678">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c93e0-679">阅读</span><span class="sxs-lookup"><span data-stu-id="c93e0-679">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="c93e0-680">示例</span><span class="sxs-lookup"><span data-stu-id="c93e0-680">Examples</span></span>

<span data-ttu-id="c93e0-681">以下代码将一个字符串传递到 `displayReplyAllForm` 函数。</span><span class="sxs-lookup"><span data-stu-id="c93e0-681">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="c93e0-682">使用空白正文答复。</span><span class="sxs-lookup"><span data-stu-id="c93e0-682">Reply with an empty body.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="c93e0-683">仅使用正文答复。</span><span class="sxs-lookup"><span data-stu-id="c93e0-683">Reply with just a body.</span></span>

```js
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="c93e0-684">使用正文和文件附件答复。</span><span class="sxs-lookup"><span data-stu-id="c93e0-684">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="c93e0-685">使用正文和项目附件答复。</span><span class="sxs-lookup"><span data-stu-id="c93e0-685">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="c93e0-686">使用正文、文件附件、项目附件和回调答复。</span><span class="sxs-lookup"><span data-stu-id="c93e0-686">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="displayreplyformformdata"></a><span data-ttu-id="c93e0-687">displayReplyForm(formData)</span><span class="sxs-lookup"><span data-stu-id="c93e0-687">displayReplyForm(formData)</span></span>

<span data-ttu-id="c93e0-688">显示答复窗体，其中仅包括所选邮件的发件人或所选约会的组织者。</span><span class="sxs-lookup"><span data-stu-id="c93e0-688">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="c93e0-689">在 Outlook for iOS 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="c93e0-689">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="c93e0-690">在 Outlook Web App 中，答复窗体显示为包含 3 列视图的弹出式窗体以及包含 2 列或 1 列视图的弹出式窗体。</span><span class="sxs-lookup"><span data-stu-id="c93e0-690">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="c93e0-691">如果任意字符串参数超出其限制，`displayReplyForm` 将引发异常。</span><span class="sxs-lookup"><span data-stu-id="c93e0-691">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="c93e0-p145">当在 `formData.attachments` 参数中指定附件时，Outlook 和 Outlook Web App 尝试下载所有附件并将其附加到答复窗体。如果无法添加任何附件，则在窗体 UI 中显示错误。如果这不可能，则不引发错误消息。</span><span class="sxs-lookup"><span data-stu-id="c93e0-p145">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c93e0-695">参数：</span><span class="sxs-lookup"><span data-stu-id="c93e0-695">Parameters:</span></span>

|<span data-ttu-id="c93e0-696">名称</span><span class="sxs-lookup"><span data-stu-id="c93e0-696">Name</span></span>| <span data-ttu-id="c93e0-697">类型</span><span class="sxs-lookup"><span data-stu-id="c93e0-697">Type</span></span>| <span data-ttu-id="c93e0-698">说明</span><span class="sxs-lookup"><span data-stu-id="c93e0-698">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="c93e0-699">字符串 &#124; 对象</span><span class="sxs-lookup"><span data-stu-id="c93e0-699">String &#124; Object</span></span>| | <span data-ttu-id="c93e0-p146">一个包含文本和 HTML 且表示答复窗体的正文的字符串。字符串限制为 32 KB。</span><span class="sxs-lookup"><span data-stu-id="c93e0-p146">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="c93e0-702">**OR**</span><span class="sxs-lookup"><span data-stu-id="c93e0-702">**OR**</span></span><br/><span data-ttu-id="c93e0-p147">包含正文或附件数据和回调函数的对象。对象定义如下。</span><span class="sxs-lookup"><span data-stu-id="c93e0-p147">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="c93e0-705">字符串</span><span class="sxs-lookup"><span data-stu-id="c93e0-705">String</span></span> | <span data-ttu-id="c93e0-706">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="c93e0-706">&lt;optional&gt;</span></span> | <span data-ttu-id="c93e0-p148">一个包含文本和 HTML 且表示答复窗体的正文的字符串。字符串限制为 32 KB。</span><span class="sxs-lookup"><span data-stu-id="c93e0-p148">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="c93e0-709">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="c93e0-709">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="c93e0-710">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="c93e0-710">&lt;optional&gt;</span></span> | <span data-ttu-id="c93e0-711">JSON 对象（是文件或项目附件）的数组。</span><span class="sxs-lookup"><span data-stu-id="c93e0-711">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="c93e0-712">字符串</span><span class="sxs-lookup"><span data-stu-id="c93e0-712">String</span></span> | | <span data-ttu-id="c93e0-p149">指示附件的类型。必须是文件附件的 `file` 或项目附件的 `item`。</span><span class="sxs-lookup"><span data-stu-id="c93e0-p149">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="c93e0-715">字符串</span><span class="sxs-lookup"><span data-stu-id="c93e0-715">String</span></span> | | <span data-ttu-id="c93e0-716">一个包含附件的名称的字符串，最多包含 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="c93e0-716">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="c93e0-717">String</span><span class="sxs-lookup"><span data-stu-id="c93e0-717">String</span></span> | | <span data-ttu-id="c93e0-p150">仅在将 `type` 设置为 `file` 时使用。文件的位置的 URI。</span><span class="sxs-lookup"><span data-stu-id="c93e0-p150">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="c93e0-720">String</span><span class="sxs-lookup"><span data-stu-id="c93e0-720">String</span></span> | | <span data-ttu-id="c93e0-p151">仅在将 `type` 设置为 `item` 时使用。附件的 EWS 项目 ID。字符串最长为 100 个字符。</span><span class="sxs-lookup"><span data-stu-id="c93e0-p151">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="c93e0-724">函数</span><span class="sxs-lookup"><span data-stu-id="c93e0-724">function</span></span> | <span data-ttu-id="c93e0-725">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="c93e0-725">&lt;optional&gt;</span></span> | <span data-ttu-id="c93e0-726">方法完成后，使用单个参数 `asyncResult`（一个 [AsyncResult](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="c93e0-726">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="c93e0-727">要求</span><span class="sxs-lookup"><span data-stu-id="c93e0-727">Requirements</span></span>

|<span data-ttu-id="c93e0-728">要求</span><span class="sxs-lookup"><span data-stu-id="c93e0-728">Requirement</span></span>| <span data-ttu-id="c93e0-729">值</span><span class="sxs-lookup"><span data-stu-id="c93e0-729">Value</span></span>|
|---|---|
|[<span data-ttu-id="c93e0-730">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="c93e0-730">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c93e0-731">1.0</span><span class="sxs-lookup"><span data-stu-id="c93e0-731">1.0</span></span>|
|[<span data-ttu-id="c93e0-732">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="c93e0-732">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c93e0-733">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c93e0-733">ReadItem</span></span>|
|[<span data-ttu-id="c93e0-734">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c93e0-734">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c93e0-735">阅读</span><span class="sxs-lookup"><span data-stu-id="c93e0-735">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="c93e0-736">示例</span><span class="sxs-lookup"><span data-stu-id="c93e0-736">Examples</span></span>

<span data-ttu-id="c93e0-737">以下代码将一个字符串传递到 `displayReplyForm` 函数。</span><span class="sxs-lookup"><span data-stu-id="c93e0-737">The following code passes a string to the `displayReplyForm` function.</span></span>

```js
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="c93e0-738">使用空白正文答复。</span><span class="sxs-lookup"><span data-stu-id="c93e0-738">Reply with an empty body.</span></span>

```js
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="c93e0-739">仅使用正文答复。</span><span class="sxs-lookup"><span data-stu-id="c93e0-739">Reply with just a body.</span></span>

```js
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="c93e0-740">使用正文和文件附件答复。</span><span class="sxs-lookup"><span data-stu-id="c93e0-740">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="c93e0-741">使用正文和项目附件答复。</span><span class="sxs-lookup"><span data-stu-id="c93e0-741">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="c93e0-742">使用正文、文件附件、项目附件和回调答复。</span><span class="sxs-lookup"><span data-stu-id="c93e0-742">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="getentities--entitiesjavascriptapioutlook14officeentities"></a><span data-ttu-id="c93e0-743">getEntities() → {[Entities](/javascript/api/outlook_1_4/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="c93e0-743">getEntities() → {[Entities](/javascript/api/outlook_1_4/office.entities)}</span></span>

<span data-ttu-id="c93e0-744">获取在所选项目的正文中找到的实体。</span><span class="sxs-lookup"><span data-stu-id="c93e0-744">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="c93e0-745">在 Outlook for iOS 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="c93e0-745">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="c93e0-746">要求</span><span class="sxs-lookup"><span data-stu-id="c93e0-746">Requirements</span></span>

|<span data-ttu-id="c93e0-747">要求</span><span class="sxs-lookup"><span data-stu-id="c93e0-747">Requirement</span></span>| <span data-ttu-id="c93e0-748">值</span><span class="sxs-lookup"><span data-stu-id="c93e0-748">Value</span></span>|
|---|---|
|[<span data-ttu-id="c93e0-749">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="c93e0-749">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c93e0-750">1.0</span><span class="sxs-lookup"><span data-stu-id="c93e0-750">1.0</span></span>|
|[<span data-ttu-id="c93e0-751">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="c93e0-751">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c93e0-752">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c93e0-752">ReadItem</span></span>|
|[<span data-ttu-id="c93e0-753">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c93e0-753">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c93e0-754">阅读</span><span class="sxs-lookup"><span data-stu-id="c93e0-754">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="c93e0-755">返回：</span><span class="sxs-lookup"><span data-stu-id="c93e0-755">Returns:</span></span>

<span data-ttu-id="c93e0-756">类型：[Entities](/javascript/api/outlook_1_4/office.entities)</span><span class="sxs-lookup"><span data-stu-id="c93e0-756">Type: [Entities](/javascript/api/outlook_1_4/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="c93e0-757">示例</span><span class="sxs-lookup"><span data-stu-id="c93e0-757">Example</span></span>

<span data-ttu-id="c93e0-758">以下示例访问当前项目的正文中的联系人实体。</span><span class="sxs-lookup"><span data-stu-id="c93e0-758">The following example accesses the contacts entities in the current item's body.</span></span>

```js
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlook14officecontactmeetingsuggestionjavascriptapioutlook14officemeetingsuggestionphonenumberjavascriptapioutlook14officephonenumbertasksuggestionjavascriptapioutlook14officetasksuggestion"></a><span data-ttu-id="c93e0-759">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_4/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_4/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_4/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_4/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="c93e0-759">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_4/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_4/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_4/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_4/office.tasksuggestion))>}</span></span>

<span data-ttu-id="c93e0-760">获取所选项目的正文中找到的指定实体类型的所有实体的数组。</span><span class="sxs-lookup"><span data-stu-id="c93e0-760">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="c93e0-761">在 Outlook for iOS 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="c93e0-761">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c93e0-762">参数：</span><span class="sxs-lookup"><span data-stu-id="c93e0-762">Parameters:</span></span>

|<span data-ttu-id="c93e0-763">名称</span><span class="sxs-lookup"><span data-stu-id="c93e0-763">Name</span></span>| <span data-ttu-id="c93e0-764">类型</span><span class="sxs-lookup"><span data-stu-id="c93e0-764">Type</span></span>| <span data-ttu-id="c93e0-765">说明</span><span class="sxs-lookup"><span data-stu-id="c93e0-765">Description</span></span>|
|---|---|---|
|`entityType`| [<span data-ttu-id="c93e0-766">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="c93e0-766">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook_1_4/office.mailboxenums.entitytype)|<span data-ttu-id="c93e0-767">EntityType 枚举值之一。</span><span class="sxs-lookup"><span data-stu-id="c93e0-767">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c93e0-768">要求</span><span class="sxs-lookup"><span data-stu-id="c93e0-768">Requirements</span></span>

|<span data-ttu-id="c93e0-769">要求</span><span class="sxs-lookup"><span data-stu-id="c93e0-769">Requirement</span></span>| <span data-ttu-id="c93e0-770">值</span><span class="sxs-lookup"><span data-stu-id="c93e0-770">Value</span></span>|
|---|---|
|[<span data-ttu-id="c93e0-771">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="c93e0-771">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c93e0-772">1.0</span><span class="sxs-lookup"><span data-stu-id="c93e0-772">1.0</span></span>|
|[<span data-ttu-id="c93e0-773">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="c93e0-773">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c93e0-774">受限</span><span class="sxs-lookup"><span data-stu-id="c93e0-774">Restricted</span></span>|
|[<span data-ttu-id="c93e0-775">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c93e0-775">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c93e0-776">阅读</span><span class="sxs-lookup"><span data-stu-id="c93e0-776">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="c93e0-777">返回：</span><span class="sxs-lookup"><span data-stu-id="c93e0-777">Returns:</span></span>

<span data-ttu-id="c93e0-778">如果在 `entityType` 中传递的值不是 `EntityType` 枚举的有效成员，该方法返回 null。</span><span class="sxs-lookup"><span data-stu-id="c93e0-778">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="c93e0-779">如果指定类型的任何实体都不存在于该项目的正文中，该方法将返回空数组。</span><span class="sxs-lookup"><span data-stu-id="c93e0-779">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="c93e0-780">否则，返回的数组中对象的类型取决于 `entityType` 参数中请求实体的类型。</span><span class="sxs-lookup"><span data-stu-id="c93e0-780">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="c93e0-781">当使用此方法的最低权限级别**受限**时，某些实体类型需要 **ReadItem** 才能进行访问，如下表中所指定。</span><span class="sxs-lookup"><span data-stu-id="c93e0-781">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

| <span data-ttu-id="c93e0-782">`entityType` 的值</span><span class="sxs-lookup"><span data-stu-id="c93e0-782">Value of `entityType`</span></span> | <span data-ttu-id="c93e0-783">返回的数组中对象的类型</span><span class="sxs-lookup"><span data-stu-id="c93e0-783">Type of objects in returned array</span></span> | <span data-ttu-id="c93e0-784">所需权限级别</span><span class="sxs-lookup"><span data-stu-id="c93e0-784">Required Permission Level</span></span> |
| --- | --- | --- |
| `Address` | <span data-ttu-id="c93e0-785">String</span><span class="sxs-lookup"><span data-stu-id="c93e0-785">String</span></span> | <span data-ttu-id="c93e0-786">**受限**</span><span class="sxs-lookup"><span data-stu-id="c93e0-786">**Restricted**</span></span> |
| `Contact` | <span data-ttu-id="c93e0-787">Contact</span><span class="sxs-lookup"><span data-stu-id="c93e0-787">Contact</span></span> | <span data-ttu-id="c93e0-788">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="c93e0-788">**ReadItem**</span></span> |
| `EmailAddress` | <span data-ttu-id="c93e0-789">String</span><span class="sxs-lookup"><span data-stu-id="c93e0-789">String</span></span> | <span data-ttu-id="c93e0-790">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="c93e0-790">**ReadItem**</span></span> |
| `MeetingSuggestion` | <span data-ttu-id="c93e0-791">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="c93e0-791">MeetingSuggestion</span></span> | <span data-ttu-id="c93e0-792">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="c93e0-792">**ReadItem**</span></span> |
| `PhoneNumber` | <span data-ttu-id="c93e0-793">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="c93e0-793">PhoneNumber</span></span> | <span data-ttu-id="c93e0-794">**受限**</span><span class="sxs-lookup"><span data-stu-id="c93e0-794">**Restricted**</span></span> |
| `TaskSuggestion` | <span data-ttu-id="c93e0-795">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="c93e0-795">TaskSuggestion</span></span> | <span data-ttu-id="c93e0-796">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="c93e0-796">**ReadItem**</span></span> |
| `URL` | <span data-ttu-id="c93e0-797">String</span><span class="sxs-lookup"><span data-stu-id="c93e0-797">String</span></span> | <span data-ttu-id="c93e0-798">**受限**</span><span class="sxs-lookup"><span data-stu-id="c93e0-798">**Restricted**</span></span> |

<span data-ttu-id="c93e0-799">类型：Array.<(String|[Contact](/javascript/api/outlook_1_4/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_4/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_4/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_4/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="c93e0-799">Type: Array.<(String|[Contact](/javascript/api/outlook_1_4/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_4/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_4/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_4/office.tasksuggestion))></span></span>

##### <a name="example"></a><span data-ttu-id="c93e0-800">示例</span><span class="sxs-lookup"><span data-stu-id="c93e0-800">Example</span></span>

<span data-ttu-id="c93e0-801">以下示例显示了如何访问表示当前项目的正文中的邮政地址的字符串数组。</span><span class="sxs-lookup"><span data-stu-id="c93e0-801">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlook14officecontactmeetingsuggestionjavascriptapioutlook14officemeetingsuggestionphonenumberjavascriptapioutlook14officephonenumbertasksuggestionjavascriptapioutlook14officetasksuggestion"></a><span data-ttu-id="c93e0-802">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_4/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_4/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_4/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_4/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="c93e0-802">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_4/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_4/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_4/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_4/office.tasksuggestion))>}</span></span>

<span data-ttu-id="c93e0-803">返回传递清单 XML 文件中定义的命名筛选器的所选项目中的已知实体。</span><span class="sxs-lookup"><span data-stu-id="c93e0-803">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="c93e0-804">在 Outlook for iOS 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="c93e0-804">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="c93e0-805">`getFilteredEntitiesByName` 方法返回匹配在具有指定 `FilterName` 元素值的清单 XML 文件中的 [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) 规则元素中定义的正则表达式的实体。</span><span class="sxs-lookup"><span data-stu-id="c93e0-805">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c93e0-806">参数：</span><span class="sxs-lookup"><span data-stu-id="c93e0-806">Parameters:</span></span>

|<span data-ttu-id="c93e0-807">名称</span><span class="sxs-lookup"><span data-stu-id="c93e0-807">Name</span></span>| <span data-ttu-id="c93e0-808">类型</span><span class="sxs-lookup"><span data-stu-id="c93e0-808">Type</span></span>| <span data-ttu-id="c93e0-809">说明</span><span class="sxs-lookup"><span data-stu-id="c93e0-809">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="c93e0-810">字符串</span><span class="sxs-lookup"><span data-stu-id="c93e0-810">String</span></span>|<span data-ttu-id="c93e0-811">定义筛选器匹配的 `ItemHasKnownEntity` 规则元素的名称。</span><span class="sxs-lookup"><span data-stu-id="c93e0-811">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c93e0-812">要求</span><span class="sxs-lookup"><span data-stu-id="c93e0-812">Requirements</span></span>

|<span data-ttu-id="c93e0-813">要求</span><span class="sxs-lookup"><span data-stu-id="c93e0-813">Requirement</span></span>| <span data-ttu-id="c93e0-814">值</span><span class="sxs-lookup"><span data-stu-id="c93e0-814">Value</span></span>|
|---|---|
|[<span data-ttu-id="c93e0-815">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="c93e0-815">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c93e0-816">1.0</span><span class="sxs-lookup"><span data-stu-id="c93e0-816">1.0</span></span>|
|[<span data-ttu-id="c93e0-817">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="c93e0-817">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c93e0-818">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c93e0-818">ReadItem</span></span>|
|[<span data-ttu-id="c93e0-819">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c93e0-819">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c93e0-820">阅读</span><span class="sxs-lookup"><span data-stu-id="c93e0-820">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="c93e0-821">返回：</span><span class="sxs-lookup"><span data-stu-id="c93e0-821">Returns:</span></span>

<span data-ttu-id="c93e0-p153">如果具有匹配 `name` 参数的 `FilterName` 元素值的清单中没有任何 `ItemHasKnownEntity` 元素，则该方法返回 `null`。如果 `name` 参数匹配清单中的 `ItemHasKnownEntity` 元素，但在匹配的当前项目中没有实体，则该方法返回一个空数组。</span><span class="sxs-lookup"><span data-stu-id="c93e0-p153">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="c93e0-824">类型：Array.<(String|[Contact](/javascript/api/outlook_1_4/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_4/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_4/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_4/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="c93e0-824">Type: Array.<(String|[Contact](/javascript/api/outlook_1_4/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_4/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_4/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_4/office.tasksuggestion))></span></span>

#### <a name="getregexmatches--object"></a><span data-ttu-id="c93e0-825">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="c93e0-825">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="c93e0-826">返回所选项目中匹配在清单 XML 文件中定义的正则表达式的字符串值。</span><span class="sxs-lookup"><span data-stu-id="c93e0-826">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="c93e0-827">在 Outlook for iOS 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="c93e0-827">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="c93e0-p154">`getRegExMatches` 方法返回匹配在清单 XML 文件中的每个 `ItemHasRegularExpressionMatch` 或 `ItemHasKnownEntity` 规则元素中定义的正则表达式的字符串。对于 `ItemHasRegularExpressionMatch` 规则，匹配字符串必须发生在该规则指定的项目的属性中。`PropertyName` 简单类型定义支持的属性。</span><span class="sxs-lookup"><span data-stu-id="c93e0-p154">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="c93e0-831">例如，考虑一个外接程序清单具有以下 `Rule` 元素：</span><span class="sxs-lookup"><span data-stu-id="c93e0-831">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="c93e0-832">从 `getRegExMatches` 返回的对象应有两个属性：`fruits` 和 `veggies`。</span><span class="sxs-lookup"><span data-stu-id="c93e0-832">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

<span data-ttu-id="c93e0-p155">如果在项目的正文属性上指定 `ItemHasRegularExpressionMatch` 规则，则正则表达式应进一步筛选正文，不应尝试返回该项目的整个正文。使用正则表达式（如 `.*`）获取项目的整个正文并不总是返回预期的结果。而是使用 [`Body.getAsync`](/javascript/api/outlook_1_4/office.body#getasync-coerciontype--options--callback-) 方法检索整个正文。</span><span class="sxs-lookup"><span data-stu-id="c93e0-p155">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results. Instead, use the [`Body.getAsync`](/javascript/api/outlook_1_4/office.body#getasync-coerciontype--options--callback-) method to retrieve the entire body.</span></span>

##### <a name="requirements"></a><span data-ttu-id="c93e0-836">Requirements</span><span class="sxs-lookup"><span data-stu-id="c93e0-836">Requirements</span></span>

|<span data-ttu-id="c93e0-837">要求</span><span class="sxs-lookup"><span data-stu-id="c93e0-837">Requirement</span></span>| <span data-ttu-id="c93e0-838">值</span><span class="sxs-lookup"><span data-stu-id="c93e0-838">Value</span></span>|
|---|---|
|[<span data-ttu-id="c93e0-839">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="c93e0-839">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c93e0-840">1.0</span><span class="sxs-lookup"><span data-stu-id="c93e0-840">1.0</span></span>|
|[<span data-ttu-id="c93e0-841">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="c93e0-841">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c93e0-842">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c93e0-842">ReadItem</span></span>|
|[<span data-ttu-id="c93e0-843">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c93e0-843">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c93e0-844">阅读</span><span class="sxs-lookup"><span data-stu-id="c93e0-844">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="c93e0-845">返回：</span><span class="sxs-lookup"><span data-stu-id="c93e0-845">Returns:</span></span>

<span data-ttu-id="c93e0-p156">一个包含匹配在清单 XML 文件中定义的正则表达式的字符串数组的对象。每个数组的名称等于匹配 `ItemHasRegularExpressionMatch` 规则的 `RegExName` 属性或匹配 `ItemHasKnownEntity` 规则的 `FilterName` 属性的相应值。</span><span class="sxs-lookup"><span data-stu-id="c93e0-p156">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<dl class="param-type"><span data-ttu-id="c93e0-848">

<dt>
类型</dt>


</span><span class="sxs-lookup"><span data-stu-id="c93e0-848">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="c93e0-849">对象</span><span class="sxs-lookup"><span data-stu-id="c93e0-849">Object</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="c93e0-850">示例</span><span class="sxs-lookup"><span data-stu-id="c93e0-850">Example</span></span>

<span data-ttu-id="c93e0-851">以下示例显示了如何访问正则表达式 <rule> 元素 `fruits` 和 `veggies` 的匹配项的数组，这些元素在清单中指定。</rule></span><span class="sxs-lookup"><span data-stu-id="c93e0-851">The following example shows how to access the array of matches for the regular expression <rule>elements `fruits` and `veggies`, which are specified in the manifest.</rule></span></span>

```js
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veges = allMatches.veggies;
```

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="c93e0-852">getRegExMatchesByName(name) → (nullable) {Array.<String>}</span><span class="sxs-lookup"><span data-stu-id="c93e0-852">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="c93e0-853">返回匹配在清单 XML 文件中定义的命名正则表达式的所选项目中的字符串值。</span><span class="sxs-lookup"><span data-stu-id="c93e0-853">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="c93e0-854">在 Outlook for iOS 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="c93e0-854">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="c93e0-855">`getRegExMatchesByName` 方法返回匹配在具有指定 `RegExName` 元素值的清单 XML 文件中的 `ItemHasRegularExpressionMatch` 规则元素中定义的正则表达式的字符串。</span><span class="sxs-lookup"><span data-stu-id="c93e0-855">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="c93e0-p157">如果在项目的正文属性上指定 `ItemHasRegularExpressionMatch` 规则，则正则表达式应进一步筛选正文，不应尝试返回该项目的整个正文。使用正则表达式（如 `.*`）获取项目的整个正文并不总是返回预期的结果。</span><span class="sxs-lookup"><span data-stu-id="c93e0-p157">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c93e0-858">参数：</span><span class="sxs-lookup"><span data-stu-id="c93e0-858">Parameters:</span></span>

|<span data-ttu-id="c93e0-859">名称</span><span class="sxs-lookup"><span data-stu-id="c93e0-859">Name</span></span>| <span data-ttu-id="c93e0-860">类型</span><span class="sxs-lookup"><span data-stu-id="c93e0-860">Type</span></span>| <span data-ttu-id="c93e0-861">说明</span><span class="sxs-lookup"><span data-stu-id="c93e0-861">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="c93e0-862">字符串</span><span class="sxs-lookup"><span data-stu-id="c93e0-862">String</span></span>|<span data-ttu-id="c93e0-863">定义筛选器匹配的 `ItemHasRegularExpressionMatch` 规则元素的名称。</span><span class="sxs-lookup"><span data-stu-id="c93e0-863">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c93e0-864">要求</span><span class="sxs-lookup"><span data-stu-id="c93e0-864">Requirements</span></span>

|<span data-ttu-id="c93e0-865">要求</span><span class="sxs-lookup"><span data-stu-id="c93e0-865">Requirement</span></span>| <span data-ttu-id="c93e0-866">值</span><span class="sxs-lookup"><span data-stu-id="c93e0-866">Value</span></span>|
|---|---|
|[<span data-ttu-id="c93e0-867">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="c93e0-867">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c93e0-868">1.0</span><span class="sxs-lookup"><span data-stu-id="c93e0-868">1.0</span></span>|
|[<span data-ttu-id="c93e0-869">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="c93e0-869">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c93e0-870">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c93e0-870">ReadItem</span></span>|
|[<span data-ttu-id="c93e0-871">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c93e0-871">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c93e0-872">阅读</span><span class="sxs-lookup"><span data-stu-id="c93e0-872">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="c93e0-873">返回：</span><span class="sxs-lookup"><span data-stu-id="c93e0-873">Returns:</span></span>

<span data-ttu-id="c93e0-874">一个包含匹配在清单 XML 文件中定义的正则表达式的字符串的数组。</span><span class="sxs-lookup"><span data-stu-id="c93e0-874">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<dl class="param-type"><span data-ttu-id="c93e0-875">

<dt>类型</dt>

</span><span class="sxs-lookup"><span data-stu-id="c93e0-875">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="c93e0-876">Array.< String ></span><span class="sxs-lookup"><span data-stu-id="c93e0-876">Array.< String ></span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="c93e0-877">示例</span><span class="sxs-lookup"><span data-stu-id="c93e0-877">Example</span></span>

```js
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

####  <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="c93e0-878">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="c93e0-878">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="c93e0-879">以异步方式返回邮件的主题或正文中选定的数据。</span><span class="sxs-lookup"><span data-stu-id="c93e0-879">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="c93e0-p158">如果没有选定内容，但光标位于正文或主题中，此方法将会为所选数据返回 null。如果选定的是字段，而不是正文或主题，则此方法返回 `InvalidSelection` 错误。</span><span class="sxs-lookup"><span data-stu-id="c93e0-p158">If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c93e0-882">参数：</span><span class="sxs-lookup"><span data-stu-id="c93e0-882">Parameters:</span></span>

|<span data-ttu-id="c93e0-883">名称</span><span class="sxs-lookup"><span data-stu-id="c93e0-883">Name</span></span>| <span data-ttu-id="c93e0-884">类型</span><span class="sxs-lookup"><span data-stu-id="c93e0-884">Type</span></span>| <span data-ttu-id="c93e0-885">属性</span><span class="sxs-lookup"><span data-stu-id="c93e0-885">Attributes</span></span>| <span data-ttu-id="c93e0-886">说明</span><span class="sxs-lookup"><span data-stu-id="c93e0-886">Description</span></span>|
|---|---|---|---|
|`coercionType`| [<span data-ttu-id="c93e0-887">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="c93e0-887">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="c93e0-p159">请求数据的格式。如果为文本，则此方法返回纯文本作为字符串，删除任何显示的 HTML 标记。如果为 HTML，则此方法返回所选文本，不论是纯文本还是 HTML。</span><span class="sxs-lookup"><span data-stu-id="c93e0-p159">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`| <span data-ttu-id="c93e0-891">对象</span><span class="sxs-lookup"><span data-stu-id="c93e0-891">Object</span></span>| <span data-ttu-id="c93e0-892">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="c93e0-892">&lt;optional&gt;</span></span>|<span data-ttu-id="c93e0-893">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="c93e0-893">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="c93e0-894">对象</span><span class="sxs-lookup"><span data-stu-id="c93e0-894">Object</span></span>| <span data-ttu-id="c93e0-895">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="c93e0-895">&lt;optional&gt;</span></span>|<span data-ttu-id="c93e0-896">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="c93e0-896">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="c93e0-897">函数</span><span class="sxs-lookup"><span data-stu-id="c93e0-897">function</span></span>||<span data-ttu-id="c93e0-898">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="c93e0-898">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="c93e0-899">若要从回调方法访问所选数据，请调用 `asyncResult.value.data`。</span><span class="sxs-lookup"><span data-stu-id="c93e0-899">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="c93e0-900">若要访问选定内容的源属性，请调用 `asyncResult.value.sourceProperty`，这将为 `body` 或 `subject`。</span><span class="sxs-lookup"><span data-stu-id="c93e0-900">To access the source property that the selection comes from, call `asyncResult.value.sourceProperty`, which will be either `body` or `subject`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c93e0-901">要求</span><span class="sxs-lookup"><span data-stu-id="c93e0-901">Requirements</span></span>

|<span data-ttu-id="c93e0-902">要求</span><span class="sxs-lookup"><span data-stu-id="c93e0-902">Requirement</span></span>| <span data-ttu-id="c93e0-903">值</span><span class="sxs-lookup"><span data-stu-id="c93e0-903">Value</span></span>|
|---|---|
|[<span data-ttu-id="c93e0-904">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="c93e0-904">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c93e0-905">1.2</span><span class="sxs-lookup"><span data-stu-id="c93e0-905">1.2</span></span>|
|[<span data-ttu-id="c93e0-906">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="c93e0-906">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c93e0-907">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="c93e0-907">ReadWriteItem</span></span>|
|[<span data-ttu-id="c93e0-908">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c93e0-908">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c93e0-909">撰写</span><span class="sxs-lookup"><span data-stu-id="c93e0-909">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="c93e0-910">返回：</span><span class="sxs-lookup"><span data-stu-id="c93e0-910">Returns:</span></span>

<span data-ttu-id="c93e0-911">作为字符串的所选数据的格式由 `coercionType` 确定。</span><span class="sxs-lookup"><span data-stu-id="c93e0-911">The selected data as a string with format determined by `coercionType`.</span></span>

<dl class="param-type"><span data-ttu-id="c93e0-912">

<dt>
类型</dt>


</span><span class="sxs-lookup"><span data-stu-id="c93e0-912">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="c93e0-913">字符串</span><span class="sxs-lookup"><span data-stu-id="c93e0-913">String</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="c93e0-914">示例</span><span class="sxs-lookup"><span data-stu-id="c93e0-914">Example</span></span>

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

####  <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="c93e0-915">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="c93e0-915">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="c93e0-916">异步加载所选项目上此外接程序的自定义属性。</span><span class="sxs-lookup"><span data-stu-id="c93e0-916">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="c93e0-p161">自定义属性基于每个应用、每个项目存储为键/值对。此方法在回调中返回 `CustomProperties` 对象，该回调提供访问特定于当前项目和当前外接程序的自定义属性的方法。自定义属性未在项目上加密，因此这不应用作安全存储。</span><span class="sxs-lookup"><span data-stu-id="c93e0-p161">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c93e0-920">参数：</span><span class="sxs-lookup"><span data-stu-id="c93e0-920">Parameters:</span></span>

|<span data-ttu-id="c93e0-921">名称</span><span class="sxs-lookup"><span data-stu-id="c93e0-921">Name</span></span>| <span data-ttu-id="c93e0-922">类型</span><span class="sxs-lookup"><span data-stu-id="c93e0-922">Type</span></span>| <span data-ttu-id="c93e0-923">属性</span><span class="sxs-lookup"><span data-stu-id="c93e0-923">Attributes</span></span>| <span data-ttu-id="c93e0-924">说明</span><span class="sxs-lookup"><span data-stu-id="c93e0-924">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="c93e0-925">函数</span><span class="sxs-lookup"><span data-stu-id="c93e0-925">function</span></span>||<span data-ttu-id="c93e0-926">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="c93e0-926">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="c93e0-927">自定义属性作为 `asyncResult.value` 属性中的 [`CustomProperties`](/javascript/api/outlook_1_4/office.customproperties) 对象提供。</span><span class="sxs-lookup"><span data-stu-id="c93e0-927">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook_1_4/office.customproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="c93e0-928">此对象可用于获取、设置以及从项目中删除自定义属性，并将自定义属性集的更改重新保存到服务器。</span><span class="sxs-lookup"><span data-stu-id="c93e0-928">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`| <span data-ttu-id="c93e0-929">对象</span><span class="sxs-lookup"><span data-stu-id="c93e0-929">Object</span></span>| <span data-ttu-id="c93e0-930">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="c93e0-930">&lt;optional&gt;</span></span>|<span data-ttu-id="c93e0-931">开发人员可以提供他们想要在回调函数中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="c93e0-931">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="c93e0-932">此对象可以通过回调函数中的 `asyncResult.asyncContext` 属性进行访问。</span><span class="sxs-lookup"><span data-stu-id="c93e0-932">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c93e0-933">要求</span><span class="sxs-lookup"><span data-stu-id="c93e0-933">Requirements</span></span>

|<span data-ttu-id="c93e0-934">要求</span><span class="sxs-lookup"><span data-stu-id="c93e0-934">Requirement</span></span>| <span data-ttu-id="c93e0-935">值</span><span class="sxs-lookup"><span data-stu-id="c93e0-935">Value</span></span>|
|---|---|
|[<span data-ttu-id="c93e0-936">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="c93e0-936">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c93e0-937">1.0</span><span class="sxs-lookup"><span data-stu-id="c93e0-937">1.0</span></span>|
|[<span data-ttu-id="c93e0-938">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="c93e0-938">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c93e0-939">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c93e0-939">ReadItem</span></span>|
|[<span data-ttu-id="c93e0-940">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c93e0-940">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c93e0-941">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="c93e0-941">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="c93e0-942">示例</span><span class="sxs-lookup"><span data-stu-id="c93e0-942">Example</span></span>

<span data-ttu-id="c93e0-p164">以下代码示例显示了如何使用 `loadCustomPropertiesAsync` 方法异步加载特定于当前项目的自定义属性。该示例还显示了如何使用 `CustomProperties.saveAsync` 方法将这些属性重新保存到服务器。加载自定义属性后，该代码示例将使用 `CustomProperties.get` 方法读取自定义属性 `myProp`，使用 `CustomProperties.set` 方法写入自定义属性 `otherProp`，最后调用 `saveAsync` 方法保存这些自定义属性。</span><span class="sxs-lookup"><span data-stu-id="c93e0-p164">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

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

####  <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="c93e0-946">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="c93e0-946">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="c93e0-947">将附件从邮件或约会中删除。</span><span class="sxs-lookup"><span data-stu-id="c93e0-947">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="c93e0-p165">`removeAttachmentAsync` 方法删除项目中带指定标识符的附件。最佳做法是，仅当同一个邮件应用程序在同一会话中添加了一个附件时，你才应使用该附件标识符来删除该附件。在 Outlook Web App 和适用于设备的 OWA 中，附件标识符只在同一个会话中才有效。当用户关闭应用程序，或者如果用户开始在内嵌窗体中撰写，并在随后弹出的内嵌窗体中继续在单独的窗口撰写时，会话即结束。</span><span class="sxs-lookup"><span data-stu-id="c93e0-p165">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item. As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session. In Outlook Web App and OWA for Devices, the attachment identifier is valid only within the same session. A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c93e0-952">参数：</span><span class="sxs-lookup"><span data-stu-id="c93e0-952">Parameters:</span></span>

|<span data-ttu-id="c93e0-953">名称</span><span class="sxs-lookup"><span data-stu-id="c93e0-953">Name</span></span>| <span data-ttu-id="c93e0-954">类型</span><span class="sxs-lookup"><span data-stu-id="c93e0-954">Type</span></span>| <span data-ttu-id="c93e0-955">属性</span><span class="sxs-lookup"><span data-stu-id="c93e0-955">Attributes</span></span>| <span data-ttu-id="c93e0-956">说明</span><span class="sxs-lookup"><span data-stu-id="c93e0-956">Description</span></span>|
|---|---|---|---|
|`attachmentId`| <span data-ttu-id="c93e0-957">字符串</span><span class="sxs-lookup"><span data-stu-id="c93e0-957">String</span></span>||<span data-ttu-id="c93e0-958">要删除的附件的标识符。</span><span class="sxs-lookup"><span data-stu-id="c93e0-958">The identifier of the attachment to remove.</span></span>|
|`options`| <span data-ttu-id="c93e0-959">对象</span><span class="sxs-lookup"><span data-stu-id="c93e0-959">Object</span></span>| <span data-ttu-id="c93e0-960">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="c93e0-960">&lt;optional&gt;</span></span>|<span data-ttu-id="c93e0-961">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="c93e0-961">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="c93e0-962">对象</span><span class="sxs-lookup"><span data-stu-id="c93e0-962">Object</span></span>| <span data-ttu-id="c93e0-963">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="c93e0-963">&lt;optional&gt;</span></span>|<span data-ttu-id="c93e0-964">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="c93e0-964">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="c93e0-965">函数</span><span class="sxs-lookup"><span data-stu-id="c93e0-965">function</span></span>| <span data-ttu-id="c93e0-966">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="c93e0-966">&lt;optional&gt;</span></span>|<span data-ttu-id="c93e0-967">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="c93e0-967">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="c93e0-968">如果删除附件失败，`asyncResult.error` 属性将包含一个说明失败原因的错误代码。</span><span class="sxs-lookup"><span data-stu-id="c93e0-968">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="c93e0-969">错误</span><span class="sxs-lookup"><span data-stu-id="c93e0-969">Errors</span></span>

| <span data-ttu-id="c93e0-970">错误代码</span><span class="sxs-lookup"><span data-stu-id="c93e0-970">Error code</span></span> | <span data-ttu-id="c93e0-971">说明</span><span class="sxs-lookup"><span data-stu-id="c93e0-971">Description</span></span> |
|------------|-------------|
| `InvalidAttachmentId` | <span data-ttu-id="c93e0-972">附件标识符不存在。</span><span class="sxs-lookup"><span data-stu-id="c93e0-972">The attachment identifier does not exist.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="c93e0-973">要求</span><span class="sxs-lookup"><span data-stu-id="c93e0-973">Requirements</span></span>

|<span data-ttu-id="c93e0-974">要求</span><span class="sxs-lookup"><span data-stu-id="c93e0-974">Requirement</span></span>| <span data-ttu-id="c93e0-975">值</span><span class="sxs-lookup"><span data-stu-id="c93e0-975">Value</span></span>|
|---|---|
|[<span data-ttu-id="c93e0-976">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="c93e0-976">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c93e0-977">1.1</span><span class="sxs-lookup"><span data-stu-id="c93e0-977">1.1</span></span>|
|[<span data-ttu-id="c93e0-978">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="c93e0-978">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c93e0-979">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="c93e0-979">ReadWriteItem</span></span>|
|[<span data-ttu-id="c93e0-980">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c93e0-980">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c93e0-981">撰写</span><span class="sxs-lookup"><span data-stu-id="c93e0-981">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="c93e0-982">示例</span><span class="sxs-lookup"><span data-stu-id="c93e0-982">Example</span></span>

<span data-ttu-id="c93e0-983">以下代码删除包含标识符 '0' 的附件。</span><span class="sxs-lookup"><span data-stu-id="c93e0-983">The following code removes an attachment with an identifier of '0'.</span></span>

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

####  <a name="saveasyncoptions-callback"></a><span data-ttu-id="c93e0-984">saveAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="c93e0-984">saveAsync([options], callback)</span></span>

<span data-ttu-id="c93e0-985">异步保存项目。</span><span class="sxs-lookup"><span data-stu-id="c93e0-985">Asynchronously saves an item.</span></span>

<span data-ttu-id="c93e0-p166">调用时，此方法将当前邮件保存为草稿，并通过回调方法返回项目 ID。在 Outlook Web App 或 Outlook 联机模式下，该项目被保存到服务器中。在 Outlook 缓存模式下，该项目被保存到本地缓存中。</span><span class="sxs-lookup"><span data-stu-id="c93e0-p166">When invoked, this method saves the current message as a draft and returns the item id via the callback method. In Outlook Web App or Outlook in online mode, the item is saved to the server. In Outlook in cached mode, the item is saved to the local cache.</span></span>

> [!NOTE]
> <span data-ttu-id="c93e0-989">如果加载项在撰写模式下对某个项目调用 `saveAsync` 来获得 `itemId`，以便与 EWS 或 REST API 一同使用，请注意，当 Outlook 处于高速缓存模式时，可能需要一段时间项目才能真正同步到服务器。</span><span class="sxs-lookup"><span data-stu-id="c93e0-989">If your add-in calls `saveAsync` on an item in compose mode in order to get an `itemId` to use with EWS or the REST API, be aware that when Outlook is in cached mode, it may take some time before the item is actually synced to the server.</span></span> <span data-ttu-id="c93e0-990">在项目同步前，使用 `itemId` 将返回一个错误。</span><span class="sxs-lookup"><span data-stu-id="c93e0-990">Until the item is synced, using the `itemId` will return an error.</span></span>

<span data-ttu-id="c93e0-p168">由于约会没有草稿状态，如果以撰写模式在约会中调用 `saveAsync`，则该项将被保存为用户日历中的正常约会。对于之前未保存过的新约会，则不会发送邀请。保存现有约会将向添加或删除的与会者发送更新。</span><span class="sxs-lookup"><span data-stu-id="c93e0-p168">Since appointments have no draft state, if `saveAsync` is called on an appointment in compose mode, the item will be saved as a normal appointment on the user's calendar. For new appointments that have not been saved before, no invitation will be sent. Saving an existing appointment will send an update to added or removed attendees.</span></span>

> [!NOTE]
> <span data-ttu-id="c93e0-994">以下客户端在撰写模式下对约会上的 `saveAsync` 具有不同的行为：</span><span class="sxs-lookup"><span data-stu-id="c93e0-994">The following clients have different behavior for `saveAsync` on appointments in compose mode:</span></span>
>
> - <span data-ttu-id="c93e0-995">Mac Outlook 不支持在撰写模式下对会议执行 `saveAsync` 操作。</span><span class="sxs-lookup"><span data-stu-id="c93e0-995">Mac Outlook does not support `saveAsync` on a meeting in compose mode.</span></span> <span data-ttu-id="c93e0-996">对 Mac Outlook 中的会议调用 `saveAsync` 将会返回一个错误。</span><span class="sxs-lookup"><span data-stu-id="c93e0-996">Calling `saveAsync` on a meeting in Mac Outlook will return an error.</span></span>
> - <span data-ttu-id="c93e0-997">在撰写模式下的约会上调用 `saveAsync` 时，Outlook 网页版始终发送邀请或更新。</span><span class="sxs-lookup"><span data-stu-id="c93e0-997">Outlook on the web always sends an invitation or update when `saveAsync` is called on an appointment in compose mode.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c93e0-998">参数：</span><span class="sxs-lookup"><span data-stu-id="c93e0-998">Parameters:</span></span>

|<span data-ttu-id="c93e0-999">名称</span><span class="sxs-lookup"><span data-stu-id="c93e0-999">Name</span></span>| <span data-ttu-id="c93e0-1000">类型</span><span class="sxs-lookup"><span data-stu-id="c93e0-1000">Type</span></span>| <span data-ttu-id="c93e0-1001">属性</span><span class="sxs-lookup"><span data-stu-id="c93e0-1001">Attributes</span></span>| <span data-ttu-id="c93e0-1002">说明</span><span class="sxs-lookup"><span data-stu-id="c93e0-1002">Description</span></span>|
|---|---|---|---|
|`options`| <span data-ttu-id="c93e0-1003">对象</span><span class="sxs-lookup"><span data-stu-id="c93e0-1003">Object</span></span>| <span data-ttu-id="c93e0-1004">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="c93e0-1004">&lt;optional&gt;</span></span>|<span data-ttu-id="c93e0-1005">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="c93e0-1005">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="c93e0-1006">对象</span><span class="sxs-lookup"><span data-stu-id="c93e0-1006">Object</span></span>| <span data-ttu-id="c93e0-1007">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="c93e0-1007">&lt;optional&gt;</span></span>|<span data-ttu-id="c93e0-1008">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="c93e0-1008">Developers can provide any object they wish to access in the callback method.</span></span>||
|`callback`| <span data-ttu-id="c93e0-1009">函数</span><span class="sxs-lookup"><span data-stu-id="c93e0-1009">function</span></span>||<span data-ttu-id="c93e0-1010">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="c93e0-1010">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="c93e0-1011">如果成功，该项目标识符将在 `asyncResult.value` 属性中提供。</span><span class="sxs-lookup"><span data-stu-id="c93e0-1011">On success, the item identifier is provided in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c93e0-1012">要求</span><span class="sxs-lookup"><span data-stu-id="c93e0-1012">Requirements</span></span>

|<span data-ttu-id="c93e0-1013">要求</span><span class="sxs-lookup"><span data-stu-id="c93e0-1013">Requirement</span></span>| <span data-ttu-id="c93e0-1014">值</span><span class="sxs-lookup"><span data-stu-id="c93e0-1014">Value</span></span>|
|---|---|
|[<span data-ttu-id="c93e0-1015">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="c93e0-1015">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c93e0-1016">1.3</span><span class="sxs-lookup"><span data-stu-id="c93e0-1016">1.3</span></span>|
|[<span data-ttu-id="c93e0-1017">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="c93e0-1017">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c93e0-1018">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="c93e0-1018">ReadWriteItem</span></span>|
|[<span data-ttu-id="c93e0-1019">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c93e0-1019">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c93e0-1020">撰写</span><span class="sxs-lookup"><span data-stu-id="c93e0-1020">Compose</span></span>|

##### <a name="examples"></a><span data-ttu-id="c93e0-1021">示例</span><span class="sxs-lookup"><span data-stu-id="c93e0-1021">Examples</span></span>

```js
Office.context.mailbox.item.saveAsync(
  function callback(result) {
    // Process the result
  });
```

<span data-ttu-id="c93e0-p170">下面是传递给回调函数的 `result` 参数的示例。`value` 属性包含的项目的项目 ID。</span><span class="sxs-lookup"><span data-stu-id="c93e0-p170">The following is an example of the `result` parameter passed to the callback function. The `value` property contains the item ID of the item.</span></span>

```json
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

####  <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="c93e0-1024">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="c93e0-1024">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="c93e0-1025">以异步方式将数据插入到邮件的正文或主题中。</span><span class="sxs-lookup"><span data-stu-id="c93e0-1025">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="c93e0-p171">`setSelectedDataAsync` 方法将指定字符串插入到项目主题或正文的光标位置，或者，如果在编辑器中已选择文本，则该方法将替换选择的文本。如果光标不在正文或主题字段中，则返回错误。插入之后，光标会位于插入内容的末尾。</span><span class="sxs-lookup"><span data-stu-id="c93e0-p171">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="c93e0-1029">参数：</span><span class="sxs-lookup"><span data-stu-id="c93e0-1029">Parameters:</span></span>

|<span data-ttu-id="c93e0-1030">名称</span><span class="sxs-lookup"><span data-stu-id="c93e0-1030">Name</span></span>| <span data-ttu-id="c93e0-1031">类型</span><span class="sxs-lookup"><span data-stu-id="c93e0-1031">Type</span></span>| <span data-ttu-id="c93e0-1032">属性</span><span class="sxs-lookup"><span data-stu-id="c93e0-1032">Attributes</span></span>| <span data-ttu-id="c93e0-1033">说明</span><span class="sxs-lookup"><span data-stu-id="c93e0-1033">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="c93e0-1034">字符串</span><span class="sxs-lookup"><span data-stu-id="c93e0-1034">String</span></span>||<span data-ttu-id="c93e0-p172">要插入的数据。数据不得超过 1,000,000 个字符。如果传入的数据超过 1,000,000 个字符，则会引发 `ArgumentOutOfRange` 异常。</span><span class="sxs-lookup"><span data-stu-id="c93e0-p172">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`| <span data-ttu-id="c93e0-1038">Object</span><span class="sxs-lookup"><span data-stu-id="c93e0-1038">Object</span></span>| <span data-ttu-id="c93e0-1039">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="c93e0-1039">&lt;optional&gt;</span></span>|<span data-ttu-id="c93e0-1040">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="c93e0-1040">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="c93e0-1041">对象</span><span class="sxs-lookup"><span data-stu-id="c93e0-1041">Object</span></span>| <span data-ttu-id="c93e0-1042">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="c93e0-1042">&lt;optional&gt;</span></span>|<span data-ttu-id="c93e0-1043">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="c93e0-1043">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`| [<span data-ttu-id="c93e0-1044">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="c93e0-1044">Office.CoercionType</span></span>](office.md#coerciontype-string)| <span data-ttu-id="c93e0-1045">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="c93e0-1045">&lt;optional&gt;</span></span>|<span data-ttu-id="c93e0-p173">如果为 `text`，则在 Outlook Web App 和 Outlook 中应用当前样式。如果该字段是 HTML 编辑器，则仅插入文本数据，即使数据为 HTML。</span><span class="sxs-lookup"><span data-stu-id="c93e0-p173">If `text`, the current style is applied in Outlook Web App and Outlook. If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="c93e0-p174">如果 `html` 和该字段支持 HTML（主题不支持），则在 Outlook Web App 中应用当前样式，而在 Outlook 中应用默认样式。如果该字段是文本字段，则返回 `InvalidDataFormat` 错误。</span><span class="sxs-lookup"><span data-stu-id="c93e0-p174">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook Web App and the default style is applied in Outlook. If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="c93e0-1050">如果未设置 `coercionType`，则结果取决于该字段：如果该字段是 HTML，则使用 HTML；如果该字段是文本，则使用纯文本。</span><span class="sxs-lookup"><span data-stu-id="c93e0-1050">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`| <span data-ttu-id="c93e0-1051">function</span><span class="sxs-lookup"><span data-stu-id="c93e0-1051">function</span></span>||<span data-ttu-id="c93e0-1052">方法完成后，使用单个参数 `callback`（一个 [`asyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `AsyncResult` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="c93e0-1052">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="c93e0-1053">Requirements</span><span class="sxs-lookup"><span data-stu-id="c93e0-1053">Requirements</span></span>

|<span data-ttu-id="c93e0-1054">要求</span><span class="sxs-lookup"><span data-stu-id="c93e0-1054">Requirement</span></span>| <span data-ttu-id="c93e0-1055">值</span><span class="sxs-lookup"><span data-stu-id="c93e0-1055">Value</span></span>|
|---|---|
|[<span data-ttu-id="c93e0-1056">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="c93e0-1056">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c93e0-1057">1.2</span><span class="sxs-lookup"><span data-stu-id="c93e0-1057">1.2</span></span>|
|[<span data-ttu-id="c93e0-1058">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="c93e0-1058">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c93e0-1059">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="c93e0-1059">ReadWriteItem</span></span>|
|[<span data-ttu-id="c93e0-1060">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="c93e0-1060">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c93e0-1061">撰写</span><span class="sxs-lookup"><span data-stu-id="c93e0-1061">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="c93e0-1062">示例</span><span class="sxs-lookup"><span data-stu-id="c93e0-1062">Example</span></span>

```js
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```
