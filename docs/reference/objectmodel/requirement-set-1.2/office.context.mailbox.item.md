---
title: Office.context.mailbox.item-要求设置 1.2
description: ''
ms.date: 12/18/2018
localization_priority: Normal
ms.openlocfilehash: d58a38ce045a179a7e5cdd2e15b4e16c2ac03c91
ms.sourcegitcommit: d1aa7201820176ed986b9f00bb9c88e055906c77
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 01/23/2019
ms.locfileid: "29388596"
---
# <a name="item"></a><span data-ttu-id="b4778-102">item</span><span class="sxs-lookup"><span data-stu-id="b4778-102">item</span></span>

### <span data-ttu-id="b4778-p101">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md). item</span><span class="sxs-lookup"><span data-stu-id="b4778-p101">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md). item</span></span>

<span data-ttu-id="b4778-p102">`item` 命名空间用于访问当前选定的邮件、会议请求或约会。可以通过使用 [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlook12officemailboxenumsitemtype) 属性确定 `item` 的类型。</span><span class="sxs-lookup"><span data-stu-id="b4778-p102">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlook12officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="b4778-107">Requirements</span><span class="sxs-lookup"><span data-stu-id="b4778-107">Requirements</span></span>

|<span data-ttu-id="b4778-108">要求</span><span class="sxs-lookup"><span data-stu-id="b4778-108">Requirement</span></span>| <span data-ttu-id="b4778-109">值</span><span class="sxs-lookup"><span data-stu-id="b4778-109">Value</span></span>|
|---|---|
|[<span data-ttu-id="b4778-110">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="b4778-110">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b4778-111">1.0</span><span class="sxs-lookup"><span data-stu-id="b4778-111">1.0</span></span>|
|[<span data-ttu-id="b4778-112">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="b4778-112">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b4778-113">受限</span><span class="sxs-lookup"><span data-stu-id="b4778-113">Restricted</span></span>|
|[<span data-ttu-id="b4778-114">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="b4778-114">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b4778-115">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="b4778-115">Compose or read</span></span>|

### <a name="example"></a><span data-ttu-id="b4778-116">示例</span><span class="sxs-lookup"><span data-stu-id="b4778-116">Example</span></span>

<span data-ttu-id="b4778-117">以下 JavaScript 代码示例显示了如何访问 Outlook 中当前项目的 `subject` 属性。</span><span class="sxs-lookup"><span data-stu-id="b4778-117">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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

### <a name="members"></a><span data-ttu-id="b4778-118">成员</span><span class="sxs-lookup"><span data-stu-id="b4778-118">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlook12officeattachmentdetails"></a><span data-ttu-id="b4778-119">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_2/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="b4778-119">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_2/office.attachmentdetails)></span></span>

<span data-ttu-id="b4778-p103">获取项目的附件数组。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="b4778-p103">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="b4778-122">由于潜在的安全问题，某些类型的文件会受到 Outlook 阻止，并且不会返回。</span><span class="sxs-lookup"><span data-stu-id="b4778-122">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="b4778-123">如需了解更多信息，请参阅 [Outlook 中阻止的附件](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519)。</span><span class="sxs-lookup"><span data-stu-id="b4778-123">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="b4778-124">类型：</span><span class="sxs-lookup"><span data-stu-id="b4778-124">Type:</span></span>

*   <span data-ttu-id="b4778-125">Array.<[AttachmentDetails](/javascript/api/outlook_1_2/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="b4778-125">Array.<[AttachmentDetails](/javascript/api/outlook_1_2/office.attachmentdetails)></span></span>

##### <a name="requirements"></a><span data-ttu-id="b4778-126">要求</span><span class="sxs-lookup"><span data-stu-id="b4778-126">Requirements</span></span>

|<span data-ttu-id="b4778-127">要求</span><span class="sxs-lookup"><span data-stu-id="b4778-127">Requirement</span></span>| <span data-ttu-id="b4778-128">值</span><span class="sxs-lookup"><span data-stu-id="b4778-128">Value</span></span>|
|---|---|
|[<span data-ttu-id="b4778-129">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="b4778-129">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b4778-130">1.0</span><span class="sxs-lookup"><span data-stu-id="b4778-130">1.0</span></span>|
|[<span data-ttu-id="b4778-131">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="b4778-131">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b4778-132">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b4778-132">ReadItem</span></span>|
|[<span data-ttu-id="b4778-133">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="b4778-133">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b4778-134">阅读</span><span class="sxs-lookup"><span data-stu-id="b4778-134">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="b4778-135">示例</span><span class="sxs-lookup"><span data-stu-id="b4778-135">Example</span></span>

<span data-ttu-id="b4778-136">以下代码使用当前项目上所有附件的详细信息构成 HTML 字符串。</span><span class="sxs-lookup"><span data-stu-id="b4778-136">The following code builds an HTML string with details of all attachments on the current item.</span></span>

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

####  <a name="bcc-recipientsjavascriptapioutlook12officerecipients"></a><span data-ttu-id="b4778-137">bcc :[Recipients](/javascript/api/outlook_1_2/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="b4778-137">bcc :[Recipients](/javascript/api/outlook_1_2/office.recipients)</span></span>

<span data-ttu-id="b4778-138">获取对象，该对象提供用于获取或更新邮件的密件抄送 (Bcc) 行上的收件人的方法。</span><span class="sxs-lookup"><span data-stu-id="b4778-138">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="b4778-139">仅限撰写模式。</span><span class="sxs-lookup"><span data-stu-id="b4778-139">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="b4778-140">类型：</span><span class="sxs-lookup"><span data-stu-id="b4778-140">Type:</span></span>

*   [<span data-ttu-id="b4778-141">收件人</span><span class="sxs-lookup"><span data-stu-id="b4778-141">Recipients</span></span>](/javascript/api/outlook_1_2/office.recipients)

##### <a name="requirements"></a><span data-ttu-id="b4778-142">要求</span><span class="sxs-lookup"><span data-stu-id="b4778-142">Requirements</span></span>

|<span data-ttu-id="b4778-143">要求</span><span class="sxs-lookup"><span data-stu-id="b4778-143">Requirement</span></span>| <span data-ttu-id="b4778-144">值</span><span class="sxs-lookup"><span data-stu-id="b4778-144">Value</span></span>|
|---|---|
|[<span data-ttu-id="b4778-145">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="b4778-145">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b4778-146">1.1</span><span class="sxs-lookup"><span data-stu-id="b4778-146">1.1</span></span>|
|[<span data-ttu-id="b4778-147">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="b4778-147">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b4778-148">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b4778-148">ReadItem</span></span>|
|[<span data-ttu-id="b4778-149">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="b4778-149">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b4778-150">撰写</span><span class="sxs-lookup"><span data-stu-id="b4778-150">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="b4778-151">示例</span><span class="sxs-lookup"><span data-stu-id="b4778-151">Example</span></span>

```JavaScript
Office.context.mailbox.item.bcc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.bcc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.bcc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfBccRecipients = asyncResult.value;
}
```

####  <a name="body-bodyjavascriptapioutlook12officebody"></a><span data-ttu-id="b4778-152">body :[Body](/javascript/api/outlook_1_2/office.body)</span><span class="sxs-lookup"><span data-stu-id="b4778-152">body :[Body](/javascript/api/outlook_1_2/office.body)</span></span>

<span data-ttu-id="b4778-153">获取一个提供用于处理项目正文的方法的对象。</span><span class="sxs-lookup"><span data-stu-id="b4778-153">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="b4778-154">类型：</span><span class="sxs-lookup"><span data-stu-id="b4778-154">Type:</span></span>

*   [<span data-ttu-id="b4778-155">Body</span><span class="sxs-lookup"><span data-stu-id="b4778-155">Body</span></span>](/javascript/api/outlook_1_2/office.body)

##### <a name="requirements"></a><span data-ttu-id="b4778-156">要求</span><span class="sxs-lookup"><span data-stu-id="b4778-156">Requirements</span></span>

|<span data-ttu-id="b4778-157">要求</span><span class="sxs-lookup"><span data-stu-id="b4778-157">Requirement</span></span>| <span data-ttu-id="b4778-158">值</span><span class="sxs-lookup"><span data-stu-id="b4778-158">Value</span></span>|
|---|---|
|[<span data-ttu-id="b4778-159">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="b4778-159">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b4778-160">1.1</span><span class="sxs-lookup"><span data-stu-id="b4778-160">1.1</span></span>|
|[<span data-ttu-id="b4778-161">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="b4778-161">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b4778-162">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b4778-162">ReadItem</span></span>|
|[<span data-ttu-id="b4778-163">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="b4778-163">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b4778-164">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="b4778-164">Compose or read</span></span>|

####  <a name="cc-arrayemailaddressdetailsjavascriptapioutlook12officeemailaddressdetailsrecipientsjavascriptapioutlook12officerecipients"></a><span data-ttu-id="b4778-165">cc :Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_2/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="b4778-165">cc :Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_2/office.recipients)</span></span>

<span data-ttu-id="b4778-166">提供对邮件的抄送 (Cc) 收件人的访问权限。</span><span class="sxs-lookup"><span data-stu-id="b4778-166">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="b4778-167">对象的类型和访问级别取决于当前项目的模式。</span><span class="sxs-lookup"><span data-stu-id="b4778-167">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="b4778-168">阅读模式</span><span class="sxs-lookup"><span data-stu-id="b4778-168">Read mode</span></span>

<span data-ttu-id="b4778-p107">`cc` 属性返回包含邮件的**抄送**行上所列的每个收件人的 `EmailAddressDetails` 对象的数组。集合上限为 100 个成员。</span><span class="sxs-lookup"><span data-stu-id="b4778-p107">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="b4778-171">撰写模式</span><span class="sxs-lookup"><span data-stu-id="b4778-171">Compose mode</span></span>

<span data-ttu-id="b4778-172">`cc` 属性返回一个 `Recipients` 对象，该对象提供用于获取或更新邮件的“**抄送**”行上收件人的方法。</span><span class="sxs-lookup"><span data-stu-id="b4778-172">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="b4778-173">类型：</span><span class="sxs-lookup"><span data-stu-id="b4778-173">Type:</span></span>

*   <span data-ttu-id="b4778-174">Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_2/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="b4778-174">Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_2/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="b4778-175">要求</span><span class="sxs-lookup"><span data-stu-id="b4778-175">Requirements</span></span>

|<span data-ttu-id="b4778-176">要求</span><span class="sxs-lookup"><span data-stu-id="b4778-176">Requirement</span></span>| <span data-ttu-id="b4778-177">值</span><span class="sxs-lookup"><span data-stu-id="b4778-177">Value</span></span>|
|---|---|
|[<span data-ttu-id="b4778-178">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="b4778-178">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b4778-179">1.0</span><span class="sxs-lookup"><span data-stu-id="b4778-179">1.0</span></span>|
|[<span data-ttu-id="b4778-180">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="b4778-180">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b4778-181">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b4778-181">ReadItem</span></span>|
|[<span data-ttu-id="b4778-182">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="b4778-182">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b4778-183">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="b4778-183">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="b4778-184">示例</span><span class="sxs-lookup"><span data-stu-id="b4778-184">Example</span></span>

```JavaScript
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

####  <a name="nullable-conversationid-string"></a><span data-ttu-id="b4778-185">(nullable) conversationId :String</span><span class="sxs-lookup"><span data-stu-id="b4778-185">(nullable) conversationId :String</span></span>

<span data-ttu-id="b4778-186">获取包含特定消息的电子邮件会话的标识符。</span><span class="sxs-lookup"><span data-stu-id="b4778-186">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="b4778-p108">如果在阅读窗体或撰写窗体的回复中激活邮件应用程序，则此属性可以获得一个整数值。如果用户随后更改了回复邮件的主题（若发送回复），则该邮件的对话 ID 将改变且之前获取的值将不适用。</span><span class="sxs-lookup"><span data-stu-id="b4778-p108">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="b4778-p109">对于撰写窗体的新项目，此属性获得一个 null 值。如果用户设置一个主题并保存该项目，`conversationId` 属性将返回一个值。</span><span class="sxs-lookup"><span data-stu-id="b4778-p109">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="b4778-191">类型:</span><span class="sxs-lookup"><span data-stu-id="b4778-191">Type:</span></span>

*   <span data-ttu-id="b4778-192">String</span><span class="sxs-lookup"><span data-stu-id="b4778-192">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="b4778-193">要求</span><span class="sxs-lookup"><span data-stu-id="b4778-193">Requirements</span></span>

|<span data-ttu-id="b4778-194">要求</span><span class="sxs-lookup"><span data-stu-id="b4778-194">Requirement</span></span>| <span data-ttu-id="b4778-195">值</span><span class="sxs-lookup"><span data-stu-id="b4778-195">Value</span></span>|
|---|---|
|[<span data-ttu-id="b4778-196">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="b4778-196">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b4778-197">1.0</span><span class="sxs-lookup"><span data-stu-id="b4778-197">1.0</span></span>|
|[<span data-ttu-id="b4778-198">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="b4778-198">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b4778-199">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b4778-199">ReadItem</span></span>|
|[<span data-ttu-id="b4778-200">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="b4778-200">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b4778-201">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="b4778-201">Compose or read</span></span>|

#### <a name="datetimecreated-date"></a><span data-ttu-id="b4778-202">dateTimeCreated :Date</span><span class="sxs-lookup"><span data-stu-id="b4778-202">dateTimeCreated :Date</span></span>

<span data-ttu-id="b4778-p110">获取项目创建的日期和时间。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="b4778-p110">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="b4778-205">类型：</span><span class="sxs-lookup"><span data-stu-id="b4778-205">Type:</span></span>

*   <span data-ttu-id="b4778-206">日期</span><span class="sxs-lookup"><span data-stu-id="b4778-206">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="b4778-207">要求</span><span class="sxs-lookup"><span data-stu-id="b4778-207">Requirements</span></span>

|<span data-ttu-id="b4778-208">要求</span><span class="sxs-lookup"><span data-stu-id="b4778-208">Requirement</span></span>| <span data-ttu-id="b4778-209">值</span><span class="sxs-lookup"><span data-stu-id="b4778-209">Value</span></span>|
|---|---|
|[<span data-ttu-id="b4778-210">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="b4778-210">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b4778-211">1.0</span><span class="sxs-lookup"><span data-stu-id="b4778-211">1.0</span></span>|
|[<span data-ttu-id="b4778-212">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="b4778-212">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b4778-213">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b4778-213">ReadItem</span></span>|
|[<span data-ttu-id="b4778-214">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="b4778-214">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b4778-215">阅读</span><span class="sxs-lookup"><span data-stu-id="b4778-215">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="b4778-216">示例</span><span class="sxs-lookup"><span data-stu-id="b4778-216">Example</span></span>

```JavaScript
var created = Office.context.mailbox.item.dateTimeCreated;
```

#### <a name="datetimemodified-date"></a><span data-ttu-id="b4778-217">dateTimeModified :Date</span><span class="sxs-lookup"><span data-stu-id="b4778-217">dateTimeModified :Date</span></span>

<span data-ttu-id="b4778-p111">获取项目最近一次修改的日期和时间。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="b4778-p111">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="b4778-220">在 Outlook for iOS 或 Outlook for Android 中不支持此成员。</span><span class="sxs-lookup"><span data-stu-id="b4778-220">This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="type"></a><span data-ttu-id="b4778-221">类型:</span><span class="sxs-lookup"><span data-stu-id="b4778-221">Type:</span></span>

*   <span data-ttu-id="b4778-222">日期</span><span class="sxs-lookup"><span data-stu-id="b4778-222">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="b4778-223">要求</span><span class="sxs-lookup"><span data-stu-id="b4778-223">Requirements</span></span>

|<span data-ttu-id="b4778-224">要求</span><span class="sxs-lookup"><span data-stu-id="b4778-224">Requirement</span></span>| <span data-ttu-id="b4778-225">值</span><span class="sxs-lookup"><span data-stu-id="b4778-225">Value</span></span>|
|---|---|
|[<span data-ttu-id="b4778-226">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="b4778-226">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b4778-227">1.0</span><span class="sxs-lookup"><span data-stu-id="b4778-227">1.0</span></span>|
|[<span data-ttu-id="b4778-228">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="b4778-228">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b4778-229">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b4778-229">ReadItem</span></span>|
|[<span data-ttu-id="b4778-230">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="b4778-230">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b4778-231">阅读</span><span class="sxs-lookup"><span data-stu-id="b4778-231">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="b4778-232">示例</span><span class="sxs-lookup"><span data-stu-id="b4778-232">Example</span></span>

```JavaScript
var modified = Office.context.mailbox.item.dateTimeModified;
```

####  <a name="end-datetimejavascriptapioutlook12officetime"></a><span data-ttu-id="b4778-233">end :Date|[Time](/javascript/api/outlook_1_2/office.time)</span><span class="sxs-lookup"><span data-stu-id="b4778-233">end :Date|[Time](/javascript/api/outlook_1_2/office.time)</span></span>

<span data-ttu-id="b4778-234">获取或设置约会结束的日期和时间。</span><span class="sxs-lookup"><span data-stu-id="b4778-234">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="b4778-p112">将 `end` 属性表示为协调世界时 (UTC) 的日期和时间值。可使用 [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook12officelocalclienttime) 方法将 end 属性值转换为客户端的本地日期和时间。</span><span class="sxs-lookup"><span data-stu-id="b4778-p112">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook12officelocalclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="b4778-237">阅读模式</span><span class="sxs-lookup"><span data-stu-id="b4778-237">Read mode</span></span>

<span data-ttu-id="b4778-238">`end` 属性返回 `Date` 对象。</span><span class="sxs-lookup"><span data-stu-id="b4778-238">The `end` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="b4778-239">撰写模式</span><span class="sxs-lookup"><span data-stu-id="b4778-239">Compose mode</span></span>

<span data-ttu-id="b4778-240">`end` 属性返回 `Time` 对象。</span><span class="sxs-lookup"><span data-stu-id="b4778-240">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="b4778-241">使用 [`Time.setAsync`](/javascript/api/outlook_1_2/office.time#setasync-datetime--options--callback-) 方法设置结束时间时，应使用 [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) 方法将客户端的本地时间转换为服务器的 UTC。</span><span class="sxs-lookup"><span data-stu-id="b4778-241">When you use the [`Time.setAsync`](/javascript/api/outlook_1_2/office.time#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="b4778-242">类型：</span><span class="sxs-lookup"><span data-stu-id="b4778-242">Type:</span></span>

*   <span data-ttu-id="b4778-243">Date | [Time](/javascript/api/outlook_1_2/office.time)</span><span class="sxs-lookup"><span data-stu-id="b4778-243">Date | [Time](/javascript/api/outlook_1_2/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="b4778-244">要求</span><span class="sxs-lookup"><span data-stu-id="b4778-244">Requirements</span></span>

|<span data-ttu-id="b4778-245">要求</span><span class="sxs-lookup"><span data-stu-id="b4778-245">Requirement</span></span>| <span data-ttu-id="b4778-246">值</span><span class="sxs-lookup"><span data-stu-id="b4778-246">Value</span></span>|
|---|---|
|[<span data-ttu-id="b4778-247">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="b4778-247">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b4778-248">1.0</span><span class="sxs-lookup"><span data-stu-id="b4778-248">1.0</span></span>|
|[<span data-ttu-id="b4778-249">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="b4778-249">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b4778-250">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b4778-250">ReadItem</span></span>|
|[<span data-ttu-id="b4778-251">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="b4778-251">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b4778-252">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="b4778-252">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="b4778-253">示例</span><span class="sxs-lookup"><span data-stu-id="b4778-253">Example</span></span>

<span data-ttu-id="b4778-254">以下示例通过使用 `Time` 对象的 [`setAsync`](/javascript/api/outlook_1_2/office.time#setasync-datetime--options--callback-) 方法，设置撰写模式下约会的结束时间。</span><span class="sxs-lookup"><span data-stu-id="b4778-254">The following example sets the end time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_2/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

#### <a name="from-emailaddressdetailsjavascriptapioutlook12officeemailaddressdetails"></a><span data-ttu-id="b4778-255">from :[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="b4778-255">from :[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)</span></span>

<span data-ttu-id="b4778-p113">获取邮件发件人的电子邮件地址。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="b4778-p113">Gets the email address of the sender of a message. Read mode only.</span></span>

<span data-ttu-id="b4778-p114">`from` 和 [`sender`](#sender-emailaddressdetailsjavascriptapioutlook12officeemailaddressdetails) 属性表示同一个人，邮件由代理人发送的除外。在此情况下，`from` 属性表示代理程序，而 sender 属性表示代理人。</span><span class="sxs-lookup"><span data-stu-id="b4778-p114">The `from` and [`sender`](#sender-emailaddressdetailsjavascriptapioutlook12officeemailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="b4778-260">`from` 属性中 `EmailAddressDetails` 对象的 `recipientType` 属性为 `undefined`。</span><span class="sxs-lookup"><span data-stu-id="b4778-260">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="b4778-261">类型：</span><span class="sxs-lookup"><span data-stu-id="b4778-261">Type:</span></span>

*   [<span data-ttu-id="b4778-262">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="b4778-262">EmailAddressDetails</span></span>](/javascript/api/outlook_1_2/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="b4778-263">要求</span><span class="sxs-lookup"><span data-stu-id="b4778-263">Requirements</span></span>

|<span data-ttu-id="b4778-264">要求</span><span class="sxs-lookup"><span data-stu-id="b4778-264">Requirement</span></span>| <span data-ttu-id="b4778-265">值</span><span class="sxs-lookup"><span data-stu-id="b4778-265">Value</span></span>|
|---|---|
|[<span data-ttu-id="b4778-266">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="b4778-266">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b4778-267">1.0</span><span class="sxs-lookup"><span data-stu-id="b4778-267">1.0</span></span>|
|[<span data-ttu-id="b4778-268">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="b4778-268">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b4778-269">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b4778-269">ReadItem</span></span>|
|[<span data-ttu-id="b4778-270">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="b4778-270">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b4778-271">阅读</span><span class="sxs-lookup"><span data-stu-id="b4778-271">Read</span></span>|

#### <a name="internetmessageid-string"></a><span data-ttu-id="b4778-272">internetMessageId :String</span><span class="sxs-lookup"><span data-stu-id="b4778-272">internetMessageId :String</span></span>

<span data-ttu-id="b4778-p115">获取电子邮件的 Internet 消息标识符。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="b4778-p115">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="b4778-275">类型：</span><span class="sxs-lookup"><span data-stu-id="b4778-275">Type:</span></span>

*   <span data-ttu-id="b4778-276">String</span><span class="sxs-lookup"><span data-stu-id="b4778-276">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="b4778-277">要求</span><span class="sxs-lookup"><span data-stu-id="b4778-277">Requirements</span></span>

|<span data-ttu-id="b4778-278">要求</span><span class="sxs-lookup"><span data-stu-id="b4778-278">Requirement</span></span>| <span data-ttu-id="b4778-279">值</span><span class="sxs-lookup"><span data-stu-id="b4778-279">Value</span></span>|
|---|---|
|[<span data-ttu-id="b4778-280">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="b4778-280">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b4778-281">1.0</span><span class="sxs-lookup"><span data-stu-id="b4778-281">1.0</span></span>|
|[<span data-ttu-id="b4778-282">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="b4778-282">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b4778-283">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b4778-283">ReadItem</span></span>|
|[<span data-ttu-id="b4778-284">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="b4778-284">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b4778-285">阅读</span><span class="sxs-lookup"><span data-stu-id="b4778-285">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="b4778-286">示例</span><span class="sxs-lookup"><span data-stu-id="b4778-286">Example</span></span>

```JavaScript
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

#### <a name="itemclass-string"></a><span data-ttu-id="b4778-287">itemClass :String</span><span class="sxs-lookup"><span data-stu-id="b4778-287">itemClass :String</span></span>

<span data-ttu-id="b4778-p116">获取选定项目的 Exchange Web 服务项目类。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="b4778-p116">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="b4778-p117">`itemClass` 属性指定所选项目的邮件类别。以下是邮件或约会项目的默认邮件类别。</span><span class="sxs-lookup"><span data-stu-id="b4778-p117">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

| <span data-ttu-id="b4778-292">类型</span><span class="sxs-lookup"><span data-stu-id="b4778-292">Type</span></span> | <span data-ttu-id="b4778-293">说明</span><span class="sxs-lookup"><span data-stu-id="b4778-293">Description</span></span> | <span data-ttu-id="b4778-294">项目类</span><span class="sxs-lookup"><span data-stu-id="b4778-294">item class</span></span> |
| --- | --- | --- |
| <span data-ttu-id="b4778-295">约会项目</span><span class="sxs-lookup"><span data-stu-id="b4778-295">Appointment items</span></span> | <span data-ttu-id="b4778-296">这些是项目类 `IPM.Appointment` 或 `IPM.Appointment.Occurence` 的日历项目。</span><span class="sxs-lookup"><span data-stu-id="b4778-296">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurence`.</span></span> | `IPM.Appointment`<br />`IPM.Appointment.Occurence` |
| <span data-ttu-id="b4778-297">邮件项目</span><span class="sxs-lookup"><span data-stu-id="b4778-297">Message items</span></span> | <span data-ttu-id="b4778-298">这些项目包括具有默认邮件类别 `IPM.Note` 的电子邮件，以及使用 `IPM.Schedule.Meeting` 作为基础邮件类别的会议请求、响应和取消。</span><span class="sxs-lookup"><span data-stu-id="b4778-298">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span> | `IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled` |

<span data-ttu-id="b4778-299">你可以创建用于扩展默认邮件类别的自定义邮件类别，例如，自定义约会邮件类别 `IPM.Appointment.Contoso`。</span><span class="sxs-lookup"><span data-stu-id="b4778-299">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="b4778-300">类型：</span><span class="sxs-lookup"><span data-stu-id="b4778-300">Type:</span></span>

*   <span data-ttu-id="b4778-301">String</span><span class="sxs-lookup"><span data-stu-id="b4778-301">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="b4778-302">要求</span><span class="sxs-lookup"><span data-stu-id="b4778-302">Requirements</span></span>

|<span data-ttu-id="b4778-303">要求</span><span class="sxs-lookup"><span data-stu-id="b4778-303">Requirement</span></span>| <span data-ttu-id="b4778-304">值</span><span class="sxs-lookup"><span data-stu-id="b4778-304">Value</span></span>|
|---|---|
|[<span data-ttu-id="b4778-305">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="b4778-305">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b4778-306">1.0</span><span class="sxs-lookup"><span data-stu-id="b4778-306">1.0</span></span>|
|[<span data-ttu-id="b4778-307">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="b4778-307">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b4778-308">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b4778-308">ReadItem</span></span>|
|[<span data-ttu-id="b4778-309">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="b4778-309">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b4778-310">阅读</span><span class="sxs-lookup"><span data-stu-id="b4778-310">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="b4778-311">示例</span><span class="sxs-lookup"><span data-stu-id="b4778-311">Example</span></span>

```JavaScript
var itemClass = Office.context.mailbox.item.itemClass;
```

#### <a name="nullable-itemid-string"></a><span data-ttu-id="b4778-312">(nullable) itemId :String</span><span class="sxs-lookup"><span data-stu-id="b4778-312">(nullable) itemId :String</span></span>

<span data-ttu-id="b4778-p118">获取当前项目的 Exchange Web 服务项目标识符。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="b4778-p118">Gets the Exchange Web Services item identifier for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="b4778-315">`itemId` 属性返回的标识符与 Exchange Web 服务项目标识符相同。</span><span class="sxs-lookup"><span data-stu-id="b4778-315">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="b4778-316">`itemId` 属性与 Outlook 条目 ID 或 Outlook REST API 使用的 ID 不同。</span><span class="sxs-lookup"><span data-stu-id="b4778-316">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="b4778-317">使用此值进行 REST API 调用前，应使用 `Office.context.mailbox.convertToRestId`（可在要求集 1.3 的开头部分中找到）对它进行转换。</span><span class="sxs-lookup"><span data-stu-id="b4778-317">Before making REST API calls using this value, it should be converted using `Office.context.mailbox.convertToRestId`, which is available starting in requirement set 1.3.</span></span> <span data-ttu-id="b4778-318">有关详细信息，请参阅[从 Outlook 加载项使用 Outlook REST API](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id)。</span><span class="sxs-lookup"><span data-stu-id="b4778-318">For more details, see [Use the Outlook REST APIs from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

##### <a name="type"></a><span data-ttu-id="b4778-319">类型：</span><span class="sxs-lookup"><span data-stu-id="b4778-319">Type:</span></span>

*   <span data-ttu-id="b4778-320">String</span><span class="sxs-lookup"><span data-stu-id="b4778-320">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="b4778-321">要求</span><span class="sxs-lookup"><span data-stu-id="b4778-321">Requirements</span></span>

|<span data-ttu-id="b4778-322">要求</span><span class="sxs-lookup"><span data-stu-id="b4778-322">Requirement</span></span>| <span data-ttu-id="b4778-323">值</span><span class="sxs-lookup"><span data-stu-id="b4778-323">Value</span></span>|
|---|---|
|[<span data-ttu-id="b4778-324">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="b4778-324">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b4778-325">1.0</span><span class="sxs-lookup"><span data-stu-id="b4778-325">1.0</span></span>|
|[<span data-ttu-id="b4778-326">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="b4778-326">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b4778-327">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b4778-327">ReadItem</span></span>|
|[<span data-ttu-id="b4778-328">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="b4778-328">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b4778-329">阅读</span><span class="sxs-lookup"><span data-stu-id="b4778-329">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="b4778-330">示例</span><span class="sxs-lookup"><span data-stu-id="b4778-330">Example</span></span>

<span data-ttu-id="b4778-p120">以下代码检查项目标识符是否存在。如果 `itemId` 属性返回 `null` 或 `undefined`，则将项目保存到存储，并从异步结果中获取项目标识符。</span><span class="sxs-lookup"><span data-stu-id="b4778-p120">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

```JavaScript
var itemId = Office.context.mailbox.item.itemId;
if (itemId === null || itemId == undefined) {
  Office.context.mailbox.item.saveAsync(function(result){
    itemId = result.value;
  });
}
```

####  <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlook12officemailboxenumsitemtype"></a><span data-ttu-id="b4778-333">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_2/office.mailboxenums.itemtype)</span><span class="sxs-lookup"><span data-stu-id="b4778-333">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_2/office.mailboxenums.itemtype)</span></span>

<span data-ttu-id="b4778-334">获取实例表示的项的类型。</span><span class="sxs-lookup"><span data-stu-id="b4778-334">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="b4778-335">`itemType` 属性返回其中一个 `ItemType` 枚举值，指示 `item` 对象实例是邮件还是约会。</span><span class="sxs-lookup"><span data-stu-id="b4778-335">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="b4778-336">类型：</span><span class="sxs-lookup"><span data-stu-id="b4778-336">Type:</span></span>

*   [<span data-ttu-id="b4778-337">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="b4778-337">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook_1_2/office.mailboxenums.itemtype)

##### <a name="requirements"></a><span data-ttu-id="b4778-338">要求</span><span class="sxs-lookup"><span data-stu-id="b4778-338">Requirements</span></span>

|<span data-ttu-id="b4778-339">要求</span><span class="sxs-lookup"><span data-stu-id="b4778-339">Requirement</span></span>| <span data-ttu-id="b4778-340">值</span><span class="sxs-lookup"><span data-stu-id="b4778-340">Value</span></span>|
|---|---|
|[<span data-ttu-id="b4778-341">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="b4778-341">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b4778-342">1.0</span><span class="sxs-lookup"><span data-stu-id="b4778-342">1.0</span></span>|
|[<span data-ttu-id="b4778-343">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="b4778-343">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b4778-344">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b4778-344">ReadItem</span></span>|
|[<span data-ttu-id="b4778-345">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="b4778-345">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b4778-346">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="b4778-346">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="b4778-347">示例</span><span class="sxs-lookup"><span data-stu-id="b4778-347">Example</span></span>

```JavaScript
if (Office.context.mailbox.item.itemType == Office.MailboxEnums.ItemType.Message)
  // do something
else
  // do something else
```

####  <a name="location-stringlocationjavascriptapioutlook12officelocation"></a><span data-ttu-id="b4778-348">location :String|[Location](/javascript/api/outlook_1_2/office.location)</span><span class="sxs-lookup"><span data-stu-id="b4778-348">location :String|[Location](/javascript/api/outlook_1_2/office.location)</span></span>

<span data-ttu-id="b4778-349">获取或设置约会的位置。</span><span class="sxs-lookup"><span data-stu-id="b4778-349">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="b4778-350">阅读模式</span><span class="sxs-lookup"><span data-stu-id="b4778-350">Read mode</span></span>

<span data-ttu-id="b4778-351">`location` 属性返回一个包含约会位置的字符串。</span><span class="sxs-lookup"><span data-stu-id="b4778-351">The `location` property returns a string that contains the location of the appointment.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="b4778-352">撰写模式</span><span class="sxs-lookup"><span data-stu-id="b4778-352">Compose mode</span></span>

<span data-ttu-id="b4778-353">`location` 属性返回一个 `Location` 对象，该对象提供用于获取和设置约会位置的方法。</span><span class="sxs-lookup"><span data-stu-id="b4778-353">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="b4778-354">类型：</span><span class="sxs-lookup"><span data-stu-id="b4778-354">Type:</span></span>

*   <span data-ttu-id="b4778-355">String | [Location](/javascript/api/outlook_1_2/office.location)</span><span class="sxs-lookup"><span data-stu-id="b4778-355">String | [Location](/javascript/api/outlook_1_2/office.location)</span></span>

##### <a name="requirements"></a><span data-ttu-id="b4778-356">要求</span><span class="sxs-lookup"><span data-stu-id="b4778-356">Requirements</span></span>

|<span data-ttu-id="b4778-357">要求</span><span class="sxs-lookup"><span data-stu-id="b4778-357">Requirement</span></span>| <span data-ttu-id="b4778-358">值</span><span class="sxs-lookup"><span data-stu-id="b4778-358">Value</span></span>|
|---|---|
|[<span data-ttu-id="b4778-359">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="b4778-359">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b4778-360">1.0</span><span class="sxs-lookup"><span data-stu-id="b4778-360">1.0</span></span>|
|[<span data-ttu-id="b4778-361">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="b4778-361">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b4778-362">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b4778-362">ReadItem</span></span>|
|[<span data-ttu-id="b4778-363">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="b4778-363">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b4778-364">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="b4778-364">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="b4778-365">示例</span><span class="sxs-lookup"><span data-stu-id="b4778-365">Example</span></span>

```JavaScript
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

#### <a name="normalizedsubject-string"></a><span data-ttu-id="b4778-366">normalizedSubject :String</span><span class="sxs-lookup"><span data-stu-id="b4778-366">normalizedSubject :String</span></span>

<span data-ttu-id="b4778-p121">获取删除了所有前缀（包括 `RE:` 和 `FWD:`）的项目主题。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="b4778-p121">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="b4778-p122">normalizedSubject 属性获取包含由电子邮件程序添加的任何标准前缀（如 `RE:` 和 `FW:`）的项目主题。若要获取包含完整前缀的项目主题，请使用 [`subject`](#subject-stringsubjectjavascriptapioutlook12officesubject) 属性。</span><span class="sxs-lookup"><span data-stu-id="b4778-p122">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubjectjavascriptapioutlook12officesubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="b4778-371">类型:</span><span class="sxs-lookup"><span data-stu-id="b4778-371">Type:</span></span>

*   <span data-ttu-id="b4778-372">String</span><span class="sxs-lookup"><span data-stu-id="b4778-372">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="b4778-373">要求</span><span class="sxs-lookup"><span data-stu-id="b4778-373">Requirements</span></span>

|<span data-ttu-id="b4778-374">要求</span><span class="sxs-lookup"><span data-stu-id="b4778-374">Requirement</span></span>| <span data-ttu-id="b4778-375">值</span><span class="sxs-lookup"><span data-stu-id="b4778-375">Value</span></span>|
|---|---|
|[<span data-ttu-id="b4778-376">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="b4778-376">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b4778-377">1.0</span><span class="sxs-lookup"><span data-stu-id="b4778-377">1.0</span></span>|
|[<span data-ttu-id="b4778-378">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="b4778-378">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b4778-379">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b4778-379">ReadItem</span></span>|
|[<span data-ttu-id="b4778-380">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="b4778-380">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b4778-381">阅读</span><span class="sxs-lookup"><span data-stu-id="b4778-381">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="b4778-382">示例</span><span class="sxs-lookup"><span data-stu-id="b4778-382">Example</span></span>

```JavaScript
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
```

####  <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlook12officeemailaddressdetailsrecipientsjavascriptapioutlook12officerecipients"></a><span data-ttu-id="b4778-383">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_2/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="b4778-383">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_2/office.recipients)</span></span>

<span data-ttu-id="b4778-384">提供对事件的可选与会者的访问权限。</span><span class="sxs-lookup"><span data-stu-id="b4778-384">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="b4778-385">对象的类型和访问级别取决于当前项目的模式。</span><span class="sxs-lookup"><span data-stu-id="b4778-385">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="b4778-386">阅读模式</span><span class="sxs-lookup"><span data-stu-id="b4778-386">Read mode</span></span>

<span data-ttu-id="b4778-387">`optionalAttendees` 属性返回一个数组，其中包含每个可选与会者的 `EmailAddressDetails` 对象。</span><span class="sxs-lookup"><span data-stu-id="b4778-387">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="b4778-388">撰写模式</span><span class="sxs-lookup"><span data-stu-id="b4778-388">Compose mode</span></span>

<span data-ttu-id="b4778-389">`optionalAttendees` 属性返回一个 `Recipients` 对象，该对象提供用于获取或更新可选与会者的方法。</span><span class="sxs-lookup"><span data-stu-id="b4778-389">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="b4778-390">类型：</span><span class="sxs-lookup"><span data-stu-id="b4778-390">Type:</span></span>

*   <span data-ttu-id="b4778-391">Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_2/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="b4778-391">Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_2/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="b4778-392">要求</span><span class="sxs-lookup"><span data-stu-id="b4778-392">Requirements</span></span>

|<span data-ttu-id="b4778-393">要求</span><span class="sxs-lookup"><span data-stu-id="b4778-393">Requirement</span></span>| <span data-ttu-id="b4778-394">值</span><span class="sxs-lookup"><span data-stu-id="b4778-394">Value</span></span>|
|---|---|
|[<span data-ttu-id="b4778-395">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="b4778-395">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b4778-396">1.0</span><span class="sxs-lookup"><span data-stu-id="b4778-396">1.0</span></span>|
|[<span data-ttu-id="b4778-397">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="b4778-397">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b4778-398">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b4778-398">ReadItem</span></span>|
|[<span data-ttu-id="b4778-399">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="b4778-399">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b4778-400">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="b4778-400">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="b4778-401">示例</span><span class="sxs-lookup"><span data-stu-id="b4778-401">Example</span></span>

```JavaScript
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

#### <a name="organizer-emailaddressdetailsjavascriptapioutlook12officeemailaddressdetails"></a><span data-ttu-id="b4778-402">organizer :[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="b4778-402">organizer :[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)</span></span>

<span data-ttu-id="b4778-p124">获取指定会议的会议组织者的电子邮件地址。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="b4778-p124">Gets the email address of the meeting organizer for a specified meeting. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="b4778-405">类型:</span><span class="sxs-lookup"><span data-stu-id="b4778-405">Type:</span></span>

*   [<span data-ttu-id="b4778-406">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="b4778-406">EmailAddressDetails</span></span>](/javascript/api/outlook_1_2/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="b4778-407">要求</span><span class="sxs-lookup"><span data-stu-id="b4778-407">Requirements</span></span>

|<span data-ttu-id="b4778-408">要求</span><span class="sxs-lookup"><span data-stu-id="b4778-408">Requirement</span></span>| <span data-ttu-id="b4778-409">值</span><span class="sxs-lookup"><span data-stu-id="b4778-409">Value</span></span>|
|---|---|
|[<span data-ttu-id="b4778-410">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="b4778-410">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b4778-411">1.0</span><span class="sxs-lookup"><span data-stu-id="b4778-411">1.0</span></span>|
|[<span data-ttu-id="b4778-412">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="b4778-412">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b4778-413">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b4778-413">ReadItem</span></span>|
|[<span data-ttu-id="b4778-414">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="b4778-414">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b4778-415">阅读</span><span class="sxs-lookup"><span data-stu-id="b4778-415">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="b4778-416">示例</span><span class="sxs-lookup"><span data-stu-id="b4778-416">Example</span></span>

```JavaScript
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
```

####  <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlook12officeemailaddressdetailsrecipientsjavascriptapioutlook12officerecipients"></a><span data-ttu-id="b4778-417">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_2/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="b4778-417">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_2/office.recipients)</span></span>

<span data-ttu-id="b4778-418">提供对事件的必需与会者的访问权限。</span><span class="sxs-lookup"><span data-stu-id="b4778-418">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="b4778-419">对象的类型和访问级别取决于当前项目的模式。</span><span class="sxs-lookup"><span data-stu-id="b4778-419">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="b4778-420">阅读模式</span><span class="sxs-lookup"><span data-stu-id="b4778-420">Read mode</span></span>

<span data-ttu-id="b4778-421">`requiredAttendees` 属性返回一个数组，其中包含每个必需与会者的 `EmailAddressDetails` 对象。</span><span class="sxs-lookup"><span data-stu-id="b4778-421">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="b4778-422">撰写模式</span><span class="sxs-lookup"><span data-stu-id="b4778-422">Compose mode</span></span>

<span data-ttu-id="b4778-423">`requiredAttendees` 属性返回一个 `Recipients` 对象，该对象提供用于获取或更新必需与会者的方法。</span><span class="sxs-lookup"><span data-stu-id="b4778-423">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="b4778-424">类型：</span><span class="sxs-lookup"><span data-stu-id="b4778-424">Type:</span></span>

*   <span data-ttu-id="b4778-425">Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_2/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="b4778-425">Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_2/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="b4778-426">要求</span><span class="sxs-lookup"><span data-stu-id="b4778-426">Requirements</span></span>

|<span data-ttu-id="b4778-427">要求</span><span class="sxs-lookup"><span data-stu-id="b4778-427">Requirement</span></span>| <span data-ttu-id="b4778-428">值</span><span class="sxs-lookup"><span data-stu-id="b4778-428">Value</span></span>|
|---|---|
|[<span data-ttu-id="b4778-429">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="b4778-429">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b4778-430">1.0</span><span class="sxs-lookup"><span data-stu-id="b4778-430">1.0</span></span>|
|[<span data-ttu-id="b4778-431">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="b4778-431">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b4778-432">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b4778-432">ReadItem</span></span>|
|[<span data-ttu-id="b4778-433">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="b4778-433">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b4778-434">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="b4778-434">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="b4778-435">示例</span><span class="sxs-lookup"><span data-stu-id="b4778-435">Example</span></span>

```JavaScript
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
}
```

#### <a name="sender-emailaddressdetailsjavascriptapioutlook12officeemailaddressdetails"></a><span data-ttu-id="b4778-436">sender :[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="b4778-436">sender :[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)</span></span>

<span data-ttu-id="b4778-p126">获取电子邮件发件人的电子邮件地址。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="b4778-p126">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="b4778-p127">[`from`](#from-emailaddressdetailsjavascriptapioutlook12officeemailaddressdetails) 和 `sender` 属性表示同一个人，邮件由代理人发送的除外。在此情况下，`from` 属性表示代理程序，而 sender 属性表示代理人。</span><span class="sxs-lookup"><span data-stu-id="b4778-p127">The [`from`](#from-emailaddressdetailsjavascriptapioutlook12officeemailaddressdetails) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="b4778-441">`sender` 属性中 `EmailAddressDetails` 对象的 `recipientType` 属性为 `undefined`。</span><span class="sxs-lookup"><span data-stu-id="b4778-441">The `recipientType` property of the `EmailAddressDetails` object in the `sender` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="b4778-442">类型：</span><span class="sxs-lookup"><span data-stu-id="b4778-442">Type:</span></span>

*   [<span data-ttu-id="b4778-443">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="b4778-443">EmailAddressDetails</span></span>](/javascript/api/outlook_1_2/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="b4778-444">要求</span><span class="sxs-lookup"><span data-stu-id="b4778-444">Requirements</span></span>

|<span data-ttu-id="b4778-445">要求</span><span class="sxs-lookup"><span data-stu-id="b4778-445">Requirement</span></span>| <span data-ttu-id="b4778-446">值</span><span class="sxs-lookup"><span data-stu-id="b4778-446">Value</span></span>|
|---|---|
|[<span data-ttu-id="b4778-447">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="b4778-447">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b4778-448">1.0</span><span class="sxs-lookup"><span data-stu-id="b4778-448">1.0</span></span>|
|[<span data-ttu-id="b4778-449">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="b4778-449">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b4778-450">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b4778-450">ReadItem</span></span>|
|[<span data-ttu-id="b4778-451">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="b4778-451">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b4778-452">阅读</span><span class="sxs-lookup"><span data-stu-id="b4778-452">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="b4778-453">示例</span><span class="sxs-lookup"><span data-stu-id="b4778-453">Example</span></span>

```JavaScript
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
```

####  <a name="start-datetimejavascriptapioutlook12officetime"></a><span data-ttu-id="b4778-454">start :Date|[Time](/javascript/api/outlook_1_2/office.time)</span><span class="sxs-lookup"><span data-stu-id="b4778-454">start :Date|[Time](/javascript/api/outlook_1_2/office.time)</span></span>

<span data-ttu-id="b4778-455">获取或设置约会开始的日期和时间。</span><span class="sxs-lookup"><span data-stu-id="b4778-455">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="b4778-p128">将 `start` 属性表示为协调世界时 (UTC) 的日期和时间值。可使用 [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook12officelocalclienttime) 方法将该值转换为客户端的本地日期和时间。</span><span class="sxs-lookup"><span data-stu-id="b4778-p128">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook12officelocalclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="b4778-458">阅读模式</span><span class="sxs-lookup"><span data-stu-id="b4778-458">Read mode</span></span>

<span data-ttu-id="b4778-459">`start` 属性返回 `Date` 对象。</span><span class="sxs-lookup"><span data-stu-id="b4778-459">The `start` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="b4778-460">撰写模式</span><span class="sxs-lookup"><span data-stu-id="b4778-460">Compose mode</span></span>

<span data-ttu-id="b4778-461">`start` 属性返回 `Time` 对象。</span><span class="sxs-lookup"><span data-stu-id="b4778-461">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="b4778-462">使用 [`Time.setAsync`](/javascript/api/outlook_1_2/office.time#setasync-datetime--options--callback-) 方法设置开始时间时，应使用 [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) 方法将客户端的本地时间转换为服务器的 UTC。</span><span class="sxs-lookup"><span data-stu-id="b4778-462">When you use the [`Time.setAsync`](/javascript/api/outlook_1_2/office.time#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="b4778-463">类型：</span><span class="sxs-lookup"><span data-stu-id="b4778-463">Type:</span></span>

*   <span data-ttu-id="b4778-464">Date | [Time](/javascript/api/outlook_1_2/office.time)</span><span class="sxs-lookup"><span data-stu-id="b4778-464">Date | [Time](/javascript/api/outlook_1_2/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="b4778-465">要求</span><span class="sxs-lookup"><span data-stu-id="b4778-465">Requirements</span></span>

|<span data-ttu-id="b4778-466">要求</span><span class="sxs-lookup"><span data-stu-id="b4778-466">Requirement</span></span>| <span data-ttu-id="b4778-467">值</span><span class="sxs-lookup"><span data-stu-id="b4778-467">Value</span></span>|
|---|---|
|[<span data-ttu-id="b4778-468">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="b4778-468">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b4778-469">1.0</span><span class="sxs-lookup"><span data-stu-id="b4778-469">1.0</span></span>|
|[<span data-ttu-id="b4778-470">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="b4778-470">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b4778-471">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b4778-471">ReadItem</span></span>|
|[<span data-ttu-id="b4778-472">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="b4778-472">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b4778-473">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="b4778-473">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="b4778-474">示例</span><span class="sxs-lookup"><span data-stu-id="b4778-474">Example</span></span>

<span data-ttu-id="b4778-475">以下示例通过使用 `Time` 对象的 [`setAsync`](/javascript/api/outlook_1_2/office.time#setasync-datetime--options--callback-) 方法，设置撰写模式下约会的开始时间。</span><span class="sxs-lookup"><span data-stu-id="b4778-475">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_2/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

####  <a name="subject-stringsubjectjavascriptapioutlook12officesubject"></a><span data-ttu-id="b4778-476">subject :String|[Subject](/javascript/api/outlook_1_2/office.subject)</span><span class="sxs-lookup"><span data-stu-id="b4778-476">subject :String|[Subject](/javascript/api/outlook_1_2/office.subject)</span></span>

<span data-ttu-id="b4778-477">获取或设置显示在项目的主题字段中的说明。</span><span class="sxs-lookup"><span data-stu-id="b4778-477">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="b4778-478">`subject` 属性获取或设置由电子邮件服务器发送项目时的整个主题。</span><span class="sxs-lookup"><span data-stu-id="b4778-478">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="b4778-479">阅读模式</span><span class="sxs-lookup"><span data-stu-id="b4778-479">Read mode</span></span>

<span data-ttu-id="b4778-p129">`subject` 属性返回一个字符串。使用 [`normalizedSubject`](#normalizedsubject-string) 属性获取不带任何前导前缀（如 `RE:` 和 `FW:`）的主题。</span><span class="sxs-lookup"><span data-stu-id="b4778-p129">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

```js
var subject = Office.context.mailbox.item.subject;
```

##### <a name="compose-mode"></a><span data-ttu-id="b4778-482">撰写模式</span><span class="sxs-lookup"><span data-stu-id="b4778-482">Compose mode</span></span>

<span data-ttu-id="b4778-483">`subject` 属性返回一个 `Subject` 对象，该对象提供用于获取和设置主题的方法。</span><span class="sxs-lookup"><span data-stu-id="b4778-483">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```JavaScript
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="b4778-484">类型：</span><span class="sxs-lookup"><span data-stu-id="b4778-484">Type:</span></span>

*   <span data-ttu-id="b4778-485">String | [Subject](/javascript/api/outlook_1_2/office.subject)</span><span class="sxs-lookup"><span data-stu-id="b4778-485">String | [Subject](/javascript/api/outlook_1_2/office.subject)</span></span>

##### <a name="requirements"></a><span data-ttu-id="b4778-486">要求</span><span class="sxs-lookup"><span data-stu-id="b4778-486">Requirements</span></span>

|<span data-ttu-id="b4778-487">要求</span><span class="sxs-lookup"><span data-stu-id="b4778-487">Requirement</span></span>| <span data-ttu-id="b4778-488">值</span><span class="sxs-lookup"><span data-stu-id="b4778-488">Value</span></span>|
|---|---|
|[<span data-ttu-id="b4778-489">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="b4778-489">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b4778-490">1.0</span><span class="sxs-lookup"><span data-stu-id="b4778-490">1.0</span></span>|
|[<span data-ttu-id="b4778-491">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="b4778-491">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b4778-492">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b4778-492">ReadItem</span></span>|
|[<span data-ttu-id="b4778-493">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="b4778-493">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b4778-494">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="b4778-494">Compose or read</span></span>|

####  <a name="to-arrayemailaddressdetailsjavascriptapioutlook12officeemailaddressdetailsrecipientsjavascriptapioutlook12officerecipients"></a><span data-ttu-id="b4778-495">to :Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_2/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="b4778-495">to :Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_2/office.recipients)</span></span>

<span data-ttu-id="b4778-496">提供对邮件的“**收件人**”行上的收件人的访问权限。</span><span class="sxs-lookup"><span data-stu-id="b4778-496">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="b4778-497">对象的类型和访问级别取决于当前项目的模式。</span><span class="sxs-lookup"><span data-stu-id="b4778-497">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="b4778-498">阅读模式</span><span class="sxs-lookup"><span data-stu-id="b4778-498">Read mode</span></span>

<span data-ttu-id="b4778-p131">`to` 属性返回包含邮件的**收件人**行上所列的每个收件人的 `EmailAddressDetails` 对象的数组。集合上限为 100 个成员。</span><span class="sxs-lookup"><span data-stu-id="b4778-p131">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="b4778-501">撰写模式</span><span class="sxs-lookup"><span data-stu-id="b4778-501">Compose mode</span></span>

<span data-ttu-id="b4778-502">`to` 属性返回一个 `Recipients` 对象，该对象提供用于获取或更新邮件的“**收件人**”行上收件人的方法。</span><span class="sxs-lookup"><span data-stu-id="b4778-502">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="b4778-503">类型：</span><span class="sxs-lookup"><span data-stu-id="b4778-503">Type:</span></span>

*   <span data-ttu-id="b4778-504">Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_2/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="b4778-504">Array.<[EmailAddressDetails](/javascript/api/outlook_1_2/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_2/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="b4778-505">要求</span><span class="sxs-lookup"><span data-stu-id="b4778-505">Requirements</span></span>

|<span data-ttu-id="b4778-506">要求</span><span class="sxs-lookup"><span data-stu-id="b4778-506">Requirement</span></span>| <span data-ttu-id="b4778-507">值</span><span class="sxs-lookup"><span data-stu-id="b4778-507">Value</span></span>|
|---|---|
|[<span data-ttu-id="b4778-508">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="b4778-508">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b4778-509">1.0</span><span class="sxs-lookup"><span data-stu-id="b4778-509">1.0</span></span>|
|[<span data-ttu-id="b4778-510">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="b4778-510">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b4778-511">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b4778-511">ReadItem</span></span>|
|[<span data-ttu-id="b4778-512">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="b4778-512">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b4778-513">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="b4778-513">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="b4778-514">示例</span><span class="sxs-lookup"><span data-stu-id="b4778-514">Example</span></span>

```JavaScript
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

### <a name="methods"></a><span data-ttu-id="b4778-515">方法</span><span class="sxs-lookup"><span data-stu-id="b4778-515">Methods</span></span>

####  <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="b4778-516">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="b4778-516">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="b4778-517">将文件作为附件添加到邮件或约会。</span><span class="sxs-lookup"><span data-stu-id="b4778-517">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="b4778-518">`addFileAttachmentAsync` 方法在指定的 URI 上载文件并将其附加到撰写窗体中的项目。</span><span class="sxs-lookup"><span data-stu-id="b4778-518">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="b4778-519">你随后可以将该标识符与 [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) 方法一同使用，以删除同一个会话中的附件。</span><span class="sxs-lookup"><span data-stu-id="b4778-519">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="b4778-520">参数：</span><span class="sxs-lookup"><span data-stu-id="b4778-520">Parameters:</span></span>

|<span data-ttu-id="b4778-521">名称</span><span class="sxs-lookup"><span data-stu-id="b4778-521">Name</span></span>| <span data-ttu-id="b4778-522">类型</span><span class="sxs-lookup"><span data-stu-id="b4778-522">Type</span></span>| <span data-ttu-id="b4778-523">属性</span><span class="sxs-lookup"><span data-stu-id="b4778-523">Attributes</span></span>| <span data-ttu-id="b4778-524">说明</span><span class="sxs-lookup"><span data-stu-id="b4778-524">Description</span></span>|
|---|---|---|---|
|`uri`| <span data-ttu-id="b4778-525">字符串</span><span class="sxs-lookup"><span data-stu-id="b4778-525">String</span></span>||<span data-ttu-id="b4778-p132">提供附加到邮件或约会的文件的位置的 URI。最大长度为 2048 个字符。</span><span class="sxs-lookup"><span data-stu-id="b4778-p132">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="b4778-528">字符串</span><span class="sxs-lookup"><span data-stu-id="b4778-528">String</span></span>||<span data-ttu-id="b4778-p133">在附件上载过程中显示的附件名称。最大长度为 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="b4778-p133">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="b4778-531">对象</span><span class="sxs-lookup"><span data-stu-id="b4778-531">Object</span></span>| <span data-ttu-id="b4778-532">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="b4778-532">&lt;optional&gt;</span></span>|<span data-ttu-id="b4778-533">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="b4778-533">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="b4778-534">对象</span><span class="sxs-lookup"><span data-stu-id="b4778-534">Object</span></span>| <span data-ttu-id="b4778-535">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="b4778-535">&lt;optional&gt;</span></span>|<span data-ttu-id="b4778-536">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="b4778-536">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="b4778-537">函数</span><span class="sxs-lookup"><span data-stu-id="b4778-537">function</span></span>| <span data-ttu-id="b4778-538">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="b4778-538">&lt;optional&gt;</span></span>|<span data-ttu-id="b4778-539">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="b4778-539">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="b4778-540">如果成功，附件标识符将在 `asyncResult.value` 属性中提供。</span><span class="sxs-lookup"><span data-stu-id="b4778-540">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="b4778-541">如果上传附件失败，`asyncResult` 对象将包含一个提供错误说明的 `Error` 对象。</span><span class="sxs-lookup"><span data-stu-id="b4778-541">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="b4778-542">错误</span><span class="sxs-lookup"><span data-stu-id="b4778-542">Errors</span></span>

| <span data-ttu-id="b4778-543">错误代码</span><span class="sxs-lookup"><span data-stu-id="b4778-543">Error code</span></span> | <span data-ttu-id="b4778-544">说明</span><span class="sxs-lookup"><span data-stu-id="b4778-544">Description</span></span> |
|------------|-------------|
| `AttachmentSizeExceeded` | <span data-ttu-id="b4778-545">附件大小超过了允许的大小。</span><span class="sxs-lookup"><span data-stu-id="b4778-545">The attachment is larger than allowed.</span></span> |
| `FileTypeNotSupported` | <span data-ttu-id="b4778-546">该附件的扩展名不是允许的扩展名。</span><span class="sxs-lookup"><span data-stu-id="b4778-546">The attachment has an extension that is not allowed.</span></span> |
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="b4778-547">邮件或约会具有的附件过多。</span><span class="sxs-lookup"><span data-stu-id="b4778-547">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="b4778-548">要求</span><span class="sxs-lookup"><span data-stu-id="b4778-548">Requirements</span></span>

|<span data-ttu-id="b4778-549">要求</span><span class="sxs-lookup"><span data-stu-id="b4778-549">Requirement</span></span>| <span data-ttu-id="b4778-550">值</span><span class="sxs-lookup"><span data-stu-id="b4778-550">Value</span></span>|
|---|---|
|[<span data-ttu-id="b4778-551">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="b4778-551">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b4778-552">1.1</span><span class="sxs-lookup"><span data-stu-id="b4778-552">1.1</span></span>|
|[<span data-ttu-id="b4778-553">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="b4778-553">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b4778-554">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="b4778-554">ReadWriteItem</span></span>|
|[<span data-ttu-id="b4778-555">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="b4778-555">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b4778-556">撰写</span><span class="sxs-lookup"><span data-stu-id="b4778-556">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="b4778-557">示例</span><span class="sxs-lookup"><span data-stu-id="b4778-557">Example</span></span>

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

####  <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="b4778-558">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="b4778-558">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="b4778-559">将 Exchange 项目（如邮件）作为附件添加到邮件或约会。</span><span class="sxs-lookup"><span data-stu-id="b4778-559">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="b4778-p134">`addItemAttachmentAsync` 方法将包含指定 Exchange 标识符的项目附加到撰写窗体中的项目。如果指定一个回调方法，此方法使用 `asyncResult` 参数调用，该参数包含一个附件标识符或代码，指示附加项目过程中出现的任何错误。可以使用 `options` 参数将状态信息传递给回调方法（如果需要）。</span><span class="sxs-lookup"><span data-stu-id="b4778-p134">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="b4778-563">随后可以将该标识符与 [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) 方法一同使用，以删除同一个会话中的附件。</span><span class="sxs-lookup"><span data-stu-id="b4778-563">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="b4778-564">如果 Office 加载项在 Outlook Web App 中运行，则 `addItemAttachmentAsync` 方法可以将项目附加到项目（正在编辑的项目除外）中；然而，不支持也不建议这样做。</span><span class="sxs-lookup"><span data-stu-id="b4778-564">If your Office Add-in is running in Outlook Web App, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="b4778-565">参数：</span><span class="sxs-lookup"><span data-stu-id="b4778-565">Parameters:</span></span>

|<span data-ttu-id="b4778-566">名称</span><span class="sxs-lookup"><span data-stu-id="b4778-566">Name</span></span>| <span data-ttu-id="b4778-567">类型</span><span class="sxs-lookup"><span data-stu-id="b4778-567">Type</span></span>| <span data-ttu-id="b4778-568">属性</span><span class="sxs-lookup"><span data-stu-id="b4778-568">Attributes</span></span>| <span data-ttu-id="b4778-569">说明</span><span class="sxs-lookup"><span data-stu-id="b4778-569">Description</span></span>|
|---|---|---|---|
|`itemId`| <span data-ttu-id="b4778-570">字符串</span><span class="sxs-lookup"><span data-stu-id="b4778-570">String</span></span>||<span data-ttu-id="b4778-p135">要附加的项目的 Exchange 标识符。最大长度为 100 个字符。</span><span class="sxs-lookup"><span data-stu-id="b4778-p135">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="b4778-573">字符串</span><span class="sxs-lookup"><span data-stu-id="b4778-573">String</span></span>||<span data-ttu-id="b4778-p136">要附加的项目的主题。最大长度为 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="b4778-p136">The sujbect of the item to be attached. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="b4778-576">对象</span><span class="sxs-lookup"><span data-stu-id="b4778-576">Object</span></span>| <span data-ttu-id="b4778-577">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="b4778-577">&lt;optional&gt;</span></span>|<span data-ttu-id="b4778-578">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="b4778-578">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="b4778-579">对象</span><span class="sxs-lookup"><span data-stu-id="b4778-579">Object</span></span>| <span data-ttu-id="b4778-580">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="b4778-580">&lt;optional&gt;</span></span>|<span data-ttu-id="b4778-581">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="b4778-581">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="b4778-582">函数</span><span class="sxs-lookup"><span data-stu-id="b4778-582">function</span></span>| <span data-ttu-id="b4778-583">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="b4778-583">&lt;optional&gt;</span></span>|<span data-ttu-id="b4778-584">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="b4778-584">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="b4778-585">如果成功，附件标识符将在 `asyncResult.value` 属性中提供。</span><span class="sxs-lookup"><span data-stu-id="b4778-585">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="b4778-586">如果添加附件失败，`asyncResult` 对象将包含一个提供错误说明的 `Error` 对象。</span><span class="sxs-lookup"><span data-stu-id="b4778-586">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="b4778-587">错误</span><span class="sxs-lookup"><span data-stu-id="b4778-587">Errors</span></span>

| <span data-ttu-id="b4778-588">错误代码</span><span class="sxs-lookup"><span data-stu-id="b4778-588">Error code</span></span> | <span data-ttu-id="b4778-589">说明</span><span class="sxs-lookup"><span data-stu-id="b4778-589">Description</span></span> |
|------------|-------------|
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="b4778-590">邮件或约会具有的附件过多。</span><span class="sxs-lookup"><span data-stu-id="b4778-590">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="b4778-591">要求</span><span class="sxs-lookup"><span data-stu-id="b4778-591">Requirements</span></span>

|<span data-ttu-id="b4778-592">要求</span><span class="sxs-lookup"><span data-stu-id="b4778-592">Requirement</span></span>| <span data-ttu-id="b4778-593">值</span><span class="sxs-lookup"><span data-stu-id="b4778-593">Value</span></span>|
|---|---|
|[<span data-ttu-id="b4778-594">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="b4778-594">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b4778-595">1.1</span><span class="sxs-lookup"><span data-stu-id="b4778-595">1.1</span></span>|
|[<span data-ttu-id="b4778-596">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="b4778-596">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b4778-597">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="b4778-597">ReadWriteItem</span></span>|
|[<span data-ttu-id="b4778-598">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="b4778-598">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b4778-599">撰写</span><span class="sxs-lookup"><span data-stu-id="b4778-599">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="b4778-600">示例</span><span class="sxs-lookup"><span data-stu-id="b4778-600">Example</span></span>

<span data-ttu-id="b4778-601">以下示例将现有的 Outlook 项目添加为名为 `My Attachment` 的附件。</span><span class="sxs-lookup"><span data-stu-id="b4778-601">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

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

#### <a name="displayreplyallformformdata"></a><span data-ttu-id="b4778-602">displayReplyAllForm(formData)</span><span class="sxs-lookup"><span data-stu-id="b4778-602">displayReplyAllForm(formData)</span></span>

<span data-ttu-id="b4778-603">显示答复窗体，其中包括所选邮件的发件人和所有收件人或所选约会的组织者和所有与会者。</span><span class="sxs-lookup"><span data-stu-id="b4778-603">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="b4778-604">在 Outlook for iOS 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="b4778-604">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="b4778-605">在 Outlook Web App 中，答复窗体显示为包含 3 列视图的弹出式窗体以及包含 2 列或 1 列视图的弹出式窗体。</span><span class="sxs-lookup"><span data-stu-id="b4778-605">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="b4778-606">如果任意字符串参数超出其限制，`displayReplyAllForm` 将引发异常。</span><span class="sxs-lookup"><span data-stu-id="b4778-606">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

<span data-ttu-id="b4778-p137">当在 `formData.attachments` 参数中指定附件时，Outlook 和 Outlook Web App 尝试下载所有附件并将其附加到答复窗体。如果无法添加任何附件，则在窗体 UI 中显示错误。如果这不可能，则不引发错误消息。</span><span class="sxs-lookup"><span data-stu-id="b4778-p137">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="b4778-610">参数：</span><span class="sxs-lookup"><span data-stu-id="b4778-610">Parameters:</span></span>

|<span data-ttu-id="b4778-611">名称</span><span class="sxs-lookup"><span data-stu-id="b4778-611">Name</span></span>| <span data-ttu-id="b4778-612">类型</span><span class="sxs-lookup"><span data-stu-id="b4778-612">Type</span></span>| <span data-ttu-id="b4778-613">说明</span><span class="sxs-lookup"><span data-stu-id="b4778-613">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="b4778-614">字符串 &#124; 对象</span><span class="sxs-lookup"><span data-stu-id="b4778-614">String &#124; Object</span></span>| |<span data-ttu-id="b4778-p138">一个包含文本和 HTML 且表示答复窗体的正文的字符串。字符串限制为 32 KB。</span><span class="sxs-lookup"><span data-stu-id="b4778-p138">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="b4778-617">**OR**</span><span class="sxs-lookup"><span data-stu-id="b4778-617">**OR**</span></span><br/><span data-ttu-id="b4778-p139">包含正文或附件数据和回调函数的对象。对象定义如下。</span><span class="sxs-lookup"><span data-stu-id="b4778-p139">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="b4778-620">字符串</span><span class="sxs-lookup"><span data-stu-id="b4778-620">String</span></span> | <span data-ttu-id="b4778-621">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="b4778-621">&lt;optional&gt;</span></span> | <span data-ttu-id="b4778-p140">一个包含文本和 HTML 且表示答复窗体的正文的字符串。字符串限制为 32 KB。</span><span class="sxs-lookup"><span data-stu-id="b4778-p140">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="b4778-624">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="b4778-624">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="b4778-625">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="b4778-625">&lt;optional&gt;</span></span> | <span data-ttu-id="b4778-626">JSON 对象（是文件或项目附件）的数组。</span><span class="sxs-lookup"><span data-stu-id="b4778-626">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="b4778-627">字符串</span><span class="sxs-lookup"><span data-stu-id="b4778-627">String</span></span> | | <span data-ttu-id="b4778-p141">指示附件的类型。必须是文件附件的 `file` 或项目附件的 `item`。</span><span class="sxs-lookup"><span data-stu-id="b4778-p141">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="b4778-630">字符串</span><span class="sxs-lookup"><span data-stu-id="b4778-630">String</span></span> | | <span data-ttu-id="b4778-631">一个包含附件的名称的字符串，最多包含 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="b4778-631">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="b4778-632">String</span><span class="sxs-lookup"><span data-stu-id="b4778-632">String</span></span> | | <span data-ttu-id="b4778-p142">仅在将 `type` 设置为 `file` 时使用。文件的位置的 URI。</span><span class="sxs-lookup"><span data-stu-id="b4778-p142">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="b4778-635">String</span><span class="sxs-lookup"><span data-stu-id="b4778-635">String</span></span> | | <span data-ttu-id="b4778-p143">仅在将 `type` 设置为 `item` 时使用。附件的 EWS 项目 ID。字符串最长为 100 个字符。</span><span class="sxs-lookup"><span data-stu-id="b4778-p143">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="b4778-639">函数</span><span class="sxs-lookup"><span data-stu-id="b4778-639">function</span></span> | <span data-ttu-id="b4778-640">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="b4778-640">&lt;optional&gt;</span></span> | <span data-ttu-id="b4778-641">方法完成后，使用单个参数 `asyncResult`（一个 [AsyncResult](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="b4778-641">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="b4778-642">要求</span><span class="sxs-lookup"><span data-stu-id="b4778-642">Requirements</span></span>

|<span data-ttu-id="b4778-643">要求</span><span class="sxs-lookup"><span data-stu-id="b4778-643">Requirement</span></span>| <span data-ttu-id="b4778-644">值</span><span class="sxs-lookup"><span data-stu-id="b4778-644">Value</span></span>|
|---|---|
|[<span data-ttu-id="b4778-645">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="b4778-645">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b4778-646">1.0</span><span class="sxs-lookup"><span data-stu-id="b4778-646">1.0</span></span>|
|[<span data-ttu-id="b4778-647">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="b4778-647">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b4778-648">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b4778-648">ReadItem</span></span>|
|[<span data-ttu-id="b4778-649">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="b4778-649">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b4778-650">阅读</span><span class="sxs-lookup"><span data-stu-id="b4778-650">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="b4778-651">示例</span><span class="sxs-lookup"><span data-stu-id="b4778-651">Examples</span></span>

<span data-ttu-id="b4778-652">以下代码将一个字符串传递到 `displayReplyAllForm` 函数。</span><span class="sxs-lookup"><span data-stu-id="b4778-652">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="b4778-653">使用空白正文答复。</span><span class="sxs-lookup"><span data-stu-id="b4778-653">Reply with an empty body.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="b4778-654">仅使用正文答复。</span><span class="sxs-lookup"><span data-stu-id="b4778-654">Reply with just a body.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="b4778-655">使用正文和文件附件答复。</span><span class="sxs-lookup"><span data-stu-id="b4778-655">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="b4778-656">使用正文和项目附件答复。</span><span class="sxs-lookup"><span data-stu-id="b4778-656">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="b4778-657">使用正文、文件附件、项目附件和回调答复。</span><span class="sxs-lookup"><span data-stu-id="b4778-657">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="displayreplyformformdata"></a><span data-ttu-id="b4778-658">displayReplyForm(formData)</span><span class="sxs-lookup"><span data-stu-id="b4778-658">displayReplyForm(formData)</span></span>

<span data-ttu-id="b4778-659">显示答复窗体，其中仅包括所选邮件的发件人或所选约会的组织者。</span><span class="sxs-lookup"><span data-stu-id="b4778-659">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="b4778-660">在 Outlook for iOS 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="b4778-660">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="b4778-661">在 Outlook Web App 中，答复窗体显示为包含 3 列视图的弹出式窗体以及包含 2 列或 1 列视图的弹出式窗体。</span><span class="sxs-lookup"><span data-stu-id="b4778-661">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="b4778-662">如果任意字符串参数超出其限制，`displayReplyForm` 将引发异常。</span><span class="sxs-lookup"><span data-stu-id="b4778-662">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

<span data-ttu-id="b4778-p144">当在 `formData.attachments` 参数中指定附件时，Outlook 和 Outlook Web App 尝试下载所有附件并将其附加到答复窗体。如果无法添加任何附件，则在窗体 UI 中显示错误。如果这不可能，则不引发错误消息。</span><span class="sxs-lookup"><span data-stu-id="b4778-p144">When attachments are specified in the `formData.attachments` parameter, Outlook and Outlook Web App attempt to download all attachments and attach them to the reply form. If any attachments fail to be added, an error is shown in the form UI. If this isn't possible, then no error message is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="b4778-666">参数：</span><span class="sxs-lookup"><span data-stu-id="b4778-666">Parameters:</span></span>

|<span data-ttu-id="b4778-667">名称</span><span class="sxs-lookup"><span data-stu-id="b4778-667">Name</span></span>| <span data-ttu-id="b4778-668">类型</span><span class="sxs-lookup"><span data-stu-id="b4778-668">Type</span></span>| <span data-ttu-id="b4778-669">说明</span><span class="sxs-lookup"><span data-stu-id="b4778-669">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="b4778-670">字符串 &#124; 对象</span><span class="sxs-lookup"><span data-stu-id="b4778-670">String &#124; Object</span></span>| | <span data-ttu-id="b4778-p145">一个包含文本和 HTML 且表示答复窗体的正文的字符串。字符串限制为 32 KB。</span><span class="sxs-lookup"><span data-stu-id="b4778-p145">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="b4778-673">**OR**</span><span class="sxs-lookup"><span data-stu-id="b4778-673">**OR**</span></span><br/><span data-ttu-id="b4778-p146">包含正文或附件数据和回调函数的对象。对象定义如下。</span><span class="sxs-lookup"><span data-stu-id="b4778-p146">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="b4778-676">字符串</span><span class="sxs-lookup"><span data-stu-id="b4778-676">String</span></span> | <span data-ttu-id="b4778-677">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="b4778-677">&lt;optional&gt;</span></span> | <span data-ttu-id="b4778-p147">一个包含文本和 HTML 且表示答复窗体的正文的字符串。字符串限制为 32 KB。</span><span class="sxs-lookup"><span data-stu-id="b4778-p147">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `formData.attachments` | <span data-ttu-id="b4778-680">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="b4778-680">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="b4778-681">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="b4778-681">&lt;optional&gt;</span></span> | <span data-ttu-id="b4778-682">JSON 对象（是文件或项目附件）的数组。</span><span class="sxs-lookup"><span data-stu-id="b4778-682">An array of JSON objects that are either file or item attachments.</span></span> |
| `formData.attachments.type` | <span data-ttu-id="b4778-683">字符串</span><span class="sxs-lookup"><span data-stu-id="b4778-683">String</span></span> | | <span data-ttu-id="b4778-p148">指示附件的类型。必须是文件附件的 `file` 或项目附件的 `item`。</span><span class="sxs-lookup"><span data-stu-id="b4778-p148">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `formData.attachments.name` | <span data-ttu-id="b4778-686">字符串</span><span class="sxs-lookup"><span data-stu-id="b4778-686">String</span></span> | | <span data-ttu-id="b4778-687">一个包含附件的名称的字符串，最多包含 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="b4778-687">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `formData.attachments.url` | <span data-ttu-id="b4778-688">String</span><span class="sxs-lookup"><span data-stu-id="b4778-688">String</span></span> | | <span data-ttu-id="b4778-p149">仅在将 `type` 设置为 `file` 时使用。文件的位置的 URI。</span><span class="sxs-lookup"><span data-stu-id="b4778-p149">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `formData.attachments.itemId` | <span data-ttu-id="b4778-691">String</span><span class="sxs-lookup"><span data-stu-id="b4778-691">String</span></span> | | <span data-ttu-id="b4778-p150">仅在将 `type` 设置为 `item` 时使用。附件的 EWS 项目 ID。字符串最长为 100 个字符。</span><span class="sxs-lookup"><span data-stu-id="b4778-p150">Only used if `type` is set to `item`. The EWS item id of the attachment. This is a string up to 100 characters.</span></span> |
| `callback` | <span data-ttu-id="b4778-695">函数</span><span class="sxs-lookup"><span data-stu-id="b4778-695">function</span></span> | <span data-ttu-id="b4778-696">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="b4778-696">&lt;optional&gt;</span></span> | <span data-ttu-id="b4778-697">方法完成后，使用单个参数 `asyncResult`（一个 [AsyncResult](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="b4778-697">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="b4778-698">要求</span><span class="sxs-lookup"><span data-stu-id="b4778-698">Requirements</span></span>

|<span data-ttu-id="b4778-699">要求</span><span class="sxs-lookup"><span data-stu-id="b4778-699">Requirement</span></span>| <span data-ttu-id="b4778-700">值</span><span class="sxs-lookup"><span data-stu-id="b4778-700">Value</span></span>|
|---|---|
|[<span data-ttu-id="b4778-701">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="b4778-701">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b4778-702">1.0</span><span class="sxs-lookup"><span data-stu-id="b4778-702">1.0</span></span>|
|[<span data-ttu-id="b4778-703">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="b4778-703">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b4778-704">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b4778-704">ReadItem</span></span>|
|[<span data-ttu-id="b4778-705">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="b4778-705">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b4778-706">阅读</span><span class="sxs-lookup"><span data-stu-id="b4778-706">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="b4778-707">示例</span><span class="sxs-lookup"><span data-stu-id="b4778-707">Examples</span></span>

<span data-ttu-id="b4778-708">以下代码将一个字符串传递到 `displayReplyForm` 函数。</span><span class="sxs-lookup"><span data-stu-id="b4778-708">The following code passes a string to the `displayReplyForm` function.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="b4778-709">使用空白正文答复。</span><span class="sxs-lookup"><span data-stu-id="b4778-709">Reply with an empty body.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="b4778-710">仅使用正文答复。</span><span class="sxs-lookup"><span data-stu-id="b4778-710">Reply with just a body.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="b4778-711">使用正文和文件附件答复。</span><span class="sxs-lookup"><span data-stu-id="b4778-711">Reply with a body and a file attachment.</span></span>

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

<span data-ttu-id="b4778-712">使用正文和项目附件答复。</span><span class="sxs-lookup"><span data-stu-id="b4778-712">Reply with a body and an item attachment.</span></span>

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

<span data-ttu-id="b4778-713">使用正文、文件附件、项目附件和回调答复。</span><span class="sxs-lookup"><span data-stu-id="b4778-713">Reply with a body, file attachment, item attachment, and a callback.</span></span>

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

#### <a name="getentities--entitiesjavascriptapioutlook12officeentities"></a><span data-ttu-id="b4778-714">getEntities() → {[Entities](/javascript/api/outlook_1_2/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="b4778-714">getEntities() → {[Entities](/javascript/api/outlook_1_2/office.entities)}</span></span>

<span data-ttu-id="b4778-715">获取在所选项目的正文中找到的实体。</span><span class="sxs-lookup"><span data-stu-id="b4778-715">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="b4778-716">在 Outlook for iOS 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="b4778-716">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="b4778-717">要求</span><span class="sxs-lookup"><span data-stu-id="b4778-717">Requirements</span></span>

|<span data-ttu-id="b4778-718">要求</span><span class="sxs-lookup"><span data-stu-id="b4778-718">Requirement</span></span>| <span data-ttu-id="b4778-719">值</span><span class="sxs-lookup"><span data-stu-id="b4778-719">Value</span></span>|
|---|---|
|[<span data-ttu-id="b4778-720">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="b4778-720">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b4778-721">1.0</span><span class="sxs-lookup"><span data-stu-id="b4778-721">1.0</span></span>|
|[<span data-ttu-id="b4778-722">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="b4778-722">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b4778-723">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b4778-723">ReadItem</span></span>|
|[<span data-ttu-id="b4778-724">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="b4778-724">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b4778-725">阅读</span><span class="sxs-lookup"><span data-stu-id="b4778-725">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="b4778-726">返回：</span><span class="sxs-lookup"><span data-stu-id="b4778-726">Returns:</span></span>

<span data-ttu-id="b4778-727">类型：[Entities](/javascript/api/outlook_1_2/office.entities)</span><span class="sxs-lookup"><span data-stu-id="b4778-727">Type: [Entities](/javascript/api/outlook_1_2/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="b4778-728">示例</span><span class="sxs-lookup"><span data-stu-id="b4778-728">Example</span></span>

<span data-ttu-id="b4778-729">以下示例访问当前项目的正文中的联系人实体。</span><span class="sxs-lookup"><span data-stu-id="b4778-729">The following example accesses the contacts entities in the current item's body.</span></span>

```js
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlook12officecontactmeetingsuggestionjavascriptapioutlook12officemeetingsuggestionphonenumberjavascriptapioutlook12officephonenumbertasksuggestionjavascriptapioutlook12officetasksuggestion"></a><span data-ttu-id="b4778-730">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_2/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_2/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_2/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_2/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="b4778-730">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_2/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_2/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_2/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_2/office.tasksuggestion))>}</span></span>

<span data-ttu-id="b4778-731">获取所选项目的正文中找到的指定实体类型的所有实体的数组。</span><span class="sxs-lookup"><span data-stu-id="b4778-731">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="b4778-732">在 Outlook for iOS 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="b4778-732">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="b4778-733">参数：</span><span class="sxs-lookup"><span data-stu-id="b4778-733">Parameters:</span></span>

|<span data-ttu-id="b4778-734">名称</span><span class="sxs-lookup"><span data-stu-id="b4778-734">Name</span></span>| <span data-ttu-id="b4778-735">类型</span><span class="sxs-lookup"><span data-stu-id="b4778-735">Type</span></span>| <span data-ttu-id="b4778-736">说明</span><span class="sxs-lookup"><span data-stu-id="b4778-736">Description</span></span>|
|---|---|---|
|`entityType`| [<span data-ttu-id="b4778-737">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="b4778-737">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook_1_2/office.mailboxenums.entitytype)|<span data-ttu-id="b4778-738">EntityType 枚举值之一。</span><span class="sxs-lookup"><span data-stu-id="b4778-738">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="b4778-739">要求</span><span class="sxs-lookup"><span data-stu-id="b4778-739">Requirements</span></span>

|<span data-ttu-id="b4778-740">要求</span><span class="sxs-lookup"><span data-stu-id="b4778-740">Requirement</span></span>| <span data-ttu-id="b4778-741">值</span><span class="sxs-lookup"><span data-stu-id="b4778-741">Value</span></span>|
|---|---|
|[<span data-ttu-id="b4778-742">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="b4778-742">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b4778-743">1.0</span><span class="sxs-lookup"><span data-stu-id="b4778-743">1.0</span></span>|
|[<span data-ttu-id="b4778-744">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="b4778-744">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b4778-745">受限</span><span class="sxs-lookup"><span data-stu-id="b4778-745">Restricted</span></span>|
|[<span data-ttu-id="b4778-746">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="b4778-746">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b4778-747">阅读</span><span class="sxs-lookup"><span data-stu-id="b4778-747">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="b4778-748">返回：</span><span class="sxs-lookup"><span data-stu-id="b4778-748">Returns:</span></span>

<span data-ttu-id="b4778-749">如果在 `entityType` 中传递的值不是 `EntityType` 枚举的有效成员，该方法返回 null。</span><span class="sxs-lookup"><span data-stu-id="b4778-749">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="b4778-750">如果指定类型的任何实体都不存在于该项目的正文中，该方法将返回空数组。</span><span class="sxs-lookup"><span data-stu-id="b4778-750">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="b4778-751">否则，返回的数组中对象的类型取决于 `entityType` 参数中请求实体的类型。</span><span class="sxs-lookup"><span data-stu-id="b4778-751">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="b4778-752">当使用此方法的最低权限级别**受限**时，某些实体类型需要 **ReadItem** 才能进行访问，如下表中所指定。</span><span class="sxs-lookup"><span data-stu-id="b4778-752">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

| <span data-ttu-id="b4778-753">`entityType` 的值</span><span class="sxs-lookup"><span data-stu-id="b4778-753">Value of `entityType`</span></span> | <span data-ttu-id="b4778-754">返回的数组中对象的类型</span><span class="sxs-lookup"><span data-stu-id="b4778-754">Type of objects in returned array</span></span> | <span data-ttu-id="b4778-755">所需权限级别</span><span class="sxs-lookup"><span data-stu-id="b4778-755">Required Permission Level</span></span> |
| --- | --- | --- |
| `Address` | <span data-ttu-id="b4778-756">String</span><span class="sxs-lookup"><span data-stu-id="b4778-756">String</span></span> | <span data-ttu-id="b4778-757">**受限**</span><span class="sxs-lookup"><span data-stu-id="b4778-757">**Restricted**</span></span> |
| `Contact` | <span data-ttu-id="b4778-758">Contact</span><span class="sxs-lookup"><span data-stu-id="b4778-758">Contact</span></span> | <span data-ttu-id="b4778-759">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="b4778-759">**ReadItem**</span></span> |
| `EmailAddress` | <span data-ttu-id="b4778-760">String</span><span class="sxs-lookup"><span data-stu-id="b4778-760">String</span></span> | <span data-ttu-id="b4778-761">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="b4778-761">**ReadItem**</span></span> |
| `MeetingSuggestion` | <span data-ttu-id="b4778-762">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="b4778-762">MeetingSuggestion</span></span> | <span data-ttu-id="b4778-763">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="b4778-763">**ReadItem**</span></span> |
| `PhoneNumber` | <span data-ttu-id="b4778-764">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="b4778-764">PhoneNumber</span></span> | <span data-ttu-id="b4778-765">**受限**</span><span class="sxs-lookup"><span data-stu-id="b4778-765">**Restricted**</span></span> |
| `TaskSuggestion` | <span data-ttu-id="b4778-766">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="b4778-766">TaskSuggestion</span></span> | <span data-ttu-id="b4778-767">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="b4778-767">**ReadItem**</span></span> |
| `URL` | <span data-ttu-id="b4778-768">String</span><span class="sxs-lookup"><span data-stu-id="b4778-768">String</span></span> | <span data-ttu-id="b4778-769">**受限**</span><span class="sxs-lookup"><span data-stu-id="b4778-769">**Restricted**</span></span> |

<span data-ttu-id="b4778-770">类型：Array.<(String|[Contact](/javascript/api/outlook_1_2/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_2/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_2/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_2/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="b4778-770">Type: Array.<(String|[Contact](/javascript/api/outlook_1_2/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_2/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_2/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_2/office.tasksuggestion))></span></span>

##### <a name="example"></a><span data-ttu-id="b4778-771">示例</span><span class="sxs-lookup"><span data-stu-id="b4778-771">Example</span></span>

<span data-ttu-id="b4778-772">以下示例显示了如何访问表示当前项目的正文中的邮政地址的字符串数组。</span><span class="sxs-lookup"><span data-stu-id="b4778-772">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlook12officecontactmeetingsuggestionjavascriptapioutlook12officemeetingsuggestionphonenumberjavascriptapioutlook12officephonenumbertasksuggestionjavascriptapioutlook12officetasksuggestion"></a><span data-ttu-id="b4778-773">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_2/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_2/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_2/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_2/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="b4778-773">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_2/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_2/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_2/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_2/office.tasksuggestion))>}</span></span>

<span data-ttu-id="b4778-774">返回传递清单 XML 文件中定义的命名筛选器的所选项目中的已知实体。</span><span class="sxs-lookup"><span data-stu-id="b4778-774">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="b4778-775">在 Outlook for iOS 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="b4778-775">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="b4778-776">`getFilteredEntitiesByName` 方法返回匹配在具有指定 `FilterName` 元素值的清单 XML 文件中的 [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) 规则元素中定义的正则表达式的实体。</span><span class="sxs-lookup"><span data-stu-id="b4778-776">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="b4778-777">参数：</span><span class="sxs-lookup"><span data-stu-id="b4778-777">Parameters:</span></span>

|<span data-ttu-id="b4778-778">名称</span><span class="sxs-lookup"><span data-stu-id="b4778-778">Name</span></span>| <span data-ttu-id="b4778-779">类型</span><span class="sxs-lookup"><span data-stu-id="b4778-779">Type</span></span>| <span data-ttu-id="b4778-780">说明</span><span class="sxs-lookup"><span data-stu-id="b4778-780">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="b4778-781">字符串</span><span class="sxs-lookup"><span data-stu-id="b4778-781">String</span></span>|<span data-ttu-id="b4778-782">定义筛选器匹配的 `ItemHasKnownEntity` 规则元素的名称。</span><span class="sxs-lookup"><span data-stu-id="b4778-782">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="b4778-783">要求</span><span class="sxs-lookup"><span data-stu-id="b4778-783">Requirements</span></span>

|<span data-ttu-id="b4778-784">要求</span><span class="sxs-lookup"><span data-stu-id="b4778-784">Requirement</span></span>| <span data-ttu-id="b4778-785">值</span><span class="sxs-lookup"><span data-stu-id="b4778-785">Value</span></span>|
|---|---|
|[<span data-ttu-id="b4778-786">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="b4778-786">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b4778-787">1.0</span><span class="sxs-lookup"><span data-stu-id="b4778-787">1.0</span></span>|
|[<span data-ttu-id="b4778-788">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="b4778-788">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b4778-789">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b4778-789">ReadItem</span></span>|
|[<span data-ttu-id="b4778-790">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="b4778-790">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b4778-791">阅读</span><span class="sxs-lookup"><span data-stu-id="b4778-791">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="b4778-792">返回：</span><span class="sxs-lookup"><span data-stu-id="b4778-792">Returns:</span></span>

<span data-ttu-id="b4778-p152">如果具有匹配 `name` 参数的 `FilterName` 元素值的清单中没有任何 `ItemHasKnownEntity` 元素，则该方法返回 `null`。如果 `name` 参数匹配清单中的 `ItemHasKnownEntity` 元素，但在匹配的当前项目中没有实体，则该方法返回一个空数组。</span><span class="sxs-lookup"><span data-stu-id="b4778-p152">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>

<span data-ttu-id="b4778-795">类型：Array.<(String|[Contact](/javascript/api/outlook_1_2/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_2/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_2/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_2/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="b4778-795">Type: Array.<(String|[Contact](/javascript/api/outlook_1_2/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_2/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_2/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_2/office.tasksuggestion))></span></span>

#### <a name="getregexmatches--object"></a><span data-ttu-id="b4778-796">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="b4778-796">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="b4778-797">返回所选项目中匹配在清单 XML 文件中定义的正则表达式的字符串值。</span><span class="sxs-lookup"><span data-stu-id="b4778-797">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="b4778-798">在 Outlook for iOS 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="b4778-798">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="b4778-p153">`getRegExMatches` 方法返回匹配在清单 XML 文件中的每个 `ItemHasRegularExpressionMatch` 或 `ItemHasKnownEntity` 规则元素中定义的正则表达式的字符串。对于 `ItemHasRegularExpressionMatch` 规则，匹配字符串必须发生在该规则指定的项目的属性中。`PropertyName` 简单类型定义支持的属性。</span><span class="sxs-lookup"><span data-stu-id="b4778-p153">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="b4778-802">例如，考虑一个外接程序清单具有以下 `Rule` 元素：</span><span class="sxs-lookup"><span data-stu-id="b4778-802">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="b4778-803">从 `getRegExMatches` 返回的对象应有两个属性：`fruits` 和 `veggies`。</span><span class="sxs-lookup"><span data-stu-id="b4778-803">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

> [!NOTE]
> <span data-ttu-id="b4778-p154">如果在项目的正文属性上指定 `ItemHasRegularExpressionMatch` 规则，则正则表达式应进一步筛选正文并应尝试返回该项目的整个正文。使用正则表达式（如 `.*`）获取项目的整个正文并不总是返回预期的结果。</span><span class="sxs-lookup"><span data-stu-id="b4778-p154">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="requirements"></a><span data-ttu-id="b4778-806">要求</span><span class="sxs-lookup"><span data-stu-id="b4778-806">Requirements</span></span>

|<span data-ttu-id="b4778-807">要求</span><span class="sxs-lookup"><span data-stu-id="b4778-807">Requirement</span></span>| <span data-ttu-id="b4778-808">值</span><span class="sxs-lookup"><span data-stu-id="b4778-808">Value</span></span>|
|---|---|
|[<span data-ttu-id="b4778-809">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="b4778-809">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b4778-810">1.0</span><span class="sxs-lookup"><span data-stu-id="b4778-810">1.0</span></span>|
|[<span data-ttu-id="b4778-811">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="b4778-811">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b4778-812">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b4778-812">ReadItem</span></span>|
|[<span data-ttu-id="b4778-813">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="b4778-813">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b4778-814">阅读</span><span class="sxs-lookup"><span data-stu-id="b4778-814">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="b4778-815">返回：</span><span class="sxs-lookup"><span data-stu-id="b4778-815">Returns:</span></span>

<span data-ttu-id="b4778-p155">一个包含匹配在清单 XML 文件中定义的正则表达式的字符串数组的对象。每个数组的名称等于匹配 `ItemHasRegularExpressionMatch` 规则的 `RegExName` 属性或匹配 `ItemHasKnownEntity` 规则的 `FilterName` 属性的相应值。</span><span class="sxs-lookup"><span data-stu-id="b4778-p155">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<dl class="param-type"><span data-ttu-id="b4778-818">

<dt>
类型</dt>


</span><span class="sxs-lookup"><span data-stu-id="b4778-818">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="b4778-819">对象</span><span class="sxs-lookup"><span data-stu-id="b4778-819">Object</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="b4778-820">示例</span><span class="sxs-lookup"><span data-stu-id="b4778-820">Example</span></span>

<span data-ttu-id="b4778-821">以下示例显示了如何访问正则表达式 <rule> 元素 `fruits` 和 `veggies` 的匹配项的数组，这些元素在清单中指定。</rule></span><span class="sxs-lookup"><span data-stu-id="b4778-821">The following example shows how to access the array of matches for the regular expression <rule>elements `fruits` and `veggies`, which are specified in the manifest.</rule></span></span>

```JavaScript
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veges = allMatches.veggies;
```

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="b4778-822">getRegExMatchesByName(name) → (nullable) {Array.<String>}</span><span class="sxs-lookup"><span data-stu-id="b4778-822">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="b4778-823">返回匹配在清单 XML 文件中定义的命名正则表达式的所选项目中的字符串值。</span><span class="sxs-lookup"><span data-stu-id="b4778-823">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="b4778-824">在 Outlook for iOS 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="b4778-824">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="b4778-825">`getRegExMatchesByName` 方法返回匹配在具有指定 `RegExName` 元素值的清单 XML 文件中的 `ItemHasRegularExpressionMatch` 规则元素中定义的正则表达式的字符串。</span><span class="sxs-lookup"><span data-stu-id="b4778-825">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="b4778-p156">如果在项目的正文属性上指定 `ItemHasRegularExpressionMatch` 规则，则正则表达式应进一步筛选正文，不应尝试返回该项目的整个正文。使用正则表达式（如 `.*`）获取项目的整个正文并不总是返回预期的结果。</span><span class="sxs-lookup"><span data-stu-id="b4778-p156">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="b4778-828">参数：</span><span class="sxs-lookup"><span data-stu-id="b4778-828">Parameters:</span></span>

|<span data-ttu-id="b4778-829">名称</span><span class="sxs-lookup"><span data-stu-id="b4778-829">Name</span></span>| <span data-ttu-id="b4778-830">类型</span><span class="sxs-lookup"><span data-stu-id="b4778-830">Type</span></span>| <span data-ttu-id="b4778-831">说明</span><span class="sxs-lookup"><span data-stu-id="b4778-831">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="b4778-832">字符串</span><span class="sxs-lookup"><span data-stu-id="b4778-832">String</span></span>|<span data-ttu-id="b4778-833">定义筛选器匹配的 `ItemHasRegularExpressionMatch` 规则元素的名称。</span><span class="sxs-lookup"><span data-stu-id="b4778-833">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="b4778-834">要求</span><span class="sxs-lookup"><span data-stu-id="b4778-834">Requirements</span></span>

|<span data-ttu-id="b4778-835">要求</span><span class="sxs-lookup"><span data-stu-id="b4778-835">Requirement</span></span>| <span data-ttu-id="b4778-836">值</span><span class="sxs-lookup"><span data-stu-id="b4778-836">Value</span></span>|
|---|---|
|[<span data-ttu-id="b4778-837">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="b4778-837">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b4778-838">1.0</span><span class="sxs-lookup"><span data-stu-id="b4778-838">1.0</span></span>|
|[<span data-ttu-id="b4778-839">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="b4778-839">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b4778-840">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b4778-840">ReadItem</span></span>|
|[<span data-ttu-id="b4778-841">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="b4778-841">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b4778-842">阅读</span><span class="sxs-lookup"><span data-stu-id="b4778-842">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="b4778-843">返回：</span><span class="sxs-lookup"><span data-stu-id="b4778-843">Returns:</span></span>

<span data-ttu-id="b4778-844">一个包含匹配在清单 XML 文件中定义的正则表达式的字符串的数组。</span><span class="sxs-lookup"><span data-stu-id="b4778-844">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<dl class="param-type"><span data-ttu-id="b4778-845">

<dt>类型</dt>

</span><span class="sxs-lookup"><span data-stu-id="b4778-845">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="b4778-846">Array.< String ></span><span class="sxs-lookup"><span data-stu-id="b4778-846">Array.< String ></span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="b4778-847">示例</span><span class="sxs-lookup"><span data-stu-id="b4778-847">Example</span></span>

```JavaScript
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

####  <a name="getselecteddataasynccoerciontype-options-callback--string"></a><span data-ttu-id="b4778-848">getSelectedDataAsync(coercionType, [options], callback) → {String}</span><span class="sxs-lookup"><span data-stu-id="b4778-848">getSelectedDataAsync(coercionType, [options], callback) → {String}</span></span>

<span data-ttu-id="b4778-849">以异步方式返回邮件的主题或正文中选定的数据。</span><span class="sxs-lookup"><span data-stu-id="b4778-849">Asynchronously returns selected data from the subject or body of a message.</span></span>

<span data-ttu-id="b4778-p157">如果没有选定内容，但光标位于正文或主题中，此方法将会为所选数据返回 null。如果选定的是字段，而不是正文或主题，则此方法返回 `InvalidSelection` 错误。</span><span class="sxs-lookup"><span data-stu-id="b4778-p157">If there is no selection but the cursor is in the body or subject, the method returns null for the selected data. If a field other than the body or subject is selected, the method returns the `InvalidSelection` error.</span></span>

##### <a name="parameters"></a><span data-ttu-id="b4778-852">参数：</span><span class="sxs-lookup"><span data-stu-id="b4778-852">Parameters:</span></span>

|<span data-ttu-id="b4778-853">名称</span><span class="sxs-lookup"><span data-stu-id="b4778-853">Name</span></span>| <span data-ttu-id="b4778-854">类型</span><span class="sxs-lookup"><span data-stu-id="b4778-854">Type</span></span>| <span data-ttu-id="b4778-855">属性</span><span class="sxs-lookup"><span data-stu-id="b4778-855">Attributes</span></span>| <span data-ttu-id="b4778-856">说明</span><span class="sxs-lookup"><span data-stu-id="b4778-856">Description</span></span>|
|---|---|---|---|
|`coercionType`| [<span data-ttu-id="b4778-857">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="b4778-857">Office.CoercionType</span></span>](office.md#coerciontype-string)||<span data-ttu-id="b4778-p158">请求数据的格式。如果为文本，则此方法返回纯文本作为字符串，删除任何显示的 HTML 标记。如果为 HTML，则此方法返回所选文本，不论是纯文本还是 HTML。</span><span class="sxs-lookup"><span data-stu-id="b4778-p158">Requests a format for the data. If Text, the method returns the plain text as a string , removing any HTML tags present. If HTML, the method returns the selected text, whether it is plaintext or HTML.</span></span>|
|`options`| <span data-ttu-id="b4778-861">对象</span><span class="sxs-lookup"><span data-stu-id="b4778-861">Object</span></span>| <span data-ttu-id="b4778-862">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="b4778-862">&lt;optional&gt;</span></span>|<span data-ttu-id="b4778-863">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="b4778-863">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="b4778-864">对象</span><span class="sxs-lookup"><span data-stu-id="b4778-864">Object</span></span>| <span data-ttu-id="b4778-865">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="b4778-865">&lt;optional&gt;</span></span>|<span data-ttu-id="b4778-866">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="b4778-866">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="b4778-867">函数</span><span class="sxs-lookup"><span data-stu-id="b4778-867">function</span></span>||<span data-ttu-id="b4778-868">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="b4778-868">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="b4778-869">若要从回调方法访问所选数据，请调用 `asyncResult.value.data`。</span><span class="sxs-lookup"><span data-stu-id="b4778-869">To access the selected data from the callback method, call `asyncResult.value.data`.</span></span> <span data-ttu-id="b4778-870">若要访问选定内容的源属性，请调用 `asyncResult.value.sourceProperty`，这将为 `body` 或 `subject`。</span><span class="sxs-lookup"><span data-stu-id="b4778-870">To access the source property that the selection comes from, call `asyncResult.value.sourceProperty`, which will be either `body` or `subject`.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="b4778-871">要求</span><span class="sxs-lookup"><span data-stu-id="b4778-871">Requirements</span></span>

|<span data-ttu-id="b4778-872">要求</span><span class="sxs-lookup"><span data-stu-id="b4778-872">Requirement</span></span>| <span data-ttu-id="b4778-873">值</span><span class="sxs-lookup"><span data-stu-id="b4778-873">Value</span></span>|
|---|---|
|[<span data-ttu-id="b4778-874">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="b4778-874">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b4778-875">1.2</span><span class="sxs-lookup"><span data-stu-id="b4778-875">1.2</span></span>|
|[<span data-ttu-id="b4778-876">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="b4778-876">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b4778-877">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="b4778-877">ReadWriteItem</span></span>|
|[<span data-ttu-id="b4778-878">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="b4778-878">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b4778-879">撰写</span><span class="sxs-lookup"><span data-stu-id="b4778-879">Compose</span></span>|

##### <a name="returns"></a><span data-ttu-id="b4778-880">返回：</span><span class="sxs-lookup"><span data-stu-id="b4778-880">Returns:</span></span>

<span data-ttu-id="b4778-881">作为字符串的所选数据的格式由 `coercionType` 确定。</span><span class="sxs-lookup"><span data-stu-id="b4778-881">The selected data as a string with format determined by `coercionType`.</span></span>

<dl class="param-type"><span data-ttu-id="b4778-882">

<dt>
类型</dt>


</span><span class="sxs-lookup"><span data-stu-id="b4778-882">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="b4778-883">字符串</span><span class="sxs-lookup"><span data-stu-id="b4778-883">String</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="b4778-884">示例</span><span class="sxs-lookup"><span data-stu-id="b4778-884">Example</span></span>

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

####  <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="b4778-885">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="b4778-885">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="b4778-886">异步加载所选项目上此外接程序的自定义属性。</span><span class="sxs-lookup"><span data-stu-id="b4778-886">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="b4778-p160">自定义属性基于每个应用、每个项目存储为键/值对。此方法在回调中返回 `CustomProperties` 对象，该回调提供访问特定于当前项目和当前外接程序的自定义属性的方法。自定义属性未在项目上加密，因此这不应用作安全存储。</span><span class="sxs-lookup"><span data-stu-id="b4778-p160">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="b4778-890">参数：</span><span class="sxs-lookup"><span data-stu-id="b4778-890">Parameters:</span></span>

|<span data-ttu-id="b4778-891">名称</span><span class="sxs-lookup"><span data-stu-id="b4778-891">Name</span></span>| <span data-ttu-id="b4778-892">类型</span><span class="sxs-lookup"><span data-stu-id="b4778-892">Type</span></span>| <span data-ttu-id="b4778-893">属性</span><span class="sxs-lookup"><span data-stu-id="b4778-893">Attributes</span></span>| <span data-ttu-id="b4778-894">说明</span><span class="sxs-lookup"><span data-stu-id="b4778-894">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="b4778-895">函数</span><span class="sxs-lookup"><span data-stu-id="b4778-895">function</span></span>||<span data-ttu-id="b4778-896">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="b4778-896">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="b4778-897">自定义属性作为 `asyncResult.value` 属性中的 [`CustomProperties`](/javascript/api/outlook_1_2/office.customproperties) 对象提供。</span><span class="sxs-lookup"><span data-stu-id="b4778-897">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook_1_2/office.customproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="b4778-898">此对象可用于获取、设置以及从项目中删除自定义属性，并将自定义属性集的更改重新保存到服务器。</span><span class="sxs-lookup"><span data-stu-id="b4778-898">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`| <span data-ttu-id="b4778-899">对象</span><span class="sxs-lookup"><span data-stu-id="b4778-899">Object</span></span>| <span data-ttu-id="b4778-900">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="b4778-900">&lt;optional&gt;</span></span>|<span data-ttu-id="b4778-901">开发人员可以提供他们想要在回调函数中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="b4778-901">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="b4778-902">此对象可以通过回调函数中的 `asyncResult.asyncContext` 属性进行访问。</span><span class="sxs-lookup"><span data-stu-id="b4778-902">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="b4778-903">要求</span><span class="sxs-lookup"><span data-stu-id="b4778-903">Requirements</span></span>

|<span data-ttu-id="b4778-904">要求</span><span class="sxs-lookup"><span data-stu-id="b4778-904">Requirement</span></span>| <span data-ttu-id="b4778-905">值</span><span class="sxs-lookup"><span data-stu-id="b4778-905">Value</span></span>|
|---|---|
|[<span data-ttu-id="b4778-906">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="b4778-906">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b4778-907">1.0</span><span class="sxs-lookup"><span data-stu-id="b4778-907">1.0</span></span>|
|[<span data-ttu-id="b4778-908">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="b4778-908">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b4778-909">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b4778-909">ReadItem</span></span>|
|[<span data-ttu-id="b4778-910">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="b4778-910">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b4778-911">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="b4778-911">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="b4778-912">示例</span><span class="sxs-lookup"><span data-stu-id="b4778-912">Example</span></span>

<span data-ttu-id="b4778-p163">以下代码示例显示了如何使用 `loadCustomPropertiesAsync` 方法异步加载特定于当前项目的自定义属性。该示例还显示了如何使用 `CustomProperties.saveAsync` 方法将这些属性重新保存到服务器。加载自定义属性后，该代码示例将使用 `CustomProperties.get` 方法读取自定义属性 `myProp`，使用 `CustomProperties.set` 方法写入自定义属性 `otherProp`，最后调用 `saveAsync` 方法保存这些自定义属性。</span><span class="sxs-lookup"><span data-stu-id="b4778-p163">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

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

####  <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="b4778-916">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="b4778-916">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="b4778-917">将附件从邮件或约会中删除。</span><span class="sxs-lookup"><span data-stu-id="b4778-917">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="b4778-p164">`removeAttachmentAsync` 方法删除项目中带指定标识符的附件。最佳做法是，仅当同一个邮件应用程序在同一会话中添加了一个附件时，你才应使用该附件标识符来删除该附件。在 Outlook Web App 和适用于设备的 OWA 中，附件标识符只在同一个会话中才有效。当用户关闭应用程序，或者如果用户开始在内嵌窗体中撰写，并在随后弹出的内嵌窗体中继续在单独的窗口撰写时，会话即结束。</span><span class="sxs-lookup"><span data-stu-id="b4778-p164">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item. As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session. In Outlook Web App and OWA for Devices, the attachment identifier is valid only within the same session. A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="b4778-922">参数：</span><span class="sxs-lookup"><span data-stu-id="b4778-922">Parameters:</span></span>

|<span data-ttu-id="b4778-923">名称</span><span class="sxs-lookup"><span data-stu-id="b4778-923">Name</span></span>| <span data-ttu-id="b4778-924">类型</span><span class="sxs-lookup"><span data-stu-id="b4778-924">Type</span></span>| <span data-ttu-id="b4778-925">属性</span><span class="sxs-lookup"><span data-stu-id="b4778-925">Attributes</span></span>| <span data-ttu-id="b4778-926">说明</span><span class="sxs-lookup"><span data-stu-id="b4778-926">Description</span></span>|
|---|---|---|---|
|`attachmentId`| <span data-ttu-id="b4778-927">字符串</span><span class="sxs-lookup"><span data-stu-id="b4778-927">String</span></span>||<span data-ttu-id="b4778-928">要删除的附件的标识符。</span><span class="sxs-lookup"><span data-stu-id="b4778-928">The identifier of the attachment to remove.</span></span>|
|`options`| <span data-ttu-id="b4778-929">对象</span><span class="sxs-lookup"><span data-stu-id="b4778-929">Object</span></span>| <span data-ttu-id="b4778-930">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="b4778-930">&lt;optional&gt;</span></span>|<span data-ttu-id="b4778-931">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="b4778-931">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="b4778-932">对象</span><span class="sxs-lookup"><span data-stu-id="b4778-932">Object</span></span>| <span data-ttu-id="b4778-933">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="b4778-933">&lt;optional&gt;</span></span>|<span data-ttu-id="b4778-934">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="b4778-934">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="b4778-935">函数</span><span class="sxs-lookup"><span data-stu-id="b4778-935">function</span></span>| <span data-ttu-id="b4778-936">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="b4778-936">&lt;optional&gt;</span></span>|<span data-ttu-id="b4778-937">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="b4778-937">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="b4778-938">如果删除附件失败，`asyncResult.error` 属性将包含一个说明失败原因的错误代码。</span><span class="sxs-lookup"><span data-stu-id="b4778-938">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="b4778-939">错误</span><span class="sxs-lookup"><span data-stu-id="b4778-939">Errors</span></span>

| <span data-ttu-id="b4778-940">错误代码</span><span class="sxs-lookup"><span data-stu-id="b4778-940">Error code</span></span> | <span data-ttu-id="b4778-941">说明</span><span class="sxs-lookup"><span data-stu-id="b4778-941">Description</span></span> |
|------------|-------------|
| `InvalidAttachmentId` | <span data-ttu-id="b4778-942">附件标识符不存在。</span><span class="sxs-lookup"><span data-stu-id="b4778-942">The attachment identifier does not exist.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="b4778-943">要求</span><span class="sxs-lookup"><span data-stu-id="b4778-943">Requirements</span></span>

|<span data-ttu-id="b4778-944">要求</span><span class="sxs-lookup"><span data-stu-id="b4778-944">Requirement</span></span>| <span data-ttu-id="b4778-945">值</span><span class="sxs-lookup"><span data-stu-id="b4778-945">Value</span></span>|
|---|---|
|[<span data-ttu-id="b4778-946">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="b4778-946">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b4778-947">1.1</span><span class="sxs-lookup"><span data-stu-id="b4778-947">1.1</span></span>|
|[<span data-ttu-id="b4778-948">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="b4778-948">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b4778-949">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="b4778-949">ReadWriteItem</span></span>|
|[<span data-ttu-id="b4778-950">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="b4778-950">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b4778-951">撰写</span><span class="sxs-lookup"><span data-stu-id="b4778-951">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="b4778-952">示例</span><span class="sxs-lookup"><span data-stu-id="b4778-952">Example</span></span>

<span data-ttu-id="b4778-953">以下代码删除包含标识符 '0' 的附件。</span><span class="sxs-lookup"><span data-stu-id="b4778-953">The following code removes an attachment with an identifier of '0'.</span></span>

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

####  <a name="setselecteddataasyncdata-options-callback"></a><span data-ttu-id="b4778-954">setSelectedDataAsync(data, [options], callback)</span><span class="sxs-lookup"><span data-stu-id="b4778-954">setSelectedDataAsync(data, [options], callback)</span></span>

<span data-ttu-id="b4778-955">以异步方式将数据插入到邮件的正文或主题中。</span><span class="sxs-lookup"><span data-stu-id="b4778-955">Asynchronously inserts data into the body or subject of a message.</span></span>

<span data-ttu-id="b4778-p165">`setSelectedDataAsync` 方法将指定字符串插入到项目主题或正文的光标位置，或者，如果在编辑器中已选择文本，则该方法将替换选择的文本。如果光标不在正文或主题字段中，则返回错误。插入之后，光标会位于插入内容的末尾。</span><span class="sxs-lookup"><span data-stu-id="b4778-p165">The `setSelectedDataAsync` method inserts the specified string at the cursor location in the subject or body of the item, or, if text is selected in the editor, it replaces the selected text. If the cursor is not in the body or subject field, an error is returned. After insertion, the cursor is placed at the end of the inserted content.</span></span>

##### <a name="parameters"></a><span data-ttu-id="b4778-959">参数：</span><span class="sxs-lookup"><span data-stu-id="b4778-959">Parameters:</span></span>

|<span data-ttu-id="b4778-960">名称</span><span class="sxs-lookup"><span data-stu-id="b4778-960">Name</span></span>| <span data-ttu-id="b4778-961">类型</span><span class="sxs-lookup"><span data-stu-id="b4778-961">Type</span></span>| <span data-ttu-id="b4778-962">属性</span><span class="sxs-lookup"><span data-stu-id="b4778-962">Attributes</span></span>| <span data-ttu-id="b4778-963">说明</span><span class="sxs-lookup"><span data-stu-id="b4778-963">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="b4778-964">字符串</span><span class="sxs-lookup"><span data-stu-id="b4778-964">String</span></span>||<span data-ttu-id="b4778-p166">要插入的数据。数据不得超过 1,000,000 个字符。如果传入的数据超过 1,000,000 个字符，则会引发 `ArgumentOutOfRange` 异常。</span><span class="sxs-lookup"><span data-stu-id="b4778-p166">The data to be inserted. Data is not to exceed 1,000,000 characters. If more than 1,000,000 characters are passed in, an `ArgumentOutOfRange` exception is thrown.</span></span>|
|`options`| <span data-ttu-id="b4778-968">Object</span><span class="sxs-lookup"><span data-stu-id="b4778-968">Object</span></span>| <span data-ttu-id="b4778-969">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="b4778-969">&lt;optional&gt;</span></span>|<span data-ttu-id="b4778-970">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="b4778-970">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="b4778-971">对象</span><span class="sxs-lookup"><span data-stu-id="b4778-971">Object</span></span>| <span data-ttu-id="b4778-972">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="b4778-972">&lt;optional&gt;</span></span>|<span data-ttu-id="b4778-973">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="b4778-973">Developers can provide any object they wish to access in the callback method.</span></span>|
|`options.coercionType`| [<span data-ttu-id="b4778-974">Office.CoercionType</span><span class="sxs-lookup"><span data-stu-id="b4778-974">Office.CoercionType</span></span>](office.md#coerciontype-string)| <span data-ttu-id="b4778-975">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="b4778-975">&lt;optional&gt;</span></span>|<span data-ttu-id="b4778-p167">如果为 `text`，则在 Outlook Web App 和 Outlook 中应用当前样式。如果该字段是 HTML 编辑器，则仅插入文本数据，即使数据为 HTML。</span><span class="sxs-lookup"><span data-stu-id="b4778-p167">If `text`, the current style is applied in Outlook Web App and Outlook. If the field is an HTML editor, only the text data is inserted, even if the data is HTML.</span></span><br/><br/><span data-ttu-id="b4778-p168">如果 `html` 和该字段支持 HTML（主题不支持），则在 Outlook Web App 中应用当前样式，而在 Outlook 中应用默认样式。如果该字段是文本字段，则返回 `InvalidDataFormat` 错误。</span><span class="sxs-lookup"><span data-stu-id="b4778-p168">If `html` and the field supports HTML (the subject doesn't), the current style is applied in Outlook Web App and the default style is applied in Outlook. If the field is a text field, an `InvalidDataFormat` error is returned.</span></span><br/><br/><span data-ttu-id="b4778-980">如果未设置 `coercionType`，则结果取决于该字段：如果该字段是 HTML，则使用 HTML；如果该字段是文本，则使用纯文本。</span><span class="sxs-lookup"><span data-stu-id="b4778-980">If `coercionType` is not set, the result depends on the field: if the field is HTML then HTML is used; if the field is text, then plain text is used.</span></span>|
|`callback`| <span data-ttu-id="b4778-981">function</span><span class="sxs-lookup"><span data-stu-id="b4778-981">function</span></span>||<span data-ttu-id="b4778-982">方法完成后，使用单个参数 `callback`（一个 [`asyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `AsyncResult` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="b4778-982">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="b4778-983">Requirements</span><span class="sxs-lookup"><span data-stu-id="b4778-983">Requirements</span></span>

|<span data-ttu-id="b4778-984">要求</span><span class="sxs-lookup"><span data-stu-id="b4778-984">Requirement</span></span>| <span data-ttu-id="b4778-985">值</span><span class="sxs-lookup"><span data-stu-id="b4778-985">Value</span></span>|
|---|---|
|[<span data-ttu-id="b4778-986">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="b4778-986">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b4778-987">1.2</span><span class="sxs-lookup"><span data-stu-id="b4778-987">1.2</span></span>|
|[<span data-ttu-id="b4778-988">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="b4778-988">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b4778-989">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="b4778-989">ReadWriteItem</span></span>|
|[<span data-ttu-id="b4778-990">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="b4778-990">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b4778-991">撰写</span><span class="sxs-lookup"><span data-stu-id="b4778-991">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="b4778-992">示例</span><span class="sxs-lookup"><span data-stu-id="b4778-992">Example</span></span>

```JavaScript
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```
