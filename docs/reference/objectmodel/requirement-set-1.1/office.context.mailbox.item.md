---
title: Office.context.mailbox.item-要求设置 1.1
description: ''
ms.date: 12/18/2018
localization_priority: Normal
ms.openlocfilehash: 63460494a049bb83d3af69f6808396e426842f1e
ms.sourcegitcommit: d1aa7201820176ed986b9f00bb9c88e055906c77
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 01/23/2019
ms.locfileid: "29389576"
---
# <a name="item"></a><span data-ttu-id="8972a-102">item</span><span class="sxs-lookup"><span data-stu-id="8972a-102">item</span></span>

### <span data-ttu-id="8972a-p101">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md). item</span><span class="sxs-lookup"><span data-stu-id="8972a-p101">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md). item</span></span>

<span data-ttu-id="8972a-p102">`item` 命名空间用于访问当前选定的邮件、会议请求或约会。可以通过使用 [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlook11officemailboxenumsitemtype) 属性确定 `item` 的类型。</span><span class="sxs-lookup"><span data-stu-id="8972a-p102">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlook11officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="8972a-107">Requirements</span><span class="sxs-lookup"><span data-stu-id="8972a-107">Requirements</span></span>

|<span data-ttu-id="8972a-108">要求</span><span class="sxs-lookup"><span data-stu-id="8972a-108">Requirement</span></span>| <span data-ttu-id="8972a-109">值</span><span class="sxs-lookup"><span data-stu-id="8972a-109">Value</span></span>|
|---|---|
|[<span data-ttu-id="8972a-110">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="8972a-110">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8972a-111">1.0</span><span class="sxs-lookup"><span data-stu-id="8972a-111">1.0</span></span>|
|[<span data-ttu-id="8972a-112">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="8972a-112">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8972a-113">受限</span><span class="sxs-lookup"><span data-stu-id="8972a-113">Restricted</span></span>|
|[<span data-ttu-id="8972a-114">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="8972a-114">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="8972a-115">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="8972a-115">Compose or read</span></span>|

### <a name="example"></a><span data-ttu-id="8972a-116">示例</span><span class="sxs-lookup"><span data-stu-id="8972a-116">Example</span></span>

<span data-ttu-id="8972a-117">以下 JavaScript 代码示例显示了如何访问 Outlook 中当前项目的 `subject` 属性。</span><span class="sxs-lookup"><span data-stu-id="8972a-117">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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

### <a name="members"></a><span data-ttu-id="8972a-118">成员</span><span class="sxs-lookup"><span data-stu-id="8972a-118">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlook11officeattachmentdetails"></a><span data-ttu-id="8972a-119">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_1/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="8972a-119">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_1/office.attachmentdetails)></span></span>

<span data-ttu-id="8972a-p103">获取项目的附件数组。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="8972a-p103">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="8972a-122">由于潜在的安全问题，某些类型的文件会受到 Outlook 阻止，并且不会返回。</span><span class="sxs-lookup"><span data-stu-id="8972a-122">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="8972a-123">如需了解更多信息，请参阅 [Outlook 中阻止的附件](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519)。</span><span class="sxs-lookup"><span data-stu-id="8972a-123">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="8972a-124">类型：</span><span class="sxs-lookup"><span data-stu-id="8972a-124">Type:</span></span>

*   <span data-ttu-id="8972a-125">Array.<[AttachmentDetails](/javascript/api/outlook_1_1/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="8972a-125">Array.<[AttachmentDetails](/javascript/api/outlook_1_1/office.attachmentdetails)></span></span>

##### <a name="requirements"></a><span data-ttu-id="8972a-126">要求</span><span class="sxs-lookup"><span data-stu-id="8972a-126">Requirements</span></span>

|<span data-ttu-id="8972a-127">要求</span><span class="sxs-lookup"><span data-stu-id="8972a-127">Requirement</span></span>| <span data-ttu-id="8972a-128">值</span><span class="sxs-lookup"><span data-stu-id="8972a-128">Value</span></span>|
|---|---|
|[<span data-ttu-id="8972a-129">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="8972a-129">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8972a-130">1.0</span><span class="sxs-lookup"><span data-stu-id="8972a-130">1.0</span></span>|
|[<span data-ttu-id="8972a-131">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="8972a-131">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8972a-132">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8972a-132">ReadItem</span></span>|
|[<span data-ttu-id="8972a-133">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="8972a-133">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="8972a-134">阅读</span><span class="sxs-lookup"><span data-stu-id="8972a-134">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="8972a-135">示例</span><span class="sxs-lookup"><span data-stu-id="8972a-135">Example</span></span>

<span data-ttu-id="8972a-136">以下代码使用当前项目上所有附件的详细信息构成 HTML 字符串。</span><span class="sxs-lookup"><span data-stu-id="8972a-136">The following code builds an HTML string with details of all attachments on the current item.</span></span>

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

####  <a name="bcc-recipientsjavascriptapioutlook11officerecipients"></a><span data-ttu-id="8972a-137">bcc :[Recipients](/javascript/api/outlook_1_1/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="8972a-137">bcc :[Recipients](/javascript/api/outlook_1_1/office.recipients)</span></span>

<span data-ttu-id="8972a-138">获取对象，该对象提供用于获取或更新邮件的密件抄送 (Bcc) 行上的收件人的方法。</span><span class="sxs-lookup"><span data-stu-id="8972a-138">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="8972a-139">仅限撰写模式。</span><span class="sxs-lookup"><span data-stu-id="8972a-139">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="8972a-140">类型：</span><span class="sxs-lookup"><span data-stu-id="8972a-140">Type:</span></span>

*   [<span data-ttu-id="8972a-141">收件人</span><span class="sxs-lookup"><span data-stu-id="8972a-141">Recipients</span></span>](/javascript/api/outlook_1_1/office.recipients)

##### <a name="requirements"></a><span data-ttu-id="8972a-142">要求</span><span class="sxs-lookup"><span data-stu-id="8972a-142">Requirements</span></span>

|<span data-ttu-id="8972a-143">要求</span><span class="sxs-lookup"><span data-stu-id="8972a-143">Requirement</span></span>| <span data-ttu-id="8972a-144">值</span><span class="sxs-lookup"><span data-stu-id="8972a-144">Value</span></span>|
|---|---|
|[<span data-ttu-id="8972a-145">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="8972a-145">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8972a-146">1.1</span><span class="sxs-lookup"><span data-stu-id="8972a-146">1.1</span></span>|
|[<span data-ttu-id="8972a-147">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="8972a-147">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8972a-148">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8972a-148">ReadItem</span></span>|
|[<span data-ttu-id="8972a-149">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="8972a-149">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="8972a-150">撰写</span><span class="sxs-lookup"><span data-stu-id="8972a-150">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="8972a-151">示例</span><span class="sxs-lookup"><span data-stu-id="8972a-151">Example</span></span>

```JavaScript
Office.context.mailbox.item.bcc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.bcc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.bcc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfBccRecipients = asyncResult.value;
}
```

####  <a name="body-bodyjavascriptapioutlook11officebody"></a><span data-ttu-id="8972a-152">body :[Body](/javascript/api/outlook_1_1/office.body)</span><span class="sxs-lookup"><span data-stu-id="8972a-152">body :[Body](/javascript/api/outlook_1_1/office.body)</span></span>

<span data-ttu-id="8972a-153">获取一个提供用于处理项目正文的方法的对象。</span><span class="sxs-lookup"><span data-stu-id="8972a-153">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="8972a-154">类型：</span><span class="sxs-lookup"><span data-stu-id="8972a-154">Type:</span></span>

*   [<span data-ttu-id="8972a-155">Body</span><span class="sxs-lookup"><span data-stu-id="8972a-155">Body</span></span>](/javascript/api/outlook_1_1/office.body)

##### <a name="requirements"></a><span data-ttu-id="8972a-156">要求</span><span class="sxs-lookup"><span data-stu-id="8972a-156">Requirements</span></span>

|<span data-ttu-id="8972a-157">要求</span><span class="sxs-lookup"><span data-stu-id="8972a-157">Requirement</span></span>| <span data-ttu-id="8972a-158">值</span><span class="sxs-lookup"><span data-stu-id="8972a-158">Value</span></span>|
|---|---|
|[<span data-ttu-id="8972a-159">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="8972a-159">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8972a-160">1.1</span><span class="sxs-lookup"><span data-stu-id="8972a-160">1.1</span></span>|
|[<span data-ttu-id="8972a-161">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="8972a-161">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8972a-162">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8972a-162">ReadItem</span></span>|
|[<span data-ttu-id="8972a-163">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="8972a-163">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="8972a-164">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="8972a-164">Compose or read</span></span>|

####  <a name="cc-arrayemailaddressdetailsjavascriptapioutlook11officeemailaddressdetailsrecipientsjavascriptapioutlook11officerecipients"></a><span data-ttu-id="8972a-165">cc :Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_1/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="8972a-165">cc :Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_1/office.recipients)</span></span>

<span data-ttu-id="8972a-166">提供对邮件的抄送 (Cc) 收件人的访问权限。</span><span class="sxs-lookup"><span data-stu-id="8972a-166">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="8972a-167">对象的类型和访问级别取决于当前项目的模式。</span><span class="sxs-lookup"><span data-stu-id="8972a-167">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="8972a-168">阅读模式</span><span class="sxs-lookup"><span data-stu-id="8972a-168">Read mode</span></span>

<span data-ttu-id="8972a-p107">`cc` 属性返回包含邮件的**抄送**行上所列的每个收件人的 `EmailAddressDetails` 对象的数组。集合上限为 100 个成员。</span><span class="sxs-lookup"><span data-stu-id="8972a-p107">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="8972a-171">撰写模式</span><span class="sxs-lookup"><span data-stu-id="8972a-171">Compose mode</span></span>

<span data-ttu-id="8972a-172">`cc` 属性返回一个 `Recipients` 对象，该对象提供用于获取或更新邮件的“**抄送**”行上收件人的方法。</span><span class="sxs-lookup"><span data-stu-id="8972a-172">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="8972a-173">类型：</span><span class="sxs-lookup"><span data-stu-id="8972a-173">Type:</span></span>

*   <span data-ttu-id="8972a-174">Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_1/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="8972a-174">Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_1/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="8972a-175">要求</span><span class="sxs-lookup"><span data-stu-id="8972a-175">Requirements</span></span>

|<span data-ttu-id="8972a-176">要求</span><span class="sxs-lookup"><span data-stu-id="8972a-176">Requirement</span></span>| <span data-ttu-id="8972a-177">值</span><span class="sxs-lookup"><span data-stu-id="8972a-177">Value</span></span>|
|---|---|
|[<span data-ttu-id="8972a-178">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="8972a-178">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8972a-179">1.0</span><span class="sxs-lookup"><span data-stu-id="8972a-179">1.0</span></span>|
|[<span data-ttu-id="8972a-180">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="8972a-180">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8972a-181">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8972a-181">ReadItem</span></span>|
|[<span data-ttu-id="8972a-182">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="8972a-182">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="8972a-183">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="8972a-183">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="8972a-184">示例</span><span class="sxs-lookup"><span data-stu-id="8972a-184">Example</span></span>

```JavaScript
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

####  <a name="nullable-conversationid-string"></a><span data-ttu-id="8972a-185">(nullable) conversationId :String</span><span class="sxs-lookup"><span data-stu-id="8972a-185">(nullable) conversationId :String</span></span>

<span data-ttu-id="8972a-186">获取包含特定消息的电子邮件会话的标识符。</span><span class="sxs-lookup"><span data-stu-id="8972a-186">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="8972a-p108">如果在阅读窗体或撰写窗体的回复中激活邮件应用程序，则此属性可以获得一个整数值。如果用户随后更改了回复邮件的主题（若发送回复），则该邮件的对话 ID 将改变且之前获取的值将不适用。</span><span class="sxs-lookup"><span data-stu-id="8972a-p108">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="8972a-p109">对于撰写窗体的新项目，此属性获得一个 null 值。如果用户设置一个主题并保存该项目，`conversationId` 属性将返回一个值。</span><span class="sxs-lookup"><span data-stu-id="8972a-p109">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="8972a-191">类型:</span><span class="sxs-lookup"><span data-stu-id="8972a-191">Type:</span></span>

*   <span data-ttu-id="8972a-192">String</span><span class="sxs-lookup"><span data-stu-id="8972a-192">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="8972a-193">要求</span><span class="sxs-lookup"><span data-stu-id="8972a-193">Requirements</span></span>

|<span data-ttu-id="8972a-194">要求</span><span class="sxs-lookup"><span data-stu-id="8972a-194">Requirement</span></span>| <span data-ttu-id="8972a-195">值</span><span class="sxs-lookup"><span data-stu-id="8972a-195">Value</span></span>|
|---|---|
|[<span data-ttu-id="8972a-196">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="8972a-196">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8972a-197">1.0</span><span class="sxs-lookup"><span data-stu-id="8972a-197">1.0</span></span>|
|[<span data-ttu-id="8972a-198">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="8972a-198">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8972a-199">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8972a-199">ReadItem</span></span>|
|[<span data-ttu-id="8972a-200">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="8972a-200">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="8972a-201">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="8972a-201">Compose or read</span></span>|

#### <a name="datetimecreated-date"></a><span data-ttu-id="8972a-202">dateTimeCreated :Date</span><span class="sxs-lookup"><span data-stu-id="8972a-202">dateTimeCreated :Date</span></span>

<span data-ttu-id="8972a-p110">获取项目创建的日期和时间。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="8972a-p110">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="8972a-205">类型：</span><span class="sxs-lookup"><span data-stu-id="8972a-205">Type:</span></span>

*   <span data-ttu-id="8972a-206">日期</span><span class="sxs-lookup"><span data-stu-id="8972a-206">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="8972a-207">要求</span><span class="sxs-lookup"><span data-stu-id="8972a-207">Requirements</span></span>

|<span data-ttu-id="8972a-208">要求</span><span class="sxs-lookup"><span data-stu-id="8972a-208">Requirement</span></span>| <span data-ttu-id="8972a-209">值</span><span class="sxs-lookup"><span data-stu-id="8972a-209">Value</span></span>|
|---|---|
|[<span data-ttu-id="8972a-210">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="8972a-210">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8972a-211">1.0</span><span class="sxs-lookup"><span data-stu-id="8972a-211">1.0</span></span>|
|[<span data-ttu-id="8972a-212">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="8972a-212">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8972a-213">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8972a-213">ReadItem</span></span>|
|[<span data-ttu-id="8972a-214">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="8972a-214">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="8972a-215">阅读</span><span class="sxs-lookup"><span data-stu-id="8972a-215">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="8972a-216">示例</span><span class="sxs-lookup"><span data-stu-id="8972a-216">Example</span></span>

```JavaScript
var created = Office.context.mailbox.item.dateTimeCreated;
```

#### <a name="datetimemodified-date"></a><span data-ttu-id="8972a-217">dateTimeModified :Date</span><span class="sxs-lookup"><span data-stu-id="8972a-217">dateTimeModified :Date</span></span>

<span data-ttu-id="8972a-p111">获取项目最近一次修改的日期和时间。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="8972a-p111">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="8972a-220">在 Outlook for iOS 或 Outlook for Android 中不支持此成员。</span><span class="sxs-lookup"><span data-stu-id="8972a-220">This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="type"></a><span data-ttu-id="8972a-221">类型:</span><span class="sxs-lookup"><span data-stu-id="8972a-221">Type:</span></span>

*   <span data-ttu-id="8972a-222">日期</span><span class="sxs-lookup"><span data-stu-id="8972a-222">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="8972a-223">要求</span><span class="sxs-lookup"><span data-stu-id="8972a-223">Requirements</span></span>

|<span data-ttu-id="8972a-224">要求</span><span class="sxs-lookup"><span data-stu-id="8972a-224">Requirement</span></span>| <span data-ttu-id="8972a-225">值</span><span class="sxs-lookup"><span data-stu-id="8972a-225">Value</span></span>|
|---|---|
|[<span data-ttu-id="8972a-226">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="8972a-226">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8972a-227">1.0</span><span class="sxs-lookup"><span data-stu-id="8972a-227">1.0</span></span>|
|[<span data-ttu-id="8972a-228">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="8972a-228">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8972a-229">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8972a-229">ReadItem</span></span>|
|[<span data-ttu-id="8972a-230">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="8972a-230">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="8972a-231">阅读</span><span class="sxs-lookup"><span data-stu-id="8972a-231">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="8972a-232">示例</span><span class="sxs-lookup"><span data-stu-id="8972a-232">Example</span></span>

```JavaScript
var modified = Office.context.mailbox.item.dateTimeModified;
```

####  <a name="end-datetimejavascriptapioutlook11officetime"></a><span data-ttu-id="8972a-233">end :Date|[Time](/javascript/api/outlook_1_1/office.time)</span><span class="sxs-lookup"><span data-stu-id="8972a-233">end :Date|[Time](/javascript/api/outlook_1_1/office.time)</span></span>

<span data-ttu-id="8972a-234">获取或设置约会结束的日期和时间。</span><span class="sxs-lookup"><span data-stu-id="8972a-234">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="8972a-p112">将 `end` 属性表示为协调世界时 (UTC) 的日期和时间值。可使用 [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook11officelocalclienttime) 方法将 end 属性值转换为客户端的本地日期和时间。</span><span class="sxs-lookup"><span data-stu-id="8972a-p112">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook11officelocalclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="8972a-237">阅读模式</span><span class="sxs-lookup"><span data-stu-id="8972a-237">Read mode</span></span>

<span data-ttu-id="8972a-238">`end` 属性返回 `Date` 对象。</span><span class="sxs-lookup"><span data-stu-id="8972a-238">The `end` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="8972a-239">撰写模式</span><span class="sxs-lookup"><span data-stu-id="8972a-239">Compose mode</span></span>

<span data-ttu-id="8972a-240">`end` 属性返回 `Time` 对象。</span><span class="sxs-lookup"><span data-stu-id="8972a-240">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="8972a-241">使用 [`Time.setAsync`](/javascript/api/outlook_1_1/office.time#setasync-datetime--options--callback-) 方法设置结束时间时，应使用 [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) 方法将客户端的本地时间转换为服务器的 UTC。</span><span class="sxs-lookup"><span data-stu-id="8972a-241">When you use the [`Time.setAsync`](/javascript/api/outlook_1_1/office.time#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="8972a-242">类型：</span><span class="sxs-lookup"><span data-stu-id="8972a-242">Type:</span></span>

*   <span data-ttu-id="8972a-243">Date | [Time](/javascript/api/outlook_1_1/office.time)</span><span class="sxs-lookup"><span data-stu-id="8972a-243">Date | [Time](/javascript/api/outlook_1_1/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="8972a-244">要求</span><span class="sxs-lookup"><span data-stu-id="8972a-244">Requirements</span></span>

|<span data-ttu-id="8972a-245">要求</span><span class="sxs-lookup"><span data-stu-id="8972a-245">Requirement</span></span>| <span data-ttu-id="8972a-246">值</span><span class="sxs-lookup"><span data-stu-id="8972a-246">Value</span></span>|
|---|---|
|[<span data-ttu-id="8972a-247">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="8972a-247">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8972a-248">1.0</span><span class="sxs-lookup"><span data-stu-id="8972a-248">1.0</span></span>|
|[<span data-ttu-id="8972a-249">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="8972a-249">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8972a-250">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8972a-250">ReadItem</span></span>|
|[<span data-ttu-id="8972a-251">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="8972a-251">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="8972a-252">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="8972a-252">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="8972a-253">示例</span><span class="sxs-lookup"><span data-stu-id="8972a-253">Example</span></span>

<span data-ttu-id="8972a-254">以下示例通过使用 `Time` 对象的 [`setAsync`](/javascript/api/outlook_1_1/office.time#setasync-datetime--options--callback-) 方法，设置撰写模式下约会的结束时间。</span><span class="sxs-lookup"><span data-stu-id="8972a-254">The following example sets the end time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_1/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

#### <a name="from-emailaddressdetailsjavascriptapioutlook11officeemailaddressdetails"></a><span data-ttu-id="8972a-255">from :[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="8972a-255">from :[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)</span></span>

<span data-ttu-id="8972a-p113">获取邮件发件人的电子邮件地址。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="8972a-p113">Gets the email address of the sender of a message. Read mode only.</span></span>

<span data-ttu-id="8972a-p114">`from` 和 [`sender`](#sender-emailaddressdetailsjavascriptapioutlook11officeemailaddressdetails) 属性表示同一个人，邮件由代理人发送的除外。在此情况下，`from` 属性表示代理程序，而 sender 属性表示代理人。</span><span class="sxs-lookup"><span data-stu-id="8972a-p114">The `from` and [`sender`](#sender-emailaddressdetailsjavascriptapioutlook11officeemailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="8972a-260">`from` 属性中 `EmailAddressDetails` 对象的 `recipientType` 属性为 `undefined`。</span><span class="sxs-lookup"><span data-stu-id="8972a-260">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="8972a-261">类型：</span><span class="sxs-lookup"><span data-stu-id="8972a-261">Type:</span></span>

*   [<span data-ttu-id="8972a-262">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="8972a-262">EmailAddressDetails</span></span>](/javascript/api/outlook_1_1/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="8972a-263">要求</span><span class="sxs-lookup"><span data-stu-id="8972a-263">Requirements</span></span>

|<span data-ttu-id="8972a-264">要求</span><span class="sxs-lookup"><span data-stu-id="8972a-264">Requirement</span></span>| <span data-ttu-id="8972a-265">值</span><span class="sxs-lookup"><span data-stu-id="8972a-265">Value</span></span>|
|---|---|
|[<span data-ttu-id="8972a-266">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="8972a-266">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8972a-267">1.0</span><span class="sxs-lookup"><span data-stu-id="8972a-267">1.0</span></span>|
|[<span data-ttu-id="8972a-268">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="8972a-268">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8972a-269">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8972a-269">ReadItem</span></span>|
|[<span data-ttu-id="8972a-270">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="8972a-270">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="8972a-271">阅读</span><span class="sxs-lookup"><span data-stu-id="8972a-271">Read</span></span>|

#### <a name="internetmessageid-string"></a><span data-ttu-id="8972a-272">internetMessageId :String</span><span class="sxs-lookup"><span data-stu-id="8972a-272">internetMessageId :String</span></span>

<span data-ttu-id="8972a-p115">获取电子邮件的 Internet 消息标识符。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="8972a-p115">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="8972a-275">类型：</span><span class="sxs-lookup"><span data-stu-id="8972a-275">Type:</span></span>

*   <span data-ttu-id="8972a-276">String</span><span class="sxs-lookup"><span data-stu-id="8972a-276">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="8972a-277">要求</span><span class="sxs-lookup"><span data-stu-id="8972a-277">Requirements</span></span>

|<span data-ttu-id="8972a-278">要求</span><span class="sxs-lookup"><span data-stu-id="8972a-278">Requirement</span></span>| <span data-ttu-id="8972a-279">值</span><span class="sxs-lookup"><span data-stu-id="8972a-279">Value</span></span>|
|---|---|
|[<span data-ttu-id="8972a-280">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="8972a-280">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8972a-281">1.0</span><span class="sxs-lookup"><span data-stu-id="8972a-281">1.0</span></span>|
|[<span data-ttu-id="8972a-282">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="8972a-282">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8972a-283">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8972a-283">ReadItem</span></span>|
|[<span data-ttu-id="8972a-284">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="8972a-284">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="8972a-285">阅读</span><span class="sxs-lookup"><span data-stu-id="8972a-285">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="8972a-286">示例</span><span class="sxs-lookup"><span data-stu-id="8972a-286">Example</span></span>

```JavaScript
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

#### <a name="itemclass-string"></a><span data-ttu-id="8972a-287">itemClass :String</span><span class="sxs-lookup"><span data-stu-id="8972a-287">itemClass :String</span></span>

<span data-ttu-id="8972a-p116">获取选定项目的 Exchange Web 服务项目类。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="8972a-p116">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="8972a-p117">`itemClass` 属性指定所选项目的邮件类别。以下是邮件或约会项目的默认邮件类别。</span><span class="sxs-lookup"><span data-stu-id="8972a-p117">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

| <span data-ttu-id="8972a-292">类型</span><span class="sxs-lookup"><span data-stu-id="8972a-292">Type</span></span> | <span data-ttu-id="8972a-293">说明</span><span class="sxs-lookup"><span data-stu-id="8972a-293">Description</span></span> | <span data-ttu-id="8972a-294">项目类</span><span class="sxs-lookup"><span data-stu-id="8972a-294">item class</span></span> |
| --- | --- | --- |
| <span data-ttu-id="8972a-295">约会项目</span><span class="sxs-lookup"><span data-stu-id="8972a-295">Appointment items</span></span> | <span data-ttu-id="8972a-296">这些是项目类 `IPM.Appointment` 或 `IPM.Appointment.Occurence` 的日历项目。</span><span class="sxs-lookup"><span data-stu-id="8972a-296">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurence`.</span></span> | `IPM.Appointment`<br />`IPM.Appointment.Occurence` |
| <span data-ttu-id="8972a-297">邮件项目</span><span class="sxs-lookup"><span data-stu-id="8972a-297">Message items</span></span> | <span data-ttu-id="8972a-298">这些项目包括具有默认邮件类别 `IPM.Note` 的电子邮件，以及使用 `IPM.Schedule.Meeting` 作为基础邮件类别的会议请求、响应和取消。</span><span class="sxs-lookup"><span data-stu-id="8972a-298">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span> | `IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled` |

<span data-ttu-id="8972a-299">你可以创建用于扩展默认邮件类别的自定义邮件类别，例如，自定义约会邮件类别 `IPM.Appointment.Contoso`。</span><span class="sxs-lookup"><span data-stu-id="8972a-299">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="8972a-300">类型：</span><span class="sxs-lookup"><span data-stu-id="8972a-300">Type:</span></span>

*   <span data-ttu-id="8972a-301">String</span><span class="sxs-lookup"><span data-stu-id="8972a-301">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="8972a-302">要求</span><span class="sxs-lookup"><span data-stu-id="8972a-302">Requirements</span></span>

|<span data-ttu-id="8972a-303">要求</span><span class="sxs-lookup"><span data-stu-id="8972a-303">Requirement</span></span>| <span data-ttu-id="8972a-304">值</span><span class="sxs-lookup"><span data-stu-id="8972a-304">Value</span></span>|
|---|---|
|[<span data-ttu-id="8972a-305">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="8972a-305">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8972a-306">1.0</span><span class="sxs-lookup"><span data-stu-id="8972a-306">1.0</span></span>|
|[<span data-ttu-id="8972a-307">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="8972a-307">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8972a-308">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8972a-308">ReadItem</span></span>|
|[<span data-ttu-id="8972a-309">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="8972a-309">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="8972a-310">阅读</span><span class="sxs-lookup"><span data-stu-id="8972a-310">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="8972a-311">示例</span><span class="sxs-lookup"><span data-stu-id="8972a-311">Example</span></span>

```JavaScript
var itemClass = Office.context.mailbox.item.itemClass;
```

#### <a name="nullable-itemid-string"></a><span data-ttu-id="8972a-312">(nullable) itemId :String</span><span class="sxs-lookup"><span data-stu-id="8972a-312">(nullable) itemId :String</span></span>

<span data-ttu-id="8972a-p118">获取当前项目的 Exchange Web 服务项目标识符。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="8972a-p118">Gets the Exchange Web Services item identifier for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="8972a-315">`itemId` 属性返回的标识符与 Exchange Web 服务项目标识符相同。</span><span class="sxs-lookup"><span data-stu-id="8972a-315">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="8972a-316">`itemId` 属性与 Outlook 条目 ID 或 Outlook REST API 使用的 ID 不同。</span><span class="sxs-lookup"><span data-stu-id="8972a-316">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="8972a-317">使用此值进行 REST API 调用前，应使用 `Office.context.mailbox.convertToRestId`（可在要求集 1.3 的开头部分中找到）对它进行转换。</span><span class="sxs-lookup"><span data-stu-id="8972a-317">Before making REST API calls using this value, it should be converted using `Office.context.mailbox.convertToRestId`, which is available starting in requirement set 1.3.</span></span> <span data-ttu-id="8972a-318">有关详细信息，请参阅[从 Outlook 加载项使用 Outlook REST API](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id)。</span><span class="sxs-lookup"><span data-stu-id="8972a-318">For more details, see [Use the Outlook REST APIs from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

##### <a name="type"></a><span data-ttu-id="8972a-319">类型：</span><span class="sxs-lookup"><span data-stu-id="8972a-319">Type:</span></span>

*   <span data-ttu-id="8972a-320">String</span><span class="sxs-lookup"><span data-stu-id="8972a-320">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="8972a-321">要求</span><span class="sxs-lookup"><span data-stu-id="8972a-321">Requirements</span></span>

|<span data-ttu-id="8972a-322">要求</span><span class="sxs-lookup"><span data-stu-id="8972a-322">Requirement</span></span>| <span data-ttu-id="8972a-323">值</span><span class="sxs-lookup"><span data-stu-id="8972a-323">Value</span></span>|
|---|---|
|[<span data-ttu-id="8972a-324">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="8972a-324">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8972a-325">1.0</span><span class="sxs-lookup"><span data-stu-id="8972a-325">1.0</span></span>|
|[<span data-ttu-id="8972a-326">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="8972a-326">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8972a-327">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8972a-327">ReadItem</span></span>|
|[<span data-ttu-id="8972a-328">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="8972a-328">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="8972a-329">阅读</span><span class="sxs-lookup"><span data-stu-id="8972a-329">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="8972a-330">示例</span><span class="sxs-lookup"><span data-stu-id="8972a-330">Example</span></span>

<span data-ttu-id="8972a-p120">以下代码检查项目标识符是否存在。如果 `itemId` 属性返回 `null` 或 `undefined`，则将项目保存到存储，并从异步结果中获取项目标识符。</span><span class="sxs-lookup"><span data-stu-id="8972a-p120">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

```JavaScript
var itemId = Office.context.mailbox.item.itemId;
if (itemId === null || itemId == undefined) {
  Office.context.mailbox.item.saveAsync(function(result){
    itemId = result.value;
  });
}
```

####  <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlook11officemailboxenumsitemtype"></a><span data-ttu-id="8972a-333">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_1/office.mailboxenums.itemtype)</span><span class="sxs-lookup"><span data-stu-id="8972a-333">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_1/office.mailboxenums.itemtype)</span></span>

<span data-ttu-id="8972a-334">获取实例表示的项的类型。</span><span class="sxs-lookup"><span data-stu-id="8972a-334">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="8972a-335">`itemType` 属性返回其中一个 `ItemType` 枚举值，指示 `item` 对象实例是邮件还是约会。</span><span class="sxs-lookup"><span data-stu-id="8972a-335">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="8972a-336">类型：</span><span class="sxs-lookup"><span data-stu-id="8972a-336">Type:</span></span>

*   [<span data-ttu-id="8972a-337">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="8972a-337">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook_1_1/office.mailboxenums.itemtype)

##### <a name="requirements"></a><span data-ttu-id="8972a-338">要求</span><span class="sxs-lookup"><span data-stu-id="8972a-338">Requirements</span></span>

|<span data-ttu-id="8972a-339">要求</span><span class="sxs-lookup"><span data-stu-id="8972a-339">Requirement</span></span>| <span data-ttu-id="8972a-340">值</span><span class="sxs-lookup"><span data-stu-id="8972a-340">Value</span></span>|
|---|---|
|[<span data-ttu-id="8972a-341">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="8972a-341">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8972a-342">1.0</span><span class="sxs-lookup"><span data-stu-id="8972a-342">1.0</span></span>|
|[<span data-ttu-id="8972a-343">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="8972a-343">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8972a-344">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8972a-344">ReadItem</span></span>|
|[<span data-ttu-id="8972a-345">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="8972a-345">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="8972a-346">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="8972a-346">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="8972a-347">示例</span><span class="sxs-lookup"><span data-stu-id="8972a-347">Example</span></span>

```JavaScript
if (Office.context.mailbox.item.itemType == Office.MailboxEnums.ItemType.Message)
  // do something
else
  // do something else
```

####  <a name="location-stringlocationjavascriptapioutlook11officelocation"></a><span data-ttu-id="8972a-348">location :String|[Location](/javascript/api/outlook_1_1/office.location)</span><span class="sxs-lookup"><span data-stu-id="8972a-348">location :String|[Location](/javascript/api/outlook_1_1/office.location)</span></span>

<span data-ttu-id="8972a-349">获取或设置约会的位置。</span><span class="sxs-lookup"><span data-stu-id="8972a-349">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="8972a-350">阅读模式</span><span class="sxs-lookup"><span data-stu-id="8972a-350">Read mode</span></span>

<span data-ttu-id="8972a-351">`location` 属性返回一个包含约会位置的字符串。</span><span class="sxs-lookup"><span data-stu-id="8972a-351">The `location` property returns a string that contains the location of the appointment.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="8972a-352">撰写模式</span><span class="sxs-lookup"><span data-stu-id="8972a-352">Compose mode</span></span>

<span data-ttu-id="8972a-353">`location` 属性返回一个 `Location` 对象，该对象提供用于获取和设置约会位置的方法。</span><span class="sxs-lookup"><span data-stu-id="8972a-353">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="8972a-354">类型：</span><span class="sxs-lookup"><span data-stu-id="8972a-354">Type:</span></span>

*   <span data-ttu-id="8972a-355">String | [Location](/javascript/api/outlook_1_1/office.location)</span><span class="sxs-lookup"><span data-stu-id="8972a-355">String | [Location](/javascript/api/outlook_1_1/office.location)</span></span>

##### <a name="requirements"></a><span data-ttu-id="8972a-356">要求</span><span class="sxs-lookup"><span data-stu-id="8972a-356">Requirements</span></span>

|<span data-ttu-id="8972a-357">要求</span><span class="sxs-lookup"><span data-stu-id="8972a-357">Requirement</span></span>| <span data-ttu-id="8972a-358">值</span><span class="sxs-lookup"><span data-stu-id="8972a-358">Value</span></span>|
|---|---|
|[<span data-ttu-id="8972a-359">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="8972a-359">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8972a-360">1.0</span><span class="sxs-lookup"><span data-stu-id="8972a-360">1.0</span></span>|
|[<span data-ttu-id="8972a-361">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="8972a-361">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8972a-362">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8972a-362">ReadItem</span></span>|
|[<span data-ttu-id="8972a-363">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="8972a-363">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="8972a-364">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="8972a-364">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="8972a-365">示例</span><span class="sxs-lookup"><span data-stu-id="8972a-365">Example</span></span>

```JavaScript
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

#### <a name="normalizedsubject-string"></a><span data-ttu-id="8972a-366">normalizedSubject :String</span><span class="sxs-lookup"><span data-stu-id="8972a-366">normalizedSubject :String</span></span>

<span data-ttu-id="8972a-p121">获取删除了所有前缀（包括 `RE:` 和 `FWD:`）的项目主题。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="8972a-p121">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="8972a-p122">normalizedSubject 属性获取包含由电子邮件程序添加的任何标准前缀（如 `RE:` 和 `FW:`）的项目主题。若要获取包含完整前缀的项目主题，请使用 [`subject`](#subject-stringsubjectjavascriptapioutlook11officesubject) 属性。</span><span class="sxs-lookup"><span data-stu-id="8972a-p122">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubjectjavascriptapioutlook11officesubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="8972a-371">类型:</span><span class="sxs-lookup"><span data-stu-id="8972a-371">Type:</span></span>

*   <span data-ttu-id="8972a-372">String</span><span class="sxs-lookup"><span data-stu-id="8972a-372">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="8972a-373">要求</span><span class="sxs-lookup"><span data-stu-id="8972a-373">Requirements</span></span>

|<span data-ttu-id="8972a-374">要求</span><span class="sxs-lookup"><span data-stu-id="8972a-374">Requirement</span></span>| <span data-ttu-id="8972a-375">值</span><span class="sxs-lookup"><span data-stu-id="8972a-375">Value</span></span>|
|---|---|
|[<span data-ttu-id="8972a-376">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="8972a-376">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8972a-377">1.0</span><span class="sxs-lookup"><span data-stu-id="8972a-377">1.0</span></span>|
|[<span data-ttu-id="8972a-378">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="8972a-378">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8972a-379">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8972a-379">ReadItem</span></span>|
|[<span data-ttu-id="8972a-380">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="8972a-380">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="8972a-381">阅读</span><span class="sxs-lookup"><span data-stu-id="8972a-381">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="8972a-382">示例</span><span class="sxs-lookup"><span data-stu-id="8972a-382">Example</span></span>

```JavaScript
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
```

####  <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlook11officeemailaddressdetailsrecipientsjavascriptapioutlook11officerecipients"></a><span data-ttu-id="8972a-383">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_1/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="8972a-383">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_1/office.recipients)</span></span>

<span data-ttu-id="8972a-384">提供对事件的可选与会者的访问权限。</span><span class="sxs-lookup"><span data-stu-id="8972a-384">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="8972a-385">对象的类型和访问级别取决于当前项目的模式。</span><span class="sxs-lookup"><span data-stu-id="8972a-385">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="8972a-386">阅读模式</span><span class="sxs-lookup"><span data-stu-id="8972a-386">Read mode</span></span>

<span data-ttu-id="8972a-387">`optionalAttendees` 属性返回一个数组，其中包含每个可选与会者的 `EmailAddressDetails` 对象。</span><span class="sxs-lookup"><span data-stu-id="8972a-387">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="8972a-388">撰写模式</span><span class="sxs-lookup"><span data-stu-id="8972a-388">Compose mode</span></span>

<span data-ttu-id="8972a-389">`optionalAttendees` 属性返回一个 `Recipients` 对象，该对象提供用于获取或更新可选与会者的方法。</span><span class="sxs-lookup"><span data-stu-id="8972a-389">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="8972a-390">类型：</span><span class="sxs-lookup"><span data-stu-id="8972a-390">Type:</span></span>

*   <span data-ttu-id="8972a-391">Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_1/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="8972a-391">Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_1/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="8972a-392">要求</span><span class="sxs-lookup"><span data-stu-id="8972a-392">Requirements</span></span>

|<span data-ttu-id="8972a-393">要求</span><span class="sxs-lookup"><span data-stu-id="8972a-393">Requirement</span></span>| <span data-ttu-id="8972a-394">值</span><span class="sxs-lookup"><span data-stu-id="8972a-394">Value</span></span>|
|---|---|
|[<span data-ttu-id="8972a-395">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="8972a-395">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8972a-396">1.0</span><span class="sxs-lookup"><span data-stu-id="8972a-396">1.0</span></span>|
|[<span data-ttu-id="8972a-397">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="8972a-397">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8972a-398">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8972a-398">ReadItem</span></span>|
|[<span data-ttu-id="8972a-399">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="8972a-399">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="8972a-400">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="8972a-400">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="8972a-401">示例</span><span class="sxs-lookup"><span data-stu-id="8972a-401">Example</span></span>

```JavaScript
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

#### <a name="organizer-emailaddressdetailsjavascriptapioutlook11officeemailaddressdetails"></a><span data-ttu-id="8972a-402">organizer :[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="8972a-402">organizer :[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)</span></span>

<span data-ttu-id="8972a-p124">获取指定会议的会议组织者的电子邮件地址。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="8972a-p124">Gets the email address of the meeting organizer for a specified meeting. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="8972a-405">类型:</span><span class="sxs-lookup"><span data-stu-id="8972a-405">Type:</span></span>

*   [<span data-ttu-id="8972a-406">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="8972a-406">EmailAddressDetails</span></span>](/javascript/api/outlook_1_1/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="8972a-407">要求</span><span class="sxs-lookup"><span data-stu-id="8972a-407">Requirements</span></span>

|<span data-ttu-id="8972a-408">要求</span><span class="sxs-lookup"><span data-stu-id="8972a-408">Requirement</span></span>| <span data-ttu-id="8972a-409">值</span><span class="sxs-lookup"><span data-stu-id="8972a-409">Value</span></span>|
|---|---|
|[<span data-ttu-id="8972a-410">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="8972a-410">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8972a-411">1.0</span><span class="sxs-lookup"><span data-stu-id="8972a-411">1.0</span></span>|
|[<span data-ttu-id="8972a-412">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="8972a-412">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8972a-413">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8972a-413">ReadItem</span></span>|
|[<span data-ttu-id="8972a-414">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="8972a-414">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="8972a-415">阅读</span><span class="sxs-lookup"><span data-stu-id="8972a-415">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="8972a-416">示例</span><span class="sxs-lookup"><span data-stu-id="8972a-416">Example</span></span>

```JavaScript
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
```

####  <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlook11officeemailaddressdetailsrecipientsjavascriptapioutlook11officerecipients"></a><span data-ttu-id="8972a-417">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_1/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="8972a-417">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_1/office.recipients)</span></span>

<span data-ttu-id="8972a-418">提供对事件的必需与会者的访问权限。</span><span class="sxs-lookup"><span data-stu-id="8972a-418">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="8972a-419">对象的类型和访问级别取决于当前项目的模式。</span><span class="sxs-lookup"><span data-stu-id="8972a-419">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="8972a-420">阅读模式</span><span class="sxs-lookup"><span data-stu-id="8972a-420">Read mode</span></span>

<span data-ttu-id="8972a-421">`requiredAttendees` 属性返回一个数组，其中包含每个必需与会者的 `EmailAddressDetails` 对象。</span><span class="sxs-lookup"><span data-stu-id="8972a-421">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="8972a-422">撰写模式</span><span class="sxs-lookup"><span data-stu-id="8972a-422">Compose mode</span></span>

<span data-ttu-id="8972a-423">`requiredAttendees` 属性返回一个 `Recipients` 对象，该对象提供用于获取或更新必需与会者的方法。</span><span class="sxs-lookup"><span data-stu-id="8972a-423">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="8972a-424">类型：</span><span class="sxs-lookup"><span data-stu-id="8972a-424">Type:</span></span>

*   <span data-ttu-id="8972a-425">Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_1/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="8972a-425">Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_1/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="8972a-426">要求</span><span class="sxs-lookup"><span data-stu-id="8972a-426">Requirements</span></span>

|<span data-ttu-id="8972a-427">要求</span><span class="sxs-lookup"><span data-stu-id="8972a-427">Requirement</span></span>| <span data-ttu-id="8972a-428">值</span><span class="sxs-lookup"><span data-stu-id="8972a-428">Value</span></span>|
|---|---|
|[<span data-ttu-id="8972a-429">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="8972a-429">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8972a-430">1.0</span><span class="sxs-lookup"><span data-stu-id="8972a-430">1.0</span></span>|
|[<span data-ttu-id="8972a-431">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="8972a-431">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8972a-432">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8972a-432">ReadItem</span></span>|
|[<span data-ttu-id="8972a-433">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="8972a-433">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="8972a-434">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="8972a-434">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="8972a-435">示例</span><span class="sxs-lookup"><span data-stu-id="8972a-435">Example</span></span>

```JavaScript
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
}
```

#### <a name="sender-emailaddressdetailsjavascriptapioutlook11officeemailaddressdetails"></a><span data-ttu-id="8972a-436">sender :[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="8972a-436">sender :[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)</span></span>

<span data-ttu-id="8972a-p126">获取电子邮件发件人的电子邮件地址。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="8972a-p126">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="8972a-p127">[`from`](#from-emailaddressdetailsjavascriptapioutlook11officeemailaddressdetails) 和 `sender` 属性表示同一个人，邮件由代理人发送的除外。在此情况下，`from` 属性表示代理程序，而 sender 属性表示代理人。</span><span class="sxs-lookup"><span data-stu-id="8972a-p127">The [`from`](#from-emailaddressdetailsjavascriptapioutlook11officeemailaddressdetails) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="8972a-441">`from` 属性中 `EmailAddressDetails` 对象的 `recipientType` 属性为 `undefined`。</span><span class="sxs-lookup"><span data-stu-id="8972a-441">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="8972a-442">类型：</span><span class="sxs-lookup"><span data-stu-id="8972a-442">Type:</span></span>

*   [<span data-ttu-id="8972a-443">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="8972a-443">EmailAddressDetails</span></span>](/javascript/api/outlook_1_1/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="8972a-444">要求</span><span class="sxs-lookup"><span data-stu-id="8972a-444">Requirements</span></span>

|<span data-ttu-id="8972a-445">要求</span><span class="sxs-lookup"><span data-stu-id="8972a-445">Requirement</span></span>| <span data-ttu-id="8972a-446">值</span><span class="sxs-lookup"><span data-stu-id="8972a-446">Value</span></span>|
|---|---|
|[<span data-ttu-id="8972a-447">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="8972a-447">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8972a-448">1.0</span><span class="sxs-lookup"><span data-stu-id="8972a-448">1.0</span></span>|
|[<span data-ttu-id="8972a-449">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="8972a-449">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8972a-450">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8972a-450">ReadItem</span></span>|
|[<span data-ttu-id="8972a-451">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="8972a-451">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="8972a-452">阅读</span><span class="sxs-lookup"><span data-stu-id="8972a-452">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="8972a-453">示例</span><span class="sxs-lookup"><span data-stu-id="8972a-453">Example</span></span>

```JavaScript
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
```

####  <a name="start-datetimejavascriptapioutlook11officetime"></a><span data-ttu-id="8972a-454">start :Date|[Time](/javascript/api/outlook_1_1/office.time)</span><span class="sxs-lookup"><span data-stu-id="8972a-454">start :Date|[Time](/javascript/api/outlook_1_1/office.time)</span></span>

<span data-ttu-id="8972a-455">获取或设置约会开始的日期和时间。</span><span class="sxs-lookup"><span data-stu-id="8972a-455">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="8972a-p128">将 `start` 属性表示为协调世界时 (UTC) 的日期和时间值。可使用 [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook11officelocalclienttime) 方法将该值转换为客户端的本地日期和时间。</span><span class="sxs-lookup"><span data-stu-id="8972a-p128">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook11officelocalclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="8972a-458">阅读模式</span><span class="sxs-lookup"><span data-stu-id="8972a-458">Read mode</span></span>

<span data-ttu-id="8972a-459">`start` 属性返回 `Date` 对象。</span><span class="sxs-lookup"><span data-stu-id="8972a-459">The `start` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="8972a-460">撰写模式</span><span class="sxs-lookup"><span data-stu-id="8972a-460">Compose mode</span></span>

<span data-ttu-id="8972a-461">`start` 属性返回 `Time` 对象。</span><span class="sxs-lookup"><span data-stu-id="8972a-461">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="8972a-462">使用 [`Time.setAsync`](/javascript/api/outlook_1_1/office.time#setasync-datetime--options--callback-) 方法设置开始时间时，应使用 [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) 方法将客户端的本地时间转换为服务器的 UTC。</span><span class="sxs-lookup"><span data-stu-id="8972a-462">When you use the [`Time.setAsync`](/javascript/api/outlook_1_1/office.time#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="8972a-463">类型：</span><span class="sxs-lookup"><span data-stu-id="8972a-463">Type:</span></span>

*   <span data-ttu-id="8972a-464">Date | [Time](/javascript/api/outlook_1_1/office.time)</span><span class="sxs-lookup"><span data-stu-id="8972a-464">Date | [Time](/javascript/api/outlook_1_1/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="8972a-465">要求</span><span class="sxs-lookup"><span data-stu-id="8972a-465">Requirements</span></span>

|<span data-ttu-id="8972a-466">要求</span><span class="sxs-lookup"><span data-stu-id="8972a-466">Requirement</span></span>| <span data-ttu-id="8972a-467">值</span><span class="sxs-lookup"><span data-stu-id="8972a-467">Value</span></span>|
|---|---|
|[<span data-ttu-id="8972a-468">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="8972a-468">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8972a-469">1.0</span><span class="sxs-lookup"><span data-stu-id="8972a-469">1.0</span></span>|
|[<span data-ttu-id="8972a-470">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="8972a-470">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8972a-471">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8972a-471">ReadItem</span></span>|
|[<span data-ttu-id="8972a-472">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="8972a-472">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="8972a-473">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="8972a-473">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="8972a-474">示例</span><span class="sxs-lookup"><span data-stu-id="8972a-474">Example</span></span>

<span data-ttu-id="8972a-475">以下示例通过使用 `Time` 对象的 [`setAsync`](/javascript/api/outlook_1_1/office.time#setasync-datetime--options--callback-) 方法，设置撰写模式下约会的开始时间。</span><span class="sxs-lookup"><span data-stu-id="8972a-475">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_1/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

####  <a name="subject-stringsubjectjavascriptapioutlook11officesubject"></a><span data-ttu-id="8972a-476">subject :String|[Subject](/javascript/api/outlook_1_1/office.subject)</span><span class="sxs-lookup"><span data-stu-id="8972a-476">subject :String|[Subject](/javascript/api/outlook_1_1/office.subject)</span></span>

<span data-ttu-id="8972a-477">获取或设置显示在项目的主题字段中的说明。</span><span class="sxs-lookup"><span data-stu-id="8972a-477">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="8972a-478">`subject` 属性获取或设置由电子邮件服务器发送项目时的整个主题。</span><span class="sxs-lookup"><span data-stu-id="8972a-478">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="8972a-479">阅读模式</span><span class="sxs-lookup"><span data-stu-id="8972a-479">Read mode</span></span>

<span data-ttu-id="8972a-p129">`subject` 属性返回一个字符串。使用 [`normalizedSubject`](#normalizedsubject-string) 属性获取不带任何前导前缀（如 `RE:` 和 `FW:`）的主题。</span><span class="sxs-lookup"><span data-stu-id="8972a-p129">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

```js
var subject = Office.context.mailbox.item.subject;
```

##### <a name="compose-mode"></a><span data-ttu-id="8972a-482">撰写模式</span><span class="sxs-lookup"><span data-stu-id="8972a-482">Compose mode</span></span>

<span data-ttu-id="8972a-483">`subject` 属性返回一个 `Subject` 对象，该对象提供用于获取和设置主题的方法。</span><span class="sxs-lookup"><span data-stu-id="8972a-483">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```JavaScript
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="8972a-484">类型：</span><span class="sxs-lookup"><span data-stu-id="8972a-484">Type:</span></span>

*   <span data-ttu-id="8972a-485">String | [Subject](/javascript/api/outlook_1_1/office.subject)</span><span class="sxs-lookup"><span data-stu-id="8972a-485">String | [Subject](/javascript/api/outlook_1_1/office.subject)</span></span>

##### <a name="requirements"></a><span data-ttu-id="8972a-486">要求</span><span class="sxs-lookup"><span data-stu-id="8972a-486">Requirements</span></span>

|<span data-ttu-id="8972a-487">要求</span><span class="sxs-lookup"><span data-stu-id="8972a-487">Requirement</span></span>| <span data-ttu-id="8972a-488">值</span><span class="sxs-lookup"><span data-stu-id="8972a-488">Value</span></span>|
|---|---|
|[<span data-ttu-id="8972a-489">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="8972a-489">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8972a-490">1.0</span><span class="sxs-lookup"><span data-stu-id="8972a-490">1.0</span></span>|
|[<span data-ttu-id="8972a-491">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="8972a-491">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8972a-492">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8972a-492">ReadItem</span></span>|
|[<span data-ttu-id="8972a-493">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="8972a-493">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="8972a-494">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="8972a-494">Compose or read</span></span>|

####  <a name="to-arrayemailaddressdetailsjavascriptapioutlook11officeemailaddressdetailsrecipientsjavascriptapioutlook11officerecipients"></a><span data-ttu-id="8972a-495">to :Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_1/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="8972a-495">to :Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_1/office.recipients)</span></span>

<span data-ttu-id="8972a-496">提供对邮件的“**收件人**”行上的收件人的访问权限。</span><span class="sxs-lookup"><span data-stu-id="8972a-496">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="8972a-497">对象的类型和访问级别取决于当前项目的模式。</span><span class="sxs-lookup"><span data-stu-id="8972a-497">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="8972a-498">阅读模式</span><span class="sxs-lookup"><span data-stu-id="8972a-498">Read mode</span></span>

<span data-ttu-id="8972a-p131">`to` 属性返回包含邮件的**收件人**行上所列的每个收件人的 `EmailAddressDetails` 对象的数组。集合上限为 100 个成员。</span><span class="sxs-lookup"><span data-stu-id="8972a-p131">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="8972a-501">撰写模式</span><span class="sxs-lookup"><span data-stu-id="8972a-501">Compose mode</span></span>

<span data-ttu-id="8972a-502">`to` 属性返回一个 `Recipients` 对象，该对象提供用于获取或更新邮件的“**收件人**”行上收件人的方法。</span><span class="sxs-lookup"><span data-stu-id="8972a-502">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="8972a-503">类型：</span><span class="sxs-lookup"><span data-stu-id="8972a-503">Type:</span></span>

*   <span data-ttu-id="8972a-504">Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_1/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="8972a-504">Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_1/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="8972a-505">要求</span><span class="sxs-lookup"><span data-stu-id="8972a-505">Requirements</span></span>

|<span data-ttu-id="8972a-506">要求</span><span class="sxs-lookup"><span data-stu-id="8972a-506">Requirement</span></span>| <span data-ttu-id="8972a-507">值</span><span class="sxs-lookup"><span data-stu-id="8972a-507">Value</span></span>|
|---|---|
|[<span data-ttu-id="8972a-508">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="8972a-508">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8972a-509">1.0</span><span class="sxs-lookup"><span data-stu-id="8972a-509">1.0</span></span>|
|[<span data-ttu-id="8972a-510">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="8972a-510">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8972a-511">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8972a-511">ReadItem</span></span>|
|[<span data-ttu-id="8972a-512">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="8972a-512">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="8972a-513">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="8972a-513">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="8972a-514">示例</span><span class="sxs-lookup"><span data-stu-id="8972a-514">Example</span></span>

```JavaScript
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

### <a name="methods"></a><span data-ttu-id="8972a-515">方法</span><span class="sxs-lookup"><span data-stu-id="8972a-515">Methods</span></span>

####  <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="8972a-516">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="8972a-516">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="8972a-517">将文件作为附件添加到邮件或约会。</span><span class="sxs-lookup"><span data-stu-id="8972a-517">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="8972a-518">`addFileAttachmentAsync` 方法在指定的 URI 上载文件并将其附加到撰写窗体中的项目。</span><span class="sxs-lookup"><span data-stu-id="8972a-518">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="8972a-519">你随后可以将该标识符与 [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) 方法一同使用，以删除同一个会话中的附件。</span><span class="sxs-lookup"><span data-stu-id="8972a-519">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="8972a-520">参数：</span><span class="sxs-lookup"><span data-stu-id="8972a-520">Parameters:</span></span>

|<span data-ttu-id="8972a-521">名称</span><span class="sxs-lookup"><span data-stu-id="8972a-521">Name</span></span>| <span data-ttu-id="8972a-522">类型</span><span class="sxs-lookup"><span data-stu-id="8972a-522">Type</span></span>| <span data-ttu-id="8972a-523">属性</span><span class="sxs-lookup"><span data-stu-id="8972a-523">Attributes</span></span>| <span data-ttu-id="8972a-524">说明</span><span class="sxs-lookup"><span data-stu-id="8972a-524">Description</span></span>|
|---|---|---|---|
|`uri`| <span data-ttu-id="8972a-525">字符串</span><span class="sxs-lookup"><span data-stu-id="8972a-525">String</span></span>||<span data-ttu-id="8972a-p132">提供附加到邮件或约会的文件的位置的 URI。最大长度为 2048 个字符。</span><span class="sxs-lookup"><span data-stu-id="8972a-p132">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="8972a-528">字符串</span><span class="sxs-lookup"><span data-stu-id="8972a-528">String</span></span>||<span data-ttu-id="8972a-p133">在附件上载过程中显示的附件名称。最大长度为 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="8972a-p133">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="8972a-531">对象</span><span class="sxs-lookup"><span data-stu-id="8972a-531">Object</span></span>| <span data-ttu-id="8972a-532">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="8972a-532">&lt;optional&gt;</span></span>|<span data-ttu-id="8972a-533">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="8972a-533">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="8972a-534">对象</span><span class="sxs-lookup"><span data-stu-id="8972a-534">Object</span></span>| <span data-ttu-id="8972a-535">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="8972a-535">&lt;optional&gt;</span></span>|<span data-ttu-id="8972a-536">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="8972a-536">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="8972a-537">函数</span><span class="sxs-lookup"><span data-stu-id="8972a-537">function</span></span>| <span data-ttu-id="8972a-538">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="8972a-538">&lt;optional&gt;</span></span>|<span data-ttu-id="8972a-539">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="8972a-539">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="8972a-540">如果成功，附件标识符将在 `asyncResult.value` 属性中提供。</span><span class="sxs-lookup"><span data-stu-id="8972a-540">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="8972a-541">如果上传附件失败，`asyncResult` 对象将包含一个提供错误说明的 `Error` 对象。</span><span class="sxs-lookup"><span data-stu-id="8972a-541">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="8972a-542">错误</span><span class="sxs-lookup"><span data-stu-id="8972a-542">Errors</span></span>

| <span data-ttu-id="8972a-543">错误代码</span><span class="sxs-lookup"><span data-stu-id="8972a-543">Error code</span></span> | <span data-ttu-id="8972a-544">说明</span><span class="sxs-lookup"><span data-stu-id="8972a-544">Description</span></span> |
|------------|-------------|
| `AttachmentSizeExceeded` | <span data-ttu-id="8972a-545">附件大小超过了允许的大小。</span><span class="sxs-lookup"><span data-stu-id="8972a-545">The attachment is larger than allowed.</span></span> |
| `FileTypeNotSupported` | <span data-ttu-id="8972a-546">该附件的扩展名不是允许的扩展名。</span><span class="sxs-lookup"><span data-stu-id="8972a-546">The attachment has an extension that is not allowed.</span></span> |
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="8972a-547">邮件或约会具有的附件过多。</span><span class="sxs-lookup"><span data-stu-id="8972a-547">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="8972a-548">要求</span><span class="sxs-lookup"><span data-stu-id="8972a-548">Requirements</span></span>

|<span data-ttu-id="8972a-549">要求</span><span class="sxs-lookup"><span data-stu-id="8972a-549">Requirement</span></span>| <span data-ttu-id="8972a-550">值</span><span class="sxs-lookup"><span data-stu-id="8972a-550">Value</span></span>|
|---|---|
|[<span data-ttu-id="8972a-551">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="8972a-551">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8972a-552">1.1</span><span class="sxs-lookup"><span data-stu-id="8972a-552">1.1</span></span>|
|[<span data-ttu-id="8972a-553">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="8972a-553">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8972a-554">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="8972a-554">ReadWriteItem</span></span>|
|[<span data-ttu-id="8972a-555">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="8972a-555">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="8972a-556">撰写</span><span class="sxs-lookup"><span data-stu-id="8972a-556">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="8972a-557">示例</span><span class="sxs-lookup"><span data-stu-id="8972a-557">Example</span></span>

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

####  <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="8972a-558">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="8972a-558">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="8972a-559">将 Exchange 项目（如邮件）作为附件添加到邮件或约会。</span><span class="sxs-lookup"><span data-stu-id="8972a-559">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="8972a-p134">`addItemAttachmentAsync` 方法将包含指定 Exchange 标识符的项目附加到撰写窗体中的项目。如果指定一个回调方法，此方法使用 `asyncResult` 参数调用，该参数包含一个附件标识符或代码，指示附加项目过程中出现的任何错误。可以使用 `options` 参数将状态信息传递给回调方法（如果需要）。</span><span class="sxs-lookup"><span data-stu-id="8972a-p134">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="8972a-563">随后可以将该标识符与 [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) 方法一同使用，以删除同一个会话中的附件。</span><span class="sxs-lookup"><span data-stu-id="8972a-563">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="8972a-564">如果 Office 加载项在 Outlook Web App 中运行，则 `addItemAttachmentAsync` 方法可以将项目附加到项目（正在编辑的项目除外）中；然而，不支持也不建议这样做。</span><span class="sxs-lookup"><span data-stu-id="8972a-564">If your Office Add-in is running in Outlook Web App, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="8972a-565">参数：</span><span class="sxs-lookup"><span data-stu-id="8972a-565">Parameters:</span></span>

|<span data-ttu-id="8972a-566">名称</span><span class="sxs-lookup"><span data-stu-id="8972a-566">Name</span></span>| <span data-ttu-id="8972a-567">类型</span><span class="sxs-lookup"><span data-stu-id="8972a-567">Type</span></span>| <span data-ttu-id="8972a-568">属性</span><span class="sxs-lookup"><span data-stu-id="8972a-568">Attributes</span></span>| <span data-ttu-id="8972a-569">说明</span><span class="sxs-lookup"><span data-stu-id="8972a-569">Description</span></span>|
|---|---|---|---|
|`itemId`| <span data-ttu-id="8972a-570">字符串</span><span class="sxs-lookup"><span data-stu-id="8972a-570">String</span></span>||<span data-ttu-id="8972a-p135">要附加的项目的 Exchange 标识符。最大长度为 100 个字符。</span><span class="sxs-lookup"><span data-stu-id="8972a-p135">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="8972a-573">字符串</span><span class="sxs-lookup"><span data-stu-id="8972a-573">String</span></span>||<span data-ttu-id="8972a-p136">要附加的项目的主题。最大长度为 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="8972a-p136">The sujbect of the item to be attached. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="8972a-576">对象</span><span class="sxs-lookup"><span data-stu-id="8972a-576">Object</span></span>| <span data-ttu-id="8972a-577">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="8972a-577">&lt;optional&gt;</span></span>|<span data-ttu-id="8972a-578">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="8972a-578">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="8972a-579">对象</span><span class="sxs-lookup"><span data-stu-id="8972a-579">Object</span></span>| <span data-ttu-id="8972a-580">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="8972a-580">&lt;optional&gt;</span></span>|<span data-ttu-id="8972a-581">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="8972a-581">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="8972a-582">函数</span><span class="sxs-lookup"><span data-stu-id="8972a-582">function</span></span>| <span data-ttu-id="8972a-583">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="8972a-583">&lt;optional&gt;</span></span>|<span data-ttu-id="8972a-584">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="8972a-584">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="8972a-585">如果成功，附件标识符将在 `asyncResult.value` 属性中提供。</span><span class="sxs-lookup"><span data-stu-id="8972a-585">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="8972a-586">如果添加附件失败，`asyncResult` 对象将包含一个提供错误说明的 `Error` 对象。</span><span class="sxs-lookup"><span data-stu-id="8972a-586">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="8972a-587">错误</span><span class="sxs-lookup"><span data-stu-id="8972a-587">Errors</span></span>

| <span data-ttu-id="8972a-588">错误代码</span><span class="sxs-lookup"><span data-stu-id="8972a-588">Error code</span></span> | <span data-ttu-id="8972a-589">说明</span><span class="sxs-lookup"><span data-stu-id="8972a-589">Description</span></span> |
|------------|-------------|
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="8972a-590">邮件或约会具有的附件过多。</span><span class="sxs-lookup"><span data-stu-id="8972a-590">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="8972a-591">要求</span><span class="sxs-lookup"><span data-stu-id="8972a-591">Requirements</span></span>

|<span data-ttu-id="8972a-592">要求</span><span class="sxs-lookup"><span data-stu-id="8972a-592">Requirement</span></span>| <span data-ttu-id="8972a-593">值</span><span class="sxs-lookup"><span data-stu-id="8972a-593">Value</span></span>|
|---|---|
|[<span data-ttu-id="8972a-594">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="8972a-594">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8972a-595">1.1</span><span class="sxs-lookup"><span data-stu-id="8972a-595">1.1</span></span>|
|[<span data-ttu-id="8972a-596">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="8972a-596">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8972a-597">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="8972a-597">ReadWriteItem</span></span>|
|[<span data-ttu-id="8972a-598">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="8972a-598">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="8972a-599">撰写</span><span class="sxs-lookup"><span data-stu-id="8972a-599">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="8972a-600">示例</span><span class="sxs-lookup"><span data-stu-id="8972a-600">Example</span></span>

<span data-ttu-id="8972a-601">以下示例将现有的 Outlook 项目添加为名为 `My Attachment` 的附件。</span><span class="sxs-lookup"><span data-stu-id="8972a-601">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

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

#### <a name="displayreplyallformformdata"></a><span data-ttu-id="8972a-602">displayReplyAllForm(formData)</span><span class="sxs-lookup"><span data-stu-id="8972a-602">displayReplyAllForm(formData)</span></span>

<span data-ttu-id="8972a-603">显示答复窗体，其中包括所选邮件的发件人和所有收件人或所选约会的组织者和所有与会者。</span><span class="sxs-lookup"><span data-stu-id="8972a-603">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="8972a-604">在 Outlook for iOS 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="8972a-604">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="8972a-605">在 Outlook Web App 中，答复窗体显示为包含 3 列视图的弹出式窗体以及包含 2 列或 1 列视图的弹出式窗体。</span><span class="sxs-lookup"><span data-stu-id="8972a-605">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="8972a-606">如果任意字符串参数超出其限制，`displayReplyAllForm` 将引发异常。</span><span class="sxs-lookup"><span data-stu-id="8972a-606">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

> [!NOTE]
> <span data-ttu-id="8972a-607">要求集 1.1 不支持 `displayReplyAllForm` 在调用中包括附件的功能。</span><span class="sxs-lookup"><span data-stu-id="8972a-607">The ability to include attachments in the call to `displayReplyAllForm` is not supported in requirement set 1.1.</span></span> <span data-ttu-id="8972a-608">附件支持已添加到要求集 1.2 及以上的 `displayReplyAllForm` 中。</span><span class="sxs-lookup"><span data-stu-id="8972a-608">Attachment support was added to `displayReplyAllForm` in requirement set 1.2 and above.</span></span>

##### <a name="parameters"></a><span data-ttu-id="8972a-609">参数：</span><span class="sxs-lookup"><span data-stu-id="8972a-609">Parameters:</span></span>

|<span data-ttu-id="8972a-610">名称</span><span class="sxs-lookup"><span data-stu-id="8972a-610">Name</span></span>| <span data-ttu-id="8972a-611">类型</span><span class="sxs-lookup"><span data-stu-id="8972a-611">Type</span></span>| <span data-ttu-id="8972a-612">说明</span><span class="sxs-lookup"><span data-stu-id="8972a-612">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="8972a-613">字符串 &#124; 对象</span><span class="sxs-lookup"><span data-stu-id="8972a-613">String &#124; Object</span></span>| |<span data-ttu-id="8972a-p138">一个包含文本和 HTML 且表示答复窗体的正文的字符串。字符串限制为 32 KB。</span><span class="sxs-lookup"><span data-stu-id="8972a-p138">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="8972a-616">**OR**</span><span class="sxs-lookup"><span data-stu-id="8972a-616">**OR**</span></span><br/><span data-ttu-id="8972a-p139">包含正文或附件数据和回调函数的对象。对象定义如下。</span><span class="sxs-lookup"><span data-stu-id="8972a-p139">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="8972a-619">字符串</span><span class="sxs-lookup"><span data-stu-id="8972a-619">String</span></span> | <span data-ttu-id="8972a-620">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="8972a-620">&lt;optional&gt;</span></span> | <span data-ttu-id="8972a-p140">一个包含文本和 HTML 且表示答复窗体的正文的字符串。字符串限制为 32 KB。</span><span class="sxs-lookup"><span data-stu-id="8972a-p140">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `callback` | <span data-ttu-id="8972a-623">函数</span><span class="sxs-lookup"><span data-stu-id="8972a-623">function</span></span> | <span data-ttu-id="8972a-624">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="8972a-624">&lt;optional&gt;</span></span> | <span data-ttu-id="8972a-625">方法完成后，使用单个参数 `asyncResult`（一个 [AsyncResult](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="8972a-625">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="8972a-626">要求</span><span class="sxs-lookup"><span data-stu-id="8972a-626">Requirements</span></span>

|<span data-ttu-id="8972a-627">要求</span><span class="sxs-lookup"><span data-stu-id="8972a-627">Requirement</span></span>| <span data-ttu-id="8972a-628">值</span><span class="sxs-lookup"><span data-stu-id="8972a-628">Value</span></span>|
|---|---|
|[<span data-ttu-id="8972a-629">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="8972a-629">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8972a-630">1.0</span><span class="sxs-lookup"><span data-stu-id="8972a-630">1.0</span></span>|
|[<span data-ttu-id="8972a-631">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="8972a-631">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8972a-632">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8972a-632">ReadItem</span></span>|
|[<span data-ttu-id="8972a-633">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="8972a-633">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="8972a-634">阅读</span><span class="sxs-lookup"><span data-stu-id="8972a-634">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="8972a-635">示例</span><span class="sxs-lookup"><span data-stu-id="8972a-635">Examples</span></span>

<span data-ttu-id="8972a-636">以下代码将一个字符串传递到 `displayReplyAllForm` 函数。</span><span class="sxs-lookup"><span data-stu-id="8972a-636">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="8972a-637">使用空白正文答复。</span><span class="sxs-lookup"><span data-stu-id="8972a-637">Reply with an empty body.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="8972a-638">仅使用正文答复。</span><span class="sxs-lookup"><span data-stu-id="8972a-638">Reply with just a body.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="8972a-639">使用正文和回调答复。</span><span class="sxs-lookup"><span data-stu-id="8972a-639">Reply with a body and a callback.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi',
  'callback' : function(asyncResult)
  {
    console.log(asyncResult.value);
  }
});
```

#### <a name="displayreplyformformdata"></a><span data-ttu-id="8972a-640">displayReplyForm(formData)</span><span class="sxs-lookup"><span data-stu-id="8972a-640">displayReplyForm(formData)</span></span>

<span data-ttu-id="8972a-641">显示答复窗体，其中仅包括所选邮件的发件人或所选约会的组织者。</span><span class="sxs-lookup"><span data-stu-id="8972a-641">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="8972a-642">在 Outlook for iOS 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="8972a-642">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="8972a-643">在 Outlook Web App 中，答复窗体显示为包含 3 列视图的弹出式窗体以及包含 2 列或 1 列视图的弹出式窗体。</span><span class="sxs-lookup"><span data-stu-id="8972a-643">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="8972a-644">如果任意字符串参数超出其限制，`displayReplyForm` 将引发异常。</span><span class="sxs-lookup"><span data-stu-id="8972a-644">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

> [!NOTE]
> <span data-ttu-id="8972a-645">要求集 1.1 不支持 `displayReplyForm` 在调用中包括附件的功能。</span><span class="sxs-lookup"><span data-stu-id="8972a-645">The ability to include attachments in the call to `displayReplyForm` is not supported in requirement set 1.1.</span></span> <span data-ttu-id="8972a-646">附件支持已添加到要求集 1.2 及以上的 `displayReplyForm` 中。</span><span class="sxs-lookup"><span data-stu-id="8972a-646">Attachment support was added to `displayReplyForm` in requirement set 1.2 and above.</span></span>

##### <a name="parameters"></a><span data-ttu-id="8972a-647">参数：</span><span class="sxs-lookup"><span data-stu-id="8972a-647">Parameters:</span></span>

|<span data-ttu-id="8972a-648">名称</span><span class="sxs-lookup"><span data-stu-id="8972a-648">Name</span></span>| <span data-ttu-id="8972a-649">类型</span><span class="sxs-lookup"><span data-stu-id="8972a-649">Type</span></span>| <span data-ttu-id="8972a-650">说明</span><span class="sxs-lookup"><span data-stu-id="8972a-650">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="8972a-651">字符串 &#124; 对象</span><span class="sxs-lookup"><span data-stu-id="8972a-651">String &#124; Object</span></span>| | <span data-ttu-id="8972a-p142">一个包含文本和 HTML 且表示答复窗体的正文的字符串。字符串限制为 32 KB。</span><span class="sxs-lookup"><span data-stu-id="8972a-p142">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="8972a-654">**OR**</span><span class="sxs-lookup"><span data-stu-id="8972a-654">**OR**</span></span><br/><span data-ttu-id="8972a-p143">包含正文或附件数据和回调函数的对象。对象定义如下。</span><span class="sxs-lookup"><span data-stu-id="8972a-p143">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="8972a-657">字符串</span><span class="sxs-lookup"><span data-stu-id="8972a-657">String</span></span> | <span data-ttu-id="8972a-658">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="8972a-658">&lt;optional&gt;</span></span> | <span data-ttu-id="8972a-p144">一个包含文本和 HTML 且表示答复窗体的正文的字符串。字符串限制为 32 KB。</span><span class="sxs-lookup"><span data-stu-id="8972a-p144">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `callback` | <span data-ttu-id="8972a-661">函数</span><span class="sxs-lookup"><span data-stu-id="8972a-661">function</span></span> | <span data-ttu-id="8972a-662">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="8972a-662">&lt;optional&gt;</span></span> | <span data-ttu-id="8972a-663">方法完成后，使用单个参数 `asyncResult`（一个 [AsyncResult](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="8972a-663">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="8972a-664">要求</span><span class="sxs-lookup"><span data-stu-id="8972a-664">Requirements</span></span>

|<span data-ttu-id="8972a-665">要求</span><span class="sxs-lookup"><span data-stu-id="8972a-665">Requirement</span></span>| <span data-ttu-id="8972a-666">值</span><span class="sxs-lookup"><span data-stu-id="8972a-666">Value</span></span>|
|---|---|
|[<span data-ttu-id="8972a-667">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="8972a-667">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8972a-668">1.0</span><span class="sxs-lookup"><span data-stu-id="8972a-668">1.0</span></span>|
|[<span data-ttu-id="8972a-669">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="8972a-669">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8972a-670">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8972a-670">ReadItem</span></span>|
|[<span data-ttu-id="8972a-671">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="8972a-671">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="8972a-672">阅读</span><span class="sxs-lookup"><span data-stu-id="8972a-672">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="8972a-673">示例</span><span class="sxs-lookup"><span data-stu-id="8972a-673">Examples</span></span>

<span data-ttu-id="8972a-674">以下代码将一个字符串传递到 `displayReplyForm` 函数。</span><span class="sxs-lookup"><span data-stu-id="8972a-674">The following code passes a string to the `displayReplyForm` function.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="8972a-675">使用空白正文答复。</span><span class="sxs-lookup"><span data-stu-id="8972a-675">Reply with an empty body.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="8972a-676">仅使用正文答复。</span><span class="sxs-lookup"><span data-stu-id="8972a-676">Reply with just a body.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="8972a-677">使用正文和回调答复。</span><span class="sxs-lookup"><span data-stu-id="8972a-677">Reply with a body and a callback.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi',
  'callback' : function(asyncResult)
  {
    console.log(asyncResult.value);
  }
});
```

#### <a name="getentities--entitiesjavascriptapioutlook11officeentities"></a><span data-ttu-id="8972a-678">getEntities() → {[Entities](/javascript/api/outlook_1_1/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="8972a-678">getEntities() → {[Entities](/javascript/api/outlook_1_1/office.entities)}</span></span>

<span data-ttu-id="8972a-679">获取在所选项目的正文中找到的实体。</span><span class="sxs-lookup"><span data-stu-id="8972a-679">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="8972a-680">在 Outlook for iOS 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="8972a-680">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="8972a-681">要求</span><span class="sxs-lookup"><span data-stu-id="8972a-681">Requirements</span></span>

|<span data-ttu-id="8972a-682">要求</span><span class="sxs-lookup"><span data-stu-id="8972a-682">Requirement</span></span>| <span data-ttu-id="8972a-683">值</span><span class="sxs-lookup"><span data-stu-id="8972a-683">Value</span></span>|
|---|---|
|[<span data-ttu-id="8972a-684">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="8972a-684">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8972a-685">1.0</span><span class="sxs-lookup"><span data-stu-id="8972a-685">1.0</span></span>|
|[<span data-ttu-id="8972a-686">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="8972a-686">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8972a-687">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8972a-687">ReadItem</span></span>|
|[<span data-ttu-id="8972a-688">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="8972a-688">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="8972a-689">阅读</span><span class="sxs-lookup"><span data-stu-id="8972a-689">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="8972a-690">返回：</span><span class="sxs-lookup"><span data-stu-id="8972a-690">Returns:</span></span>

<span data-ttu-id="8972a-691">类型：[Entities](/javascript/api/outlook_1_1/office.entities)</span><span class="sxs-lookup"><span data-stu-id="8972a-691">Type: [Entities](/javascript/api/outlook_1_1/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="8972a-692">示例</span><span class="sxs-lookup"><span data-stu-id="8972a-692">Example</span></span>

<span data-ttu-id="8972a-693">以下示例访问当前项目的正文中的联系人实体。</span><span class="sxs-lookup"><span data-stu-id="8972a-693">The following example accesses the contacts entities in the current item's body.</span></span>

```JavaScript
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlook11officecontactmeetingsuggestionjavascriptapioutlook11officemeetingsuggestionphonenumberjavascriptapioutlook11officephonenumbertasksuggestionjavascriptapioutlook11officetasksuggestion"></a><span data-ttu-id="8972a-694">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_1/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_1/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_1/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_1/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="8972a-694">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_1/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_1/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_1/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_1/office.tasksuggestion))>}</span></span>

<span data-ttu-id="8972a-695">获取所选项目的正文中找到的指定实体类型的所有实体的数组。</span><span class="sxs-lookup"><span data-stu-id="8972a-695">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="8972a-696">在 Outlook for iOS 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="8972a-696">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="8972a-697">参数：</span><span class="sxs-lookup"><span data-stu-id="8972a-697">Parameters:</span></span>

|<span data-ttu-id="8972a-698">名称</span><span class="sxs-lookup"><span data-stu-id="8972a-698">Name</span></span>| <span data-ttu-id="8972a-699">类型</span><span class="sxs-lookup"><span data-stu-id="8972a-699">Type</span></span>| <span data-ttu-id="8972a-700">说明</span><span class="sxs-lookup"><span data-stu-id="8972a-700">Description</span></span>|
|---|---|---|
|`entityType`| [<span data-ttu-id="8972a-701">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="8972a-701">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook_1_1/office.MailboxEnums.entitytype)|<span data-ttu-id="8972a-702">EntityType 枚举值之一。</span><span class="sxs-lookup"><span data-stu-id="8972a-702">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="8972a-703">要求</span><span class="sxs-lookup"><span data-stu-id="8972a-703">Requirements</span></span>

|<span data-ttu-id="8972a-704">要求</span><span class="sxs-lookup"><span data-stu-id="8972a-704">Requirement</span></span>| <span data-ttu-id="8972a-705">值</span><span class="sxs-lookup"><span data-stu-id="8972a-705">Value</span></span>|
|---|---|
|[<span data-ttu-id="8972a-706">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="8972a-706">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8972a-707">1.0</span><span class="sxs-lookup"><span data-stu-id="8972a-707">1.0</span></span>|
|[<span data-ttu-id="8972a-708">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="8972a-708">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8972a-709">受限</span><span class="sxs-lookup"><span data-stu-id="8972a-709">Restricted</span></span>|
|[<span data-ttu-id="8972a-710">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="8972a-710">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="8972a-711">阅读</span><span class="sxs-lookup"><span data-stu-id="8972a-711">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="8972a-712">返回：</span><span class="sxs-lookup"><span data-stu-id="8972a-712">Returns:</span></span>

<span data-ttu-id="8972a-713">如果在 `entityType` 中传递的值不是 `EntityType` 枚举的有效成员，该方法返回 null。</span><span class="sxs-lookup"><span data-stu-id="8972a-713">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="8972a-714">如果指定类型的任何实体都不存在于该项目的正文中，该方法将返回空数组。</span><span class="sxs-lookup"><span data-stu-id="8972a-714">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="8972a-715">否则，返回的数组中对象的类型取决于 `entityType` 参数中请求实体的类型。</span><span class="sxs-lookup"><span data-stu-id="8972a-715">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="8972a-716">当使用此方法的最低权限级别**受限**时，某些实体类型需要 **ReadItem** 才能进行访问，如下表中所指定。</span><span class="sxs-lookup"><span data-stu-id="8972a-716">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

| <span data-ttu-id="8972a-717">`entityType` 的值</span><span class="sxs-lookup"><span data-stu-id="8972a-717">Value of `entityType`</span></span> | <span data-ttu-id="8972a-718">返回的数组中对象的类型</span><span class="sxs-lookup"><span data-stu-id="8972a-718">Type of objects in returned array</span></span> | <span data-ttu-id="8972a-719">所需权限级别</span><span class="sxs-lookup"><span data-stu-id="8972a-719">Required Permission Level</span></span> |
| --- | --- | --- |
| `Address` | <span data-ttu-id="8972a-720">String</span><span class="sxs-lookup"><span data-stu-id="8972a-720">String</span></span> | <span data-ttu-id="8972a-721">**受限**</span><span class="sxs-lookup"><span data-stu-id="8972a-721">**Restricted**</span></span> |
| `Contact` | <span data-ttu-id="8972a-722">Contact</span><span class="sxs-lookup"><span data-stu-id="8972a-722">Contact</span></span> | <span data-ttu-id="8972a-723">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="8972a-723">**ReadItem**</span></span> |
| `EmailAddress` | <span data-ttu-id="8972a-724">String</span><span class="sxs-lookup"><span data-stu-id="8972a-724">String</span></span> | <span data-ttu-id="8972a-725">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="8972a-725">**ReadItem**</span></span> |
| `MeetingSuggestion` | <span data-ttu-id="8972a-726">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="8972a-726">MeetingSuggestion</span></span> | <span data-ttu-id="8972a-727">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="8972a-727">**ReadItem**</span></span> |
| `PhoneNumber` | <span data-ttu-id="8972a-728">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="8972a-728">PhoneNumber</span></span> | <span data-ttu-id="8972a-729">**受限**</span><span class="sxs-lookup"><span data-stu-id="8972a-729">**Restricted**</span></span> |
| `TaskSuggestion` | <span data-ttu-id="8972a-730">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="8972a-730">TaskSuggestion</span></span> | <span data-ttu-id="8972a-731">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="8972a-731">**ReadItem**</span></span> |
| `URL` | <span data-ttu-id="8972a-732">String</span><span class="sxs-lookup"><span data-stu-id="8972a-732">String</span></span> | <span data-ttu-id="8972a-733">**受限**</span><span class="sxs-lookup"><span data-stu-id="8972a-733">**Restricted**</span></span> |

<span data-ttu-id="8972a-734">类型：Array.<(String|[Contact](/javascript/api/outlook_1_1/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_1/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_1/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_1/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="8972a-734">Type:  Array.<(String|[Contact](/javascript/api/outlook_1_1/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_1/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_1/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_1/office.tasksuggestion))></span></span>


##### <a name="example"></a><span data-ttu-id="8972a-735">示例</span><span class="sxs-lookup"><span data-stu-id="8972a-735">Example</span></span>

<span data-ttu-id="8972a-736">以下示例显示了如何访问表示当前项目的正文中的邮政地址的字符串数组。</span><span class="sxs-lookup"><span data-stu-id="8972a-736">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlook11officecontactmeetingsuggestionjavascriptapioutlook11officemeetingsuggestionphonenumberjavascriptapioutlook11officephonenumbertasksuggestionjavascriptapioutlook11officetasksuggestion"></a><span data-ttu-id="8972a-737">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_1/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_1/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_1/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_1/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="8972a-737">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_1/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_1/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_1/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_1/office.tasksuggestion))>}</span></span>

<span data-ttu-id="8972a-738">返回传递清单 XML 文件中定义的命名筛选器的所选项目中的已知实体。</span><span class="sxs-lookup"><span data-stu-id="8972a-738">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="8972a-739">在 Outlook for iOS 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="8972a-739">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="8972a-740">`getFilteredEntitiesByName` 方法返回匹配在具有指定 `FilterName` 元素值的清单 XML 文件中的 [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) 规则元素中定义的正则表达式的实体。</span><span class="sxs-lookup"><span data-stu-id="8972a-740">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="8972a-741">参数：</span><span class="sxs-lookup"><span data-stu-id="8972a-741">Parameters:</span></span>

|<span data-ttu-id="8972a-742">名称</span><span class="sxs-lookup"><span data-stu-id="8972a-742">Name</span></span>| <span data-ttu-id="8972a-743">类型</span><span class="sxs-lookup"><span data-stu-id="8972a-743">Type</span></span>| <span data-ttu-id="8972a-744">说明</span><span class="sxs-lookup"><span data-stu-id="8972a-744">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="8972a-745">字符串</span><span class="sxs-lookup"><span data-stu-id="8972a-745">String</span></span>|<span data-ttu-id="8972a-746">定义筛选器匹配的 `ItemHasKnownEntity` 规则元素的名称。</span><span class="sxs-lookup"><span data-stu-id="8972a-746">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="8972a-747">要求</span><span class="sxs-lookup"><span data-stu-id="8972a-747">Requirements</span></span>

|<span data-ttu-id="8972a-748">要求</span><span class="sxs-lookup"><span data-stu-id="8972a-748">Requirement</span></span>| <span data-ttu-id="8972a-749">值</span><span class="sxs-lookup"><span data-stu-id="8972a-749">Value</span></span>|
|---|---|
|[<span data-ttu-id="8972a-750">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="8972a-750">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8972a-751">1.0</span><span class="sxs-lookup"><span data-stu-id="8972a-751">1.0</span></span>|
|[<span data-ttu-id="8972a-752">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="8972a-752">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8972a-753">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8972a-753">ReadItem</span></span>|
|[<span data-ttu-id="8972a-754">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="8972a-754">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="8972a-755">阅读</span><span class="sxs-lookup"><span data-stu-id="8972a-755">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="8972a-756">返回：</span><span class="sxs-lookup"><span data-stu-id="8972a-756">Returns:</span></span>

<span data-ttu-id="8972a-p146">如果具有匹配 `name` 参数的 `FilterName` 元素值的清单中没有任何 `ItemHasKnownEntity` 元素，则该方法返回 `null`。如果 `name` 参数匹配清单中的 `ItemHasKnownEntity` 元素，但在匹配的当前项目中没有实体，则该方法返回一个空数组。</span><span class="sxs-lookup"><span data-stu-id="8972a-p146">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>


<span data-ttu-id="8972a-759">类型：Array.<(String|[Contact](/javascript/api/outlook_1_1/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_1/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_1/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_1/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="8972a-759">Type: Array.<(String|[Contact](/javascript/api/outlook_1_1/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_1/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_1/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_1/office.tasksuggestion))></span></span>


#### <a name="getregexmatches--object"></a><span data-ttu-id="8972a-760">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="8972a-760">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="8972a-761">返回所选项目中匹配在清单 XML 文件中定义的正则表达式的字符串值。</span><span class="sxs-lookup"><span data-stu-id="8972a-761">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="8972a-762">在 Outlook for iOS 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="8972a-762">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="8972a-p147">`getRegExMatches` 方法返回匹配在清单 XML 文件中的每个 `ItemHasRegularExpressionMatch` 或 `ItemHasKnownEntity` 规则元素中定义的正则表达式的字符串。对于 `ItemHasRegularExpressionMatch` 规则，匹配字符串必须发生在该规则指定的项目的属性中。`PropertyName` 简单类型定义支持的属性。</span><span class="sxs-lookup"><span data-stu-id="8972a-p147">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="8972a-766">例如，考虑一个外接程序清单具有以下 `Rule` 元素：</span><span class="sxs-lookup"><span data-stu-id="8972a-766">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="8972a-767">从 `getRegExMatches` 返回的对象应有两个属性：`fruits` 和 `veggies`。</span><span class="sxs-lookup"><span data-stu-id="8972a-767">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

> [!NOTE]
> <span data-ttu-id="8972a-p148">如果在项目的正文属性上指定 `ItemHasRegularExpressionMatch` 规则，则正则表达式应进一步筛选正文并应尝试返回该项目的整个正文。使用正则表达式（如 `.*`）获取项目的整个正文并不总是返回预期的结果。</span><span class="sxs-lookup"><span data-stu-id="8972a-p148">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="requirements"></a><span data-ttu-id="8972a-770">要求</span><span class="sxs-lookup"><span data-stu-id="8972a-770">Requirements</span></span>

|<span data-ttu-id="8972a-771">要求</span><span class="sxs-lookup"><span data-stu-id="8972a-771">Requirement</span></span>| <span data-ttu-id="8972a-772">值</span><span class="sxs-lookup"><span data-stu-id="8972a-772">Value</span></span>|
|---|---|
|[<span data-ttu-id="8972a-773">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="8972a-773">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8972a-774">1.0</span><span class="sxs-lookup"><span data-stu-id="8972a-774">1.0</span></span>|
|[<span data-ttu-id="8972a-775">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="8972a-775">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8972a-776">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8972a-776">ReadItem</span></span>|
|[<span data-ttu-id="8972a-777">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="8972a-777">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="8972a-778">阅读</span><span class="sxs-lookup"><span data-stu-id="8972a-778">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="8972a-779">返回：</span><span class="sxs-lookup"><span data-stu-id="8972a-779">Returns:</span></span>

<span data-ttu-id="8972a-p149">一个包含匹配在清单 XML 文件中定义的正则表达式的字符串数组的对象。每个数组的名称等于匹配 `ItemHasRegularExpressionMatch` 规则的 `RegExName` 属性或匹配 `ItemHasKnownEntity` 规则的 `FilterName` 属性的相应值。</span><span class="sxs-lookup"><span data-stu-id="8972a-p149">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<dl class="param-type"><span data-ttu-id="8972a-782">

<dt>
类型</dt>


</span><span class="sxs-lookup"><span data-stu-id="8972a-782">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="8972a-783">对象</span><span class="sxs-lookup"><span data-stu-id="8972a-783">Object</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="8972a-784">示例</span><span class="sxs-lookup"><span data-stu-id="8972a-784">Example</span></span>

<span data-ttu-id="8972a-785">以下示例显示了如何访问正则表达式 <rule> 元素 `fruits` 和 `veggies` 的匹配项的数组，这些元素在清单中指定。</rule></span><span class="sxs-lookup"><span data-stu-id="8972a-785">The following example shows how to access the array of matches for the regular expression <rule>elements `fruits` and `veggies`, which are specified in the manifest.</rule></span></span>

```JavaScript
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veges = allMatches.veggies;
```

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="8972a-786">getRegExMatchesByName(name) → (nullable) {Array.<String>}</span><span class="sxs-lookup"><span data-stu-id="8972a-786">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="8972a-787">返回匹配在清单 XML 文件中定义的命名正则表达式的所选项目中的字符串值。</span><span class="sxs-lookup"><span data-stu-id="8972a-787">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="8972a-788">在 Outlook for iOS 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="8972a-788">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="8972a-789">`getRegExMatchesByName` 方法返回匹配在具有指定 `RegExName` 元素值的清单 XML 文件中的 `ItemHasRegularExpressionMatch` 规则元素中定义的正则表达式的字符串。</span><span class="sxs-lookup"><span data-stu-id="8972a-789">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="8972a-p150">如果在项目的正文属性上指定 `ItemHasRegularExpressionMatch` 规则，则正则表达式应进一步筛选正文，不应尝试返回该项目的整个正文。使用正则表达式（如 `.*`）获取项目的整个正文并不总是返回预期的结果。</span><span class="sxs-lookup"><span data-stu-id="8972a-p150">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="8972a-792">参数：</span><span class="sxs-lookup"><span data-stu-id="8972a-792">Parameters:</span></span>

|<span data-ttu-id="8972a-793">名称</span><span class="sxs-lookup"><span data-stu-id="8972a-793">Name</span></span>| <span data-ttu-id="8972a-794">类型</span><span class="sxs-lookup"><span data-stu-id="8972a-794">Type</span></span>| <span data-ttu-id="8972a-795">说明</span><span class="sxs-lookup"><span data-stu-id="8972a-795">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="8972a-796">字符串</span><span class="sxs-lookup"><span data-stu-id="8972a-796">String</span></span>|<span data-ttu-id="8972a-797">定义筛选器匹配的 `ItemHasRegularExpressionMatch` 规则元素的名称。</span><span class="sxs-lookup"><span data-stu-id="8972a-797">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="8972a-798">要求</span><span class="sxs-lookup"><span data-stu-id="8972a-798">Requirements</span></span>

|<span data-ttu-id="8972a-799">要求</span><span class="sxs-lookup"><span data-stu-id="8972a-799">Requirement</span></span>| <span data-ttu-id="8972a-800">值</span><span class="sxs-lookup"><span data-stu-id="8972a-800">Value</span></span>|
|---|---|
|[<span data-ttu-id="8972a-801">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="8972a-801">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8972a-802">1.0</span><span class="sxs-lookup"><span data-stu-id="8972a-802">1.0</span></span>|
|[<span data-ttu-id="8972a-803">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="8972a-803">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8972a-804">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8972a-804">ReadItem</span></span>|
|[<span data-ttu-id="8972a-805">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="8972a-805">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="8972a-806">阅读</span><span class="sxs-lookup"><span data-stu-id="8972a-806">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="8972a-807">返回：</span><span class="sxs-lookup"><span data-stu-id="8972a-807">Returns:</span></span>

<span data-ttu-id="8972a-808">一个包含匹配在清单 XML 文件中定义的正则表达式的字符串的数组。</span><span class="sxs-lookup"><span data-stu-id="8972a-808">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<dl class="param-type"><span data-ttu-id="8972a-809">

<dt>类型</dt>

</span><span class="sxs-lookup"><span data-stu-id="8972a-809">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="8972a-810">Array.< String ></span><span class="sxs-lookup"><span data-stu-id="8972a-810">Array.< String ></span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="8972a-811">示例</span><span class="sxs-lookup"><span data-stu-id="8972a-811">Example</span></span>

```JavaScript
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

####  <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="8972a-812">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="8972a-812">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="8972a-813">异步加载所选项目上此外接程序的自定义属性。</span><span class="sxs-lookup"><span data-stu-id="8972a-813">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="8972a-p151">自定义属性基于每个应用、每个项目存储为键/值对。此方法在回调中返回 `CustomProperties` 对象，该回调提供访问特定于当前项目和当前外接程序的自定义属性的方法。自定义属性未在项目上加密，因此这不应用作安全存储。</span><span class="sxs-lookup"><span data-stu-id="8972a-p151">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="8972a-817">参数：</span><span class="sxs-lookup"><span data-stu-id="8972a-817">Parameters:</span></span>

|<span data-ttu-id="8972a-818">名称</span><span class="sxs-lookup"><span data-stu-id="8972a-818">Name</span></span>| <span data-ttu-id="8972a-819">类型</span><span class="sxs-lookup"><span data-stu-id="8972a-819">Type</span></span>| <span data-ttu-id="8972a-820">属性</span><span class="sxs-lookup"><span data-stu-id="8972a-820">Attributes</span></span>| <span data-ttu-id="8972a-821">说明</span><span class="sxs-lookup"><span data-stu-id="8972a-821">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="8972a-822">函数</span><span class="sxs-lookup"><span data-stu-id="8972a-822">function</span></span>||<span data-ttu-id="8972a-823">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="8972a-823">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="8972a-824">自定义属性作为 `asyncResult.value` 属性中的 [`CustomProperties`](/javascript/api/outlook_1_1/office.customproperties) 对象提供。</span><span class="sxs-lookup"><span data-stu-id="8972a-824">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook_1_1/office.customproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="8972a-825">此对象可用于获取、设置以及从项目中删除自定义属性，并将自定义属性集的更改重新保存到服务器。</span><span class="sxs-lookup"><span data-stu-id="8972a-825">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`| <span data-ttu-id="8972a-826">对象</span><span class="sxs-lookup"><span data-stu-id="8972a-826">Object</span></span>| <span data-ttu-id="8972a-827">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="8972a-827">&lt;optional&gt;</span></span>|<span data-ttu-id="8972a-828">开发人员可以提供他们想要在回调函数中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="8972a-828">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="8972a-829">此对象可以通过回调函数中的 `asyncResult.asyncContext` 属性进行访问。</span><span class="sxs-lookup"><span data-stu-id="8972a-829">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="8972a-830">要求</span><span class="sxs-lookup"><span data-stu-id="8972a-830">Requirements</span></span>

|<span data-ttu-id="8972a-831">要求</span><span class="sxs-lookup"><span data-stu-id="8972a-831">Requirement</span></span>| <span data-ttu-id="8972a-832">值</span><span class="sxs-lookup"><span data-stu-id="8972a-832">Value</span></span>|
|---|---|
|[<span data-ttu-id="8972a-833">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="8972a-833">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8972a-834">1.0</span><span class="sxs-lookup"><span data-stu-id="8972a-834">1.0</span></span>|
|[<span data-ttu-id="8972a-835">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="8972a-835">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8972a-836">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8972a-836">ReadItem</span></span>|
|[<span data-ttu-id="8972a-837">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="8972a-837">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="8972a-838">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="8972a-838">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="8972a-839">示例</span><span class="sxs-lookup"><span data-stu-id="8972a-839">Example</span></span>

<span data-ttu-id="8972a-p154">以下代码示例显示了如何使用 `loadCustomPropertiesAsync` 方法异步加载特定于当前项目的自定义属性。该示例还显示了如何使用 `CustomProperties.saveAsync` 方法将这些属性重新保存到服务器。加载自定义属性后，该代码示例将使用 `CustomProperties.get` 方法读取自定义属性 `myProp`，使用 `CustomProperties.set` 方法写入自定义属性 `otherProp`，最后调用 `saveAsync` 方法保存这些自定义属性。</span><span class="sxs-lookup"><span data-stu-id="8972a-p154">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

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

####  <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="8972a-843">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="8972a-843">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="8972a-844">将附件从邮件或约会中删除。</span><span class="sxs-lookup"><span data-stu-id="8972a-844">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="8972a-p155">`removeAttachmentAsync` 方法删除项目中带指定标识符的附件。最佳做法是，仅当同一个邮件应用程序在同一会话中添加了一个附件时，你才应使用该附件标识符来删除该附件。在 Outlook Web App 和适用于设备的 OWA 中，附件标识符只在同一个会话中才有效。当用户关闭应用程序，或者如果用户开始在内嵌窗体中撰写，并在随后弹出的内嵌窗体中继续在单独的窗口撰写时，会话即结束。</span><span class="sxs-lookup"><span data-stu-id="8972a-p155">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item. As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session. In Outlook Web App and OWA for Devices, the attachment identifier is valid only within the same session. A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="8972a-849">参数：</span><span class="sxs-lookup"><span data-stu-id="8972a-849">Parameters:</span></span>

|<span data-ttu-id="8972a-850">名称</span><span class="sxs-lookup"><span data-stu-id="8972a-850">Name</span></span>| <span data-ttu-id="8972a-851">类型</span><span class="sxs-lookup"><span data-stu-id="8972a-851">Type</span></span>| <span data-ttu-id="8972a-852">属性</span><span class="sxs-lookup"><span data-stu-id="8972a-852">Attributes</span></span>| <span data-ttu-id="8972a-853">说明</span><span class="sxs-lookup"><span data-stu-id="8972a-853">Description</span></span>|
|---|---|---|---|
|`attachmentId`| <span data-ttu-id="8972a-854">字符串</span><span class="sxs-lookup"><span data-stu-id="8972a-854">String</span></span>||<span data-ttu-id="8972a-855">要删除的附件的标识符。</span><span class="sxs-lookup"><span data-stu-id="8972a-855">The identifier of the attachment to remove.</span></span>|
|`options`| <span data-ttu-id="8972a-856">对象</span><span class="sxs-lookup"><span data-stu-id="8972a-856">Object</span></span>| <span data-ttu-id="8972a-857">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="8972a-857">&lt;optional&gt;</span></span>|<span data-ttu-id="8972a-858">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="8972a-858">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="8972a-859">对象</span><span class="sxs-lookup"><span data-stu-id="8972a-859">Object</span></span>| <span data-ttu-id="8972a-860">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="8972a-860">&lt;optional&gt;</span></span>|<span data-ttu-id="8972a-861">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="8972a-861">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="8972a-862">函数</span><span class="sxs-lookup"><span data-stu-id="8972a-862">function</span></span>| <span data-ttu-id="8972a-863">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="8972a-863">&lt;optional&gt;</span></span>|<span data-ttu-id="8972a-864">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="8972a-864">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="8972a-865">如果删除附件失败，`asyncResult.error` 属性将包含一个说明失败原因的错误代码。</span><span class="sxs-lookup"><span data-stu-id="8972a-865">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="8972a-866">错误</span><span class="sxs-lookup"><span data-stu-id="8972a-866">Errors</span></span>

| <span data-ttu-id="8972a-867">错误代码</span><span class="sxs-lookup"><span data-stu-id="8972a-867">Error code</span></span> | <span data-ttu-id="8972a-868">说明</span><span class="sxs-lookup"><span data-stu-id="8972a-868">Description</span></span> |
|------------|-------------|
| `InvalidAttachmentId` | <span data-ttu-id="8972a-869">附件标识符不存在。</span><span class="sxs-lookup"><span data-stu-id="8972a-869">The attachment identifier does not exist.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="8972a-870">要求</span><span class="sxs-lookup"><span data-stu-id="8972a-870">Requirements</span></span>

|<span data-ttu-id="8972a-871">要求</span><span class="sxs-lookup"><span data-stu-id="8972a-871">Requirement</span></span>| <span data-ttu-id="8972a-872">值</span><span class="sxs-lookup"><span data-stu-id="8972a-872">Value</span></span>|
|---|---|
|[<span data-ttu-id="8972a-873">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="8972a-873">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8972a-874">1.1</span><span class="sxs-lookup"><span data-stu-id="8972a-874">1.1</span></span>|
|[<span data-ttu-id="8972a-875">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="8972a-875">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8972a-876">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="8972a-876">ReadWriteItem</span></span>|
|[<span data-ttu-id="8972a-877">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="8972a-877">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="8972a-878">撰写</span><span class="sxs-lookup"><span data-stu-id="8972a-878">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="8972a-879">示例</span><span class="sxs-lookup"><span data-stu-id="8972a-879">Example</span></span>

<span data-ttu-id="8972a-880">以下代码删除包含标识符 '0' 的附件。</span><span class="sxs-lookup"><span data-stu-id="8972a-880">The following code removes an attachment with an identifier of '0'.</span></span>

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
