---
title: Office.context.mailbox.item-要求设置 1.1
description: ''
ms.date: 01/30/2019
localization_priority: Normal
ms.openlocfilehash: ce8c10987c08609eba90a3a957b372114e62cd81
ms.sourcegitcommit: bf5c56d9b8c573e42bf2268e10ca3fd4d2bb4ff9
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 02/01/2019
ms.locfileid: "29701874"
---
# <a name="item"></a><span data-ttu-id="ae6be-102">item</span><span class="sxs-lookup"><span data-stu-id="ae6be-102">item</span></span>

### <span data-ttu-id="ae6be-p101">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md). item</span><span class="sxs-lookup"><span data-stu-id="ae6be-p101">[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md). item</span></span>

<span data-ttu-id="ae6be-p102">`item` 命名空间用于访问当前选定的邮件、会议请求或约会。可以通过使用 [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlook11officemailboxenumsitemtype) 属性确定 `item` 的类型。</span><span class="sxs-lookup"><span data-stu-id="ae6be-p102">The `item` namespace is used to access the currently selected message, meeting request, or appointment. You can determine the type of the `item` by using the [itemType](#itemtype-officemailboxenumsitemtypejavascriptapioutlook11officemailboxenumsitemtype) property.</span></span>

##### <a name="requirements"></a><span data-ttu-id="ae6be-107">Requirements</span><span class="sxs-lookup"><span data-stu-id="ae6be-107">Requirements</span></span>

|<span data-ttu-id="ae6be-108">要求</span><span class="sxs-lookup"><span data-stu-id="ae6be-108">Requirement</span></span>| <span data-ttu-id="ae6be-109">值</span><span class="sxs-lookup"><span data-stu-id="ae6be-109">Value</span></span>|
|---|---|
|[<span data-ttu-id="ae6be-110">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="ae6be-110">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ae6be-111">1.0</span><span class="sxs-lookup"><span data-stu-id="ae6be-111">1.0</span></span>|
|[<span data-ttu-id="ae6be-112">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="ae6be-112">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ae6be-113">受限</span><span class="sxs-lookup"><span data-stu-id="ae6be-113">Restricted</span></span>|
|[<span data-ttu-id="ae6be-114">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="ae6be-114">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ae6be-115">Compose 或 Read</span><span class="sxs-lookup"><span data-stu-id="ae6be-115">Compose or read</span></span>|

### <a name="example"></a><span data-ttu-id="ae6be-116">示例</span><span class="sxs-lookup"><span data-stu-id="ae6be-116">Example</span></span>

<span data-ttu-id="ae6be-117">以下 JavaScript 代码示例显示了如何访问 Outlook 中当前项目的 `subject` 属性。</span><span class="sxs-lookup"><span data-stu-id="ae6be-117">The following JavaScript code example shows how to access the `subject` property of the current item in Outlook.</span></span>

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

### <a name="members"></a><span data-ttu-id="ae6be-118">成员</span><span class="sxs-lookup"><span data-stu-id="ae6be-118">Members</span></span>

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlook11officeattachmentdetails"></a><span data-ttu-id="ae6be-119">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_1/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="ae6be-119">attachments :Array.<[AttachmentDetails](/javascript/api/outlook_1_1/office.attachmentdetails)></span></span>

<span data-ttu-id="ae6be-p103">获取项目的附件数组。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="ae6be-p103">Gets an array of attachments for the item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="ae6be-122">由于潜在的安全问题，某些类型的文件会受到 Outlook 阻止，并且不会返回。</span><span class="sxs-lookup"><span data-stu-id="ae6be-122">Certain types of files are blocked by Outlook due to potential security issues and are therefore not returned.</span></span> <span data-ttu-id="ae6be-123">如需了解更多信息，请参阅 [Outlook 中阻止的附件](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519)。</span><span class="sxs-lookup"><span data-stu-id="ae6be-123">For more information, see [Blocked attachments in Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).</span></span>

##### <a name="type"></a><span data-ttu-id="ae6be-124">类型：</span><span class="sxs-lookup"><span data-stu-id="ae6be-124">Type:</span></span>

*   <span data-ttu-id="ae6be-125">Array.<[AttachmentDetails](/javascript/api/outlook_1_1/office.attachmentdetails)></span><span class="sxs-lookup"><span data-stu-id="ae6be-125">Array.<[AttachmentDetails](/javascript/api/outlook_1_1/office.attachmentdetails)></span></span>

##### <a name="requirements"></a><span data-ttu-id="ae6be-126">Requirements</span><span class="sxs-lookup"><span data-stu-id="ae6be-126">Requirements</span></span>

|<span data-ttu-id="ae6be-127">要求</span><span class="sxs-lookup"><span data-stu-id="ae6be-127">Requirement</span></span>| <span data-ttu-id="ae6be-128">值</span><span class="sxs-lookup"><span data-stu-id="ae6be-128">Value</span></span>|
|---|---|
|[<span data-ttu-id="ae6be-129">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="ae6be-129">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ae6be-130">1.0</span><span class="sxs-lookup"><span data-stu-id="ae6be-130">1.0</span></span>|
|[<span data-ttu-id="ae6be-131">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="ae6be-131">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ae6be-132">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ae6be-132">ReadItem</span></span>|
|[<span data-ttu-id="ae6be-133">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="ae6be-133">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ae6be-134">阅读</span><span class="sxs-lookup"><span data-stu-id="ae6be-134">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="ae6be-135">示例</span><span class="sxs-lookup"><span data-stu-id="ae6be-135">Example</span></span>

<span data-ttu-id="ae6be-136">以下代码使用当前项目上所有附件的详细信息构成 HTML 字符串。</span><span class="sxs-lookup"><span data-stu-id="ae6be-136">The following code builds an HTML string with details of all attachments on the current item.</span></span>

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

####  <a name="bcc-recipientsjavascriptapioutlook11officerecipients"></a><span data-ttu-id="ae6be-137">bcc :[Recipients](/javascript/api/outlook_1_1/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="ae6be-137">bcc :[Recipients](/javascript/api/outlook_1_1/office.recipients)</span></span>

<span data-ttu-id="ae6be-138">获取对象，该对象提供用于获取或更新邮件的密件抄送 (Bcc) 行上的收件人的方法。</span><span class="sxs-lookup"><span data-stu-id="ae6be-138">Gets an object that provides methods to get or update the recipients on the Bcc (blind carbon copy) line of a message.</span></span> <span data-ttu-id="ae6be-139">仅限撰写模式。</span><span class="sxs-lookup"><span data-stu-id="ae6be-139">Compose mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="ae6be-140">类型:</span><span class="sxs-lookup"><span data-stu-id="ae6be-140">Type:</span></span>

*   [<span data-ttu-id="ae6be-141">收件人</span><span class="sxs-lookup"><span data-stu-id="ae6be-141">Recipients</span></span>](/javascript/api/outlook_1_1/office.recipients)

##### <a name="requirements"></a><span data-ttu-id="ae6be-142">Requirements</span><span class="sxs-lookup"><span data-stu-id="ae6be-142">Requirements</span></span>

|<span data-ttu-id="ae6be-143">要求</span><span class="sxs-lookup"><span data-stu-id="ae6be-143">Requirement</span></span>| <span data-ttu-id="ae6be-144">值</span><span class="sxs-lookup"><span data-stu-id="ae6be-144">Value</span></span>|
|---|---|
|[<span data-ttu-id="ae6be-145">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="ae6be-145">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ae6be-146">1.1</span><span class="sxs-lookup"><span data-stu-id="ae6be-146">1.1</span></span>|
|[<span data-ttu-id="ae6be-147">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="ae6be-147">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ae6be-148">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ae6be-148">ReadItem</span></span>|
|[<span data-ttu-id="ae6be-149">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="ae6be-149">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ae6be-150">撰写</span><span class="sxs-lookup"><span data-stu-id="ae6be-150">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="ae6be-151">示例</span><span class="sxs-lookup"><span data-stu-id="ae6be-151">Example</span></span>

```JavaScript
Office.context.mailbox.item.bcc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.bcc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.bcc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfBccRecipients = asyncResult.value;
}
```

####  <a name="body-bodyjavascriptapioutlook11officebody"></a><span data-ttu-id="ae6be-152">body :[Body](/javascript/api/outlook_1_1/office.body)</span><span class="sxs-lookup"><span data-stu-id="ae6be-152">body :[Body](/javascript/api/outlook_1_1/office.body)</span></span>

<span data-ttu-id="ae6be-153">获取一个提供用于处理项目正文的方法的对象。</span><span class="sxs-lookup"><span data-stu-id="ae6be-153">Gets an object that provides methods for manipulating the body of an item.</span></span>

##### <a name="type"></a><span data-ttu-id="ae6be-154">类型：</span><span class="sxs-lookup"><span data-stu-id="ae6be-154">Type:</span></span>

*   [<span data-ttu-id="ae6be-155">Body</span><span class="sxs-lookup"><span data-stu-id="ae6be-155">Body</span></span>](/javascript/api/outlook_1_1/office.body)

##### <a name="requirements"></a><span data-ttu-id="ae6be-156">要求</span><span class="sxs-lookup"><span data-stu-id="ae6be-156">Requirements</span></span>

|<span data-ttu-id="ae6be-157">要求</span><span class="sxs-lookup"><span data-stu-id="ae6be-157">Requirement</span></span>| <span data-ttu-id="ae6be-158">值</span><span class="sxs-lookup"><span data-stu-id="ae6be-158">Value</span></span>|
|---|---|
|[<span data-ttu-id="ae6be-159">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="ae6be-159">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ae6be-160">1.1</span><span class="sxs-lookup"><span data-stu-id="ae6be-160">1.1</span></span>|
|[<span data-ttu-id="ae6be-161">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="ae6be-161">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ae6be-162">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ae6be-162">ReadItem</span></span>|
|[<span data-ttu-id="ae6be-163">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="ae6be-163">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ae6be-164">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="ae6be-164">Compose or read</span></span>|

####  <a name="cc-arrayemailaddressdetailsjavascriptapioutlook11officeemailaddressdetailsrecipientsjavascriptapioutlook11officerecipients"></a><span data-ttu-id="ae6be-165">cc :Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_1/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="ae6be-165">cc :Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_1/office.recipients)</span></span>

<span data-ttu-id="ae6be-166">提供对邮件的抄送 (Cc) 收件人的访问权限。</span><span class="sxs-lookup"><span data-stu-id="ae6be-166">Provides access to the Cc (carbon copy) recipients of a message.</span></span> <span data-ttu-id="ae6be-167">对象的类型和访问级别取决于当前项目的模式。</span><span class="sxs-lookup"><span data-stu-id="ae6be-167">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="ae6be-168">阅读模式</span><span class="sxs-lookup"><span data-stu-id="ae6be-168">Read mode</span></span>

<span data-ttu-id="ae6be-p107">`cc` 属性返回包含邮件的**抄送**行上所列的每个收件人的 `EmailAddressDetails` 对象的数组。集合上限为 100 个成员。</span><span class="sxs-lookup"><span data-stu-id="ae6be-p107">The `cc` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **Cc** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="ae6be-171">撰写模式</span><span class="sxs-lookup"><span data-stu-id="ae6be-171">Compose mode</span></span>

<span data-ttu-id="ae6be-172">`cc` 属性返回一个 `Recipients` 对象，该对象提供用于获取或更新邮件的“**抄送**”行上收件人的方法。</span><span class="sxs-lookup"><span data-stu-id="ae6be-172">The `cc` property returns a `Recipients` object that provides methods to get or update the recipients on the **Cc** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="ae6be-173">类型：</span><span class="sxs-lookup"><span data-stu-id="ae6be-173">Type:</span></span>

*   <span data-ttu-id="ae6be-174">Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_1/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="ae6be-174">Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_1/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="ae6be-175">要求</span><span class="sxs-lookup"><span data-stu-id="ae6be-175">Requirements</span></span>

|<span data-ttu-id="ae6be-176">要求</span><span class="sxs-lookup"><span data-stu-id="ae6be-176">Requirement</span></span>| <span data-ttu-id="ae6be-177">值</span><span class="sxs-lookup"><span data-stu-id="ae6be-177">Value</span></span>|
|---|---|
|[<span data-ttu-id="ae6be-178">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="ae6be-178">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ae6be-179">1.0</span><span class="sxs-lookup"><span data-stu-id="ae6be-179">1.0</span></span>|
|[<span data-ttu-id="ae6be-180">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="ae6be-180">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ae6be-181">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ae6be-181">ReadItem</span></span>|
|[<span data-ttu-id="ae6be-182">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="ae6be-182">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ae6be-183">Compose 或 Read</span><span class="sxs-lookup"><span data-stu-id="ae6be-183">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="ae6be-184">示例</span><span class="sxs-lookup"><span data-stu-id="ae6be-184">Example</span></span>

```JavaScript
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

####  <a name="nullable-conversationid-string"></a><span data-ttu-id="ae6be-185">(nullable) conversationId :String</span><span class="sxs-lookup"><span data-stu-id="ae6be-185">(nullable) conversationId :String</span></span>

<span data-ttu-id="ae6be-186">获取包含特定消息的电子邮件会话的标识符。</span><span class="sxs-lookup"><span data-stu-id="ae6be-186">Gets an identifier for the email conversation that contains a particular message.</span></span>

<span data-ttu-id="ae6be-p108">如果在阅读窗体或撰写窗体的回复中激活邮件应用程序，则此属性可以获得一个整数值。如果用户随后更改了回复邮件的主题（若发送回复），则该邮件的对话 ID 将改变且之前获取的值将不适用。</span><span class="sxs-lookup"><span data-stu-id="ae6be-p108">You can get an integer for this property if your mail app is activated in read forms or responses in compose forms. If subsequently the user changes the subject of the reply message, upon sending the reply, the conversation ID for that message will change and that value you obtained earlier will no longer apply.</span></span>

<span data-ttu-id="ae6be-p109">对于撰写窗体的新项目，此属性获得一个 null 值。如果用户设置一个主题并保存该项目，`conversationId` 属性将返回一个值。</span><span class="sxs-lookup"><span data-stu-id="ae6be-p109">You get null for this property for a new item in a compose form. If the user sets a subject and saves the item, the `conversationId` property will return a value.</span></span>

##### <a name="type"></a><span data-ttu-id="ae6be-191">类型：</span><span class="sxs-lookup"><span data-stu-id="ae6be-191">Type:</span></span>

*   <span data-ttu-id="ae6be-192">String</span><span class="sxs-lookup"><span data-stu-id="ae6be-192">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="ae6be-193">要求</span><span class="sxs-lookup"><span data-stu-id="ae6be-193">Requirements</span></span>

|<span data-ttu-id="ae6be-194">要求</span><span class="sxs-lookup"><span data-stu-id="ae6be-194">Requirement</span></span>| <span data-ttu-id="ae6be-195">值</span><span class="sxs-lookup"><span data-stu-id="ae6be-195">Value</span></span>|
|---|---|
|[<span data-ttu-id="ae6be-196">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="ae6be-196">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ae6be-197">1.0</span><span class="sxs-lookup"><span data-stu-id="ae6be-197">1.0</span></span>|
|[<span data-ttu-id="ae6be-198">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="ae6be-198">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ae6be-199">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ae6be-199">ReadItem</span></span>|
|[<span data-ttu-id="ae6be-200">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="ae6be-200">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ae6be-201">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="ae6be-201">Compose or read</span></span>|

#### <a name="datetimecreated-date"></a><span data-ttu-id="ae6be-202">dateTimeCreated :Date</span><span class="sxs-lookup"><span data-stu-id="ae6be-202">dateTimeCreated :Date</span></span>

<span data-ttu-id="ae6be-p110">获取项目创建的日期和时间。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="ae6be-p110">Gets the date and time that an item was created. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="ae6be-205">类型：</span><span class="sxs-lookup"><span data-stu-id="ae6be-205">Type:</span></span>

*   <span data-ttu-id="ae6be-206">日期</span><span class="sxs-lookup"><span data-stu-id="ae6be-206">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="ae6be-207">Requirements</span><span class="sxs-lookup"><span data-stu-id="ae6be-207">Requirements</span></span>

|<span data-ttu-id="ae6be-208">要求</span><span class="sxs-lookup"><span data-stu-id="ae6be-208">Requirement</span></span>| <span data-ttu-id="ae6be-209">值</span><span class="sxs-lookup"><span data-stu-id="ae6be-209">Value</span></span>|
|---|---|
|[<span data-ttu-id="ae6be-210">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="ae6be-210">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ae6be-211">1.0</span><span class="sxs-lookup"><span data-stu-id="ae6be-211">1.0</span></span>|
|[<span data-ttu-id="ae6be-212">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="ae6be-212">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ae6be-213">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ae6be-213">ReadItem</span></span>|
|[<span data-ttu-id="ae6be-214">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="ae6be-214">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ae6be-215">阅读</span><span class="sxs-lookup"><span data-stu-id="ae6be-215">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="ae6be-216">示例</span><span class="sxs-lookup"><span data-stu-id="ae6be-216">Example</span></span>

```JavaScript
var created = Office.context.mailbox.item.dateTimeCreated;
```

#### <a name="datetimemodified-date"></a><span data-ttu-id="ae6be-217">dateTimeModified :Date</span><span class="sxs-lookup"><span data-stu-id="ae6be-217">dateTimeModified :Date</span></span>

<span data-ttu-id="ae6be-p111">获取项目最近一次修改的日期和时间。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="ae6be-p111">Gets the date and time that an item was last modified. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="ae6be-220">在 Outlook for iOS 或 Outlook for Android 中不支持此成员。</span><span class="sxs-lookup"><span data-stu-id="ae6be-220">This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="type"></a><span data-ttu-id="ae6be-221">类型：</span><span class="sxs-lookup"><span data-stu-id="ae6be-221">Type:</span></span>

*   <span data-ttu-id="ae6be-222">日期</span><span class="sxs-lookup"><span data-stu-id="ae6be-222">Date</span></span>

##### <a name="requirements"></a><span data-ttu-id="ae6be-223">Requirements</span><span class="sxs-lookup"><span data-stu-id="ae6be-223">Requirements</span></span>

|<span data-ttu-id="ae6be-224">要求</span><span class="sxs-lookup"><span data-stu-id="ae6be-224">Requirement</span></span>| <span data-ttu-id="ae6be-225">值</span><span class="sxs-lookup"><span data-stu-id="ae6be-225">Value</span></span>|
|---|---|
|[<span data-ttu-id="ae6be-226">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="ae6be-226">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ae6be-227">1.0</span><span class="sxs-lookup"><span data-stu-id="ae6be-227">1.0</span></span>|
|[<span data-ttu-id="ae6be-228">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="ae6be-228">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ae6be-229">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ae6be-229">ReadItem</span></span>|
|[<span data-ttu-id="ae6be-230">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="ae6be-230">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ae6be-231">阅读</span><span class="sxs-lookup"><span data-stu-id="ae6be-231">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="ae6be-232">示例</span><span class="sxs-lookup"><span data-stu-id="ae6be-232">Example</span></span>

```JavaScript
var modified = Office.context.mailbox.item.dateTimeModified;
```

####  <a name="end-datetimejavascriptapioutlook11officetime"></a><span data-ttu-id="ae6be-233">end :Date|[Time](/javascript/api/outlook_1_1/office.time)</span><span class="sxs-lookup"><span data-stu-id="ae6be-233">end :Date|[Time](/javascript/api/outlook_1_1/office.time)</span></span>

<span data-ttu-id="ae6be-234">获取或设置约会结束的日期和时间。</span><span class="sxs-lookup"><span data-stu-id="ae6be-234">Gets or sets the date and time that the appointment is to end.</span></span>

<span data-ttu-id="ae6be-p112">将 `end` 属性表示为协调世界时 (UTC) 的日期和时间值。可使用 [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook11officelocalclienttime) 方法将 end 属性值转换为客户端的本地日期和时间。</span><span class="sxs-lookup"><span data-stu-id="ae6be-p112">The `end` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook11officelocalclienttime) method to convert the end property value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="ae6be-237">阅读模式</span><span class="sxs-lookup"><span data-stu-id="ae6be-237">Read mode</span></span>

<span data-ttu-id="ae6be-238">`end` 属性返回 `Date` 对象。</span><span class="sxs-lookup"><span data-stu-id="ae6be-238">The `end` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="ae6be-239">撰写模式</span><span class="sxs-lookup"><span data-stu-id="ae6be-239">Compose mode</span></span>

<span data-ttu-id="ae6be-240">`end` 属性返回 `Time` 对象。</span><span class="sxs-lookup"><span data-stu-id="ae6be-240">The `end` property returns a `Time` object.</span></span>

<span data-ttu-id="ae6be-241">使用 [`Time.setAsync`](/javascript/api/outlook_1_1/office.time#setasync-datetime--options--callback-) 方法设置结束时间时，应使用 [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) 方法将客户端的本地时间转换为服务器的 UTC。</span><span class="sxs-lookup"><span data-stu-id="ae6be-241">When you use the [`Time.setAsync`](/javascript/api/outlook_1_1/office.time#setasync-datetime--options--callback-) method to set the end time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="ae6be-242">类型：</span><span class="sxs-lookup"><span data-stu-id="ae6be-242">Type:</span></span>

*   <span data-ttu-id="ae6be-243">Date | [Time](/javascript/api/outlook_1_1/office.time)</span><span class="sxs-lookup"><span data-stu-id="ae6be-243">Date | [Time](/javascript/api/outlook_1_1/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="ae6be-244">Requirements</span><span class="sxs-lookup"><span data-stu-id="ae6be-244">Requirements</span></span>

|<span data-ttu-id="ae6be-245">要求</span><span class="sxs-lookup"><span data-stu-id="ae6be-245">Requirement</span></span>| <span data-ttu-id="ae6be-246">值</span><span class="sxs-lookup"><span data-stu-id="ae6be-246">Value</span></span>|
|---|---|
|[<span data-ttu-id="ae6be-247">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="ae6be-247">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ae6be-248">1.0</span><span class="sxs-lookup"><span data-stu-id="ae6be-248">1.0</span></span>|
|[<span data-ttu-id="ae6be-249">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="ae6be-249">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ae6be-250">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ae6be-250">ReadItem</span></span>|
|[<span data-ttu-id="ae6be-251">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="ae6be-251">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ae6be-252">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="ae6be-252">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="ae6be-253">示例</span><span class="sxs-lookup"><span data-stu-id="ae6be-253">Example</span></span>

<span data-ttu-id="ae6be-254">以下示例通过使用 `Time` 对象的 [`setAsync`](/javascript/api/outlook_1_1/office.time#setasync-datetime--options--callback-) 方法，设置撰写模式下约会的结束时间。</span><span class="sxs-lookup"><span data-stu-id="ae6be-254">The following example sets the end time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_1/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

#### <a name="from-emailaddressdetailsjavascriptapioutlook11officeemailaddressdetails"></a><span data-ttu-id="ae6be-255">from :[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="ae6be-255">from :[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)</span></span>

<span data-ttu-id="ae6be-p113">获取邮件发件人的电子邮件地址。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="ae6be-p113">Gets the email address of the sender of a message. Read mode only.</span></span>

<span data-ttu-id="ae6be-p114">`from` 和 [`sender`](#sender-emailaddressdetailsjavascriptapioutlook11officeemailaddressdetails) 属性表示同一个人，邮件由代理人发送的除外。在此情况下，`from` 属性表示代理程序，而 sender 属性表示代理人。</span><span class="sxs-lookup"><span data-stu-id="ae6be-p114">The `from` and [`sender`](#sender-emailaddressdetailsjavascriptapioutlook11officeemailaddressdetails) properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="ae6be-260">`from` 属性中 `EmailAddressDetails` 对象的 `recipientType` 属性为 `undefined`。</span><span class="sxs-lookup"><span data-stu-id="ae6be-260">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="ae6be-261">类型：</span><span class="sxs-lookup"><span data-stu-id="ae6be-261">Type:</span></span>

*   [<span data-ttu-id="ae6be-262">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="ae6be-262">EmailAddressDetails</span></span>](/javascript/api/outlook_1_1/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="ae6be-263">要求</span><span class="sxs-lookup"><span data-stu-id="ae6be-263">Requirements</span></span>

|<span data-ttu-id="ae6be-264">要求</span><span class="sxs-lookup"><span data-stu-id="ae6be-264">Requirement</span></span>| <span data-ttu-id="ae6be-265">值</span><span class="sxs-lookup"><span data-stu-id="ae6be-265">Value</span></span>|
|---|---|
|[<span data-ttu-id="ae6be-266">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="ae6be-266">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ae6be-267">1.0</span><span class="sxs-lookup"><span data-stu-id="ae6be-267">1.0</span></span>|
|[<span data-ttu-id="ae6be-268">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="ae6be-268">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ae6be-269">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ae6be-269">ReadItem</span></span>|
|[<span data-ttu-id="ae6be-270">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="ae6be-270">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ae6be-271">阅读</span><span class="sxs-lookup"><span data-stu-id="ae6be-271">Read</span></span>|

#### <a name="internetmessageid-string"></a><span data-ttu-id="ae6be-272">internetMessageId :String</span><span class="sxs-lookup"><span data-stu-id="ae6be-272">internetMessageId :String</span></span>

<span data-ttu-id="ae6be-p115">获取电子邮件的 Internet 消息标识符。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="ae6be-p115">Gets the Internet message identifier for an email message. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="ae6be-275">类型：</span><span class="sxs-lookup"><span data-stu-id="ae6be-275">Type:</span></span>

*   <span data-ttu-id="ae6be-276">String</span><span class="sxs-lookup"><span data-stu-id="ae6be-276">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="ae6be-277">要求</span><span class="sxs-lookup"><span data-stu-id="ae6be-277">Requirements</span></span>

|<span data-ttu-id="ae6be-278">要求</span><span class="sxs-lookup"><span data-stu-id="ae6be-278">Requirement</span></span>| <span data-ttu-id="ae6be-279">值</span><span class="sxs-lookup"><span data-stu-id="ae6be-279">Value</span></span>|
|---|---|
|[<span data-ttu-id="ae6be-280">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="ae6be-280">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ae6be-281">1.0</span><span class="sxs-lookup"><span data-stu-id="ae6be-281">1.0</span></span>|
|[<span data-ttu-id="ae6be-282">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="ae6be-282">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ae6be-283">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ae6be-283">ReadItem</span></span>|
|[<span data-ttu-id="ae6be-284">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="ae6be-284">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ae6be-285">阅读</span><span class="sxs-lookup"><span data-stu-id="ae6be-285">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="ae6be-286">示例</span><span class="sxs-lookup"><span data-stu-id="ae6be-286">Example</span></span>

```JavaScript
var internetMessageId = Office.context.mailbox.item.internetMessageId;
```

#### <a name="itemclass-string"></a><span data-ttu-id="ae6be-287">itemClass :String</span><span class="sxs-lookup"><span data-stu-id="ae6be-287">itemClass :String</span></span>

<span data-ttu-id="ae6be-p116">获取选定项目的 Exchange Web 服务项目类。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="ae6be-p116">Gets the Exchange Web Services item class of the selected item. Read mode only.</span></span>

<span data-ttu-id="ae6be-p117">`itemClass` 属性指定所选项目的邮件类别。以下是邮件或约会项目的默认邮件类别。</span><span class="sxs-lookup"><span data-stu-id="ae6be-p117">The `itemClass` property specifies the message class of the selected item. The following are the default message classes for the message or appointment item.</span></span>

| <span data-ttu-id="ae6be-292">类型</span><span class="sxs-lookup"><span data-stu-id="ae6be-292">Type</span></span> | <span data-ttu-id="ae6be-293">说明</span><span class="sxs-lookup"><span data-stu-id="ae6be-293">Description</span></span> | <span data-ttu-id="ae6be-294">项目类</span><span class="sxs-lookup"><span data-stu-id="ae6be-294">item class</span></span> |
| --- | --- | --- |
| <span data-ttu-id="ae6be-295">约会项目</span><span class="sxs-lookup"><span data-stu-id="ae6be-295">Appointment items</span></span> | <span data-ttu-id="ae6be-296">这些是项目类 `IPM.Appointment` 或 `IPM.Appointment.Occurrence` 的日历项目。</span><span class="sxs-lookup"><span data-stu-id="ae6be-296">These are calendar items of the item class `IPM.Appointment` or `IPM.Appointment.Occurrence`.</span></span> | `IPM.Appointment`<br />`IPM.Appointment.Occurrence` |
| <span data-ttu-id="ae6be-297">邮件项目</span><span class="sxs-lookup"><span data-stu-id="ae6be-297">Message items</span></span> | <span data-ttu-id="ae6be-298">这些项目包括具有默认邮件类别 `IPM.Note` 的电子邮件，以及使用 `IPM.Schedule.Meeting` 作为基础邮件类别的会议请求、响应和取消。</span><span class="sxs-lookup"><span data-stu-id="ae6be-298">These include email messages that have the default message class `IPM.Note`, and meeting requests, responses, and cancellations, that use `IPM.Schedule.Meeting` as the base message class.</span></span> | `IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled` |

<span data-ttu-id="ae6be-299">你可以创建用于扩展默认邮件类别的自定义邮件类别，例如，自定义约会邮件类别 `IPM.Appointment.Contoso`。</span><span class="sxs-lookup"><span data-stu-id="ae6be-299">You can create custom message classes that extends a default message class, for example, a custom appointment message class `IPM.Appointment.Contoso`.</span></span>

##### <a name="type"></a><span data-ttu-id="ae6be-300">类型：</span><span class="sxs-lookup"><span data-stu-id="ae6be-300">Type:</span></span>

*   <span data-ttu-id="ae6be-301">String</span><span class="sxs-lookup"><span data-stu-id="ae6be-301">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="ae6be-302">要求</span><span class="sxs-lookup"><span data-stu-id="ae6be-302">Requirements</span></span>

|<span data-ttu-id="ae6be-303">要求</span><span class="sxs-lookup"><span data-stu-id="ae6be-303">Requirement</span></span>| <span data-ttu-id="ae6be-304">值</span><span class="sxs-lookup"><span data-stu-id="ae6be-304">Value</span></span>|
|---|---|
|[<span data-ttu-id="ae6be-305">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="ae6be-305">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ae6be-306">1.0</span><span class="sxs-lookup"><span data-stu-id="ae6be-306">1.0</span></span>|
|[<span data-ttu-id="ae6be-307">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="ae6be-307">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ae6be-308">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ae6be-308">ReadItem</span></span>|
|[<span data-ttu-id="ae6be-309">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="ae6be-309">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ae6be-310">阅读</span><span class="sxs-lookup"><span data-stu-id="ae6be-310">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="ae6be-311">示例</span><span class="sxs-lookup"><span data-stu-id="ae6be-311">Example</span></span>

```JavaScript
var itemClass = Office.context.mailbox.item.itemClass;
```

#### <a name="nullable-itemid-string"></a><span data-ttu-id="ae6be-312">(nullable) itemId :String</span><span class="sxs-lookup"><span data-stu-id="ae6be-312">(nullable) itemId :String</span></span>

<span data-ttu-id="ae6be-p118">获取当前项目的 Exchange Web 服务项目标识符。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="ae6be-p118">Gets the Exchange Web Services item identifier for the current item. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="ae6be-315">`itemId` 属性返回的标识符与 Exchange Web 服务项目标识符相同。</span><span class="sxs-lookup"><span data-stu-id="ae6be-315">The identifier returned by the `itemId` property is the same as the Exchange Web Services item identifier.</span></span> <span data-ttu-id="ae6be-316">`itemId` 属性与 Outlook 条目 ID 或 Outlook REST API 使用的 ID 不同。</span><span class="sxs-lookup"><span data-stu-id="ae6be-316">The `itemId` property is not identical to the Outlook Entry ID or the ID used by the Outlook REST API.</span></span> <span data-ttu-id="ae6be-317">使用此值进行 REST API 调用前，应使用 `Office.context.mailbox.convertToRestId`（可在要求集 1.3 的开头部分中找到）对它进行转换。</span><span class="sxs-lookup"><span data-stu-id="ae6be-317">Before making REST API calls using this value, it should be converted using `Office.context.mailbox.convertToRestId`, which is available starting in requirement set 1.3.</span></span> <span data-ttu-id="ae6be-318">有关详细信息，请参阅[从 Outlook 加载项使用 Outlook REST API](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id)。</span><span class="sxs-lookup"><span data-stu-id="ae6be-318">For more details, see [Use the Outlook REST APIs from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/use-rest-api#get-the-item-id).</span></span>

##### <a name="type"></a><span data-ttu-id="ae6be-319">类型：</span><span class="sxs-lookup"><span data-stu-id="ae6be-319">Type:</span></span>

*   <span data-ttu-id="ae6be-320">String</span><span class="sxs-lookup"><span data-stu-id="ae6be-320">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="ae6be-321">要求</span><span class="sxs-lookup"><span data-stu-id="ae6be-321">Requirements</span></span>

|<span data-ttu-id="ae6be-322">要求</span><span class="sxs-lookup"><span data-stu-id="ae6be-322">Requirement</span></span>| <span data-ttu-id="ae6be-323">值</span><span class="sxs-lookup"><span data-stu-id="ae6be-323">Value</span></span>|
|---|---|
|[<span data-ttu-id="ae6be-324">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="ae6be-324">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ae6be-325">1.0</span><span class="sxs-lookup"><span data-stu-id="ae6be-325">1.0</span></span>|
|[<span data-ttu-id="ae6be-326">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="ae6be-326">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ae6be-327">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ae6be-327">ReadItem</span></span>|
|[<span data-ttu-id="ae6be-328">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="ae6be-328">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ae6be-329">阅读</span><span class="sxs-lookup"><span data-stu-id="ae6be-329">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="ae6be-330">示例</span><span class="sxs-lookup"><span data-stu-id="ae6be-330">Example</span></span>

<span data-ttu-id="ae6be-p120">以下代码检查项目标识符是否存在。如果 `itemId` 属性返回 `null` 或 `undefined`，则将项目保存到存储，并从异步结果中获取项目标识符。</span><span class="sxs-lookup"><span data-stu-id="ae6be-p120">The following code checks for the presence of an item identifier. If the `itemId` property returns `null` or `undefined`, it saves the item to the store and gets the item identifier from the asynchronous result.</span></span>

```JavaScript
var itemId = Office.context.mailbox.item.itemId;
if (itemId === null || itemId == undefined) {
  Office.context.mailbox.item.saveAsync(function(result){
    itemId = result.value;
  });
}
```

####  <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlook11officemailboxenumsitemtype"></a><span data-ttu-id="ae6be-333">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_1/office.mailboxenums.itemtype)</span><span class="sxs-lookup"><span data-stu-id="ae6be-333">itemType :[Office.MailboxEnums.ItemType](/javascript/api/outlook_1_1/office.mailboxenums.itemtype)</span></span>

<span data-ttu-id="ae6be-334">获取实例表示的项的类型。</span><span class="sxs-lookup"><span data-stu-id="ae6be-334">Gets the type of item that an instance represents.</span></span>

<span data-ttu-id="ae6be-335">`itemType` 属性返回其中一个 `ItemType` 枚举值，指示 `item` 对象实例是邮件还是约会。</span><span class="sxs-lookup"><span data-stu-id="ae6be-335">The `itemType` property returns one of the `ItemType` enumeration values, indicating whether the `item` object instance is a message or an appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="ae6be-336">类型：</span><span class="sxs-lookup"><span data-stu-id="ae6be-336">Type:</span></span>

*   [<span data-ttu-id="ae6be-337">Office.MailboxEnums.ItemType</span><span class="sxs-lookup"><span data-stu-id="ae6be-337">Office.MailboxEnums.ItemType</span></span>](/javascript/api/outlook_1_1/office.mailboxenums.itemtype)

##### <a name="requirements"></a><span data-ttu-id="ae6be-338">要求</span><span class="sxs-lookup"><span data-stu-id="ae6be-338">Requirements</span></span>

|<span data-ttu-id="ae6be-339">要求</span><span class="sxs-lookup"><span data-stu-id="ae6be-339">Requirement</span></span>| <span data-ttu-id="ae6be-340">值</span><span class="sxs-lookup"><span data-stu-id="ae6be-340">Value</span></span>|
|---|---|
|[<span data-ttu-id="ae6be-341">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="ae6be-341">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ae6be-342">1.0</span><span class="sxs-lookup"><span data-stu-id="ae6be-342">1.0</span></span>|
|[<span data-ttu-id="ae6be-343">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="ae6be-343">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ae6be-344">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ae6be-344">ReadItem</span></span>|
|[<span data-ttu-id="ae6be-345">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="ae6be-345">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ae6be-346">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="ae6be-346">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="ae6be-347">示例</span><span class="sxs-lookup"><span data-stu-id="ae6be-347">Example</span></span>

```JavaScript
if (Office.context.mailbox.item.itemType == Office.MailboxEnums.ItemType.Message)
  // do something
else
  // do something else
```

####  <a name="location-stringlocationjavascriptapioutlook11officelocation"></a><span data-ttu-id="ae6be-348">location :String|[Location](/javascript/api/outlook_1_1/office.location)</span><span class="sxs-lookup"><span data-stu-id="ae6be-348">location :String|[Location](/javascript/api/outlook_1_1/office.location)</span></span>

<span data-ttu-id="ae6be-349">获取或设置约会的位置。</span><span class="sxs-lookup"><span data-stu-id="ae6be-349">Gets or sets the location of an appointment.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="ae6be-350">阅读模式</span><span class="sxs-lookup"><span data-stu-id="ae6be-350">Read mode</span></span>

<span data-ttu-id="ae6be-351">`location` 属性返回一个包含约会位置的字符串。</span><span class="sxs-lookup"><span data-stu-id="ae6be-351">The `location` property returns a string that contains the location of the appointment.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="ae6be-352">撰写模式</span><span class="sxs-lookup"><span data-stu-id="ae6be-352">Compose mode</span></span>

<span data-ttu-id="ae6be-353">`location` 属性返回一个 `Location` 对象，该对象提供用于获取和设置约会位置的方法。</span><span class="sxs-lookup"><span data-stu-id="ae6be-353">The `location` property returns a `Location` object that provides methods that are used to get and set the location of the appointment.</span></span>

##### <a name="type"></a><span data-ttu-id="ae6be-354">类型：</span><span class="sxs-lookup"><span data-stu-id="ae6be-354">Type:</span></span>

*   <span data-ttu-id="ae6be-355">String | [Location](/javascript/api/outlook_1_1/office.location)</span><span class="sxs-lookup"><span data-stu-id="ae6be-355">String | [Location](/javascript/api/outlook_1_1/office.location)</span></span>

##### <a name="requirements"></a><span data-ttu-id="ae6be-356">要求</span><span class="sxs-lookup"><span data-stu-id="ae6be-356">Requirements</span></span>

|<span data-ttu-id="ae6be-357">要求</span><span class="sxs-lookup"><span data-stu-id="ae6be-357">Requirement</span></span>| <span data-ttu-id="ae6be-358">值</span><span class="sxs-lookup"><span data-stu-id="ae6be-358">Value</span></span>|
|---|---|
|[<span data-ttu-id="ae6be-359">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="ae6be-359">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ae6be-360">1.0</span><span class="sxs-lookup"><span data-stu-id="ae6be-360">1.0</span></span>|
|[<span data-ttu-id="ae6be-361">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="ae6be-361">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ae6be-362">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ae6be-362">ReadItem</span></span>|
|[<span data-ttu-id="ae6be-363">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="ae6be-363">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ae6be-364">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="ae6be-364">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="ae6be-365">示例</span><span class="sxs-lookup"><span data-stu-id="ae6be-365">Example</span></span>

```JavaScript
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

#### <a name="normalizedsubject-string"></a><span data-ttu-id="ae6be-366">normalizedSubject :String</span><span class="sxs-lookup"><span data-stu-id="ae6be-366">normalizedSubject :String</span></span>

<span data-ttu-id="ae6be-p121">获取删除了所有前缀（包括 `RE:` 和 `FWD:`）的项目主题。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="ae6be-p121">Gets the subject of an item, with all prefixes removed (including `RE:` and `FWD:`). Read mode only.</span></span>

<span data-ttu-id="ae6be-p122">normalizedSubject 属性获取包含由电子邮件程序添加的任何标准前缀（如 `RE:` 和 `FW:`）的项目主题。若要获取包含完整前缀的项目主题，请使用 [`subject`](#subject-stringsubjectjavascriptapioutlook11officesubject) 属性。</span><span class="sxs-lookup"><span data-stu-id="ae6be-p122">The normalizedSubject property gets the subject of the item, with any standard prefixes (such as `RE:` and `FW:`) that are added by email programs. To get the subject of the item with the prefixes intact, use the [`subject`](#subject-stringsubjectjavascriptapioutlook11officesubject) property.</span></span>

##### <a name="type"></a><span data-ttu-id="ae6be-371">类型：</span><span class="sxs-lookup"><span data-stu-id="ae6be-371">Type:</span></span>

*   <span data-ttu-id="ae6be-372">String</span><span class="sxs-lookup"><span data-stu-id="ae6be-372">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="ae6be-373">要求</span><span class="sxs-lookup"><span data-stu-id="ae6be-373">Requirements</span></span>

|<span data-ttu-id="ae6be-374">要求</span><span class="sxs-lookup"><span data-stu-id="ae6be-374">Requirement</span></span>| <span data-ttu-id="ae6be-375">值</span><span class="sxs-lookup"><span data-stu-id="ae6be-375">Value</span></span>|
|---|---|
|[<span data-ttu-id="ae6be-376">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="ae6be-376">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ae6be-377">1.0</span><span class="sxs-lookup"><span data-stu-id="ae6be-377">1.0</span></span>|
|[<span data-ttu-id="ae6be-378">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="ae6be-378">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ae6be-379">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ae6be-379">ReadItem</span></span>|
|[<span data-ttu-id="ae6be-380">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="ae6be-380">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ae6be-381">阅读</span><span class="sxs-lookup"><span data-stu-id="ae6be-381">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="ae6be-382">示例</span><span class="sxs-lookup"><span data-stu-id="ae6be-382">Example</span></span>

```JavaScript
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
```

####  <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlook11officeemailaddressdetailsrecipientsjavascriptapioutlook11officerecipients"></a><span data-ttu-id="ae6be-383">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_1/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="ae6be-383">optionalAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_1/office.recipients)</span></span>

<span data-ttu-id="ae6be-384">提供对事件的可选与会者的访问权限。</span><span class="sxs-lookup"><span data-stu-id="ae6be-384">Provides access to the optional attendees of an event.</span></span> <span data-ttu-id="ae6be-385">对象的类型和访问级别取决于当前项目的模式。</span><span class="sxs-lookup"><span data-stu-id="ae6be-385">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="ae6be-386">阅读模式</span><span class="sxs-lookup"><span data-stu-id="ae6be-386">Read mode</span></span>

<span data-ttu-id="ae6be-387">`optionalAttendees` 属性返回一个数组，其中包含每个可选与会者的 `EmailAddressDetails` 对象。</span><span class="sxs-lookup"><span data-stu-id="ae6be-387">The `optionalAttendees` property returns an array that contains an `EmailAddressDetails` object for each optional attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="ae6be-388">撰写模式</span><span class="sxs-lookup"><span data-stu-id="ae6be-388">Compose mode</span></span>

<span data-ttu-id="ae6be-389">`optionalAttendees` 属性返回一个 `Recipients` 对象，该对象提供用于获取或更新可选与会者的方法。</span><span class="sxs-lookup"><span data-stu-id="ae6be-389">The `optionalAttendees` property returns a `Recipients` object that provides methods to get or update the optional attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="ae6be-390">类型：</span><span class="sxs-lookup"><span data-stu-id="ae6be-390">Type:</span></span>

*   <span data-ttu-id="ae6be-391">Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_1/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="ae6be-391">Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_1/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="ae6be-392">Requirements</span><span class="sxs-lookup"><span data-stu-id="ae6be-392">Requirements</span></span>

|<span data-ttu-id="ae6be-393">要求</span><span class="sxs-lookup"><span data-stu-id="ae6be-393">Requirement</span></span>| <span data-ttu-id="ae6be-394">值</span><span class="sxs-lookup"><span data-stu-id="ae6be-394">Value</span></span>|
|---|---|
|[<span data-ttu-id="ae6be-395">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="ae6be-395">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ae6be-396">1.0</span><span class="sxs-lookup"><span data-stu-id="ae6be-396">1.0</span></span>|
|[<span data-ttu-id="ae6be-397">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="ae6be-397">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ae6be-398">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ae6be-398">ReadItem</span></span>|
|[<span data-ttu-id="ae6be-399">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="ae6be-399">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ae6be-400">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="ae6be-400">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="ae6be-401">示例</span><span class="sxs-lookup"><span data-stu-id="ae6be-401">Example</span></span>

```JavaScript
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

#### <a name="organizer-emailaddressdetailsjavascriptapioutlook11officeemailaddressdetails"></a><span data-ttu-id="ae6be-402">organizer :[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="ae6be-402">organizer :[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)</span></span>

<span data-ttu-id="ae6be-p124">获取指定会议的会议组织者的电子邮件地址。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="ae6be-p124">Gets the email address of the meeting organizer for a specified meeting. Read mode only.</span></span>

##### <a name="type"></a><span data-ttu-id="ae6be-405">类型:</span><span class="sxs-lookup"><span data-stu-id="ae6be-405">Type:</span></span>

*   [<span data-ttu-id="ae6be-406">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="ae6be-406">EmailAddressDetails</span></span>](/javascript/api/outlook_1_1/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="ae6be-407">Requirements</span><span class="sxs-lookup"><span data-stu-id="ae6be-407">Requirements</span></span>

|<span data-ttu-id="ae6be-408">要求</span><span class="sxs-lookup"><span data-stu-id="ae6be-408">Requirement</span></span>| <span data-ttu-id="ae6be-409">值</span><span class="sxs-lookup"><span data-stu-id="ae6be-409">Value</span></span>|
|---|---|
|[<span data-ttu-id="ae6be-410">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="ae6be-410">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ae6be-411">1.0</span><span class="sxs-lookup"><span data-stu-id="ae6be-411">1.0</span></span>|
|[<span data-ttu-id="ae6be-412">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="ae6be-412">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ae6be-413">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ae6be-413">ReadItem</span></span>|
|[<span data-ttu-id="ae6be-414">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="ae6be-414">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ae6be-415">阅读</span><span class="sxs-lookup"><span data-stu-id="ae6be-415">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="ae6be-416">示例</span><span class="sxs-lookup"><span data-stu-id="ae6be-416">Example</span></span>

```JavaScript
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
```

####  <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlook11officeemailaddressdetailsrecipientsjavascriptapioutlook11officerecipients"></a><span data-ttu-id="ae6be-417">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_1/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="ae6be-417">requiredAttendees :Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_1/office.recipients)</span></span>

<span data-ttu-id="ae6be-418">提供对事件的必需与会者的访问权限。</span><span class="sxs-lookup"><span data-stu-id="ae6be-418">Provides access to the required attendees of an event.</span></span> <span data-ttu-id="ae6be-419">对象的类型和访问级别取决于当前项目的模式。</span><span class="sxs-lookup"><span data-stu-id="ae6be-419">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="ae6be-420">阅读模式</span><span class="sxs-lookup"><span data-stu-id="ae6be-420">Read mode</span></span>

<span data-ttu-id="ae6be-421">`requiredAttendees` 属性返回一个数组，其中包含每个必需与会者的 `EmailAddressDetails` 对象。</span><span class="sxs-lookup"><span data-stu-id="ae6be-421">The `requiredAttendees` property returns an array that contains an `EmailAddressDetails` object for each required attendee to the meeting.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="ae6be-422">撰写模式</span><span class="sxs-lookup"><span data-stu-id="ae6be-422">Compose mode</span></span>

<span data-ttu-id="ae6be-423">`requiredAttendees` 属性返回一个 `Recipients` 对象，该对象提供用于获取或更新必需与会者的方法。</span><span class="sxs-lookup"><span data-stu-id="ae6be-423">The `requiredAttendees` property returns a `Recipients` object that provides methods to get or update the required attendees for a meeting.</span></span>

##### <a name="type"></a><span data-ttu-id="ae6be-424">类型：</span><span class="sxs-lookup"><span data-stu-id="ae6be-424">Type:</span></span>

*   <span data-ttu-id="ae6be-425">Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_1/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="ae6be-425">Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_1/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="ae6be-426">Requirements</span><span class="sxs-lookup"><span data-stu-id="ae6be-426">Requirements</span></span>

|<span data-ttu-id="ae6be-427">要求</span><span class="sxs-lookup"><span data-stu-id="ae6be-427">Requirement</span></span>| <span data-ttu-id="ae6be-428">值</span><span class="sxs-lookup"><span data-stu-id="ae6be-428">Value</span></span>|
|---|---|
|[<span data-ttu-id="ae6be-429">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="ae6be-429">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ae6be-430">1.0</span><span class="sxs-lookup"><span data-stu-id="ae6be-430">1.0</span></span>|
|[<span data-ttu-id="ae6be-431">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="ae6be-431">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ae6be-432">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ae6be-432">ReadItem</span></span>|
|[<span data-ttu-id="ae6be-433">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="ae6be-433">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ae6be-434">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="ae6be-434">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="ae6be-435">示例</span><span class="sxs-lookup"><span data-stu-id="ae6be-435">Example</span></span>

```JavaScript
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
}
```

#### <a name="sender-emailaddressdetailsjavascriptapioutlook11officeemailaddressdetails"></a><span data-ttu-id="ae6be-436">sender :[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)</span><span class="sxs-lookup"><span data-stu-id="ae6be-436">sender :[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)</span></span>

<span data-ttu-id="ae6be-p126">获取电子邮件发件人的电子邮件地址。仅限阅读模式。</span><span class="sxs-lookup"><span data-stu-id="ae6be-p126">Gets the email address of the sender of an email message. Read mode only.</span></span>

<span data-ttu-id="ae6be-p127">[`from`](#from-emailaddressdetailsjavascriptapioutlook11officeemailaddressdetails) 和 `sender` 属性表示同一个人，邮件由代理人发送的除外。在此情况下，`from` 属性表示代理程序，而 sender 属性表示代理人。</span><span class="sxs-lookup"><span data-stu-id="ae6be-p127">The [`from`](#from-emailaddressdetailsjavascriptapioutlook11officeemailaddressdetails) and `sender` properties represent the same person unless the message is sent by a delegate. In that case, the `from` property represents the delegator, and the sender property represents the delegate.</span></span>

> [!NOTE]
> <span data-ttu-id="ae6be-441">`from` 属性中 `EmailAddressDetails` 对象的 `recipientType` 属性为 `undefined`。</span><span class="sxs-lookup"><span data-stu-id="ae6be-441">The `recipientType` property of the `EmailAddressDetails` object in the `from` property is `undefined`.</span></span>

##### <a name="type"></a><span data-ttu-id="ae6be-442">类型:</span><span class="sxs-lookup"><span data-stu-id="ae6be-442">Type:</span></span>

*   [<span data-ttu-id="ae6be-443">EmailAddressDetails</span><span class="sxs-lookup"><span data-stu-id="ae6be-443">EmailAddressDetails</span></span>](/javascript/api/outlook_1_1/office.emailaddressdetails)

##### <a name="requirements"></a><span data-ttu-id="ae6be-444">Requirements</span><span class="sxs-lookup"><span data-stu-id="ae6be-444">Requirements</span></span>

|<span data-ttu-id="ae6be-445">要求</span><span class="sxs-lookup"><span data-stu-id="ae6be-445">Requirement</span></span>| <span data-ttu-id="ae6be-446">值</span><span class="sxs-lookup"><span data-stu-id="ae6be-446">Value</span></span>|
|---|---|
|[<span data-ttu-id="ae6be-447">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="ae6be-447">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ae6be-448">1.0</span><span class="sxs-lookup"><span data-stu-id="ae6be-448">1.0</span></span>|
|[<span data-ttu-id="ae6be-449">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="ae6be-449">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ae6be-450">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ae6be-450">ReadItem</span></span>|
|[<span data-ttu-id="ae6be-451">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="ae6be-451">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ae6be-452">阅读</span><span class="sxs-lookup"><span data-stu-id="ae6be-452">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="ae6be-453">示例</span><span class="sxs-lookup"><span data-stu-id="ae6be-453">Example</span></span>

```JavaScript
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
```

####  <a name="start-datetimejavascriptapioutlook11officetime"></a><span data-ttu-id="ae6be-454">start :Date|[Time](/javascript/api/outlook_1_1/office.time)</span><span class="sxs-lookup"><span data-stu-id="ae6be-454">start :Date|[Time](/javascript/api/outlook_1_1/office.time)</span></span>

<span data-ttu-id="ae6be-455">获取或设置约会开始的日期和时间。</span><span class="sxs-lookup"><span data-stu-id="ae6be-455">Gets or sets the date and time that the appointment is to begin.</span></span>

<span data-ttu-id="ae6be-p128">将 `start` 属性表示为协调世界时 (UTC) 的日期和时间值。可使用 [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook11officelocalclienttime) 方法将该值转换为客户端的本地日期和时间。</span><span class="sxs-lookup"><span data-stu-id="ae6be-p128">The `start` property is expressed as a Coordinated Universal Time (UTC) date and time value. You can use the [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook11officelocalclienttime) method to convert the value to the client’s local date and time.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="ae6be-458">阅读模式</span><span class="sxs-lookup"><span data-stu-id="ae6be-458">Read mode</span></span>

<span data-ttu-id="ae6be-459">`start` 属性返回 `Date` 对象。</span><span class="sxs-lookup"><span data-stu-id="ae6be-459">The `start` property returns a `Date` object.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="ae6be-460">撰写模式</span><span class="sxs-lookup"><span data-stu-id="ae6be-460">Compose mode</span></span>

<span data-ttu-id="ae6be-461">`start` 属性返回 `Time` 对象。</span><span class="sxs-lookup"><span data-stu-id="ae6be-461">The `start` property returns a `Time` object.</span></span>

<span data-ttu-id="ae6be-462">使用 [`Time.setAsync`](/javascript/api/outlook_1_1/office.time#setasync-datetime--options--callback-) 方法设置开始时间时，应使用 [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) 方法将客户端的本地时间转换为服务器的 UTC。</span><span class="sxs-lookup"><span data-stu-id="ae6be-462">When you use the [`Time.setAsync`](/javascript/api/outlook_1_1/office.time#setasync-datetime--options--callback-) method to set the start time, you should use the [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) method to convert the local time on the client to UTC for the server.</span></span>

##### <a name="type"></a><span data-ttu-id="ae6be-463">类型：</span><span class="sxs-lookup"><span data-stu-id="ae6be-463">Type:</span></span>

*   <span data-ttu-id="ae6be-464">Date | [Time](/javascript/api/outlook_1_1/office.time)</span><span class="sxs-lookup"><span data-stu-id="ae6be-464">Date | [Time](/javascript/api/outlook_1_1/office.time)</span></span>

##### <a name="requirements"></a><span data-ttu-id="ae6be-465">Requirements</span><span class="sxs-lookup"><span data-stu-id="ae6be-465">Requirements</span></span>

|<span data-ttu-id="ae6be-466">要求</span><span class="sxs-lookup"><span data-stu-id="ae6be-466">Requirement</span></span>| <span data-ttu-id="ae6be-467">值</span><span class="sxs-lookup"><span data-stu-id="ae6be-467">Value</span></span>|
|---|---|
|[<span data-ttu-id="ae6be-468">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="ae6be-468">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ae6be-469">1.0</span><span class="sxs-lookup"><span data-stu-id="ae6be-469">1.0</span></span>|
|[<span data-ttu-id="ae6be-470">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="ae6be-470">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ae6be-471">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ae6be-471">ReadItem</span></span>|
|[<span data-ttu-id="ae6be-472">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="ae6be-472">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ae6be-473">Compose 或 Read</span><span class="sxs-lookup"><span data-stu-id="ae6be-473">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="ae6be-474">示例</span><span class="sxs-lookup"><span data-stu-id="ae6be-474">Example</span></span>

<span data-ttu-id="ae6be-475">以下示例通过使用 `Time` 对象的 [`setAsync`](/javascript/api/outlook_1_1/office.time#setasync-datetime--options--callback-) 方法，设置撰写模式下约会的开始时间。</span><span class="sxs-lookup"><span data-stu-id="ae6be-475">The following example sets the start time of an appointment in compose mode by using the [`setAsync`](/javascript/api/outlook_1_1/office.time#setasync-datetime--options--callback-) method of the `Time` object.</span></span>

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

####  <a name="subject-stringsubjectjavascriptapioutlook11officesubject"></a><span data-ttu-id="ae6be-476">subject :String|[Subject](/javascript/api/outlook_1_1/office.subject)</span><span class="sxs-lookup"><span data-stu-id="ae6be-476">subject :String|[Subject](/javascript/api/outlook_1_1/office.subject)</span></span>

<span data-ttu-id="ae6be-477">获取或设置显示在项目的主题字段中的说明。</span><span class="sxs-lookup"><span data-stu-id="ae6be-477">Gets or sets the description that appears in the subject field of an item.</span></span>

<span data-ttu-id="ae6be-478">`subject` 属性获取或设置由电子邮件服务器发送项目时的整个主题。</span><span class="sxs-lookup"><span data-stu-id="ae6be-478">The `subject` property gets or sets the entire subject of the item, as sent by the email server.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="ae6be-479">阅读模式</span><span class="sxs-lookup"><span data-stu-id="ae6be-479">Read mode</span></span>

<span data-ttu-id="ae6be-p129">`subject` 属性返回一个字符串。使用 [`normalizedSubject`](#normalizedsubject-string) 属性获取不带任何前导前缀（如 `RE:` 和 `FW:`）的主题。</span><span class="sxs-lookup"><span data-stu-id="ae6be-p129">The `subject` property returns a string. Use the [`normalizedSubject`](#normalizedsubject-string) property to get the subject minus any leading prefixes such as `RE:` and `FW:`.</span></span>

```js
var subject = Office.context.mailbox.item.subject;
```

##### <a name="compose-mode"></a><span data-ttu-id="ae6be-482">撰写模式</span><span class="sxs-lookup"><span data-stu-id="ae6be-482">Compose mode</span></span>

<span data-ttu-id="ae6be-483">`subject` 属性返回一个 `Subject` 对象，该对象提供用于获取和设置主题的方法。</span><span class="sxs-lookup"><span data-stu-id="ae6be-483">The `subject` property returns a `Subject` object that provides methods to get and set the subject.</span></span>

```JavaScript
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
}
```

##### <a name="type"></a><span data-ttu-id="ae6be-484">类型：</span><span class="sxs-lookup"><span data-stu-id="ae6be-484">Type:</span></span>

*   <span data-ttu-id="ae6be-485">String | [Subject](/javascript/api/outlook_1_1/office.subject)</span><span class="sxs-lookup"><span data-stu-id="ae6be-485">String | [Subject](/javascript/api/outlook_1_1/office.subject)</span></span>

##### <a name="requirements"></a><span data-ttu-id="ae6be-486">要求</span><span class="sxs-lookup"><span data-stu-id="ae6be-486">Requirements</span></span>

|<span data-ttu-id="ae6be-487">要求</span><span class="sxs-lookup"><span data-stu-id="ae6be-487">Requirement</span></span>| <span data-ttu-id="ae6be-488">值</span><span class="sxs-lookup"><span data-stu-id="ae6be-488">Value</span></span>|
|---|---|
|[<span data-ttu-id="ae6be-489">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="ae6be-489">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ae6be-490">1.0</span><span class="sxs-lookup"><span data-stu-id="ae6be-490">1.0</span></span>|
|[<span data-ttu-id="ae6be-491">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="ae6be-491">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ae6be-492">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ae6be-492">ReadItem</span></span>|
|[<span data-ttu-id="ae6be-493">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="ae6be-493">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ae6be-494">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="ae6be-494">Compose or read</span></span>|

####  <a name="to-arrayemailaddressdetailsjavascriptapioutlook11officeemailaddressdetailsrecipientsjavascriptapioutlook11officerecipients"></a><span data-ttu-id="ae6be-495">to :Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_1/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="ae6be-495">to :Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook_1_1/office.recipients)</span></span>

<span data-ttu-id="ae6be-496">提供对邮件的“**收件人**”行上的收件人的访问权限。</span><span class="sxs-lookup"><span data-stu-id="ae6be-496">Provides access to the recipients on the **To** line of a message.</span></span> <span data-ttu-id="ae6be-497">对象的类型和访问级别取决于当前项目的模式。</span><span class="sxs-lookup"><span data-stu-id="ae6be-497">The type of object and level of access depends on the mode of the current item.</span></span>

##### <a name="read-mode"></a><span data-ttu-id="ae6be-498">阅读模式</span><span class="sxs-lookup"><span data-stu-id="ae6be-498">Read mode</span></span>

<span data-ttu-id="ae6be-p131">`to` 属性返回包含邮件的**收件人**行上所列的每个收件人的 `EmailAddressDetails` 对象的数组。集合上限为 100 个成员。</span><span class="sxs-lookup"><span data-stu-id="ae6be-p131">The `to` property returns an array that contains an `EmailAddressDetails` object for each recipient listed on the **To** line of the message. The collection is limited to a maximum of 100 members.</span></span>

##### <a name="compose-mode"></a><span data-ttu-id="ae6be-501">撰写模式</span><span class="sxs-lookup"><span data-stu-id="ae6be-501">Compose mode</span></span>

<span data-ttu-id="ae6be-502">`to` 属性返回一个 `Recipients` 对象，该对象提供用于获取或更新邮件的“**收件人**”行上收件人的方法。</span><span class="sxs-lookup"><span data-stu-id="ae6be-502">The `to` property returns a `Recipients` object that provides methods to get or update the recipients on the **To** line of the message.</span></span>

##### <a name="type"></a><span data-ttu-id="ae6be-503">类型：</span><span class="sxs-lookup"><span data-stu-id="ae6be-503">Type:</span></span>

*   <span data-ttu-id="ae6be-504">Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_1/office.recipients)</span><span class="sxs-lookup"><span data-stu-id="ae6be-504">Array.<[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook_1_1/office.recipients)</span></span>

##### <a name="requirements"></a><span data-ttu-id="ae6be-505">Requirements</span><span class="sxs-lookup"><span data-stu-id="ae6be-505">Requirements</span></span>

|<span data-ttu-id="ae6be-506">要求</span><span class="sxs-lookup"><span data-stu-id="ae6be-506">Requirement</span></span>| <span data-ttu-id="ae6be-507">值</span><span class="sxs-lookup"><span data-stu-id="ae6be-507">Value</span></span>|
|---|---|
|[<span data-ttu-id="ae6be-508">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="ae6be-508">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ae6be-509">1.0</span><span class="sxs-lookup"><span data-stu-id="ae6be-509">1.0</span></span>|
|[<span data-ttu-id="ae6be-510">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="ae6be-510">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ae6be-511">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ae6be-511">ReadItem</span></span>|
|[<span data-ttu-id="ae6be-512">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="ae6be-512">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ae6be-513">Compose 或 Read</span><span class="sxs-lookup"><span data-stu-id="ae6be-513">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="ae6be-514">示例</span><span class="sxs-lookup"><span data-stu-id="ae6be-514">Example</span></span>

```JavaScript
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

### <a name="methods"></a><span data-ttu-id="ae6be-515">方法</span><span class="sxs-lookup"><span data-stu-id="ae6be-515">Methods</span></span>

####  <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a><span data-ttu-id="ae6be-516">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="ae6be-516">addFileAttachmentAsync(uri, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="ae6be-517">将文件作为附件添加到邮件或约会。</span><span class="sxs-lookup"><span data-stu-id="ae6be-517">Adds a file to a message or appointment as an attachment.</span></span>

<span data-ttu-id="ae6be-518">`addFileAttachmentAsync` 方法在指定的 URI 上载文件并将其附加到撰写窗体中的项目。</span><span class="sxs-lookup"><span data-stu-id="ae6be-518">The `addFileAttachmentAsync` method uploads the file at the specified URI and attaches it to the item in the compose form.</span></span>

<span data-ttu-id="ae6be-519">你随后可以将该标识符与 [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) 方法一同使用，以删除同一个会话中的附件。</span><span class="sxs-lookup"><span data-stu-id="ae6be-519">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

##### <a name="parameters"></a><span data-ttu-id="ae6be-520">参数：</span><span class="sxs-lookup"><span data-stu-id="ae6be-520">Parameters:</span></span>

|<span data-ttu-id="ae6be-521">名称</span><span class="sxs-lookup"><span data-stu-id="ae6be-521">Name</span></span>| <span data-ttu-id="ae6be-522">类型</span><span class="sxs-lookup"><span data-stu-id="ae6be-522">Type</span></span>| <span data-ttu-id="ae6be-523">属性</span><span class="sxs-lookup"><span data-stu-id="ae6be-523">Attributes</span></span>| <span data-ttu-id="ae6be-524">说明</span><span class="sxs-lookup"><span data-stu-id="ae6be-524">Description</span></span>|
|---|---|---|---|
|`uri`| <span data-ttu-id="ae6be-525">String</span><span class="sxs-lookup"><span data-stu-id="ae6be-525">String</span></span>||<span data-ttu-id="ae6be-p132">提供附加到邮件或约会的文件的位置的 URI。最大长度为 2048 个字符。</span><span class="sxs-lookup"><span data-stu-id="ae6be-p132">The URI that provides the location of the file to attach to the message or appointment. The maximum length is 2048 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="ae6be-528">字符串</span><span class="sxs-lookup"><span data-stu-id="ae6be-528">String</span></span>||<span data-ttu-id="ae6be-p133">在附件上载过程中显示的附件名称。最大长度为 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="ae6be-p133">The name of the attachment that is shown while the attachment is uploading. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="ae6be-531">Object</span><span class="sxs-lookup"><span data-stu-id="ae6be-531">Object</span></span>| <span data-ttu-id="ae6be-532">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="ae6be-532">&lt;optional&gt;</span></span>|<span data-ttu-id="ae6be-533">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="ae6be-533">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="ae6be-534">对象</span><span class="sxs-lookup"><span data-stu-id="ae6be-534">Object</span></span>| <span data-ttu-id="ae6be-535">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="ae6be-535">&lt;optional&gt;</span></span>|<span data-ttu-id="ae6be-536">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="ae6be-536">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="ae6be-537">函数</span><span class="sxs-lookup"><span data-stu-id="ae6be-537">function</span></span>| <span data-ttu-id="ae6be-538">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="ae6be-538">&lt;optional&gt;</span></span>|<span data-ttu-id="ae6be-539">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="ae6be-539">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="ae6be-540">如果成功，附件标识符将在 `asyncResult.value` 属性中提供。</span><span class="sxs-lookup"><span data-stu-id="ae6be-540">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="ae6be-541">如果上传附件失败，`asyncResult` 对象将包含一个提供错误说明的 `Error` 对象。</span><span class="sxs-lookup"><span data-stu-id="ae6be-541">If uploading the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="ae6be-542">错误</span><span class="sxs-lookup"><span data-stu-id="ae6be-542">Errors</span></span>

| <span data-ttu-id="ae6be-543">错误代码</span><span class="sxs-lookup"><span data-stu-id="ae6be-543">Error code</span></span> | <span data-ttu-id="ae6be-544">说明</span><span class="sxs-lookup"><span data-stu-id="ae6be-544">Description</span></span> |
|------------|-------------|
| `AttachmentSizeExceeded` | <span data-ttu-id="ae6be-545">附件大小超过了允许的大小。</span><span class="sxs-lookup"><span data-stu-id="ae6be-545">The attachment is larger than allowed.</span></span> |
| `FileTypeNotSupported` | <span data-ttu-id="ae6be-546">该附件的扩展名不是允许的扩展名。</span><span class="sxs-lookup"><span data-stu-id="ae6be-546">The attachment has an extension that is not allowed.</span></span> |
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="ae6be-547">邮件或约会具有的附件过多。</span><span class="sxs-lookup"><span data-stu-id="ae6be-547">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="ae6be-548">要求</span><span class="sxs-lookup"><span data-stu-id="ae6be-548">Requirements</span></span>

|<span data-ttu-id="ae6be-549">要求</span><span class="sxs-lookup"><span data-stu-id="ae6be-549">Requirement</span></span>| <span data-ttu-id="ae6be-550">值</span><span class="sxs-lookup"><span data-stu-id="ae6be-550">Value</span></span>|
|---|---|
|[<span data-ttu-id="ae6be-551">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="ae6be-551">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ae6be-552">1.1</span><span class="sxs-lookup"><span data-stu-id="ae6be-552">1.1</span></span>|
|[<span data-ttu-id="ae6be-553">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="ae6be-553">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ae6be-554">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="ae6be-554">ReadWriteItem</span></span>|
|[<span data-ttu-id="ae6be-555">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="ae6be-555">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ae6be-556">撰写</span><span class="sxs-lookup"><span data-stu-id="ae6be-556">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="ae6be-557">示例</span><span class="sxs-lookup"><span data-stu-id="ae6be-557">Example</span></span>

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

####  <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a><span data-ttu-id="ae6be-558">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="ae6be-558">addItemAttachmentAsync(itemId, attachmentName, [options], [callback])</span></span>

<span data-ttu-id="ae6be-559">将 Exchange 项目（如邮件）作为附件添加到邮件或约会。</span><span class="sxs-lookup"><span data-stu-id="ae6be-559">Adds an Exchange item, such as a message, as an attachment to the message or appointment.</span></span>

<span data-ttu-id="ae6be-p134">`addItemAttachmentAsync` 方法将包含指定 Exchange 标识符的项目附加到撰写窗体中的项目。如果指定一个回调方法，此方法使用 `asyncResult` 参数调用，该参数包含一个附件标识符或代码，指示附加项目过程中出现的任何错误。可以使用 `options` 参数将状态信息传递给回调方法（如果需要）。</span><span class="sxs-lookup"><span data-stu-id="ae6be-p134">The `addItemAttachmentAsync` method attaches the item with the specified Exchange identifier to the item in the compose form. If you specify a callback method, the method is called with one parameter, `asyncResult`, which contains either the attachment identifier or a code that indicates any error that occurred while attaching the item. You can use the `options` parameter to pass state information to the callback method, if needed.</span></span>

<span data-ttu-id="ae6be-563">随后可以将该标识符与 [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) 方法一同使用，以删除同一个会话中的附件。</span><span class="sxs-lookup"><span data-stu-id="ae6be-563">You can subsequently use the identifier with the [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) method to remove the attachment in the same session.</span></span>

<span data-ttu-id="ae6be-564">如果 Office 加载项在 Outlook Web App 中运行，则 `addItemAttachmentAsync` 方法可以将项目附加到项目（正在编辑的项目除外）中；然而，不支持也不建议这样做。</span><span class="sxs-lookup"><span data-stu-id="ae6be-564">If your Office Add-in is running in Outlook Web App, the `addItemAttachmentAsync` method can attach items to items other than the item that you are editing; however, this is not supported and is not recommended.</span></span>

##### <a name="parameters"></a><span data-ttu-id="ae6be-565">参数：</span><span class="sxs-lookup"><span data-stu-id="ae6be-565">Parameters:</span></span>

|<span data-ttu-id="ae6be-566">名称</span><span class="sxs-lookup"><span data-stu-id="ae6be-566">Name</span></span>| <span data-ttu-id="ae6be-567">类型</span><span class="sxs-lookup"><span data-stu-id="ae6be-567">Type</span></span>| <span data-ttu-id="ae6be-568">属性</span><span class="sxs-lookup"><span data-stu-id="ae6be-568">Attributes</span></span>| <span data-ttu-id="ae6be-569">说明</span><span class="sxs-lookup"><span data-stu-id="ae6be-569">Description</span></span>|
|---|---|---|---|
|`itemId`| <span data-ttu-id="ae6be-570">String</span><span class="sxs-lookup"><span data-stu-id="ae6be-570">String</span></span>||<span data-ttu-id="ae6be-p135">要附加的项目的 Exchange 标识符。最大长度为 100 个字符。</span><span class="sxs-lookup"><span data-stu-id="ae6be-p135">The Exchange identifier of the item to attach. The maximum length is 100 characters.</span></span>|
|`attachmentName`| <span data-ttu-id="ae6be-573">String</span><span class="sxs-lookup"><span data-stu-id="ae6be-573">String</span></span>||<span data-ttu-id="ae6be-p136">要附加的项目的主题。最大长度为 255 个字符。</span><span class="sxs-lookup"><span data-stu-id="ae6be-p136">The sujbect of the item to be attached. The maximum length is 255 characters.</span></span>|
|`options`| <span data-ttu-id="ae6be-576">Object</span><span class="sxs-lookup"><span data-stu-id="ae6be-576">Object</span></span>| <span data-ttu-id="ae6be-577">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="ae6be-577">&lt;optional&gt;</span></span>|<span data-ttu-id="ae6be-578">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="ae6be-578">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="ae6be-579">对象</span><span class="sxs-lookup"><span data-stu-id="ae6be-579">Object</span></span>| <span data-ttu-id="ae6be-580">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="ae6be-580">&lt;optional&gt;</span></span>|<span data-ttu-id="ae6be-581">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="ae6be-581">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="ae6be-582">function</span><span class="sxs-lookup"><span data-stu-id="ae6be-582">function</span></span>| <span data-ttu-id="ae6be-583">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="ae6be-583">&lt;optional&gt;</span></span>|<span data-ttu-id="ae6be-584">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="ae6be-584">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="ae6be-585">如果成功，附件标识符将在 `asyncResult.value` 属性中提供。</span><span class="sxs-lookup"><span data-stu-id="ae6be-585">On success, the attachment identifier will be provided in the `asyncResult.value` property.</span></span><br/><span data-ttu-id="ae6be-586">如果添加附件失败，`asyncResult` 对象将包含一个提供错误说明的 `Error` 对象。</span><span class="sxs-lookup"><span data-stu-id="ae6be-586">If adding the attachment fails, the `asyncResult` object will contain an `Error` object that provides a description of the error.</span></span>|

##### <a name="errors"></a><span data-ttu-id="ae6be-587">错误</span><span class="sxs-lookup"><span data-stu-id="ae6be-587">Errors</span></span>

| <span data-ttu-id="ae6be-588">错误代码</span><span class="sxs-lookup"><span data-stu-id="ae6be-588">Error code</span></span> | <span data-ttu-id="ae6be-589">说明</span><span class="sxs-lookup"><span data-stu-id="ae6be-589">Description</span></span> |
|------------|-------------|
| `NumberOfAttachmentsExceeded` | <span data-ttu-id="ae6be-590">邮件或约会具有的附件过多。</span><span class="sxs-lookup"><span data-stu-id="ae6be-590">The message or appointment has too many attachments.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="ae6be-591">Requirements</span><span class="sxs-lookup"><span data-stu-id="ae6be-591">Requirements</span></span>

|<span data-ttu-id="ae6be-592">要求</span><span class="sxs-lookup"><span data-stu-id="ae6be-592">Requirement</span></span>| <span data-ttu-id="ae6be-593">值</span><span class="sxs-lookup"><span data-stu-id="ae6be-593">Value</span></span>|
|---|---|
|[<span data-ttu-id="ae6be-594">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="ae6be-594">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ae6be-595">1.1</span><span class="sxs-lookup"><span data-stu-id="ae6be-595">1.1</span></span>|
|[<span data-ttu-id="ae6be-596">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="ae6be-596">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ae6be-597">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="ae6be-597">ReadWriteItem</span></span>|
|[<span data-ttu-id="ae6be-598">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="ae6be-598">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ae6be-599">撰写</span><span class="sxs-lookup"><span data-stu-id="ae6be-599">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="ae6be-600">示例</span><span class="sxs-lookup"><span data-stu-id="ae6be-600">Example</span></span>

<span data-ttu-id="ae6be-601">以下示例将现有的 Outlook 项目添加为名为 `My Attachment` 的附件。</span><span class="sxs-lookup"><span data-stu-id="ae6be-601">The following example adds an existing Outlook item as an attachment with the name `My Attachment`.</span></span>

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

#### <a name="displayreplyallformformdata"></a><span data-ttu-id="ae6be-602">displayReplyAllForm(formData)</span><span class="sxs-lookup"><span data-stu-id="ae6be-602">displayReplyAllForm(formData)</span></span>

<span data-ttu-id="ae6be-603">显示答复窗体，其中包括所选邮件的发件人和所有收件人或所选约会的组织者和所有与会者。</span><span class="sxs-lookup"><span data-stu-id="ae6be-603">Displays a reply form that includes the sender and all recipients of the selected message or the organizer and all attendees of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="ae6be-604">在 Outlook for iOS 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="ae6be-604">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="ae6be-605">在 Outlook Web App 中，答复窗体显示为包含 3 列视图的弹出式窗体以及包含 2 列或 1 列视图的弹出式窗体。</span><span class="sxs-lookup"><span data-stu-id="ae6be-605">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="ae6be-606">如果任意字符串参数超出其限制，`displayReplyAllForm` 将引发异常。</span><span class="sxs-lookup"><span data-stu-id="ae6be-606">If any of the string parameters exceed their limits, `displayReplyAllForm` throws an exception.</span></span>

> [!NOTE]
> <span data-ttu-id="ae6be-607">要求集 1.1 不支持 `displayReplyAllForm` 在调用中包括附件的功能。</span><span class="sxs-lookup"><span data-stu-id="ae6be-607">The ability to include attachments in the call to `displayReplyAllForm` is not supported in requirement set 1.1.</span></span> <span data-ttu-id="ae6be-608">附件支持已添加到要求集 1.2 及以上的 `displayReplyAllForm` 中。</span><span class="sxs-lookup"><span data-stu-id="ae6be-608">Attachment support was added to `displayReplyAllForm` in requirement set 1.2 and above.</span></span>

##### <a name="parameters"></a><span data-ttu-id="ae6be-609">参数：</span><span class="sxs-lookup"><span data-stu-id="ae6be-609">Parameters:</span></span>

|<span data-ttu-id="ae6be-610">名称</span><span class="sxs-lookup"><span data-stu-id="ae6be-610">Name</span></span>| <span data-ttu-id="ae6be-611">说明</span><span class="sxs-lookup"><span data-stu-id="ae6be-611">Type</span></span>| <span data-ttu-id="ae6be-612">说明</span><span class="sxs-lookup"><span data-stu-id="ae6be-612">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="ae6be-613">字符串 &#124; 对象</span><span class="sxs-lookup"><span data-stu-id="ae6be-613">String &#124; Object</span></span>| |<span data-ttu-id="ae6be-p138">一个包含文本和 HTML 且表示答复窗体的正文的字符串。字符串限制为 32 KB。</span><span class="sxs-lookup"><span data-stu-id="ae6be-p138">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="ae6be-616">**或**</span><span class="sxs-lookup"><span data-stu-id="ae6be-616">**OR**</span></span><br/><span data-ttu-id="ae6be-p139">包含正文或附件数据和回调函数的对象。对象定义如下。</span><span class="sxs-lookup"><span data-stu-id="ae6be-p139">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="ae6be-619">字符串</span><span class="sxs-lookup"><span data-stu-id="ae6be-619">String</span></span> | <span data-ttu-id="ae6be-620">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="ae6be-620">&lt;optional&gt;</span></span> | <span data-ttu-id="ae6be-p140">一个包含文本和 HTML 且表示答复窗体的正文的字符串。字符串限制为 32 KB。</span><span class="sxs-lookup"><span data-stu-id="ae6be-p140">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `callback` | <span data-ttu-id="ae6be-623">函数</span><span class="sxs-lookup"><span data-stu-id="ae6be-623">function</span></span> | <span data-ttu-id="ae6be-624">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="ae6be-624">&lt;optional&gt;</span></span> | <span data-ttu-id="ae6be-625">方法完成后，使用单个参数 `asyncResult`（一个 [AsyncResult](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="ae6be-625">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="ae6be-626">要求</span><span class="sxs-lookup"><span data-stu-id="ae6be-626">Requirements</span></span>

|<span data-ttu-id="ae6be-627">要求</span><span class="sxs-lookup"><span data-stu-id="ae6be-627">Requirement</span></span>| <span data-ttu-id="ae6be-628">值</span><span class="sxs-lookup"><span data-stu-id="ae6be-628">Value</span></span>|
|---|---|
|[<span data-ttu-id="ae6be-629">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="ae6be-629">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ae6be-630">1.0</span><span class="sxs-lookup"><span data-stu-id="ae6be-630">1.0</span></span>|
|[<span data-ttu-id="ae6be-631">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="ae6be-631">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ae6be-632">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ae6be-632">ReadItem</span></span>|
|[<span data-ttu-id="ae6be-633">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="ae6be-633">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ae6be-634">阅读</span><span class="sxs-lookup"><span data-stu-id="ae6be-634">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="ae6be-635">示例</span><span class="sxs-lookup"><span data-stu-id="ae6be-635">Examples</span></span>

<span data-ttu-id="ae6be-636">以下代码将一个字符串传递到 `displayReplyAllForm` 函数。</span><span class="sxs-lookup"><span data-stu-id="ae6be-636">The following code passes a string to the `displayReplyAllForm` function.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

<span data-ttu-id="ae6be-637">使用空白正文答复。</span><span class="sxs-lookup"><span data-stu-id="ae6be-637">Reply with an empty body.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyAllForm({});
```

<span data-ttu-id="ae6be-638">仅使用正文答复。</span><span class="sxs-lookup"><span data-stu-id="ae6be-638">Reply with just a body.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="ae6be-639">使用正文和回调答复。</span><span class="sxs-lookup"><span data-stu-id="ae6be-639">Reply with a body and a callback.</span></span>

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

#### <a name="displayreplyformformdata"></a><span data-ttu-id="ae6be-640">displayReplyForm(formData)</span><span class="sxs-lookup"><span data-stu-id="ae6be-640">displayReplyForm(formData)</span></span>

<span data-ttu-id="ae6be-641">显示答复窗体，其中仅包括所选邮件的发件人或所选约会的组织者。</span><span class="sxs-lookup"><span data-stu-id="ae6be-641">Displays a reply form that includes only the sender of the selected message or the organizer of the selected appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="ae6be-642">在 Outlook for iOS 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="ae6be-642">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="ae6be-643">在 Outlook Web App 中，答复窗体显示为包含 3 列视图的弹出式窗体以及包含 2 列或 1 列视图的弹出式窗体。</span><span class="sxs-lookup"><span data-stu-id="ae6be-643">In Outlook Web App, the reply form is displayed as a pop-out form in the 3-column view and a pop-up form in the 2- or 1-column view.</span></span>

<span data-ttu-id="ae6be-644">如果任意字符串参数超出其限制，`displayReplyForm` 将引发异常。</span><span class="sxs-lookup"><span data-stu-id="ae6be-644">If any of the string parameters exceed their limits, `displayReplyForm` throws an exception.</span></span>

> [!NOTE]
> <span data-ttu-id="ae6be-645">要求集 1.1 不支持 `displayReplyForm` 在调用中包括附件的功能。</span><span class="sxs-lookup"><span data-stu-id="ae6be-645">The ability to include attachments in the call to `displayReplyForm` is not supported in requirement set 1.1.</span></span> <span data-ttu-id="ae6be-646">附件支持已添加到要求集 1.2 及以上的 `displayReplyForm` 中。</span><span class="sxs-lookup"><span data-stu-id="ae6be-646">Attachment support was added to `displayReplyForm` in requirement set 1.2 and above.</span></span>

##### <a name="parameters"></a><span data-ttu-id="ae6be-647">参数：</span><span class="sxs-lookup"><span data-stu-id="ae6be-647">Parameters:</span></span>

|<span data-ttu-id="ae6be-648">名称</span><span class="sxs-lookup"><span data-stu-id="ae6be-648">Name</span></span>| <span data-ttu-id="ae6be-649">说明</span><span class="sxs-lookup"><span data-stu-id="ae6be-649">Type</span></span>| <span data-ttu-id="ae6be-650">说明</span><span class="sxs-lookup"><span data-stu-id="ae6be-650">Description</span></span>|
|---|---|---|
|`formData`| <span data-ttu-id="ae6be-651">字符串 &#124; 对象</span><span class="sxs-lookup"><span data-stu-id="ae6be-651">String &#124; Object</span></span>| | <span data-ttu-id="ae6be-p142">一个包含文本和 HTML 且表示答复窗体的正文的字符串。字符串限制为 32 KB。</span><span class="sxs-lookup"><span data-stu-id="ae6be-p142">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span><br/><span data-ttu-id="ae6be-654">**或**</span><span class="sxs-lookup"><span data-stu-id="ae6be-654">**OR**</span></span><br/><span data-ttu-id="ae6be-p143">包含正文或附件数据和回调函数的对象。对象定义如下。</span><span class="sxs-lookup"><span data-stu-id="ae6be-p143">An object that contains body or attachment data and a callback function. The object is defined as follows.</span></span> |
| `formData.htmlBody` | <span data-ttu-id="ae6be-657">字符串</span><span class="sxs-lookup"><span data-stu-id="ae6be-657">String</span></span> | <span data-ttu-id="ae6be-658">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="ae6be-658">&lt;optional&gt;</span></span> | <span data-ttu-id="ae6be-p144">一个包含文本和 HTML 且表示答复窗体的正文的字符串。字符串限制为 32 KB。</span><span class="sxs-lookup"><span data-stu-id="ae6be-p144">A string that contains text and HTML and that represents the body of the reply form. The string is limited to 32 KB.</span></span>
| `callback` | <span data-ttu-id="ae6be-661">函数</span><span class="sxs-lookup"><span data-stu-id="ae6be-661">function</span></span> | <span data-ttu-id="ae6be-662">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="ae6be-662">&lt;optional&gt;</span></span> | <span data-ttu-id="ae6be-663">方法完成后，使用单个参数 `asyncResult`（一个 [AsyncResult](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="ae6be-663">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [AsyncResult](/javascript/api/office/office.asyncresult) object.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="ae6be-664">要求</span><span class="sxs-lookup"><span data-stu-id="ae6be-664">Requirements</span></span>

|<span data-ttu-id="ae6be-665">要求</span><span class="sxs-lookup"><span data-stu-id="ae6be-665">Requirement</span></span>| <span data-ttu-id="ae6be-666">值</span><span class="sxs-lookup"><span data-stu-id="ae6be-666">Value</span></span>|
|---|---|
|[<span data-ttu-id="ae6be-667">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="ae6be-667">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ae6be-668">1.0</span><span class="sxs-lookup"><span data-stu-id="ae6be-668">1.0</span></span>|
|[<span data-ttu-id="ae6be-669">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="ae6be-669">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ae6be-670">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ae6be-670">ReadItem</span></span>|
|[<span data-ttu-id="ae6be-671">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="ae6be-671">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ae6be-672">阅读</span><span class="sxs-lookup"><span data-stu-id="ae6be-672">Read</span></span>|

##### <a name="examples"></a><span data-ttu-id="ae6be-673">示例</span><span class="sxs-lookup"><span data-stu-id="ae6be-673">Examples</span></span>

<span data-ttu-id="ae6be-674">以下代码将一个字符串传递到 `displayReplyForm` 函数。</span><span class="sxs-lookup"><span data-stu-id="ae6be-674">The following code passes a string to the `displayReplyForm` function.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

<span data-ttu-id="ae6be-675">使用空白正文答复。</span><span class="sxs-lookup"><span data-stu-id="ae6be-675">Reply with an empty body.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyForm({});
```

<span data-ttu-id="ae6be-676">仅使用正文答复。</span><span class="sxs-lookup"><span data-stu-id="ae6be-676">Reply with just a body.</span></span>

```JavaScript
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

<span data-ttu-id="ae6be-677">使用正文和回调答复。</span><span class="sxs-lookup"><span data-stu-id="ae6be-677">Reply with a body and a callback.</span></span>

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

#### <a name="getentities--entitiesjavascriptapioutlook11officeentities"></a><span data-ttu-id="ae6be-678">getEntities() → {[Entities](/javascript/api/outlook_1_1/office.entities)}</span><span class="sxs-lookup"><span data-stu-id="ae6be-678">getEntities() → {[Entities](/javascript/api/outlook_1_1/office.entities)}</span></span>

<span data-ttu-id="ae6be-679">获取在所选项目的正文中找到的实体。</span><span class="sxs-lookup"><span data-stu-id="ae6be-679">Gets the entities found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="ae6be-680">在 Outlook for iOS 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="ae6be-680">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="requirements"></a><span data-ttu-id="ae6be-681">Requirements</span><span class="sxs-lookup"><span data-stu-id="ae6be-681">Requirements</span></span>

|<span data-ttu-id="ae6be-682">要求</span><span class="sxs-lookup"><span data-stu-id="ae6be-682">Requirement</span></span>| <span data-ttu-id="ae6be-683">值</span><span class="sxs-lookup"><span data-stu-id="ae6be-683">Value</span></span>|
|---|---|
|[<span data-ttu-id="ae6be-684">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="ae6be-684">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ae6be-685">1.0</span><span class="sxs-lookup"><span data-stu-id="ae6be-685">1.0</span></span>|
|[<span data-ttu-id="ae6be-686">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="ae6be-686">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ae6be-687">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ae6be-687">ReadItem</span></span>|
|[<span data-ttu-id="ae6be-688">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="ae6be-688">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ae6be-689">阅读</span><span class="sxs-lookup"><span data-stu-id="ae6be-689">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="ae6be-690">返回：</span><span class="sxs-lookup"><span data-stu-id="ae6be-690">Returns:</span></span>

<span data-ttu-id="ae6be-691">类型：[Entities](/javascript/api/outlook_1_1/office.entities)</span><span class="sxs-lookup"><span data-stu-id="ae6be-691">Type: [Entities](/javascript/api/outlook_1_1/office.entities)</span></span>

##### <a name="example"></a><span data-ttu-id="ae6be-692">示例</span><span class="sxs-lookup"><span data-stu-id="ae6be-692">Example</span></span>

<span data-ttu-id="ae6be-693">以下示例访问当前项目的正文中的联系人实体。</span><span class="sxs-lookup"><span data-stu-id="ae6be-693">The following example accesses the contacts entities in the current item's body.</span></span>

```JavaScript
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlook11officecontactmeetingsuggestionjavascriptapioutlook11officemeetingsuggestionphonenumberjavascriptapioutlook11officephonenumbertasksuggestionjavascriptapioutlook11officetasksuggestion"></a><span data-ttu-id="ae6be-694">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_1/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_1/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_1/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_1/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="ae6be-694">getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_1/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_1/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_1/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_1/office.tasksuggestion))>}</span></span>

<span data-ttu-id="ae6be-695">获取所选项目的正文中找到的指定实体类型的所有实体的数组。</span><span class="sxs-lookup"><span data-stu-id="ae6be-695">Gets an array of all the entities of the specified entity type found in the selected item's body.</span></span>

> [!NOTE]
> <span data-ttu-id="ae6be-696">在 Outlook for iOS 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="ae6be-696">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

##### <a name="parameters"></a><span data-ttu-id="ae6be-697">参数：</span><span class="sxs-lookup"><span data-stu-id="ae6be-697">Parameters:</span></span>

|<span data-ttu-id="ae6be-698">名称</span><span class="sxs-lookup"><span data-stu-id="ae6be-698">Name</span></span>| <span data-ttu-id="ae6be-699">类型</span><span class="sxs-lookup"><span data-stu-id="ae6be-699">Type</span></span>| <span data-ttu-id="ae6be-700">说明</span><span class="sxs-lookup"><span data-stu-id="ae6be-700">Description</span></span>|
|---|---|---|
|`entityType`| [<span data-ttu-id="ae6be-701">Office.MailboxEnums.EntityType</span><span class="sxs-lookup"><span data-stu-id="ae6be-701">Office.MailboxEnums.EntityType</span></span>](/javascript/api/outlook_1_1/office.MailboxEnums.entitytype)|<span data-ttu-id="ae6be-702">EntityType 枚举值之一。</span><span class="sxs-lookup"><span data-stu-id="ae6be-702">One of the EntityType enumeration values.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="ae6be-703">Requirements</span><span class="sxs-lookup"><span data-stu-id="ae6be-703">Requirements</span></span>

|<span data-ttu-id="ae6be-704">要求</span><span class="sxs-lookup"><span data-stu-id="ae6be-704">Requirement</span></span>| <span data-ttu-id="ae6be-705">值</span><span class="sxs-lookup"><span data-stu-id="ae6be-705">Value</span></span>|
|---|---|
|[<span data-ttu-id="ae6be-706">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="ae6be-706">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ae6be-707">1.0</span><span class="sxs-lookup"><span data-stu-id="ae6be-707">1.0</span></span>|
|[<span data-ttu-id="ae6be-708">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="ae6be-708">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ae6be-709">受限</span><span class="sxs-lookup"><span data-stu-id="ae6be-709">Restricted</span></span>|
|[<span data-ttu-id="ae6be-710">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="ae6be-710">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ae6be-711">阅读</span><span class="sxs-lookup"><span data-stu-id="ae6be-711">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="ae6be-712">返回：</span><span class="sxs-lookup"><span data-stu-id="ae6be-712">Returns:</span></span>

<span data-ttu-id="ae6be-713">如果在 `entityType` 中传递的值不是 `EntityType` 枚举的有效成员，该方法返回 null。</span><span class="sxs-lookup"><span data-stu-id="ae6be-713">If the value passed in `entityType` is not a valid member of the `EntityType` enumeration, the method returns null.</span></span> <span data-ttu-id="ae6be-714">如果指定类型的任何实体都不存在于该项目的正文中，该方法将返回空数组。</span><span class="sxs-lookup"><span data-stu-id="ae6be-714">If no entities of the specified type are present in the item's body, the method returns an empty array.</span></span> <span data-ttu-id="ae6be-715">否则，返回的数组中对象的类型取决于 `entityType` 参数中请求实体的类型。</span><span class="sxs-lookup"><span data-stu-id="ae6be-715">Otherwise, the type of the objects in the returned array depends on the type of entity requested in the `entityType` parameter.</span></span>

<span data-ttu-id="ae6be-716">当使用此方法的最低权限级别**受限**时，某些实体类型需要 **ReadItem** 才能进行访问，如下表中所指定。</span><span class="sxs-lookup"><span data-stu-id="ae6be-716">While the minimum permission level to use this method is **Restricted**, some entity types require **ReadItem** to access, as specified in the following table.</span></span>

| <span data-ttu-id="ae6be-717">`entityType` 的值</span><span class="sxs-lookup"><span data-stu-id="ae6be-717">Value of `entityType`</span></span> | <span data-ttu-id="ae6be-718">返回的数组中对象的类型</span><span class="sxs-lookup"><span data-stu-id="ae6be-718">Type of objects in returned array</span></span> | <span data-ttu-id="ae6be-719">所需权限级别</span><span class="sxs-lookup"><span data-stu-id="ae6be-719">Required Permission Level</span></span> |
| --- | --- | --- |
| `Address` | <span data-ttu-id="ae6be-720">字符串</span><span class="sxs-lookup"><span data-stu-id="ae6be-720">String</span></span> | <span data-ttu-id="ae6be-721">**受限**</span><span class="sxs-lookup"><span data-stu-id="ae6be-721">**Restricted**</span></span> |
| `Contact` | <span data-ttu-id="ae6be-722">Contact</span><span class="sxs-lookup"><span data-stu-id="ae6be-722">Contact</span></span> | <span data-ttu-id="ae6be-723">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="ae6be-723">**ReadItem**</span></span> |
| `EmailAddress` | <span data-ttu-id="ae6be-724">String</span><span class="sxs-lookup"><span data-stu-id="ae6be-724">String</span></span> | <span data-ttu-id="ae6be-725">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="ae6be-725">**ReadItem**</span></span> |
| `MeetingSuggestion` | <span data-ttu-id="ae6be-726">MeetingSuggestion</span><span class="sxs-lookup"><span data-stu-id="ae6be-726">MeetingSuggestion</span></span> | <span data-ttu-id="ae6be-727">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="ae6be-727">**ReadItem**</span></span> |
| `PhoneNumber` | <span data-ttu-id="ae6be-728">PhoneNumber</span><span class="sxs-lookup"><span data-stu-id="ae6be-728">PhoneNumber</span></span> | <span data-ttu-id="ae6be-729">**受限**</span><span class="sxs-lookup"><span data-stu-id="ae6be-729">**Restricted**</span></span> |
| `TaskSuggestion` | <span data-ttu-id="ae6be-730">TaskSuggestion</span><span class="sxs-lookup"><span data-stu-id="ae6be-730">TaskSuggestion</span></span> | <span data-ttu-id="ae6be-731">**ReadItem**</span><span class="sxs-lookup"><span data-stu-id="ae6be-731">**ReadItem**</span></span> |
| `URL` | <span data-ttu-id="ae6be-732">String</span><span class="sxs-lookup"><span data-stu-id="ae6be-732">String</span></span> | <span data-ttu-id="ae6be-733">**受限**</span><span class="sxs-lookup"><span data-stu-id="ae6be-733">**Restricted**</span></span> |

<span data-ttu-id="ae6be-734">类型：Array.<(String|[Contact](/javascript/api/outlook_1_1/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_1/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_1/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_1/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="ae6be-734">Type:  Array.<(String|[Contact](/javascript/api/outlook_1_1/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_1/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_1/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_1/office.tasksuggestion))></span></span>


##### <a name="example"></a><span data-ttu-id="ae6be-735">示例</span><span class="sxs-lookup"><span data-stu-id="ae6be-735">Example</span></span>

<span data-ttu-id="ae6be-736">以下示例显示了如何访问表示当前项目的正文中的邮政地址的字符串数组。</span><span class="sxs-lookup"><span data-stu-id="ae6be-736">The following example shows how to access an array of strings that represent postal addresses in the current item's body.</span></span>

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

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlook11officecontactmeetingsuggestionjavascriptapioutlook11officemeetingsuggestionphonenumberjavascriptapioutlook11officephonenumbertasksuggestionjavascriptapioutlook11officetasksuggestion"></a><span data-ttu-id="ae6be-737">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_1/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_1/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_1/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_1/office.tasksuggestion))>}</span><span class="sxs-lookup"><span data-stu-id="ae6be-737">getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook_1_1/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_1/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_1/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_1/office.tasksuggestion))>}</span></span>

<span data-ttu-id="ae6be-738">返回传递清单 XML 文件中定义的命名筛选器的所选项目中的已知实体。</span><span class="sxs-lookup"><span data-stu-id="ae6be-738">Returns well-known entities in the selected item that pass the named filter defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="ae6be-739">在 Outlook for iOS 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="ae6be-739">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="ae6be-740">`getFilteredEntitiesByName` 方法返回匹配在具有指定 `FilterName` 元素值的清单 XML 文件中的 [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) 规则元素中定义的正则表达式的实体。</span><span class="sxs-lookup"><span data-stu-id="ae6be-740">The `getFilteredEntitiesByName` method returns the entities that match the regular expression defined in the [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) rule element in the manifest XML file with the specified `FilterName` element value.</span></span>

##### <a name="parameters"></a><span data-ttu-id="ae6be-741">参数：</span><span class="sxs-lookup"><span data-stu-id="ae6be-741">Parameters:</span></span>

|<span data-ttu-id="ae6be-742">名称</span><span class="sxs-lookup"><span data-stu-id="ae6be-742">Name</span></span>| <span data-ttu-id="ae6be-743">类型</span><span class="sxs-lookup"><span data-stu-id="ae6be-743">Type</span></span>| <span data-ttu-id="ae6be-744">说明</span><span class="sxs-lookup"><span data-stu-id="ae6be-744">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="ae6be-745">字符串</span><span class="sxs-lookup"><span data-stu-id="ae6be-745">String</span></span>|<span data-ttu-id="ae6be-746">定义筛选器匹配的 `ItemHasKnownEntity` 规则元素的名称。</span><span class="sxs-lookup"><span data-stu-id="ae6be-746">The name of the `ItemHasKnownEntity` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="ae6be-747">Requirements</span><span class="sxs-lookup"><span data-stu-id="ae6be-747">Requirements</span></span>

|<span data-ttu-id="ae6be-748">要求</span><span class="sxs-lookup"><span data-stu-id="ae6be-748">Requirement</span></span>| <span data-ttu-id="ae6be-749">值</span><span class="sxs-lookup"><span data-stu-id="ae6be-749">Value</span></span>|
|---|---|
|[<span data-ttu-id="ae6be-750">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="ae6be-750">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ae6be-751">1.0</span><span class="sxs-lookup"><span data-stu-id="ae6be-751">1.0</span></span>|
|[<span data-ttu-id="ae6be-752">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="ae6be-752">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ae6be-753">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ae6be-753">ReadItem</span></span>|
|[<span data-ttu-id="ae6be-754">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="ae6be-754">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ae6be-755">阅读</span><span class="sxs-lookup"><span data-stu-id="ae6be-755">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="ae6be-756">返回：</span><span class="sxs-lookup"><span data-stu-id="ae6be-756">Returns:</span></span>

<span data-ttu-id="ae6be-p146">如果具有匹配 `name` 参数的 `FilterName` 元素值的清单中没有任何 `ItemHasKnownEntity` 元素，则该方法返回 `null`。如果 `name` 参数匹配清单中的 `ItemHasKnownEntity` 元素，但在匹配的当前项目中没有实体，则该方法返回一个空数组。</span><span class="sxs-lookup"><span data-stu-id="ae6be-p146">If there is no `ItemHasKnownEntity` element in the manifest with a `FilterName` element value that matches the `name` parameter, the method returns `null`. If the `name` parameter does match an `ItemHasKnownEntity` element in the manifest, but there are no entities in the current item that match, the method return an empty array.</span></span>


<span data-ttu-id="ae6be-759">类型：Array.<(String|[Contact](/javascript/api/outlook_1_1/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_1/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_1/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_1/office.tasksuggestion))></span><span class="sxs-lookup"><span data-stu-id="ae6be-759">Type: Array.<(String|[Contact](/javascript/api/outlook_1_1/office.contact)|[MeetingSuggestion](/javascript/api/outlook_1_1/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook_1_1/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook_1_1/office.tasksuggestion))></span></span>


#### <a name="getregexmatches--object"></a><span data-ttu-id="ae6be-760">getRegExMatches() → {Object}</span><span class="sxs-lookup"><span data-stu-id="ae6be-760">getRegExMatches() → {Object}</span></span>

<span data-ttu-id="ae6be-761">返回所选项目中匹配在清单 XML 文件中定义的正则表达式的字符串值。</span><span class="sxs-lookup"><span data-stu-id="ae6be-761">Returns string values in the selected item that match the regular expressions defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="ae6be-762">在 Outlook for iOS 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="ae6be-762">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="ae6be-p147">`getRegExMatches` 方法返回匹配在清单 XML 文件中的每个 `ItemHasRegularExpressionMatch` 或 `ItemHasKnownEntity` 规则元素中定义的正则表达式的字符串。对于 `ItemHasRegularExpressionMatch` 规则，匹配字符串必须发生在该规则指定的项目的属性中。`PropertyName` 简单类型定义支持的属性。</span><span class="sxs-lookup"><span data-stu-id="ae6be-p147">The `getRegExMatches` method returns the strings that match the regular expression defined in each `ItemHasRegularExpressionMatch` or `ItemHasKnownEntity` rule element in the manifest XML file. For an `ItemHasRegularExpressionMatch` rule, a matching string has to occur in the property of the item that is specified by that rule. The `PropertyName` simple type defines the supported properties.</span></span>

<span data-ttu-id="ae6be-766">例如，考虑一个外接程序清单具有以下 `Rule` 元素：</span><span class="sxs-lookup"><span data-stu-id="ae6be-766">For example, consider an add-in manifest has the following `Rule` element:</span></span>

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

<span data-ttu-id="ae6be-767">从 `getRegExMatches` 返回的对象应有两个属性：`fruits` 和 `veggies`。</span><span class="sxs-lookup"><span data-stu-id="ae6be-767">The object returned from `getRegExMatches` would have two properties: `fruits` and `veggies`.</span></span>

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

> [!NOTE]
> <span data-ttu-id="ae6be-p148">如果在项目的正文属性上指定 `ItemHasRegularExpressionMatch` 规则，则正则表达式应进一步筛选正文并应尝试返回该项目的整个正文。使用正则表达式（如 `.*`）获取项目的整个正文并不总是返回预期的结果。</span><span class="sxs-lookup"><span data-stu-id="ae6be-p148">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="requirements"></a><span data-ttu-id="ae6be-770">要求</span><span class="sxs-lookup"><span data-stu-id="ae6be-770">Requirements</span></span>

|<span data-ttu-id="ae6be-771">要求</span><span class="sxs-lookup"><span data-stu-id="ae6be-771">Requirement</span></span>| <span data-ttu-id="ae6be-772">值</span><span class="sxs-lookup"><span data-stu-id="ae6be-772">Value</span></span>|
|---|---|
|[<span data-ttu-id="ae6be-773">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="ae6be-773">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ae6be-774">1.0</span><span class="sxs-lookup"><span data-stu-id="ae6be-774">1.0</span></span>|
|[<span data-ttu-id="ae6be-775">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="ae6be-775">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ae6be-776">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ae6be-776">ReadItem</span></span>|
|[<span data-ttu-id="ae6be-777">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="ae6be-777">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ae6be-778">阅读</span><span class="sxs-lookup"><span data-stu-id="ae6be-778">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="ae6be-779">返回：</span><span class="sxs-lookup"><span data-stu-id="ae6be-779">Returns:</span></span>

<span data-ttu-id="ae6be-p149">一个包含匹配在清单 XML 文件中定义的正则表达式的字符串数组的对象。每个数组的名称等于匹配 `ItemHasRegularExpressionMatch` 规则的 `RegExName` 属性或匹配 `ItemHasKnownEntity` 规则的 `FilterName` 属性的相应值。</span><span class="sxs-lookup"><span data-stu-id="ae6be-p149">An object that contains arrays of strings that match the regular expressions defined in the manifest XML file. The name of each array is equal to the corresponding value of the `RegExName` attribute of the matching `ItemHasRegularExpressionMatch` rule or the `FilterName` attribute of the matching `ItemHasKnownEntity` rule.</span></span>

<dl class="param-type"><span data-ttu-id="ae6be-782">

<dt>
类型</dt>


</span><span class="sxs-lookup"><span data-stu-id="ae6be-782">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="ae6be-783">对象</span><span class="sxs-lookup"><span data-stu-id="ae6be-783">Object</span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="ae6be-784">示例</span><span class="sxs-lookup"><span data-stu-id="ae6be-784">Example</span></span>

<span data-ttu-id="ae6be-785">以下示例显示了如何访问正则表达式 <rule> 元素 `fruits` 和 `veggies` 的匹配项的数组，这些元素在清单中指定。</rule></span><span class="sxs-lookup"><span data-stu-id="ae6be-785">The following example shows how to access the array of matches for the regular expression <rule>elements `fruits` and `veggies`, which are specified in the manifest.</rule></span></span>

```JavaScript
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veges = allMatches.veggies;
```

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a><span data-ttu-id="ae6be-786">getRegExMatchesByName(name) → (nullable) {Array.<String>}</span><span class="sxs-lookup"><span data-stu-id="ae6be-786">getRegExMatchesByName(name) → (nullable) {Array.< String >}</span></span>

<span data-ttu-id="ae6be-787">返回匹配在清单 XML 文件中定义的命名正则表达式的所选项目中的字符串值。</span><span class="sxs-lookup"><span data-stu-id="ae6be-787">Returns string values in the selected item that match the named regular expression defined in the manifest XML file.</span></span>

> [!NOTE]
> <span data-ttu-id="ae6be-788">在 Outlook for iOS 或 Outlook for Android 中不支持此方法。</span><span class="sxs-lookup"><span data-stu-id="ae6be-788">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="ae6be-789">`getRegExMatchesByName` 方法返回匹配在具有指定 `RegExName` 元素值的清单 XML 文件中的 `ItemHasRegularExpressionMatch` 规则元素中定义的正则表达式的字符串。</span><span class="sxs-lookup"><span data-stu-id="ae6be-789">The `getRegExMatchesByName` method returns the strings that match the regular expression defined in the `ItemHasRegularExpressionMatch` rule element in the manifest XML file with the specified `RegExName` element value.</span></span>

<span data-ttu-id="ae6be-p150">如果在项目的正文属性上指定 `ItemHasRegularExpressionMatch` 规则，则正则表达式应进一步筛选正文，不应尝试返回该项目的整个正文。使用正则表达式（如 `.*`）获取项目的整个正文并不总是返回预期的结果。</span><span class="sxs-lookup"><span data-stu-id="ae6be-p150">If you specify an `ItemHasRegularExpressionMatch` rule on the body property of an item, the regular expression should further filter the body and should not attempt to return the entire body of the item. Using a regular expression such as `.*` to obtain the entire body of an item does not always return the expected results.</span></span>

##### <a name="parameters"></a><span data-ttu-id="ae6be-792">参数：</span><span class="sxs-lookup"><span data-stu-id="ae6be-792">Parameters:</span></span>

|<span data-ttu-id="ae6be-793">名称</span><span class="sxs-lookup"><span data-stu-id="ae6be-793">Name</span></span>| <span data-ttu-id="ae6be-794">说明</span><span class="sxs-lookup"><span data-stu-id="ae6be-794">Type</span></span>| <span data-ttu-id="ae6be-795">描述</span><span class="sxs-lookup"><span data-stu-id="ae6be-795">Description</span></span>|
|---|---|---|
|`name`| <span data-ttu-id="ae6be-796">字符串</span><span class="sxs-lookup"><span data-stu-id="ae6be-796">String</span></span>|<span data-ttu-id="ae6be-797">定义筛选器匹配的 `ItemHasRegularExpressionMatch` 规则元素的名称。</span><span class="sxs-lookup"><span data-stu-id="ae6be-797">The name of the `ItemHasRegularExpressionMatch` rule element that defines the filter to match.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="ae6be-798">Requirements</span><span class="sxs-lookup"><span data-stu-id="ae6be-798">Requirements</span></span>

|<span data-ttu-id="ae6be-799">要求</span><span class="sxs-lookup"><span data-stu-id="ae6be-799">Requirement</span></span>| <span data-ttu-id="ae6be-800">值</span><span class="sxs-lookup"><span data-stu-id="ae6be-800">Value</span></span>|
|---|---|
|[<span data-ttu-id="ae6be-801">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="ae6be-801">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ae6be-802">1.0</span><span class="sxs-lookup"><span data-stu-id="ae6be-802">1.0</span></span>|
|[<span data-ttu-id="ae6be-803">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="ae6be-803">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ae6be-804">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ae6be-804">ReadItem</span></span>|
|[<span data-ttu-id="ae6be-805">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="ae6be-805">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ae6be-806">阅读</span><span class="sxs-lookup"><span data-stu-id="ae6be-806">Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="ae6be-807">返回：</span><span class="sxs-lookup"><span data-stu-id="ae6be-807">Returns:</span></span>

<span data-ttu-id="ae6be-808">一个包含匹配在清单 XML 文件中定义的正则表达式的字符串的数组。</span><span class="sxs-lookup"><span data-stu-id="ae6be-808">An array that contains the strings that match the regular expression defined in the manifest XML file.</span></span>

<dl class="param-type"><span data-ttu-id="ae6be-809">

<dt>类型</dt>

</span><span class="sxs-lookup"><span data-stu-id="ae6be-809">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="ae6be-810">Array.< String ></span><span class="sxs-lookup"><span data-stu-id="ae6be-810">Array.< String ></span></span></dd>

</dl>

##### <a name="example"></a><span data-ttu-id="ae6be-811">示例</span><span class="sxs-lookup"><span data-stu-id="ae6be-811">Example</span></span>

```JavaScript
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

####  <a name="loadcustompropertiesasynccallback-usercontext"></a><span data-ttu-id="ae6be-812">loadCustomPropertiesAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="ae6be-812">loadCustomPropertiesAsync(callback, [userContext])</span></span>

<span data-ttu-id="ae6be-813">异步加载所选项目上此外接程序的自定义属性。</span><span class="sxs-lookup"><span data-stu-id="ae6be-813">Asynchronously loads custom properties for this add-in on the selected item.</span></span>

<span data-ttu-id="ae6be-p151">自定义属性基于每个应用、每个项目存储为键/值对。此方法在回调中返回 `CustomProperties` 对象，该回调提供访问特定于当前项目和当前外接程序的自定义属性的方法。自定义属性未在项目上加密，因此这不应用作安全存储。</span><span class="sxs-lookup"><span data-stu-id="ae6be-p151">Custom properties are stored as key/value pairs on a per-app, per-item basis. This method returns a `CustomProperties` object in the callback, which provides methods to access the custom properties specific to the current item and the current add-in. Custom properties are not encrypted on the item, so this should not be used as secure storage.</span></span>

##### <a name="parameters"></a><span data-ttu-id="ae6be-817">参数：</span><span class="sxs-lookup"><span data-stu-id="ae6be-817">Parameters:</span></span>

|<span data-ttu-id="ae6be-818">名称</span><span class="sxs-lookup"><span data-stu-id="ae6be-818">Name</span></span>| <span data-ttu-id="ae6be-819">类型</span><span class="sxs-lookup"><span data-stu-id="ae6be-819">Type</span></span>| <span data-ttu-id="ae6be-820">属性</span><span class="sxs-lookup"><span data-stu-id="ae6be-820">Attributes</span></span>| <span data-ttu-id="ae6be-821">说明</span><span class="sxs-lookup"><span data-stu-id="ae6be-821">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="ae6be-822">函数</span><span class="sxs-lookup"><span data-stu-id="ae6be-822">function</span></span>||<span data-ttu-id="ae6be-823">方法完成后，使用单个参数 `asyncResult`（一个 [`AsyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `callback` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="ae6be-823">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="ae6be-824">自定义属性作为 `asyncResult.value` 属性中的 [`CustomProperties`](/javascript/api/outlook_1_1/office.customproperties) 对象提供。</span><span class="sxs-lookup"><span data-stu-id="ae6be-824">The custom properties are provided as a [`CustomProperties`](/javascript/api/outlook_1_1/office.customproperties) object in the `asyncResult.value` property.</span></span> <span data-ttu-id="ae6be-825">此对象可用于获取、设置以及从项目中删除自定义属性，并将自定义属性集的更改重新保存到服务器。</span><span class="sxs-lookup"><span data-stu-id="ae6be-825">This object can be used to get, set, and remove custom properties from the item and save changes to the custom property set back to the server.</span></span>|
|`userContext`| <span data-ttu-id="ae6be-826">对象</span><span class="sxs-lookup"><span data-stu-id="ae6be-826">Object</span></span>| <span data-ttu-id="ae6be-827">&lt;可选&gt;</span><span class="sxs-lookup"><span data-stu-id="ae6be-827">&lt;optional&gt;</span></span>|<span data-ttu-id="ae6be-828">开发人员可以提供他们想要在回调函数中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="ae6be-828">Developers can provide any object they wish to access in the callback function.</span></span> <span data-ttu-id="ae6be-829">此对象可以通过回调函数中的 `asyncResult.asyncContext` 属性进行访问。</span><span class="sxs-lookup"><span data-stu-id="ae6be-829">This object can be accessed by the `asyncResult.asyncContext` property in the callback function.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="ae6be-830">要求</span><span class="sxs-lookup"><span data-stu-id="ae6be-830">Requirements</span></span>

|<span data-ttu-id="ae6be-831">要求</span><span class="sxs-lookup"><span data-stu-id="ae6be-831">Requirement</span></span>| <span data-ttu-id="ae6be-832">值</span><span class="sxs-lookup"><span data-stu-id="ae6be-832">Value</span></span>|
|---|---|
|[<span data-ttu-id="ae6be-833">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="ae6be-833">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ae6be-834">1.0</span><span class="sxs-lookup"><span data-stu-id="ae6be-834">1.0</span></span>|
|[<span data-ttu-id="ae6be-835">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="ae6be-835">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ae6be-836">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ae6be-836">ReadItem</span></span>|
|[<span data-ttu-id="ae6be-837">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="ae6be-837">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ae6be-838">撰写或阅读</span><span class="sxs-lookup"><span data-stu-id="ae6be-838">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="ae6be-839">示例</span><span class="sxs-lookup"><span data-stu-id="ae6be-839">Example</span></span>

<span data-ttu-id="ae6be-p154">以下代码示例显示了如何使用 `loadCustomPropertiesAsync` 方法异步加载特定于当前项目的自定义属性。该示例还显示了如何使用 `CustomProperties.saveAsync` 方法将这些属性重新保存到服务器。加载自定义属性后，该代码示例将使用 `CustomProperties.get` 方法读取自定义属性 `myProp`，使用 `CustomProperties.set` 方法写入自定义属性 `otherProp`，最后调用 `saveAsync` 方法保存这些自定义属性。</span><span class="sxs-lookup"><span data-stu-id="ae6be-p154">The following code example shows how to use the `loadCustomPropertiesAsync` method to asynchronously load custom properties that are specific to the current item. The example also shows how to use the `CustomProperties.saveAsync` method to save these properties back to the server. After loading the custom properties, the code sample uses the `CustomProperties.get` method to read the custom property `myProp`, the `CustomProperties.set` method to write the custom property `otherProp`, and then finally calls the `saveAsync` method to save the custom properties.</span></span>

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

####  <a name="removeattachmentasyncattachmentid-options-callback"></a><span data-ttu-id="ae6be-843">removeAttachmentAsync(attachmentId, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="ae6be-843">removeAttachmentAsync(attachmentId, [options], [callback])</span></span>

<span data-ttu-id="ae6be-844">将附件从邮件或约会中删除。</span><span class="sxs-lookup"><span data-stu-id="ae6be-844">Removes an attachment from a message or appointment.</span></span>

<span data-ttu-id="ae6be-p155">`removeAttachmentAsync` 方法删除项目中带指定标识符的附件。最佳做法是，仅当同一个邮件应用程序在同一会话中添加了一个附件时，你才应使用该附件标识符来删除该附件。在 Outlook Web App 和适用于设备的 OWA 中，附件标识符只在同一个会话中才有效。当用户关闭应用程序，或者如果用户开始在内嵌窗体中撰写，并在随后弹出的内嵌窗体中继续在单独的窗口撰写时，会话即结束。</span><span class="sxs-lookup"><span data-stu-id="ae6be-p155">The `removeAttachmentAsync` method removes the attachment with the specified identifier from the item. As a best practice, you should use the attachment identifier to remove an attachment only if the same mail app has added that attachment in the same session. In Outlook Web App and OWA for Devices, the attachment identifier is valid only within the same session. A session is over when the user closes the app, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

##### <a name="parameters"></a><span data-ttu-id="ae6be-849">参数：</span><span class="sxs-lookup"><span data-stu-id="ae6be-849">Parameters:</span></span>

|<span data-ttu-id="ae6be-850">名称</span><span class="sxs-lookup"><span data-stu-id="ae6be-850">Name</span></span>| <span data-ttu-id="ae6be-851">类型</span><span class="sxs-lookup"><span data-stu-id="ae6be-851">Type</span></span>| <span data-ttu-id="ae6be-852">属性</span><span class="sxs-lookup"><span data-stu-id="ae6be-852">Attributes</span></span>| <span data-ttu-id="ae6be-853">说明</span><span class="sxs-lookup"><span data-stu-id="ae6be-853">Description</span></span>|
|---|---|---|---|
|`attachmentId`| <span data-ttu-id="ae6be-854">字符串</span><span class="sxs-lookup"><span data-stu-id="ae6be-854">String</span></span>||<span data-ttu-id="ae6be-855">要删除的附件的标识符。</span><span class="sxs-lookup"><span data-stu-id="ae6be-855">The identifier of the attachment to remove.</span></span>|
|`options`| <span data-ttu-id="ae6be-856">对象</span><span class="sxs-lookup"><span data-stu-id="ae6be-856">Object</span></span>| <span data-ttu-id="ae6be-857">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="ae6be-857">&lt;optional&gt;</span></span>|<span data-ttu-id="ae6be-858">包含一个或多个以下属性的对象文本。</span><span class="sxs-lookup"><span data-stu-id="ae6be-858">An object literal that contains one or more of the following properties.</span></span>|
|`options.asyncContext`| <span data-ttu-id="ae6be-859">对象</span><span class="sxs-lookup"><span data-stu-id="ae6be-859">Object</span></span>| <span data-ttu-id="ae6be-860">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="ae6be-860">&lt;optional&gt;</span></span>|<span data-ttu-id="ae6be-861">开发人员可以提供他们想要在回调方法中访问的任何对象。</span><span class="sxs-lookup"><span data-stu-id="ae6be-861">Developers can provide any object they wish to access in the callback method.</span></span>|
|`callback`| <span data-ttu-id="ae6be-862">function</span><span class="sxs-lookup"><span data-stu-id="ae6be-862">function</span></span>| <span data-ttu-id="ae6be-863">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="ae6be-863">&lt;optional&gt;</span></span>|<span data-ttu-id="ae6be-864">方法完成后，使用单个参数 `callback`（一个 [`asyncResult`](/javascript/api/office/office.asyncresult) 对象）调用在 `AsyncResult` 参数中传递的函数。</span><span class="sxs-lookup"><span data-stu-id="ae6be-864">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span> <br/><span data-ttu-id="ae6be-865">如果删除附件失败，`asyncResult.error` 属性将包含一个说明失败原因的错误代码。</span><span class="sxs-lookup"><span data-stu-id="ae6be-865">If removing the attachment fails, the `asyncResult.error` property will contain an error code with the reason for the failure.</span></span>|

##### <a name="errors"></a><span data-ttu-id="ae6be-866">错误</span><span class="sxs-lookup"><span data-stu-id="ae6be-866">Errors</span></span>

| <span data-ttu-id="ae6be-867">错误代码</span><span class="sxs-lookup"><span data-stu-id="ae6be-867">Error code</span></span> | <span data-ttu-id="ae6be-868">说明</span><span class="sxs-lookup"><span data-stu-id="ae6be-868">Description</span></span> |
|------------|-------------|
| `InvalidAttachmentId` | <span data-ttu-id="ae6be-869">附件标识符不存在。</span><span class="sxs-lookup"><span data-stu-id="ae6be-869">The attachment identifier does not exist.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="ae6be-870">Requirements</span><span class="sxs-lookup"><span data-stu-id="ae6be-870">Requirements</span></span>

|<span data-ttu-id="ae6be-871">要求</span><span class="sxs-lookup"><span data-stu-id="ae6be-871">Requirement</span></span>| <span data-ttu-id="ae6be-872">值</span><span class="sxs-lookup"><span data-stu-id="ae6be-872">Value</span></span>|
|---|---|
|[<span data-ttu-id="ae6be-873">最低版本的邮箱要求集</span><span class="sxs-lookup"><span data-stu-id="ae6be-873">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ae6be-874">1.1</span><span class="sxs-lookup"><span data-stu-id="ae6be-874">1.1</span></span>|
|[<span data-ttu-id="ae6be-875">最低权限级别</span><span class="sxs-lookup"><span data-stu-id="ae6be-875">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ae6be-876">ReadWriteItem</span><span class="sxs-lookup"><span data-stu-id="ae6be-876">ReadWriteItem</span></span>|
|[<span data-ttu-id="ae6be-877">适用的 Outlook 模式</span><span class="sxs-lookup"><span data-stu-id="ae6be-877">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ae6be-878">撰写</span><span class="sxs-lookup"><span data-stu-id="ae6be-878">Compose</span></span>|

##### <a name="example"></a><span data-ttu-id="ae6be-879">示例</span><span class="sxs-lookup"><span data-stu-id="ae6be-879">Example</span></span>

<span data-ttu-id="ae6be-880">以下代码删除包含标识符 '0' 的附件。</span><span class="sxs-lookup"><span data-stu-id="ae6be-880">The following code removes an attachment with an identifier of '0'.</span></span>

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
