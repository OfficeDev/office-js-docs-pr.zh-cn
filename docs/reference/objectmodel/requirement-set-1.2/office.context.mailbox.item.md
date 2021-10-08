---
title: Office.context.mailbox.item - 要求集 1.2
description: Outlook邮箱 API 要求集 1.2 版本的项目对象模型。
ms.date: 07/16/2021
ms.localizationpriority: medium
ms.openlocfilehash: 3607f3ab923b8fcf7c5fd4de69e5c9c4248810a9
ms.sourcegitcommit: efd0966f6400c8e685017ce0c8c016a2cbab0d5c
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 10/08/2021
ms.locfileid: "60237313"
---
# <a name="item-mailbox-requirement-set-12"></a>item (Mailbox requirement set 1.2) 

### <a name="officecontextmailboxitem"></a>[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item

`item` 用于访问当前选择的邮件、会议请求或约会。 可以使用 属性来确定项目 `itemType` 的类型。

##### <a name="requirements"></a>要求

|要求|值|
|---|---|
|[最低版本的邮箱要求集](../../requirement-sets/outlook-api-requirement-sets.md)|1.1|
|[最低权限级别](../../../outlook/understanding-outlook-add-in-permissions.md)|受限|
|[适用的 Outlook 模式](../../../outlook/outlook-add-ins-overview.md#extension-points)|约会组织者、约会参与者、<br>邮件撰写或邮件阅读|

> [!IMPORTANT]
> Android 和 iOS：外接程序激活时间以及哪些 API 可用存在限制。 若要了解详细信息，请参阅 [将移动支持添加到 Outlook 加载项](../../../outlook/add-mobile-support.md#compose-mode-and-appointments)。

## <a name="properties"></a>属性

| 属性 | 最小值<br>权限级别 | 按模式显示的详细信息 | 返回类型 | 最小值<br>要求集 |
|---|---|---|---|:---:|
| attachments | ReadItem | [约会参与者](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.2&preserve-view=true#attachments) | Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.2&preserve-view=true)> | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [邮件阅读](/javascript/api/outlook/office.messageread?view=outlook-js-1.2&preserve-view=true#attachments) | Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.2&preserve-view=true)> | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| bcc | ReadItem | [邮件撰写](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.2&preserve-view=true#bcc) | [收件人](/javascript/api/outlook/office.recipients?view=outlook-js-1.2&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| body | ReadItem | [约会组织者](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.2&preserve-view=true#body) | [正文](/javascript/api/outlook/office.body?view=outlook-js-1.2&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [约会参与者](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.2&preserve-view=true#body) | [正文](/javascript/api/outlook/office.body?view=outlook-js-1.2&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [邮件撰写](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.2&preserve-view=true#body) | [正文](/javascript/api/outlook/office.body?view=outlook-js-1.2&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [邮件阅读](/javascript/api/outlook/office.messageread?view=outlook-js-1.2&preserve-view=true#body) | [正文](/javascript/api/outlook/office.body?view=outlook-js-1.2&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| cc | ReadItem | [邮件撰写](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.2&preserve-view=true#cc) | [收件人](/javascript/api/outlook/office.recipients?view=outlook-js-1.2&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [邮件阅读](/javascript/api/outlook/office.messageread?view=outlook-js-1.2&preserve-view=true#cc) | Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2&preserve-view=true)> | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| conversationId | ReadItem | [邮件撰写](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.2&preserve-view=true#conversationId) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [邮件阅读](/javascript/api/outlook/office.messageread?view=outlook-js-1.2&preserve-view=true#conversationId) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| dateTimeCreated | ReadItem | [约会参与者](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.2&preserve-view=true#dateTimeCreated) | 日期 | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [邮件阅读](/javascript/api/outlook/office.messageread?view=outlook-js-1.2&preserve-view=true#dateTimeCreated) | 日期 | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| dateTimeModified | ReadItem | [约会参与者](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.2&preserve-view=true#dateTimeModified) | 日期 | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [邮件阅读](/javascript/api/outlook/office.messageread?view=outlook-js-1.2&preserve-view=true#dateTimeModified) | 日期 | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| end | ReadItem | [约会组织者](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.2&preserve-view=true#end) | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.2&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [约会参与者](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.2&preserve-view=true#end) | 日期 | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [邮件阅读](/javascript/api/outlook/office.messageread?view=outlook-js-1.2&preserve-view=true#end)<br> (会议请求)  | 日期 | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| 起始数量 | ReadItem | [邮件阅读](/javascript/api/outlook/office.messageread?view=outlook-js-1.2&preserve-view=true#from) | [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| internetMessageId | ReadItem | [邮件阅读](/javascript/api/outlook/office.messageread?view=outlook-js-1.2&preserve-view=true#internetMessageId) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| itemClass | ReadItem | [约会参与者](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.2&preserve-view=true#itemClass) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [邮件阅读](/javascript/api/outlook/office.messageread?view=outlook-js-1.2&preserve-view=true#itemClass) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| itemId | ReadItem | [约会参与者](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.2&preserve-view=true#itemId) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [邮件阅读](/javascript/api/outlook/office.messageread?view=outlook-js-1.2&preserve-view=true#itemId) | 字符串 | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| itemType | ReadItem | [约会组织者](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.2&preserve-view=true#itemType) | [MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.2&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [约会参与者](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.2&preserve-view=true#itemType) | [MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.2&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [邮件撰写](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.2&preserve-view=true#itemType) | [MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.2&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [邮件阅读](/javascript/api/outlook/office.messageread?view=outlook-js-1.2&preserve-view=true#itemType) | [MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.2&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| 位置 | ReadItem | [约会组织者](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.2&preserve-view=true#location) | [位置](/javascript/api/outlook/office.location?view=outlook-js-1.2&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [约会参与者](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.2&preserve-view=true#location) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [邮件阅读](/javascript/api/outlook/office.messageread?view=outlook-js-1.2&preserve-view=true#location)<br> (会议请求)  | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| normalizedSubject | ReadItem | [约会参与者](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.2&preserve-view=true#normalizedSubject) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [邮件阅读](/javascript/api/outlook/office.messageread?view=outlook-js-1.2&preserve-view=true#normalizedSubject) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| optionalAttendees | ReadItem | [约会组织者](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.2&preserve-view=true#optionalAttendees) | [收件人](/javascript/api/outlook/office.recipients?view=outlook-js-1.2&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [约会参与者](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.2&preserve-view=true#optionalAttendees) | Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2&preserve-view=true)> | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| 组织者 | ReadItem | [约会参与者](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.2&preserve-view=true#organizer) | [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| requiredAttendees | ReadItem | [约会组织者](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.2&preserve-view=true#requiredAttendees) | [收件人](/javascript/api/outlook/office.recipients?view=outlook-js-1.2&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [约会参与者](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.2&preserve-view=true#requiredAttendees) | Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2&preserve-view=true)> | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| sender | ReadItem | [邮件阅读](/javascript/api/outlook/office.messageread?view=outlook-js-1.2&preserve-view=true#sender) | [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| start | ReadItem | [约会组织者](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.2&preserve-view=true#start) | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.2&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [约会参与者](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.2&preserve-view=true#start) | 日期 | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [邮件阅读](/javascript/api/outlook/office.messageread?view=outlook-js-1.2&preserve-view=true#start)<br> (会议请求)  | 日期 | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| subject | ReadItem | [约会组织者](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.2&preserve-view=true#subject) | [主题](/javascript/api/outlook/office.subject?view=outlook-js-1.2&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [约会参与者](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.2&preserve-view=true#subject) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [邮件撰写](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.2&preserve-view=true#subject) | [主题](/javascript/api/outlook/office.subject?view=outlook-js-1.2&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [邮件阅读](/javascript/api/outlook/office.messageread?view=outlook-js-1.2&preserve-view=true#subject) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| 更改为 | ReadItem | [邮件撰写](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.2&preserve-view=true#to) | [收件人](/javascript/api/outlook/office.recipients?view=outlook-js-1.2&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [邮件阅读](/javascript/api/outlook/office.messageread?view=outlook-js-1.2&preserve-view=true#to) | Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2&preserve-view=true)> | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="methods"></a>方法

| 方法 | 最小值<br>权限级别 | 按模式显示的详细信息 | 最小值<br>要求集 |
|---|---|---|:---:|
| addFileAttachmentAsync(uri, attachmentName, [options], [callback]) | ReadWriteItem | [约会组织者](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.2&preserve-view=true#addFileAttachmentAsync_uri__attachmentName__options__callback_) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [邮件撰写](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.2&preserve-view=true#addFileAttachmentAsync_uri__attachmentName__options__callback_) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| addItemAttachmentAsync(itemId, attachmentName, [options], [callback]) | ReadWriteItem | [约会组织者](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.2&preserve-view=true#addItemAttachmentAsync_itemId__attachmentName__options__callback_) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [邮件撰写](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.2&preserve-view=true#addItemAttachmentAsync_itemId__attachmentName__options__callback_) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| displayReplyAllForm(formData) | ReadItem | [约会参与者](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.2&preserve-view=true#displayReplyAllForm_formData_) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [邮件阅读](/javascript/api/outlook/office.messageread?view=outlook-js-1.2&preserve-view=true#displayReplyAllForm_formData_) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| displayReplyForm(formData) | ReadItem | [约会参与者](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.2&preserve-view=true#displayReplyForm_formData_) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [邮件阅读](/javascript/api/outlook/office.messageread?view=outlook-js-1.2&preserve-view=true#displayReplyForm_formData_) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| getEntities ()  | ReadItem | [约会参与者](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.2&preserve-view=true#getEntities__) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [邮件阅读](/javascript/api/outlook/office.messageread?view=outlook-js-1.2&preserve-view=true#getEntities__) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| getEntitiesByType (entityType)  | 受限 | [约会参与者](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.2&preserve-view=true#getEntitiesByType_entityType_) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [邮件阅读](/javascript/api/outlook/office.messageread?view=outlook-js-1.2&preserve-view=true#getEntitiesByType_entityType_) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| getFilteredEntitiesByName (name)  | ReadItem | [约会参与者](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.2&preserve-view=true#getFilteredEntitiesByName_name_) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [邮件阅读](/javascript/api/outlook/office.messageread?view=outlook-js-1.2&preserve-view=true#getFilteredEntitiesByName_name_) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| getRegExMatches ()  | ReadItem | [约会参与者](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.2&preserve-view=true#getRegExMatches__) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [邮件阅读](/javascript/api/outlook/office.messageread?view=outlook-js-1.2&preserve-view=true#getRegExMatches__) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| getRegExMatchesByName (name)  | ReadItem | [约会参与者](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.2&preserve-view=true#getRegExMatchesByName_name_) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [邮件阅读](/javascript/api/outlook/office.messageread?view=outlook-js-1.2&preserve-view=true#getRegExMatchesByName_name_) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| getSelectedDataAsync (coercionType， [options]， callback)  | ReadItem | [约会组织者](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.2&preserve-view=true#getSelectedDataAsync_coercionType__options__callback_) | [1.2](../requirement-set-1.2/outlook-requirement-set-1.2.md) |
| | | [邮件撰写](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.2&preserve-view=true#getSelectedDataAsync_coercionType__options__callback_) | [1.2](../requirement-set-1.2/outlook-requirement-set-1.2.md) |
| loadCustomPropertiesAsync(callback, [userContext]) | ReadItem | [约会组织者](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.2&preserve-view=true#loadCustomPropertiesAsync_callback__userContext_) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [约会参与者](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.2&preserve-view=true#loadCustomPropertiesAsync_callback__userContext_) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [邮件撰写](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.2&preserve-view=true#loadCustomPropertiesAsync_callback__userContext_) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [邮件阅读](/javascript/api/outlook/office.messageread?view=outlook-js-1.2&preserve-view=true#loadCustomPropertiesAsync_callback__userContext_) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| removeAttachmentAsync(attachmentId, [options], [callback]) | ReadWriteItem | [约会组织者](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.2&preserve-view=true#removeAttachmentAsync_attachmentId__options__callback_) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
|  |  | [邮件撰写](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.2&preserve-view=true#removeAttachmentAsync_attachmentId__options__callback_) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| setSelectedDataAsync(data, [options], callback) | ReadWriteItem | [约会组织者](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.2&preserve-view=true#setSelectedDataAsync_data__options__callback_) | [1.2](../requirement-set-1.2/outlook-requirement-set-1.2.md) |
| | | [邮件撰写](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.2&preserve-view=true#setSelectedDataAsync_data__options__callback_) | [1.2](../requirement-set-1.2/outlook-requirement-set-1.2.md) |

## <a name="example"></a>示例

以下 JavaScript 代码示例显示了如何访问 Outlook 中当前项目的 `subject` 属性。

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
};
```
