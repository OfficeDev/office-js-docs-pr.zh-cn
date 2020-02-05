---
title: "\"Context\"-\"邮箱\"。项目-要求集1。6"
description: ''
ms.date: 02/03/2020
localization_priority: Normal
ms.openlocfilehash: a2372c6cdcd22489f7e06f43cdda14f96be9fed4
ms.sourcegitcommit: c1dbea577ae6183523fb663d364422d2adbc8bcf
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 02/05/2020
ms.locfileid: "41773549"
---
# <a name="item"></a>item

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmditem"></a>[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item

`item`用于访问当前选定的邮件、会议请求或约会。您可以通过使用`itemType`属性来确定项目的类型。

##### <a name="requirements"></a>Requirements

|要求|值|
|---|---|
|[最低版本的邮箱要求集](../../requirement-sets/outlook-api-requirement-sets.md)|1.1|
|[最低权限级别](/outlook/add-ins/understanding-outlook-add-in-permissions)|受限|
|[适用的 Outlook 模式](/outlook/add-ins/#extension-points)|约会组织者、约会与会者、<br>邮件撰写或邮件读取|

## <a name="properties"></a>属性

| 属性 | 最低<br>权限级别 | 详细信息（按模式） | 返回类型 | 最低<br>要求集 |
|---|---|---|---|:---:|
| attachments | ReadItem | [约会与会者](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.6#attachments) | Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)> | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [邮件读取](/javascript/api/outlook/office.messageread?view=outlook-js-1.6#attachments) | Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)> | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| bcc | ReadItem | [邮件撰写](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.6#bcc) | [收件人](/javascript/api/outlook/office.recipients) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| body | ReadItem | [约会组织者](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.6#body) | [Body](/javascript/api/outlook/office.body) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [约会与会者](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.6#body) | [Body](/javascript/api/outlook/office.body) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [邮件撰写](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.6#body) | [Body](/javascript/api/outlook/office.body) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [邮件读取](/javascript/api/outlook/office.messageread?view=outlook-js-1.6#body) | [Body](/javascript/api/outlook/office.body) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| cc | ReadItem | [邮件撰写](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.6#cc) | [收件人](/javascript/api/outlook/office.recipients) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [邮件读取](/javascript/api/outlook/office.messageread?view=outlook-js-1.6#cc) | <[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| conversationId | ReadItem | [邮件撰写](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.6#conversationid) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [邮件读取](/javascript/api/outlook/office.messageread?view=outlook-js-1.6#conversationid) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| dateTimeCreated | ReadItem | [约会与会者](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.6#datetimecreated) | 日期 | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [邮件读取](/javascript/api/outlook/office.messageread?view=outlook-js-1.6#datetimecreated) | 日期 | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| dateTimeModified | ReadItem | [约会与会者](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.6#datetimemodified) | 日期 | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [邮件读取](/javascript/api/outlook/office.messageread?view=outlook-js-1.6#datetimemodified) | 日期 | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| end | ReadItem | [约会组织者](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.6#end) | [Time](/javascript/api/outlook/office.time) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [约会与会者](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.6#end) | 日期 | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [邮件读取](/javascript/api/outlook/office.messageread?view=outlook-js-1.6#end)<br>（会议请求） | 日期 | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| 发件人 | ReadItem | [邮件读取](/javascript/api/outlook/office.messageread?view=outlook-js-1.6#from) | [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| internetMessageId | ReadItem | [邮件读取](/javascript/api/outlook/office.messageread?view=outlook-js-1.6#internetmessageid) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| itemClass | ReadItem | [约会与会者](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.6#itemclass) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [邮件读取](/javascript/api/outlook/office.messageread?view=outlook-js-1.6#itemclass) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| itemId | ReadItem | [约会与会者](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.6#itemid) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [邮件读取](/javascript/api/outlook/office.messageread?view=outlook-js-1.6#itemid) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| itemType | ReadItem | [约会组织者](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.6#itemtype) | [MailboxEnums](/javascript/api/outlook/office.mailboxenums.itemtype) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [约会与会者](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.6#itemtype) | [MailboxEnums](/javascript/api/outlook/office.mailboxenums.itemtype) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [邮件撰写](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.6#itemtype) | [MailboxEnums](/javascript/api/outlook/office.mailboxenums.itemtype) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [邮件读取](/javascript/api/outlook/office.messageread?view=outlook-js-1.6#itemtype) | [MailboxEnums](/javascript/api/outlook/office.mailboxenums.itemtype) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| location | ReadItem | [约会组织者](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.6#location) | [Location](/javascript/api/outlook/office.location) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [约会与会者](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.6#location) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [邮件读取](/javascript/api/outlook/office.messageread?view=outlook-js-1.6#location)<br>（会议请求） | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| normalizedSubject | ReadItem | [约会与会者](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.6#normalizedsubject) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [邮件读取](/javascript/api/outlook/office.messageread?view=outlook-js-1.6#normalizedsubject) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| notificationMessages | ReadItem | [约会组织者](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.6#notificationmessages) | [NotificationMessages](/javascript/api/outlook/office.notificationmessages) | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| | | [约会与会者](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.6#notificationmessages) | [NotificationMessages](/javascript/api/outlook/office.notificationmessages) | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| | | [邮件撰写](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.6#notificationmessages) | [NotificationMessages](/javascript/api/outlook/office.notificationmessages) | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| | | [邮件读取](/javascript/api/outlook/office.messageread?view=outlook-js-1.6#notificationmessages) | [NotificationMessages](/javascript/api/outlook/office.notificationmessages) | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| optionalAttendees | ReadItem | [约会组织者](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.6#optionalattendees) | [收件人](/javascript/api/outlook/office.recipients) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [约会与会者](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.6#optionalattendees) | <[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| organizer － 组织者 | ReadItem | [约会与会者](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.6#organizer) | [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| requiredAttendees | ReadItem | [约会组织者](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.6#requiredattendees) | [收件人](/javascript/api/outlook/office.recipients) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [约会与会者](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.6#requiredattendees) | <[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| sender | ReadItem | [邮件读取](/javascript/api/outlook/office.messageread?view=outlook-js-1.6#sender) | [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| start | ReadItem | [约会组织者](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.6#start) | [Time](/javascript/api/outlook/office.time) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [约会与会者](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.6#start) | 日期 | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [邮件读取](/javascript/api/outlook/office.messageread?view=outlook-js-1.6#start)<br>（会议请求） | 日期 | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| subject | ReadItem | [约会组织者](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.6#subject) | [Subject](/javascript/api/outlook/office.subject) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [约会与会者](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.6#subject) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [邮件撰写](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.6#subject) | [Subject](/javascript/api/outlook/office.subject) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [邮件读取](/javascript/api/outlook/office.messageread?view=outlook-js-1.6#subject) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| to | ReadItem | [邮件撰写](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.6#to) | [收件人](/javascript/api/outlook/office.recipients) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [邮件读取](/javascript/api/outlook/office.messageread?view=outlook-js-1.6#to) | <[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="methods"></a>方法

| 方法 | 最低<br>权限级别 | 详细信息（按模式） | 最低<br>要求集 |
|---|---|---|:---:|
| addFileAttachmentAsync | ReadWriteItem | [约会组织者](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.6#addfileattachmentasync-uri--attachmentname--options--callback-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [邮件撰写](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.6#addfileattachmentasync-uri--attachmentname--options--callback-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| addItemAttachmentAsync | ReadWriteItem | [约会组织者](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.6#additemattachmentasync-itemid--attachmentname--options--callback-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [邮件撰写](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.6#additemattachmentasync-itemid--attachmentname--options--callback-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| 关闭 | 受限 | [约会组织者](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.6#close--) | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| | | [邮件撰写](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.6#close--) | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| displayReplyAllForm | ReadItem | [约会与会者](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.6#displayreplyallform-formdata--callback-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [邮件读取](/javascript/api/outlook/office.messageread?view=outlook-js-1.6#displayreplyallform-formdata--callback-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| displayReplyForm | ReadItem | [约会与会者](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.6#displayreplyform-formdata--callback-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [邮件读取](/javascript/api/outlook/office.messageread?view=outlook-js-1.6#displayreplyform-formdata--callback-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| getEntities | ReadItem | [约会与会者](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.6#getentities--) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [邮件读取](/javascript/api/outlook/office.messageread?view=outlook-js-1.6#getentities--) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| getEntitiesByType | 受限 | [约会与会者](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.6#getentitiesbytype-entitytype-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [邮件读取](/javascript/api/outlook/office.messageread?view=outlook-js-1.6#getentitiesbytype-entitytype-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| getFilteredEntitiesByName | ReadItem | [约会与会者](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.6#getfilteredentitiesbyname-name-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [邮件读取](/javascript/api/outlook/office.messageread?view=outlook-js-1.6#getfilteredentitiesbyname-name-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| getRegExMatches | ReadItem | [约会与会者](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.6#getregexmatches--) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [邮件读取](/javascript/api/outlook/office.messageread?view=outlook-js-1.6#getregexmatches--) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| getRegExMatchesByName | ReadItem | [约会与会者](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.6#getregexmatchesbyname-name-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [邮件读取](/javascript/api/outlook/office.messageread?view=outlook-js-1.6#getregexmatchesbyname-name-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| getSelectedDataAsync | ReadItem | [约会组织者](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.6#getselecteddataasync-coerciontype--options--callback-) | [1.2](../requirement-set-1.2/outlook-requirement-set-1.2.md) |
| | | [邮件撰写](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.6#getselecteddataasync-coerciontype--options--callback-) | [1.2](../requirement-set-1.2/outlook-requirement-set-1.2.md) |
| Office.context.mailbox.item.getselectedentities | ReadItem | [约会与会者](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.6#getselectedentities--) | [1.6](../requirement-set-1.6/outlook-requirement-set-1.6.md) |
| | | [邮件读取](/javascript/api/outlook/office.messageread?view=outlook-js-1.6#getselectedentities--) | [1.6](../requirement-set-1.6/outlook-requirement-set-1.6.md) |
| Office.context.mailbox.item.getselectedregexmatches | ReadItem | [约会与会者](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.6#getselectedregexmatches--) | [1.6](../requirement-set-1.6/outlook-requirement-set-1.6.md) |
| | | [邮件读取](/javascript/api/outlook/office.messageread?view=outlook-js-1.6#getselectedregexmatches--) | [1.6](../requirement-set-1.6/outlook-requirement-set-1.6.md) |
| loadCustomPropertiesAsync | ReadItem | [约会组织者](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.6#loadcustompropertiesasync-callback--usercontext-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [约会与会者](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.6#loadcustompropertiesasync-callback--usercontext-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [邮件撰写](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.6#loadcustompropertiesasync-callback--usercontext-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [邮件读取](/javascript/api/outlook/office.messageread?view=outlook-js-1.6#loadcustompropertiesasync-callback--usercontext-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| removeAttachmentAsync | ReadWriteItem | [约会组织者](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.6#removeattachmentasync-attachmentid--options--callback-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
|  |  | [邮件撰写](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.6#removeattachmentasync-attachmentid--options--callback-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| saveAsync | ReadWriteItem | [约会组织者](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.6#saveasync-options--callback-) | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| | | [邮件撰写](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.6#saveasync-options--callback-) | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| setSelectedDataAsync | ReadWriteItem | [约会组织者](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.6#setselecteddataasync-data--options--callback-) | [1.2](../requirement-set-1.2/outlook-requirement-set-1.2.md) |
| | | [邮件撰写](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.6#setselecteddataasync-data--options--callback-) | [1.2](../requirement-set-1.2/outlook-requirement-set-1.2.md) |

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
