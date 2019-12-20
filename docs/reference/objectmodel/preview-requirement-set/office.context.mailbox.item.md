---
title: "\"Context.subname\"-\"邮箱\"-预览要求集"
description: ''
ms.date: 12/19/2019
localization_priority: Normal
ms.openlocfilehash: 39555575c4d801e82b1bf0d56754cc67a31695eb
ms.sourcegitcommit: 8c5c5a1bd3fe8b90f6253d9850e9352ed0b283ee
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 12/19/2019
ms.locfileid: "40815065"
---
# <a name="item"></a>item

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmditem"></a>[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item

`item`用于访问当前选定的邮件、会议请求或约会。您可以通过使用`itemType`属性来确定项目的类型。

##### <a name="requirements"></a>Requirements

|要求|值|
|---|---|
|[最低版本的邮箱要求集](../../requirement-sets/outlook-api-requirement-sets.md)|1.1|
|[最低权限级别](/outlook/add-ins/understanding-outlook-add-in-permissions)|受限|
|[适用的 Outlook 模式](/outlook/add-ins/#extension-points)|撰写或阅读|

## <a name="properties"></a>属性

| 属性 | 最低<br>权限级别 | 详细信息（按模式） | 返回类型 | 最低<br>要求集 |
|---|---|---|---|:---:|
| attachments | ReadItem | [约会与会者](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview#attachments) | Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)> | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [邮件读取](/javascript/api/outlook/office.messageread?view=outlook-js-preview#attachments) | Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)> | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| bcc | ReadItem | [邮件撰写](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview#bcc) | [收件人](/javascript/api/outlook/office.recipients) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| body | ReadItem | [约会组织者](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#body) | [Body](/javascript/api/outlook/office.body) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [约会与会者](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview#body) | [Body](/javascript/api/outlook/office.body) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [邮件撰写](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview#body) | [Body](/javascript/api/outlook/office.body) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [邮件读取](/javascript/api/outlook/office.messageread?view=outlook-js-preview#body) | [Body](/javascript/api/outlook/office.body) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| categories | ReadItem | [约会组织者](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#categories) | [Categories](/javascript/api/outlook/office.categories) | [1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| | | [约会与会者](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview#categories) | [Categories](/javascript/api/outlook/office.categories) | [1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| | | [邮件撰写](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview#categories) | [Categories](/javascript/api/outlook/office.categories) | [1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| | | [邮件读取](/javascript/api/outlook/office.messageread?view=outlook-js-preview#categories) | [Categories](/javascript/api/outlook/office.categories) | [1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| cc | ReadItem | [邮件撰写](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview#cc) | [收件人](/javascript/api/outlook/office.recipients) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [邮件读取](/javascript/api/outlook/office.messageread?view=outlook-js-preview#cc) | <[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| conversationId | ReadItem | [邮件撰写](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview#conversationid) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [邮件读取](/javascript/api/outlook/office.messageread?view=outlook-js-preview#conversationid) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| dateTimeCreated | ReadItem | [约会与会者](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview#datetimecreated) | 日期 | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [邮件读取](/javascript/api/outlook/office.messageread?view=outlook-js-preview#datetimecreated) | 日期 | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| dateTimeModified | ReadItem | [约会与会者](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview#datetimemodified) | 日期 | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [邮件读取](/javascript/api/outlook/office.messageread?view=outlook-js-preview#datetimemodified) | 日期 | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| end | ReadItem | [约会组织者](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#end) | [Time](/javascript/api/outlook/office.time) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [约会与会者](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview#end) | 日期 | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [邮件读取](/javascript/api/outlook/office.messageread?view=outlook-js-preview#end)<br>（会议请求） | 日期 | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| enhancedLocation | ReadItem | [约会组织者](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#enhancedlocation) | [EnhancedLocation](/javascript/api/outlook/office.enhancedlocation) | [1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| | | [约会与会者](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview#enhancedlocation) | [EnhancedLocation](/javascript/api/outlook/office.enhancedlocation) | [1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| 发件人 | ReadWriteItem | [邮件撰写](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview#from) | [From](/javascript/api/outlook/office.from) | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | ReadItem | [邮件读取](/javascript/api/outlook/office.messageread?view=outlook-js-preview#from) | [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| internetHeaders | ReadItem | [邮件撰写](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview#internetheaders) | [InternetHeaders](/javascript/api/outlook/office.internetheaders) | [1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| internetMessageId | ReadItem | [邮件读取](/javascript/api/outlook/office.messageread?view=outlook-js-preview#internetmessageid) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| itemClass | ReadItem | [约会与会者](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview#itemclass) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [邮件读取](/javascript/api/outlook/office.messageread?view=outlook-js-preview#itemclass) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| itemId | ReadItem | [约会与会者](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview#itemid) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [邮件读取](/javascript/api/outlook/office.messageread?view=outlook-js-preview#itemid) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| itemType | ReadItem | [约会组织者](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#itemtype) | [MailboxEnums](/javascript/api/outlook/office.mailboxenums.itemtype) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [约会与会者](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview#itemtype) | [MailboxEnums](/javascript/api/outlook/office.mailboxenums.itemtype) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [邮件撰写](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview#itemtype) | [MailboxEnums](/javascript/api/outlook/office.mailboxenums.itemtype) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [邮件读取](/javascript/api/outlook/office.messageread?view=outlook-js-preview#itemtype) | [MailboxEnums](/javascript/api/outlook/office.mailboxenums.itemtype) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| location | ReadItem | [约会组织者](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#location) | [Location](/javascript/api/outlook/office.location) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [约会与会者](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview#location) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [邮件读取](/javascript/api/outlook/office.messageread?view=outlook-js-preview#location)<br>（会议请求） | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| normalizedSubject | ReadItem | [约会与会者](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview#normalizedsubject) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [邮件读取](/javascript/api/outlook/office.messageread?view=outlook-js-preview#normalizedsubject) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| notificationMessages | ReadItem | [约会组织者](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#notificationmessages) | [NotificationMessages](/javascript/api/outlook/office.notificationmessages) | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| | | [约会与会者](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview#notificationmessages) | [NotificationMessages](/javascript/api/outlook/office.notificationmessages) | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| | | [邮件撰写](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview#notificationmessages) | [NotificationMessages](/javascript/api/outlook/office.notificationmessages) | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| | | [邮件读取](/javascript/api/outlook/office.messageread?view=outlook-js-preview#notificationmessages) | [NotificationMessages](/javascript/api/outlook/office.notificationmessages) | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| optionalAttendees | ReadItem | [约会组织者](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#optionalattendees) | [收件人](/javascript/api/outlook/office.recipients) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [约会与会者](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview#optionalattendees) | <[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| 组织者 | ReadWriteItem | [约会组织者](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#organizer) | [Organizer](/javascript/api/outlook/office.organizer) | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | ReadItem | [约会与会者](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview#organizer) | [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| recurrence | ReadItem | [约会组织者](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#recurrence) | [循环](/javascript/api/outlook/office.recurrence) | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | | [约会与会者](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview#recurrence) | [循环](/javascript/api/outlook/office.recurrence) | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | | [邮件读取](/javascript/api/outlook/office.messageread?view=outlook-js-preview#recurrence)<br>（会议请求） | [循环](/javascript/api/outlook/office.recurrence) | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| requiredAttendees | ReadItem | [约会组织者](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#requiredattendees) | [收件人](/javascript/api/outlook/office.recipients) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [约会与会者](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview#requiredattendees) | <[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| sender | ReadItem | [邮件读取](/javascript/api/outlook/office.messageread?view=outlook-js-preview#sender) | [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| Webcasts&seriesid | ReadItem | [约会组织者](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#seriesid) | String | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | | [约会与会者](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview#seriesid) | String | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | | [邮件撰写](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview#seriesid) | String | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | | [邮件读取](/javascript/api/outlook/office.messageread?view=outlook-js-preview#seriesid) | String | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| start | ReadItem | [约会组织者](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#start) | [Time](/javascript/api/outlook/office.time) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [约会与会者](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview#start) | 日期 | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [邮件读取](/javascript/api/outlook/office.messageread?view=outlook-js-preview#start)<br>（会议请求） | 日期 | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| subject | ReadItem | [约会组织者](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#subject) | [Subject](/javascript/api/outlook/office.subject) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [约会与会者](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview#subject) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [邮件撰写](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview#subject) | [Subject](/javascript/api/outlook/office.subject) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [邮件读取](/javascript/api/outlook/office.messageread?view=outlook-js-preview#subject) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| to | ReadItem | [邮件撰写](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview#to) | [收件人](/javascript/api/outlook/office.recipients) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [邮件读取](/javascript/api/outlook/office.messageread?view=outlook-js-preview#to) | <[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="methods"></a>方法

| 方法 | 最低<br>权限级别 | 详细信息（按模式） | 最低<br>要求集 |
|---|---|---|:---:|
| addFileAttachmentAsync | ReadWriteItem | [约会组织者](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#addfileattachmentasync-uri--attachmentname--options--callback-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [邮件撰写](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview#addfileattachmentasync-uri--attachmentname--options--callback-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| addFileAttachmentFromBase64Async | ReadWriteItem | [约会组织者](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#addfileattachmentfrombase64async-base64file--attachmentname--options--callback-) | [1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| | | [邮件撰写](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview#addfileattachmentfrombase64async-base64file--attachmentname--options--callback-) | [1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| addHandlerAsync | ReadItem | [约会组织者](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#addhandlerasync-eventtype--handler--options--callback-) | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | | [约会与会者](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview#addhandlerasync-eventtype--handler--options--callback-) | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | | [邮件撰写](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview#addhandlerasync-eventtype--handler--options--callback-) | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | | [邮件读取](/javascript/api/outlook/office.messageread?view=outlook-js-preview#addhandlerasync-eventtype--handler--options--callback-) | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| addItemAttachmentAsync | ReadWriteItem | [约会组织者](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#additemattachmentasync-itemid--attachmentname--options--callback-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [邮件撰写](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview#additemattachmentasync-itemid--attachmentname--options--callback-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| 关闭 | 受限 | [约会组织者](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#close--) | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| | | [邮件撰写](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview#close--) | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| displayReplyAllForm | ReadItem | [约会与会者](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview#displayreplyallform-formdata--callback-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [邮件读取](/javascript/api/outlook/office.messageread?view=outlook-js-preview#displayreplyallform-formdata--callback-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| displayReplyForm | ReadItem | [约会与会者](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview#displayreplyform-formdata--callback-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [邮件读取](/javascript/api/outlook/office.messageread?view=outlook-js-preview#displayreplyform-formdata--callback-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| getAllInternetHeadersAsync | ReadItem | [邮件读取](/javascript/api/outlook/office.messageread?view=outlook-js-preview#getallinternetheadersasync-options--callback-) | [1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| getAttachmentContentAsync | ReadItem | [约会组织者](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#getattachmentcontentasync-attachmentid--options--callback-) | [1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| | | [约会与会者](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview#getattachmentcontentasync-attachmentid--options--callback-) | [1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| | | [邮件撰写](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview#getattachmentcontentasync-attachmentid--options--callback-) | [1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| | | [邮件读取](/javascript/api/outlook/office.messageread?view=outlook-js-preview#getattachmentcontentasync-attachmentid--options--callback-) | [1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| getAttachmentsAsync | ReadItem | [约会组织者](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#getattachmentsasync-options--callback-) | [1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| | | [邮件撰写](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview#getattachmentsasync-options--callback-) | [1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| getEntities | ReadItem | [约会与会者](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview#getentities--) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [邮件读取](/javascript/api/outlook/office.messageread?view=outlook-js-preview#getentities--) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| getEntitiesByType | 受限 | [约会与会者](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview#getentitiesbytype-entitytype-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [邮件读取](/javascript/api/outlook/office.messageread?view=outlook-js-preview#getentitiesbytype-entitytype-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| getFilteredEntitiesByName | ReadItem | [约会与会者](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview#getfilteredentitiesbyname-name-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [邮件读取](/javascript/api/outlook/office.messageread?view=outlook-js-preview#getfilteredentitiesbyname-name-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| Office.context.mailbox.item.getinitializationcontextasync | ReadItem | [约会组织者](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#getinitializationcontextasync-options--callback-) | [预览](../preview-requirement-set/outlook-requirement-set-preview.md) |
| | | [约会与会者](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview#getinitializationcontextasync-options--callback-) | [预览](../preview-requirement-set/outlook-requirement-set-preview.md) |
| | | [邮件撰写](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview#getinitializationcontextasync-options--callback-) | [预览](../preview-requirement-set/outlook-requirement-set-preview.md) |
| | | [邮件读取](/javascript/api/outlook/office.messageread?view=outlook-js-preview#getinitializationcontextasync-options--callback-) | [预览](../preview-requirement-set/outlook-requirement-set-preview.md) |
| getItemIdAsync | ReadItem | [约会组织者](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#getitemidasync-options--callback-) | [1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| | | [邮件撰写](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview#getitemidasync-options--callback-) | [1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| getRegExMatches | ReadItem | [约会与会者](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview#getregexmatches--) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [邮件读取](/javascript/api/outlook/office.messageread?view=outlook-js-preview#getregexmatches--) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| getRegExMatchesByName | ReadItem | [约会与会者](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview#getregexmatchesbyname-name-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [邮件读取](/javascript/api/outlook/office.messageread?view=outlook-js-preview#getregexmatchesbyname-name-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| getSelectedDataAsync | ReadItem | [约会组织者](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#getselecteddataasync-coerciontype--options--callback-) | [1.2](../requirement-set-1.2/outlook-requirement-set-1.2.md) |
| | | [邮件撰写](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview#getselecteddataasync-coerciontype--options--callback-) | [1.2](../requirement-set-1.2/outlook-requirement-set-1.2.md) |
| Office.context.mailbox.item.getselectedentities | ReadItem | [约会与会者](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview#getselectedentities--) | [1.6](../requirement-set-1.6/outlook-requirement-set-1.6.md) |
| | | [邮件读取](/javascript/api/outlook/office.messageread?view=outlook-js-preview#getselectedentities--) | [1.6](../requirement-set-1.6/outlook-requirement-set-1.6.md) |
| Office.context.mailbox.item.getselectedregexmatches | ReadItem | [约会与会者](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview#getselectedregexmatches--) | [1.6](../requirement-set-1.6/outlook-requirement-set-1.6.md) |
| | | [邮件读取](/javascript/api/outlook/office.messageread?view=outlook-js-preview#getselectedregexmatches--) | [1.6](../requirement-set-1.6/outlook-requirement-set-1.6.md) |
| getSharedPropertiesAsync | ReadItem | [约会组织者](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#getsharedpropertiesasync-options--callback-) | [1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| | | [约会与会者](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview#getsharedpropertiesasync-options--callback-) | [1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| | | [邮件撰写](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview#getsharedpropertiesasync-options--callback-) | [1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| | | [邮件读取](/javascript/api/outlook/office.messageread?view=outlook-js-preview#getsharedpropertiesasync-options--callback-) | [1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
| loadCustomPropertiesAsync | ReadItem | [约会组织者](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#loadcustompropertiesasync-callback--usercontext-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [约会与会者](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview#loadcustompropertiesasync-callback--usercontext-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [邮件撰写](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview#loadcustompropertiesasync-callback--usercontext-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [邮件读取](/javascript/api/outlook/office.messageread?view=outlook-js-preview#loadcustompropertiesasync-callback--usercontext-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| removeAttachmentAsync | ReadWriteItem | [约会组织者](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#removeattachmentasync-attachmentid--options--callback-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
|  |  | [邮件撰写](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview#removeattachmentasync-attachmentid--options--callback-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| removeHandlerAsync | ReadItem | [约会组织者](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#removehandlerasync-eventtype--options--callback-) | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | | [约会与会者](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview#removehandlerasync-eventtype--options--callback-) | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | | [邮件撰写](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview#removehandlerasync-eventtype--options--callback-) | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | | [邮件读取](/javascript/api/outlook/office.messageread?view=outlook-js-preview#removehandlerasync-eventtype--options--callback-) | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| saveAsync | ReadWriteItem | [约会组织者](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#saveasync-options--callback-) | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| | | [邮件撰写](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview#saveasync-options--callback-) | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| setSelectedDataAsync | ReadWriteItem | [约会组织者](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#setselecteddataasync-data--options--callback-) | [1.2](../requirement-set-1.2/outlook-requirement-set-1.2.md) |
| | | [邮件撰写](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview#setselecteddataasync-data--options--callback-) | [1.2](../requirement-set-1.2/outlook-requirement-set-1.2.md) |

## <a name="events"></a>事件

您可以使用`addHandlerAsync`和`removeHandlerAsync`分别订阅和取消订阅以下事件。

| 事件 | 说明 | 最低<br>要求集 |
|---|---|:---:|
|`AppointmentTimeChanged`| 所选的约会或系列的日期或时间已更改。 | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
|`AttachmentsChanged`| 已将附件添加到项目或已从项目删除附件。 | [1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
|`EnhancedLocationsChanged`| 所选约会的位置已更改。 | [1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md) |
|`RecipientsChanged`| 选定项目或约会位置的收件人列表已更改。 | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
|`RecurrenceChanged`| 选定系列的定期模式已更改。 | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |

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
