---
title: Office.context.mailbox.item - 要求集 1.7
description: Outlook邮箱 API 要求集 1.7 版本的项目对象模型。
ms.date: 05/17/2021
localization_priority: Normal
ms.openlocfilehash: 7aa50979eb8bf070d71261cbd3b047ba79f96415
ms.sourcegitcommit: ab3d38f2829e83f624bf43c49c0d267166552eec
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 06/11/2021
ms.locfileid: "52893587"
---
# <a name="item-mailbox-requirement-set-17"></a>item (Mailbox requirement set 1.7) 

### <a name="officecontextmailboxitem"></a>[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item

`item` 用于访问当前选择的邮件、会议请求或约会。 可以使用 属性来确定项目 `itemType` 的类型。

##### <a name="requirements"></a>要求

|要求|值|
|---|---|
|[最低版本的邮箱要求集](../../requirement-sets/outlook-api-requirement-sets.md)|1.1|
|[最低权限级别](../../../outlook/understanding-outlook-add-in-permissions.md)|受限|
|[适用的 Outlook 模式](../../../outlook/outlook-add-ins-overview.md#extension-points)|约会组织者、约会参与者、<br>邮件撰写或邮件阅读|

## <a name="properties"></a>属性

| 属性 | 最小值<br>权限级别 | 按模式显示的详细信息 | 返回类型 | 最小值<br>要求集 |
|---|---|---|---|:---:|
| attachments | ReadItem | [约会参与者](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.7&preserve-view=true#attachments) | Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)> | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [邮件阅读](/javascript/api/outlook/office.messageread?view=outlook-js-1.7&preserve-view=true#attachments) | Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)> | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| bcc | ReadItem | [邮件撰写](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.7&preserve-view=true#bcc) | [收件人](/javascript/api/outlook/office.recipients) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| body | ReadItem | [约会组织者](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.7&preserve-view=true#body) | [Body](/javascript/api/outlook/office.body) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [约会参与者](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.7&preserve-view=true#body) | [Body](/javascript/api/outlook/office.body) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [邮件撰写](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.7&preserve-view=true#body) | [Body](/javascript/api/outlook/office.body) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [邮件阅读](/javascript/api/outlook/office.messageread?view=outlook-js-1.7&preserve-view=true#body) | [Body](/javascript/api/outlook/office.body) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| cc | ReadItem | [邮件撰写](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.7&preserve-view=true#cc) | [收件人](/javascript/api/outlook/office.recipients) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [邮件阅读](/javascript/api/outlook/office.messageread?view=outlook-js-1.7&preserve-view=true#cc) | Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| conversationId | ReadItem | [邮件撰写](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.7&preserve-view=true#conversationid) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [邮件阅读](/javascript/api/outlook/office.messageread?view=outlook-js-1.7&preserve-view=true#conversationid) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| dateTimeCreated | ReadItem | [约会参与者](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.7&preserve-view=true#datetimecreated) | 日期 | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [邮件阅读](/javascript/api/outlook/office.messageread?view=outlook-js-1.7&preserve-view=true#datetimecreated) | 日期 | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| dateTimeModified | ReadItem | [约会参与者](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.7&preserve-view=true#datetimemodified) | 日期 | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [邮件阅读](/javascript/api/outlook/office.messageread?view=outlook-js-1.7&preserve-view=true#datetimemodified) | 日期 | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| end | ReadItem | [约会组织者](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.7&preserve-view=true#end) | [Time](/javascript/api/outlook/office.time) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [约会参与者](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.7&preserve-view=true#end) | 日期 | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [邮件阅读](/javascript/api/outlook/office.messageread?view=outlook-js-1.7&preserve-view=true#end)<br> (会议请求)  | 日期 | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| 发件人 | ReadWriteItem | [邮件撰写](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.7&preserve-view=true#from) | [From](/javascript/api/outlook/office.from) | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | ReadItem | [邮件阅读](/javascript/api/outlook/office.messageread?view=outlook-js-1.7&preserve-view=true#from) | [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| internetMessageId | ReadItem | [邮件阅读](/javascript/api/outlook/office.messageread?view=outlook-js-1.7&preserve-view=true#internetmessageid) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| itemClass | ReadItem | [约会参与者](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.7&preserve-view=true#itemclass) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [邮件阅读](/javascript/api/outlook/office.messageread?view=outlook-js-1.7&preserve-view=true#itemclass) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| itemId | ReadItem | [约会参与者](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.7&preserve-view=true#itemid) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [邮件阅读](/javascript/api/outlook/office.messageread?view=outlook-js-1.7&preserve-view=true#itemid) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| itemType | ReadItem | [约会组织者](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.7&preserve-view=true#itemtype) | [MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [约会参与者](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.7&preserve-view=true#itemtype) | [MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [邮件撰写](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.7&preserve-view=true#itemtype) | [MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [邮件阅读](/javascript/api/outlook/office.messageread?view=outlook-js-1.7&preserve-view=true#itemtype) | [MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| 位置 | ReadItem | [约会组织者](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.7&preserve-view=true#location) | [位置](/javascript/api/outlook/office.location) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [约会参与者](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.7&preserve-view=true#location) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [邮件阅读](/javascript/api/outlook/office.messageread?view=outlook-js-1.7&preserve-view=true#location)<br> (会议请求)  | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| normalizedSubject | ReadItem | [约会参与者](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.7&preserve-view=true#normalizedsubject) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [邮件阅读](/javascript/api/outlook/office.messageread?view=outlook-js-1.7&preserve-view=true#normalizedsubject) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| notificationMessages | ReadItem | [约会组织者](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.7&preserve-view=true#notificationmessages) | [NotificationMessages](/javascript/api/outlook/office.notificationmessages) | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| | | [约会参与者](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.7&preserve-view=true#notificationmessages) | [NotificationMessages](/javascript/api/outlook/office.notificationmessages) | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| | | [邮件撰写](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.7&preserve-view=true#notificationmessages) | [NotificationMessages](/javascript/api/outlook/office.notificationmessages) | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| | | [邮件阅读](/javascript/api/outlook/office.messageread?view=outlook-js-1.7&preserve-view=true#notificationmessages) | [NotificationMessages](/javascript/api/outlook/office.notificationmessages) | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| optionalAttendees | ReadItem | [约会组织者](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.7&preserve-view=true#optionalattendees) | [收件人](/javascript/api/outlook/office.recipients) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [约会参与者](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.7&preserve-view=true#optionalattendees) | Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| organizer － 组织者 | ReadWriteItem | [约会组织者](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.7&preserve-view=true#organizer) | [Organizer](/javascript/api/outlook/office.organizer) | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | ReadItem | [约会参与者](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.7&preserve-view=true#organizer) | [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| recurrence | ReadItem | [约会组织者](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.7&preserve-view=true#recurrence) | [定期](/javascript/api/outlook/office.recurrence) | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | | [约会参与者](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.7&preserve-view=true#recurrence) | [定期](/javascript/api/outlook/office.recurrence) | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | | [邮件阅读](/javascript/api/outlook/office.messageread?view=outlook-js-1.7&preserve-view=true#recurrence)<br> (会议请求)  | [定期](/javascript/api/outlook/office.recurrence) | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| requiredAttendees | ReadItem | [约会组织者](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.7&preserve-view=true#requiredattendees) | [收件人](/javascript/api/outlook/office.recipients) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [约会参与者](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.7&preserve-view=true#requiredattendees) | Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| sender | ReadItem | [邮件阅读](/javascript/api/outlook/office.messageread?view=outlook-js-1.7&preserve-view=true#sender) | [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| seriesId | ReadItem | [约会组织者](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.7&preserve-view=true#seriesid) | String | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | | [约会参与者](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.7&preserve-view=true#seriesid) | String | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | | [邮件撰写](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.7&preserve-view=true#seriesid) | String | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | | [邮件阅读](/javascript/api/outlook/office.messageread?view=outlook-js-1.7&preserve-view=true#seriesid) | String | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| start | ReadItem | [约会组织者](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.7&preserve-view=true#start) | [Time](/javascript/api/outlook/office.time) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [约会参与者](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.7&preserve-view=true#start) | 日期 | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [邮件阅读](/javascript/api/outlook/office.messageread?view=outlook-js-1.7&preserve-view=true#start)<br> (会议请求)  | 日期 | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| subject | ReadItem | [约会组织者](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.7&preserve-view=true#subject) | [主题](/javascript/api/outlook/office.subject) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [约会参与者](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.7&preserve-view=true#subject) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [邮件撰写](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.7&preserve-view=true#subject) | [主题](/javascript/api/outlook/office.subject) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [邮件阅读](/javascript/api/outlook/office.messageread?view=outlook-js-1.7&preserve-view=true#subject) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| 更改为 | ReadItem | [邮件撰写](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.7&preserve-view=true#to) | [收件人](/javascript/api/outlook/office.recipients) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [邮件阅读](/javascript/api/outlook/office.messageread?view=outlook-js-1.7&preserve-view=true#to) | Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="methods"></a>方法

| 方法 | 最小值<br>权限级别 | 按模式显示的详细信息 | 最小值<br>要求集 |
|---|---|---|:---:|
| addFileAttachmentAsync(uri, attachmentName, [options], [callback]) | ReadWriteItem | [约会组织者](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.7&preserve-view=true#addfileattachmentasync-uri--attachmentname--options--callback-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [邮件撰写](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.7&preserve-view=true#addfileattachmentasync-uri--attachmentname--options--callback-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| addHandlerAsync(eventType, handler, [options], [callback]) | ReadItem | [约会组织者](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.7&preserve-view=true#addhandlerasync-eventtype--handler--options--callback-) | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | | [约会参与者](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.7&preserve-view=true#addhandlerasync-eventtype--handler--options--callback-) | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | | [邮件撰写](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.7&preserve-view=true#addhandlerasync-eventtype--handler--options--callback-) | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | | [邮件阅读](/javascript/api/outlook/office.messageread?view=outlook-js-1.7&preserve-view=true#addhandlerasync-eventtype--handler--options--callback-) | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| addItemAttachmentAsync(itemId, attachmentName, [options], [callback]) | ReadWriteItem | [约会组织者](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.7&preserve-view=true#additemattachmentasync-itemid--attachmentname--options--callback-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [邮件撰写](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.7&preserve-view=true#additemattachmentasync-itemid--attachmentname--options--callback-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| close() | 受限 | [约会组织者](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.7&preserve-view=true#close--) | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| | | [邮件撰写](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.7&preserve-view=true#close--) | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| displayReplyAllForm(formData) | ReadItem | [约会参与者](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.7&preserve-view=true#displayreplyallform-formdata-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [邮件阅读](/javascript/api/outlook/office.messageread?view=outlook-js-1.7&preserve-view=true#displayreplyallform-formdata-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| displayReplyForm(formData) | ReadItem | [约会参与者](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.7&preserve-view=true#displayreplyform-formdata-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [邮件阅读](/javascript/api/outlook/office.messageread?view=outlook-js-1.7&preserve-view=true#displayreplyform-formdata-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| getEntities ()  | ReadItem | [约会参与者](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.7&preserve-view=true#getentities--) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [邮件阅读](/javascript/api/outlook/office.messageread?view=outlook-js-1.7&preserve-view=true#getentities--) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| getEntitiesByType (entityType)  | 受限 | [约会参与者](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.7&preserve-view=true#getentitiesbytype-entitytype-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [邮件阅读](/javascript/api/outlook/office.messageread?view=outlook-js-1.7&preserve-view=true#getentitiesbytype-entitytype-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| getFilteredEntitiesByName (name)  | ReadItem | [约会参与者](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.7&preserve-view=true#getfilteredentitiesbyname-name-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [邮件阅读](/javascript/api/outlook/office.messageread?view=outlook-js-1.7&preserve-view=true#getfilteredentitiesbyname-name-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| getRegExMatches ()  | ReadItem | [约会参与者](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.7&preserve-view=true#getregexmatches--) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [邮件阅读](/javascript/api/outlook/office.messageread?view=outlook-js-1.7&preserve-view=true#getregexmatches--) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| getRegExMatchesByName (name)  | ReadItem | [约会参与者](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.7&preserve-view=true#getregexmatchesbyname-name-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [邮件阅读](/javascript/api/outlook/office.messageread?view=outlook-js-1.7&preserve-view=true#getregexmatchesbyname-name-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| getSelectedDataAsync (coercionType， [options]， callback)  | ReadItem | [约会组织者](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.7&preserve-view=true#getselecteddataasync-coerciontype--options--callback-) | [1.2](../requirement-set-1.2/outlook-requirement-set-1.2.md) |
| | | [邮件撰写](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.7&preserve-view=true#getselecteddataasync-coerciontype--options--callback-) | [1.2](../requirement-set-1.2/outlook-requirement-set-1.2.md) |
| getSelectedEntities ()  | ReadItem | [约会参与者](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.7&preserve-view=true#getselectedentities--) | [1.6](../requirement-set-1.6/outlook-requirement-set-1.6.md) |
| | | [邮件阅读](/javascript/api/outlook/office.messageread?view=outlook-js-1.7&preserve-view=true#getselectedentities--) | [1.6](../requirement-set-1.6/outlook-requirement-set-1.6.md) |
| getSelectedRegExMatches ()  | ReadItem | [约会参与者](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.7&preserve-view=true#getselectedregexmatches--) | [1.6](../requirement-set-1.6/outlook-requirement-set-1.6.md) |
| | | [邮件阅读](/javascript/api/outlook/office.messageread?view=outlook-js-1.7&preserve-view=true#getselectedregexmatches--) | [1.6](../requirement-set-1.6/outlook-requirement-set-1.6.md) |
| loadCustomPropertiesAsync(callback, [userContext]) | ReadItem | [约会组织者](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.7&preserve-view=true#loadcustompropertiesasync-callback--usercontext-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [约会参与者](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.7&preserve-view=true#loadcustompropertiesasync-callback--usercontext-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [邮件撰写](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.7&preserve-view=true#loadcustompropertiesasync-callback--usercontext-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [邮件阅读](/javascript/api/outlook/office.messageread?view=outlook-js-1.7&preserve-view=true#loadcustompropertiesasync-callback--usercontext-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| removeAttachmentAsync(attachmentId, [options], [callback]) | ReadWriteItem | [约会组织者](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.7&preserve-view=true#removeattachmentasync-attachmentid--options--callback-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
|  |  | [邮件撰写](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.7&preserve-view=true#removeattachmentasync-attachmentid--options--callback-) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| removeHandlerAsync(eventType, [options], [callback]) | ReadItem | [约会组织者](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.7&preserve-view=true#removehandlerasync-eventtype--options--callback-) | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | | [约会参与者](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.7&preserve-view=true#removehandlerasync-eventtype--options--callback-) | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | | [邮件撰写](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.7&preserve-view=true#removehandlerasync-eventtype--options--callback-) | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| | | [邮件阅读](/javascript/api/outlook/office.messageread?view=outlook-js-1.7&preserve-view=true#removehandlerasync-eventtype--options--callback-) | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
| saveAsync([options], callback) | ReadWriteItem | [约会组织者](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.7&preserve-view=true#saveasync-options--callback-) | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| | | [邮件撰写](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.7&preserve-view=true#saveasync-options--callback-) | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| setSelectedDataAsync(data, [options], callback) | ReadWriteItem | [约会组织者](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.7&preserve-view=true#setselecteddataasync-data--options--callback-) | [1.2](../requirement-set-1.2/outlook-requirement-set-1.2.md) |
| | | [邮件撰写](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.7&preserve-view=true#setselecteddataasync-data--options--callback-) | [1.2](../requirement-set-1.2/outlook-requirement-set-1.2.md) |

## <a name="events"></a>活动

可以使用 和 分别订阅和取消订阅以下 `addHandlerAsync` `removeHandlerAsync` 事件。

> [!IMPORTANT]
> 事件仅适用于任务窗格实现。

| [Event](/javascript/api/office/office.eventtype) | 说明 | 最小值<br>要求集 |
|---|---|:---:|
|`AppointmentTimeChanged`| 所选的约会或系列的日期或时间已更改。 | [1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) |
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
