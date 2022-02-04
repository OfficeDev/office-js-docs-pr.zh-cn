---
title: Office.context.mailbox.item - 要求集 1.6
description: Outlook邮箱 API 要求集 1.6 版本的项目对象模型。
ms.date: 07/16/2021
ms.localizationpriority: medium
---

# <a name="item-mailbox-requirement-set-16"></a>item (Mailbox requirement set 1.6) 

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
| attachments | ReadItem | [约会参与者](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.6&preserve-view=true#outlook-office-appointmentread-attachments-member) | Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.6&preserve-view=true)> | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [邮件阅读](/javascript/api/outlook/office.messageread?view=outlook-js-1.6&preserve-view=true#outlook-office-messageread-attachments-member) | Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails?view=outlook-js-1.6&preserve-view=true)> | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| bcc | ReadItem | [邮件撰写](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.6&preserve-view=true#outlook-office-messagecompose-bcc-member) | [收件人](/javascript/api/outlook/office.recipients?view=outlook-js-1.6&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| body | ReadItem | [约会组织者](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.6&preserve-view=true#outlook-office-appointmentcompose-body-member) | [正文](/javascript/api/outlook/office.body?view=outlook-js-1.6&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [约会参与者](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.6&preserve-view=true#outlook-office-appointmentread-body-member) | [正文](/javascript/api/outlook/office.body?view=outlook-js-1.6&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [邮件撰写](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.6&preserve-view=true#outlook-office-messagecompose-body-member) | [正文](/javascript/api/outlook/office.body?view=outlook-js-1.6&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [邮件阅读](/javascript/api/outlook/office.messageread?view=outlook-js-1.6&preserve-view=true#outlook-office-messageread-body-member) | [正文](/javascript/api/outlook/office.body?view=outlook-js-1.6&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| cc | ReadItem | [邮件撰写](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.6&preserve-view=true#outlook-office-messagecompose-cc-member) | [收件人](/javascript/api/outlook/office.recipients?view=outlook-js-1.6&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [邮件阅读](/javascript/api/outlook/office.messageread?view=outlook-js-1.6&preserve-view=true#outlook-office-messageread-cc-member) | Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6&preserve-view=true)> | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| conversationId | ReadItem | [邮件撰写](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.6&preserve-view=true#outlook-office-messagecompose-conversationid-member) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [邮件阅读](/javascript/api/outlook/office.messageread?view=outlook-js-1.6&preserve-view=true#outlook-office-messageread-conversationid-member) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| dateTimeCreated | ReadItem | [约会参与者](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.6&preserve-view=true#outlook-office-appointmentread-datetimecreated-member) | Date | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [邮件阅读](/javascript/api/outlook/office.messageread?view=outlook-js-1.6&preserve-view=true#outlook-office-messageread-datetimecreated-member) | Date | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| dateTimeModified | ReadItem | [约会参与者](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.6&preserve-view=true#outlook-office-appointmentread-datetimemodified-member) | Date | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [邮件阅读](/javascript/api/outlook/office.messageread?view=outlook-js-1.6&preserve-view=true#outlook-office-messageread-datetimemodified-member) | Date | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| end | ReadItem | [约会组织者](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.6&preserve-view=true#outlook-office-appointmentcompose-end-member) | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.6&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [约会参与者](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.6&preserve-view=true#outlook-office-appointmentread-end-member) | Date | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [邮件阅读](/javascript/api/outlook/office.messageread?view=outlook-js-1.6&preserve-view=true#outlook-office-messageread-end-member)<br> (会议请求)  | Date | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| 起始数量 | ReadItem | [邮件阅读](/javascript/api/outlook/office.messageread?view=outlook-js-1.6&preserve-view=true#outlook-office-messageread-from-member) | [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| internetMessageId | ReadItem | [邮件阅读](/javascript/api/outlook/office.messageread?view=outlook-js-1.6&preserve-view=true#outlook-office-messageread-internetmessageid-member) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| itemClass | ReadItem | [约会参与者](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.6&preserve-view=true#outlook-office-appointmentread-itemclass-member) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [邮件阅读](/javascript/api/outlook/office.messageread?view=outlook-js-1.6&preserve-view=true#outlook-office-messageread-itemclass-member) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| itemId | ReadItem | [约会参与者](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.6&preserve-view=true#outlook-office-appointmentread-itemid-member) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [邮件阅读](/javascript/api/outlook/office.messageread?view=outlook-js-1.6&preserve-view=true#outlook-office-messageread-itemid-member) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| itemType | ReadItem | [约会组织者](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.6&preserve-view=true#outlook-office-appointmentcompose-itemtype-member) | [MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.6&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [约会参与者](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.6&preserve-view=true#outlook-office-appointmentread-itemtype-member) | [MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.6&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [邮件撰写](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.6&preserve-view=true#outlook-office-messagecompose-itemtype-member) | [MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.6&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [邮件阅读](/javascript/api/outlook/office.messageread?view=outlook-js-1.6&preserve-view=true#outlook-office-messageread-itemtype-member) | [MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype?view=outlook-js-1.6&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| 位置 | ReadItem | [约会组织者](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.6&preserve-view=true#outlook-office-appointmentcompose-location-member) | [位置](/javascript/api/outlook/office.location?view=outlook-js-1.6&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [约会参与者](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.6&preserve-view=true#outlook-office-appointmentread-location-member) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [邮件阅读](/javascript/api/outlook/office.messageread?view=outlook-js-1.6&preserve-view=true#outlook-office-messageread-location-member)<br> (会议请求)  | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| normalizedSubject | ReadItem | [约会参与者](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.6&preserve-view=true#outlook-office-appointmentread-normalizedsubject-member) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [邮件阅读](/javascript/api/outlook/office.messageread?view=outlook-js-1.6&preserve-view=true#outlook-office-messageread-normalizedsubject-member) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| notificationMessages | ReadItem | [约会组织者](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.6&preserve-view=true#outlook-office-appointmentcompose-notificationmessages-member) | [NotificationMessages](/javascript/api/outlook/office.notificationmessages?view=outlook-js-1.6&preserve-view=true) | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| | | [约会参与者](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.6&preserve-view=true#outlook-office-appointmentread-notificationmessages-member) | [NotificationMessages](/javascript/api/outlook/office.notificationmessages?view=outlook-js-1.6&preserve-view=true) | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| | | [邮件撰写](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.6&preserve-view=true#outlook-office-messagecompose-notificationmessages-member) | [NotificationMessages](/javascript/api/outlook/office.notificationmessages?view=outlook-js-1.6&preserve-view=true) | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| | | [邮件阅读](/javascript/api/outlook/office.messageread?view=outlook-js-1.6&preserve-view=true#outlook-office-messageread-notificationmessages-member) | [NotificationMessages](/javascript/api/outlook/office.notificationmessages?view=outlook-js-1.6&preserve-view=true) | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| optionalAttendees | ReadItem | [约会组织者](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.6&preserve-view=true#outlook-office-appointmentcompose-optionalattendees-member) | [收件人](/javascript/api/outlook/office.recipients?view=outlook-js-1.6&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [约会参与者](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.6&preserve-view=true#outlook-office-appointmentread-optionalattendees-member) | Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6&preserve-view=true)> | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| 组织者 | ReadItem | [约会参与者](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.6&preserve-view=true#outlook-office-appointmentread-organizer-member) | [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| requiredAttendees | ReadItem | [约会组织者](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.6&preserve-view=true#outlook-office-appointmentcompose-requiredattendees-member) | [收件人](/javascript/api/outlook/office.recipients?view=outlook-js-1.6&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [约会参与者](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.6&preserve-view=true#outlook-office-appointmentread-requiredattendees-member) | Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6&preserve-view=true)> | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| sender | ReadItem | [邮件阅读](/javascript/api/outlook/office.messageread?view=outlook-js-1.6&preserve-view=true#outlook-office-messageread-sender-member) | [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| start | ReadItem | [约会组织者](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.6&preserve-view=true#outlook-office-appointmentcompose-start-member) | [Time](/javascript/api/outlook/office.time?view=outlook-js-1.6&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [约会参与者](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.6&preserve-view=true#outlook-office-appointmentread-start-member) | Date | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [邮件阅读](/javascript/api/outlook/office.messageread?view=outlook-js-1.6&preserve-view=true#outlook-office-messageread-start-member)<br> (会议请求)  | Date | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| subject | ReadItem | [约会组织者](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.6&preserve-view=true#outlook-office-appointmentcompose-subject-member) | [主题](/javascript/api/outlook/office.subject?view=outlook-js-1.6&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [约会参与者](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.6&preserve-view=true#outlook-office-appointmentread-subject-member) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [邮件撰写](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.6&preserve-view=true#outlook-office-messagecompose-subject-member) | [主题](/javascript/api/outlook/office.subject?view=outlook-js-1.6&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [邮件阅读](/javascript/api/outlook/office.messageread?view=outlook-js-1.6&preserve-view=true#outlook-office-messageread-subject-member) | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| 更改为 | ReadItem | [邮件撰写](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.6&preserve-view=true#outlook-office-messagecompose-to-member) | [收件人](/javascript/api/outlook/office.recipients?view=outlook-js-1.6&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [邮件阅读](/javascript/api/outlook/office.messageread?view=outlook-js-1.6&preserve-view=true#outlook-office-messageread-to-member) | Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6&preserve-view=true)> | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="methods"></a>方法

| 方法 | 最小值<br>权限级别 | 按模式显示的详细信息 | 最小值<br>要求集 |
|---|---|---|:---:|
| addFileAttachmentAsync(uri, attachmentName, [options], [callback]) | ReadWriteItem | [约会组织者](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.6&preserve-view=true#outlook-office-appointmentcompose-addfileattachmentasync-member(1)) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [邮件撰写](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.6&preserve-view=true#outlook-office-messagecompose-addfileattachmentasync-member(1)) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| addItemAttachmentAsync(itemId, attachmentName, [options], [callback]) | ReadWriteItem | [约会组织者](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.6&preserve-view=true#outlook-office-appointmentcompose-additemattachmentasync-member(1)) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [邮件撰写](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.6&preserve-view=true#outlook-office-messagecompose-additemattachmentasync-member(1)) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| close() | 受限 | [约会组织者](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.6&preserve-view=true#outlook-office-appointmentcompose-close-member(1)) | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| | | [邮件撰写](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.6&preserve-view=true#outlook-office-messagecompose-close-member(1)) | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| displayReplyAllForm(formData) | ReadItem | [约会参与者](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.6&preserve-view=true#outlook-office-appointmentread-displayreplyallform-member(1)) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [邮件阅读](/javascript/api/outlook/office.messageread?view=outlook-js-1.6&preserve-view=true#outlook-office-messageread-displayreplyallform-member(1)) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| displayReplyForm(formData) | ReadItem | [约会参与者](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.6&preserve-view=true#outlook-office-appointmentread-displayreplyform-member(1)) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [邮件阅读](/javascript/api/outlook/office.messageread?view=outlook-js-1.6&preserve-view=true#outlook-office-messageread-displayreplyform-member(1)) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| getEntities ()  | ReadItem | [约会参与者](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.6&preserve-view=true#outlook-office-appointmentread-getentities-member(1)) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [邮件阅读](/javascript/api/outlook/office.messageread?view=outlook-js-1.6&preserve-view=true#outlook-office-messageread-getentities-member(1)) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| getEntitiesByType (entityType)  | 受限 | [约会参与者](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.6&preserve-view=true#outlook-office-appointmentread-getentitiesbytype-member(1)) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [邮件阅读](/javascript/api/outlook/office.messageread?view=outlook-js-1.6&preserve-view=true#outlook-office-messageread-getentitiesbytype-member(1)) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| getFilteredEntitiesByName (name)  | ReadItem | [约会参与者](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.6&preserve-view=true#outlook-office-appointmentread-getfilteredentitiesbyname-member(1)) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [邮件阅读](/javascript/api/outlook/office.messageread?view=outlook-js-1.6&preserve-view=true#outlook-office-messageread-getfilteredentitiesbyname-member(1)) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| getRegExMatches ()  | ReadItem | [约会参与者](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.6&preserve-view=true#outlook-office-appointmentread-getregexmatches-member(1)) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [邮件阅读](/javascript/api/outlook/office.messageread?view=outlook-js-1.6&preserve-view=true#outlook-office-messageread-getregexmatches-member(1)) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| getRegExMatchesByName (name)  | ReadItem | [约会参与者](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.6&preserve-view=true#outlook-office-appointmentread-getregexmatchesbyname-member(1)) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [邮件阅读](/javascript/api/outlook/office.messageread?view=outlook-js-1.6&preserve-view=true#outlook-office-messageread-getregexmatchesbyname-member(1)) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| getSelectedDataAsync (coercionType， [options]， callback)  | ReadItem | [约会组织者](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.6&preserve-view=true#outlook-office-appointmentcompose-getselecteddataasync-member(1)) | [1.2](../requirement-set-1.2/outlook-requirement-set-1.2.md) |
| | | [邮件撰写](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.6&preserve-view=true#outlook-office-messagecompose-getselecteddataasync-member(1)) | [1.2](../requirement-set-1.2/outlook-requirement-set-1.2.md) |
| getSelectedEntities ()  | ReadItem | [约会参与者](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.6&preserve-view=true#outlook-office-appointmentread-getselectedentities-member(1)) | [1.6](../requirement-set-1.6/outlook-requirement-set-1.6.md) |
| | | [邮件阅读](/javascript/api/outlook/office.messageread?view=outlook-js-1.6&preserve-view=true#outlook-office-messageread-getselectedentities-member(1)) | [1.6](../requirement-set-1.6/outlook-requirement-set-1.6.md) |
| getSelectedRegExMatches ()  | ReadItem | [约会参与者](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.6&preserve-view=true#outlook-office-appointmentread-getselectedregexmatches-member(1)) | [1.6](../requirement-set-1.6/outlook-requirement-set-1.6.md) |
| | | [邮件阅读](/javascript/api/outlook/office.messageread?view=outlook-js-1.6&preserve-view=true#outlook-office-messageread-getselectedregexmatches-member(1)) | [1.6](../requirement-set-1.6/outlook-requirement-set-1.6.md) |
| loadCustomPropertiesAsync(callback, [userContext]) | ReadItem | [约会组织者](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.6&preserve-view=true#outlook-office-appointmentcompose-loadcustompropertiesasync-member(1)) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [约会参与者](/javascript/api/outlook/office.appointmentread?view=outlook-js-1.6&preserve-view=true#outlook-office-appointmentread-loadcustompropertiesasync-member(1)) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [邮件撰写](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.6&preserve-view=true#outlook-office-messagecompose-loadcustompropertiesasync-member(1)) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| | | [邮件阅读](/javascript/api/outlook/office.messageread?view=outlook-js-1.6&preserve-view=true#outlook-office-messageread-loadcustompropertiesasync-member(1)) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| removeAttachmentAsync(attachmentId, [options], [callback]) | ReadWriteItem | [约会组织者](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.6&preserve-view=true#outlook-office-appointmentcompose-removeattachmentasync-member(1)) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
|  |  | [邮件撰写](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.6&preserve-view=true#outlook-office-messagecompose-removeattachmentasync-member(1)) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| saveAsync([options], callback) | ReadWriteItem | [约会组织者](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.6&preserve-view=true#outlook-office-appointmentcompose-saveasync-member(1)) | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| | | [邮件撰写](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.6&preserve-view=true#outlook-office-messagecompose-saveasync-member(1)) | [1.3](../requirement-set-1.3/outlook-requirement-set-1.3.md) |
| setSelectedDataAsync(data, [options], callback) | ReadWriteItem | [约会组织者](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-1.6&preserve-view=true#outlook-office-appointmentcompose-setselecteddataasync-member(1)) | [1.2](../requirement-set-1.2/outlook-requirement-set-1.2.md) |
| | | [邮件撰写](/javascript/api/outlook/office.messagecompose?view=outlook-js-1.6&preserve-view=true#outlook-office-messagecompose-setselecteddataasync-member(1)) | [1.2](../requirement-set-1.2/outlook-requirement-set-1.2.md) |

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
