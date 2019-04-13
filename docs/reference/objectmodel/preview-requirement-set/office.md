---
title: Office 命名空间 - 预览要求集
description: ''
ms.date: 04/12/2019
localization_priority: Normal
ms.openlocfilehash: 7effc930d196aa009c3c779b702e082ae388fada
ms.sourcegitcommit: 95ed6dfbfa680dbb40ff9757020fa7e5be4760b6
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 04/13/2019
ms.locfileid: "31838513"
---
# <a name="office"></a>Office

该 Office 命名空间提供所有 Office 应用中的加载项所使用的共享接口。此列表仅记录 Outlook 加载项所使用的接口。有关 Office 命名空间的完整列表，请参阅[公用 API](/javascript/api/office)。

##### <a name="requirements"></a>要求

|要求| 值|
|---|---|
|[最低版本的邮箱要求集](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[适用的 Outlook 模式](/outlook/add-ins/#extension-points)| 撰写或阅读|

##### <a name="members-and-methods"></a>成员和方法

| 成员 | 类型 |
|--------|------|
| [AsyncResultStatus](#asyncresultstatus-string) | Member |
| [CoercionType](#coerciontype-string) | Member |
| [EventType](#eventtype-string) | Member |
| [SourceProperty](#sourceproperty-string) | 成员 |

### <a name="namespaces"></a>命名空间

[context](office.context.md)：提供 Office 加载项 API 的上下文命名空间中的共享接口以便在 Outlook 加载项 API 中使用。

[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype)：包括 ItemType、EntityType、AttachmentType、RecipientType、ResponseType 和 ItemNotificationMessageType 枚举。

### <a name="members"></a>成员

####  <a name="asyncresultstatus-string"></a>AsyncResultStatus :String

指定异步调用的结果。

##### <a name="type"></a>类型

*   String

##### <a name="properties"></a>属性：

|名称| 类型| 说明|
|---|---|---|
|`Succeeded`| String|调用成功。|
|`Failed`| String|调用失败。|

##### <a name="requirements"></a>要求

|要求| 值|
|---|---|
|[最低版本的邮箱要求集](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[适用的 Outlook 模式](/outlook/add-ins/#extension-points)| 撰写或阅读|

---
---

####  <a name="coerciontype-string"></a>CoercionType :String

指定如何强制由调用方法返回或设置的数据。

##### <a name="type"></a>类型

*   String

##### <a name="properties"></a>属性：

|名称| 类型| 说明|
|---|---|---|
|`Html`| String|请求以 HTML 格式返回的数据。|
|`Text`| String|请求以文本格式返回的数据。|

##### <a name="requirements"></a>要求

|要求| 值|
|---|---|
|[最低版本的邮箱要求集](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[适用的 Outlook 模式](/outlook/add-ins/#extension-points)| 撰写或阅读|

---
---

####  <a name="eventtype-string"></a>EventType :String

指定与事件处理程序相关联的事件。

##### <a name="type"></a>类型

*   String

##### <a name="properties"></a>属性：

| 名称 | 类型 | 说明 | 最低要求集 |
|---|---|---|---|
|`AppointmentTimeChanged`| String | 所选的约会或系列的日期或时间已更改。 | 1.7 |
|`AttachmentsChanged`| String | 已将附件添加到项目或已从项目删除附件。 | 预览 |
|`EnhancedLocationsChanged`| String | 所选约会的位置已更改。 | 预览 |
|`ItemChanged`| String | 在任务窗格固定时，将选择不同的 Outlook 项进行查看。 | 1.5 |
|`OfficeThemeChanged`| String | 邮箱上的 Office 主题已更改。 | 预览 |
|`RecipientsChanged`| String | 选定项目或约会位置的收件人列表已更改。 | 1.7 |
|`RecurrenceChanged`| String | 选定系列的定期模式已更改。 | 1.7 |

##### <a name="requirements"></a>要求

|要求| 值|
|---|---|
|[最低版本的邮箱要求集](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.5 |
|[适用的 Outlook 模式](/outlook/add-ins/#extension-points)| 撰写或阅读 |

---
---

####  <a name="sourceproperty-string"></a>SourceProperty :String

指定由调用方法返回的数据源。

##### <a name="type"></a>类型

*   String

##### <a name="properties"></a>属性：

|名称| 类型| 说明|
|---|---|---|
|`Body`| String|数据源来自邮件的正文。|
|`Subject`| String|数据源来自邮件的主题。|

##### <a name="requirements"></a>要求

|要求| 值|
|---|---|
|[最低版本的邮箱要求集](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[适用的 Outlook 模式](/outlook/add-ins/#extension-points)| 撰写或阅读|
