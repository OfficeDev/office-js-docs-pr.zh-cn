---
title: Office 命名空间 - 预览要求集
description: 使用邮箱 API preview 要求集的 Outlook 外接程序可用的 Office 命名空间成员。
ms.date: 03/18/2020
localization_priority: Normal
ms.openlocfilehash: 634b8593e1d1a58b61c4a330ed96611903e4a27e
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 06/08/2020
ms.locfileid: "44611608"
---
# <a name="office-mailbox-preview-requirement-set"></a>Office （邮箱预览要求集）

该 Office 命名空间提供所有 Office 应用中的加载项所使用的共享接口。此列表仅记录 Outlook 加载项所使用的接口。有关 Office 命名空间的完整列表，请参阅[公用 API](/javascript/api/office)。

##### <a name="requirements"></a>Requirements

|要求| 值|
|---|---|
|[最低版本的邮箱要求集](../../requirement-sets/outlook-api-requirement-sets.md)| 1.1|
|[适用的 Outlook 模式](../../../outlook/outlook-add-ins-overview.md#extension-points)| 撰写或阅读|

##### <a name="properties"></a>属性

| 属性 | 型号 | 返回类型 | 最低<br>要求集 |
|---|---|---|:---:|
| [context](office.context.md) | 撰写<br>Read | [Context](/javascript/api/office/office.context?view=outlook-js-preview) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

##### <a name="enumerations"></a>枚举

| 枚举 | 型号 | 返回类型 | 最低<br>要求集 |
|---|---|---|:---:|
| [AsyncResultStatus](#asyncresultstatus-string) | 撰写<br>Read | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [CoercionType](#coerciontype-string) | 撰写<br>Read | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [EventType](#eventtype-string) | 撰写<br>Read | String | [1.5](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [SourceProperty](#sourceproperty-string) | 撰写<br>Read | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

### <a name="namespaces"></a>命名空间

[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-preview)：包含许多特定于 Outlook 的枚举，例如、、、、、 `ItemType` `EntityType` `AttachmentType` `RecipientType` `ResponseType` 和 `ItemNotificationMessageType` 。

## <a name="enumeration-details"></a>枚举详细信息

#### <a name="asyncresultstatus-string"></a>AsyncResultStatus： String

指定异步调用的结果。

##### <a name="type"></a>类型

*   String

##### <a name="properties"></a>属性：

|名称| 类型| 说明|
|---|---|---|
|`Succeeded`| String|调用成功。|
|`Failed`| String|调用失败。|

##### <a name="requirements"></a>Requirements

|要求| 值|
|---|---|
|[最低版本的邮箱要求集](../../requirement-sets/outlook-api-requirement-sets.md)| 1.1|
|[适用的 Outlook 模式](../../../outlook/outlook-add-ins-overview.md#extension-points)| 撰写或阅读|

<br>

---
---

#### <a name="coerciontype-string"></a>CoercionType： String

指定如何强制由调用方法返回或设置的数据。

##### <a name="type"></a>类型

*   String

##### <a name="properties"></a>属性：

|名称| 类型| 说明|
|---|---|---|
|`Html`| String|请求以 HTML 格式返回的数据。|
|`Text`| String|请求以文本格式返回的数据。|

##### <a name="requirements"></a>Requirements

|要求| 值|
|---|---|
|[最低版本的邮箱要求集](../../requirement-sets/outlook-api-requirement-sets.md)| 1.1|
|[适用的 Outlook 模式](../../../outlook/outlook-add-ins-overview.md#extension-points)| 撰写或阅读|

<br>

---
---

#### <a name="eventtype-string"></a>事件类型： String

指定与事件处理程序相关联的事件。

##### <a name="type"></a>类型

*   String

##### <a name="properties"></a>属性：

| 名称 | 类型 | Description | 最低要求集 |
|---|---|---|:---:|
|`AppointmentTimeChanged`| String | 所选的约会或系列的日期或时间已更改。 | 1.7 |
|`AttachmentsChanged`| String | 已将附件添加到项目或已从项目删除附件。 | 1.8 |
|`EnhancedLocationsChanged`| String | 所选约会的位置已更改。 | 1.8 |
|`ItemChanged`| String | 在任务窗格固定时，将选择不同的 Outlook 项进行查看。 | 1.5 |
|`OfficeThemeChanged`| String | 邮箱上的 Office 主题已更改。 | 预览 |
|`RecipientsChanged`| String | 选定项目或约会位置的收件人列表已更改。 | 1.7 |
|`RecurrenceChanged`| String | 选定系列的定期模式已更改。 | 1.7 |

##### <a name="requirements"></a>Requirements

|要求| 值|
|---|---|
|[最低版本的邮箱要求集](../../requirement-sets/outlook-api-requirement-sets.md)| 1.5 |
|[适用的 Outlook 模式](../../../outlook/outlook-add-ins-overview.md#extension-points)| 撰写或阅读|

<br>

---
---

#### <a name="sourceproperty-string"></a>SourceProperty： String

指定由调用方法返回的数据源。

##### <a name="type"></a>类型

*   String

##### <a name="properties"></a>属性：

|名称| 类型| 说明|
|---|---|---|
|`Body`| String|数据源来自邮件的正文。|
|`Subject`| String|数据源来自邮件的主题。|

##### <a name="requirements"></a>Requirements

|要求| 值|
|---|---|
|[最低版本的邮箱要求集](../../requirement-sets/outlook-api-requirement-sets.md)| 1.1|
|[适用的 Outlook 模式](../../../outlook/outlook-add-ins-overview.md#extension-points)| 撰写或阅读|
