---
title: Office命名空间 - 要求集 1.4
description: Office邮箱 API 要求集 1.4 Outlook外接程序可用的命名空间成员。
ms.date: 05/17/2021
localization_priority: Normal
ms.openlocfilehash: 0221ab09048719317c131f0204e2fc60c4f8f7d4
ms.sourcegitcommit: 0d9fcdc2aeb160ff475fbe817425279267c7ff31
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 05/21/2021
ms.locfileid: "52591021"
---
# <a name="office-mailbox-requirement-set-14"></a>Office (邮箱要求集 1.4) 

该 Office 命名空间提供所有 Office 应用中的加载项所使用的共享接口。此列表仅记录 Outlook 加载项所使用的接口。有关 Office 命名空间的完整列表，请参阅[公用 API](/javascript/api/office)。

##### <a name="requirements"></a>Requirements

|要求| 值|
|---|---|
|[最低版本的邮箱要求集](../../requirement-sets/outlook-api-requirement-sets.md)| 1.1|
|[适用的 Outlook 模式](../../../outlook/outlook-add-ins-overview.md#extension-points)| 撰写或阅读|

## <a name="properties"></a>属性

| 属性 | 模式 | 返回类型 | 最小值<br>要求集 |
|---|---|---|:---:|
| [context](office.context.md) | 撰写<br>阅读 | [Context](/javascript/api/office/office.context?view=outlook-js-1.4&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="enumerations"></a>枚举

| 枚举 | 模式 | 返回类型 | 最小值<br>要求集 |
|---|---|---|:---:|
| [AsyncResultStatus](#asyncresultstatus-string) | 撰写<br>阅读 | 字符串 | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [CoercionType](#coerciontype-string) | 撰写<br>阅读 | 字符串 | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [SourceProperty](#sourceproperty-string) | 撰写<br>阅读 | 字符串 | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="namespaces"></a>命名空间

[MailboxEnums：](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.4&preserve-view=true)包括许多Outlook枚举，例如、 `ItemType` `EntityType` 和 `AttachmentType` `RecipientType` `ResponseType` `ItemNotificationMessageType` 。

## <a name="enumeration-details"></a>枚举详细信息

#### <a name="asyncresultstatus-string"></a>AsyncResultStatus：String

指定异步调用的结果。

##### <a name="type"></a>类型

*   String

##### <a name="properties"></a>属性

|名称| 类型| 描述|
|---|---|---|
|`Succeeded`| 字符串|调用成功。|
|`Failed`| 字符串|调用失败。|

##### <a name="requirements"></a>Requirements

|要求| 值|
|---|---|
|[最低版本的邮箱要求集](../../requirement-sets/outlook-api-requirement-sets.md)| 1.1|
|[适用的 Outlook 模式](../../../outlook/outlook-add-ins-overview.md#extension-points)| 撰写或阅读|

<br>

---
---

#### <a name="coerciontype-string"></a>CoercionType：String

指定如何强制由调用方法返回或设置的数据。

##### <a name="type"></a>类型

*   String

##### <a name="properties"></a>属性

|名称| 类型| 描述|
|---|---|---|
|`Html`| 字符串|请求以 HTML 格式返回的数据。|
|`Text`| 字符串|请求以文本格式返回的数据。|

##### <a name="requirements"></a>Requirements

|要求| 值|
|---|---|
|[最低版本的邮箱要求集](../../requirement-sets/outlook-api-requirement-sets.md)| 1.1|
|[适用的 Outlook 模式](../../../outlook/outlook-add-ins-overview.md#extension-points)| 撰写或阅读|

<br>

---
---

#### <a name="sourceproperty-string"></a>SourceProperty：String

指定由调用方法返回的数据源。

##### <a name="type"></a>类型

*   String

##### <a name="properties"></a>属性

|名称| 类型| 描述|
|---|---|---|
|`Body`| 字符串|数据源来自邮件的正文。|
|`Subject`| String|数据源来自邮件的主题。|

##### <a name="requirements"></a>Requirements

|要求| 值|
|---|---|
|[最低版本的邮箱要求集](../../requirement-sets/outlook-api-requirement-sets.md)| 1.1|
|[适用的 Outlook 模式](../../../outlook/outlook-add-ins-overview.md#extension-points)| 撰写或阅读|
