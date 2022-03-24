---
title: Outlook外接程序 API 要求集 1.11
description: 加载项 API 要求集 1.11 Outlook要求集 1.11。
ms.date: 11/01/2021
ms.localizationpriority: medium
ms.openlocfilehash: 384e872b44b213b60a1b651f85ac315cd06cf082
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/23/2022
ms.locfileid: "63744144"
---
# <a name="outlook-add-in-api-requirement-set-111"></a>Outlook外接程序 API 要求集 1.11

Outlook JavaScript API 的 Office 加载项 API 子集包括可在加载项中Outlook的对象、方法、属性和事件。

## <a name="whats-new-in-111"></a>1.11 中的新增功能是什么？

要求集 1.11 包括要求集 [1.10 的所有功能](../requirement-set-1.10/outlook-requirement-set-1.10.md)。 它还添加了下列功能。

- 添加了用于基于事件 [激活的新事件](../../../outlook/autolaunch.md#supported-events)。
- 添加了 SessionData API。

### <a name="change-log"></a>更改日志

- 添加了 [Office.context.mailbox.item.sessionData](office.context.mailbox.item.md#properties)：添加了一个新属性，用于管理撰写模式下项目的会话数据。
- 添加了 [Office。SessionData](/javascript/api/outlook/office.sessiondata?view=outlook-js-1.11&preserve-view=true)：添加新对象，该对象表示撰写项目的会话数据。
- 添加了基于事件 [的新激活事件](../../../outlook/autolaunch.md#supported-events)：添加了对以下事件的支持。

  - `OnAppointmentAttachmentsChanged`
  - `OnAppointmentAttendeesChanged`
  - `OnAppointmentRecurrenceChanged`
  - `OnAppointmentTimeChanged`
  - `OnInfoBarDismissClicked`
  - `OnMessageAttachmentsChanged`
  - `OnMessageRecipientsChanged`

- 添加了 [Office。AppointmentTimeChangedEventArgs](/javascript/api/outlook/office.appointmenttimechangedeventargs?view=outlook-js-1.11&preserve-view=true)：添加支持该事件的对象`OnAppointmentTimeChanged`。
- 添加了 [Office。AttachmentsChangedEventArgs](/javascript/api/outlook/office.attachmentschangedeventargs?view=outlook-js-1.11&preserve-view=true)：添加支持 `OnAppointmentAttachmentsChanged` 和 事件`OnMessageAttachmentsChanged`的对象。
- 添加了 [Office。InfobarClickedEventArgs](/javascript/api/outlook/office.infobarclickedeventargs?view=outlook-js-1.11&preserve-view=true)：添加支持该事件`OnInfoBarDismissClicked`的对象。
- 添加了 [Office。RecipientsChangedEventArgs](/javascript/api/outlook/office.recipientschangedeventargs?view=outlook-js-1.11&preserve-view=true)：添加支持 `OnAppointmentAttendeesChanged` 和 事件`OnMessageRecipientsChanged`的对象。
- 添加了 [Office。RecurrenceChangedEventArgs](/javascript/api/outlook/office.recurrencechangedeventargs?view=outlook-js-1.11&preserve-view=true)：添加支持该事件的对象`OnAppointmentRecurrenceChanged`。

## <a name="see-also"></a>另请参阅

- [Outlook 加载项](../../../outlook/outlook-add-ins-overview.md)
- [Outlook 外接程序代码示例](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [入门](../../../quickstarts/outlook-quickstart.md)
- [要求集和支持的客户端](../../requirement-sets/outlook-api-requirement-sets.md)
