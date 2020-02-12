---
title: Outlook 外接程序 API 要求集 1.7
description: ''
ms.date: 12/17/2019
localization_priority: Normal
ms.openlocfilehash: ed4d62405b763429bfef130f00b884444c9733d7
ms.sourcegitcommit: d15bca2c12732f8599be2ec4b2adc7c254552f52
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 02/12/2020
ms.locfileid: "41950556"
---
# <a name="outlook-add-in-api-requirement-set-17"></a>Outlook 外接程序 API 要求集 1.7

适用于 Office 的 JavaScript API 的 Outlook 外接程序 API 子集包括您可以在 Outlook 外接程序中使用的对象、方法、属性和事件。

> [!NOTE]
> 本文档适用于最新要求集之外的[要求集](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)。

## <a name="whats-new-in-17"></a>1.7 中的新增功能有哪些？

要求集 1.7 包括[要求集 1.6](../requirement-set-1.6/outlook-requirement-set-1.6.md) 的所有功能。 它还添加了下列功能。

- 添加了有关会议请求型消息和约会的定期模式的新 API。
- 修改了 item.from 属性，使其亦可用于撰写模式。
- 添加了对 RecurrenceChanged、RecipientsChanged 和 AppointmentTimeChanged 事件的支持。

### <a name="change-log"></a>更改日志

- 添加了 [From](/javascript/api/outlook/office.from?view=outlook-js-1.7)：添加了一个新对象，该对象可提供用于获取收件人值的方法。
- 添加了 [Organizer](/javascript/api/outlook/office.organizer?view=outlook-js-1.7)：添加了一个新对象，该对象可提供用于获取组织者值的方法。
- 添加了 [Recurrence](/javascript/api/outlook/office.recurrence?view=outlook-js-1.7)：添加了一个新对象，该对象可提供用于获取和设置约会的定期模式以及仅获取会议请求型消息的定期模式的方法。
- 添加了 [RecurrenceTimeZone](/javascript/api/outlook/office.recurrencetimezone?view=outlook-js-1.7)：添加了一个新对象，该对象代表定期模式的时区配置。
- 添加了 [SeriesTime](/javascript/api/outlook/office.seriestime?view=outlook-js-1.7)：添加了一个新对象，该对象可提供用于获取和设置定期系列约会的日期和时间以及获取定期系列会议请求的日期和时间的方法。
- 添加了 [Office.context.mailbox.item.addHandlerAsync](office.context.mailbox.item.md#methods)：添加了一种新方法，该方法可添加相应支持事件的事件处理程序。
- 修改了 [Office.context.mailbox.item.from](office.context.mailbox.item.md#properties)：添加在撰写模式下获取收件人值的功能。
- 修改了 [Office.context.mailbox.item.organizer](office.context.mailbox.item.md#properties)：添加在撰写模式下获取组织者值的功能。
- 添加了 [Office.context.mailbox.item.recurrence](office.context.mailbox.item.md#properties)：添加了一个新属性，该属性用于获取或设置可提供约会项目定期模式的管理方法的对象。 该属性还可用于获取会议请求项目的定期模式。
- 添加了 [Office.context.mailbox.item.removeHandlerAsync](office.context.mailbox.item.md#methods)：添加了一种新方法，该方法可删除受支持的事件类型的事件处理程序。
- 添加了 [Office.context.mailbox.item.seriesId](office.context.mailbox.item.md#properties)，添加了一个新属性，该属性可获取事件所属系列的 ID。
- 添加了 [Office.MailboxEnums.Days](/javascript/api/outlook/office.mailboxenums.days?view=outlook-js-1.7)：添加了一个新枚举，该枚举指定星期几或日期类型。
- 添加了 [Office.MailboxEnums.Month](/javascript/api/outlook/office.mailboxenums.month?view=outlook-js-1.7)：添加了一个新枚举，该枚举指定月份。
- 添加了 [Office.MailboxEnums.RecurrenceTimeZone](/javascript/api/outlook/office.mailboxenums.recurrencetimezone?view=outlook-js-1.7)：添加了一个新枚举，该枚举指定对重复周期应用的时区。
- 添加了 [Office.MailboxEnums.RecurrenceType](/javascript/api/outlook/office.mailboxenums.recurrencetype?view=outlook-js-1.7)：添加了一个新枚举，该枚举指定重复周期的类型。
- 添加了 [Office.MailboxEnums.WeekNumber](/javascript/api/outlook/office.mailboxenums.weeknumber?view=outlook-js-1.7)：添加了一个新枚举，该枚举指定是当月的第几周。
- 修改了 [Office.EventType](/javascript/api/office/office.eventtype)：添加对 `RecurrenceChanged`、`RecipientsChanged` 和 `AppointmentTimeChanged` 事件的支持。

## <a name="see-also"></a>另请参阅

- [Outlook 加载项](/outlook/add-ins/)
- [Outlook 外接程序代码示例](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [入门](/outlook/add-ins/quick-start)
- [要求集和支持的客户端](../../requirement-sets/outlook-api-requirement-sets.md)
