# <a name="outlook-add-in-api-requirement-set-17"></a>Outlook 加载项 API 要求集 1.7

适用于 Office 的 JavaScript API 的 Outlook 加载项 API 子集包括您可以在 Outlook 加载项中使用的对象、方法、属性和事件。

## <a name="whats-new-in-17"></a>1.7 中的新增功能有哪些？

要求集 1.7 包括[要求集 1.6](../requirement-set-1.6/outlook-requirement-set-1.6.md) 的所有功能。 新增了以下功能。

- 添加了有关约会的定期模式和会议请求的消息的新 API。
- 修改 item.from 属性，也可在撰写模式下可用。
- 添加了对 RecurrenceChanged、RecipientsChanged 和 AppointmentTimeChanged 事件的支持。

### <a name="change-log"></a>更改日志

- 添加了 [From](/javascript/api/outlook_1_7/office.from)：添加一个新对象，该对象提供获取 from 值的方法。
- 添加了 [Organizer](/javascript/api/outlook_1_7/office.organizer)：添加一个新对象，该对象提供获取 organizer 值的方法。
- 添加了 [Recurrence](/javascript/api/outlook_1_7/office.recurrence)： 添加一个新对象，该对象提供获取和设置约会的定期模式的方法，但仅获取会议请求的消息的定期模式。
- 添加了 [RecurrenceTimeZone](/javascript/api/outlook_1_7/office.recurrencetimezone)： 添加一个新对象，该对象表示定期模式的时区配置。
- 添加了 [SeriesTime](/javascript/api/outlook_1_7/office.seriestime)： 添加一个新对象，该对象提供获取和设置定期系列中约会的日期和时间的方法，并获取在定期系列中会议请求的日期和时间。
- 添加了 [Office.context.mailbox.addHandlerAsync](office.context.mailbox.item.md#addhandlerasynceventtype-handler-options-callback)：添加一个新方法，为受支持的事件添加事件处理程序。
- 修改了 [Office.context.mailbox.item.from](office.context.mailbox.item.md#from-emailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsfromjavascriptapioutlook17officefrom)：修改以在撰写模式下获取 from 值。
- 修改了 [Office.context.mailbox.item.organizer](office.context.mailbox.item.md#organizer-emailaddressdetailsjavascriptapioutlook17officeemailaddressdetailsorganizerjavascriptapioutlook17officeorganizer) - 修改以在撰写模式下获取 organizer 值。
- 添加了 [Office.context.mailbox.item.recurrence](office.context.mailbox.item.md#nullable-recurrence-recurrencejavascriptapioutlook17officerecurrence)： 添加一个获取或设置对象的新属性，该对象提供管理约会项目的定期模式的方法。 此属性还可用于获取会议请求项的定期模式。
- 添加了 [Office.context.mailbox.item.removeHandlerAsync](office.context.mailbox.item.md#removehandlerasynceventtype-handler-options-callback)： 添加一个删除事件处理程序的新方法。
- 添加了 [Office.context.mailbox.item.seriesId](office.context.mailbox.item.md#nullable-seriesid-string)：添加一个获取事件所属系列的 id 的新属性。
- 添加了 [Office.MailboxEnums.Days](/javascript/api/outlook_1_7/office.mailboxenums.days)： 添加一个指定星期几或一天的类型的新枚举。
- 添加了 [Office.MailboxEnums.Month](/javascript/api/outlook_1_7/office.mailboxenums.month)： 添加一个指定月份的新枚举。
- 添加了 [Office.MailboxEnums.RecurrenceTimeZone](/javascript/api/outlook_1_7/office.mailboxenums.recurrencetimezone)： 添加一个指定应用于重复周期的时区的新枚举。
- 添加了 [Office.MailboxEnums.RecurrenceType](/javascript/api/outlook_1_7/office.mailboxenums.recurrencetype)： 添加一个指定重复周期的类型的新枚举。
- 添加了 [Office.MailboxEnums.WeekNumber](/javascript/api/outlook_1_7/office.mailboxenums.weeknumber)： 添加一个新的枚举，指定该月的一周。
- 修改 [Office.EventType](/javascript/api/office/office.eventtype)： 修改为分别通过添加 `RecurrenceChanged`、`RecipientsChanged` 和 `AppointmentTimeChanged` 条目以支持RecurrenceChanged、 RecipientsChanged 和 AppointmentTimeChanged 事件。

## <a name="see-also"></a>另请参阅

- [Outlook 加载项](https://docs.microsoft.com/outlook/add-ins/)
- [Outlook 加载项代码示例](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [入门](https://docs.microsoft.com/outlook/add-ins/quick-start)