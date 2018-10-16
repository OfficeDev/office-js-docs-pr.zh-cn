# <a name="outlook-add-in-api-requirement-set-13"></a>Outlook 加载项 API 要求集 1.3

适用于 Office 的 JavaScript API 的 Outlook 加载项 API 子集包括您可以在 Outlook 加载项中使用的对象、方法、属性和事件。

> [!NOTE]
> 本文档适用于最新要求集之外的[要求集](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)。 

## <a name="whats-new-in-13"></a>1.3 中的新增功能有哪些？

要求集 1.3 包括[要求集 1.2](../requirement-set-1.2/outlook-requirement-set-1.2.md) 的所有功能。它添加了下列功能。

- 添加了对[加载项命令](https://docs.microsoft.com/outlook/add-ins/add-in-commands-for-outlook)的支持。
- 添加了保存或关闭正在撰写的项目的功能。
- 改进了 [Body](/javascript/api/outlook_1_3/office.body) 对象，允许加载项获取或设置整个正文。
- 添加了在 EWS 和 REST 格式之间转换 ID 的转换方法。
- 添加了将通知邮件添加到项目的信息栏中的功能。

### <a name="change-log"></a>更改日志

- 添加了 [Body.getAsync](/javascript/api/outlook_1_3/office.body#getasync-coerciontype--options--callback-)：使用指定格式返回当前正文。
- 添加了 [Body.setAsync](/javascript/api/outlook_1_3/office.body#setasync-data--options--callback-)：将整个正文替换为指定文本。
- 添加了 [Office.context.officeTheme](office.context.md#officetheme-object)：提供了对 Office 主题颜色的访问权限。
- 添加了 [Event](/javascript/api/office/office.addincommands.event) 对象：作为参数传递到 Outlook 加载项中的无用户界面命令函数。用来表示处理已完成。
- 添加了 [Office.context.mailbox.item.close](office.context.mailbox.item.md#close)：关闭正在撰写的当前项。
- 添加了 [Office.context.mailbox.item.saveAsync](office.context.mailbox.item.md#saveasyncoptions-callback)：异步保存项目。
- 添加了 [Office.context.mailbox.item.notificationMessages](office.context.mailbox.item.md#notificationmessages-notificationmessagesjavascriptapioutlook13officenotificationmessages)：获取项目的通知邮件。
- 添加了 [Office.context.mailbox.convertToEwsId](office.context.mailbox.md#converttoewsiditemid-restversion--string)：将项目 ID 格式化（从 REST 转换为 EWS 格式）。
- 添加了 [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string)：将项目 ID 格式化（从 EWS 转换为 REST 格式）。
- 添加了 [Office.MailboxEnums.ItemNotificationMessageType](/javascript/api/outlook_1_3/office.mailboxenums.itemnotificationmessagetype)：为约会或邮件指定通知邮件类型。
- 添加了 [Office.MailboxEnums.RestVersion](/javascript/api/outlook_1_3/office.mailboxenums.restversion)：指定对应于 REST 格式的项目 ID 的 REST API 的版本。
- 添加了 [NotificationMessages](/javascript/api/outlook_1_3/office.notificationmessages) 对象：提供用于访问 Outlook 加载项中的通知邮件的方法。
- 添加了 [NotificationMessageDetails](/javascript/api/outlook_1_3/office.notificationmessagedetails) 类型：由 `NotificationMessages.getAllAsync` 方法返回。

## <a name="see-also"></a>另请参阅

- [Outlook 加载项](https://docs.microsoft.com/outlook/add-ins/)
- [Outlook 加载项代码示例](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [入门](https://docs.microsoft.com/outlook/add-ins/quick-start)