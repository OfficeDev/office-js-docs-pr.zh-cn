# <a name="outlook-add-in-api-requirement-set-16"></a>Outlook 加载项 API 要求集 1.6

适用于 Office 的 JavaScript API 的 Outlook 加载项 API 子集包括可以在 Outlook 加载项中使用的对象、方法、属性和事件。

> [!NOTE]
> 本文档适用于最新要求集之外的[要求集](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)。

## <a name="whats-new-in-16"></a>1.6 中的新增功能有哪些？

要求集 1.6 包括[要求集 1.5](../requirement-set-1.5/outlook-requirement-set-1.5.md) 的所有功能。 新增了以下功能。

- 为上下文加载项添加了新 API，以获取用户选择用于激活加载项的实体或 RegEx 匹配项。
- 添加了一个新 API 以打开新的邮件窗体。
- 添加了加载项的功能，以确定用户邮箱的帐户类型。

### <a name="change-log"></a>更改日志

- 添加了[Office.context.mailbox.item.getSelectedEntities](office.context.mailbox.item.md#getselectedentities--entitiesjavascriptapioutlook16officeentities)：添加可用于从用户选择的突出显示匹配项中获取实体的新函数。 突出显示的匹配项适用于上下文加载项。
- 添加了[Office.context.mailbox.item.getSelectedRegExMatches](office.context.mailbox.item.md#getselectedregexmatches--object)：添加新函数，可用于返回突出显示匹配项中与清单 XML 文件中定义的正则表达式匹配的字符串值。 突出显示的匹配项适用于上下文加载项。
- [Office.context.mailbox.displayNewMessageForm](office.context.mailbox.md#displaynewmessageformparameters)：添加一个打开新邮件窗体的新函数。
- 添加了 [Office.context.mailbox.userProfile.accountType](office.context.mailbox.userprofile.md#accounttype-string)：将新成员添加到指示用户帐户类型的用户配置文件。

## <a name="see-also"></a>另请参阅

- [Outlook 加载项](https://docs.microsoft.com/outlook/add-ins/)
- [Outlook 加载项代码示例](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [入门](https://docs.microsoft.com/outlook/add-ins/quick-start)