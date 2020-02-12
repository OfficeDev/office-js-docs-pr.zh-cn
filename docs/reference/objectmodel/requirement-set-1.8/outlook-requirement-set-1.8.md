---
title: Outlook 加载项 API 要求集 1.8
description: ''
ms.date: 12/17/2019
localization_priority: Normal
ms.openlocfilehash: 4fc47a5aec1296a88654de44856ecd69ab23cc3a
ms.sourcegitcommit: d15bca2c12732f8599be2ec4b2adc7c254552f52
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 02/12/2020
ms.locfileid: "41950983"
---
# <a name="outlook-add-in-api-requirement-set-18"></a>Outlook 加载项 API 要求集 1.8

适用于 Office 的 JavaScript API 的 Outlook 外接程序 API 子集包括您可以在 Outlook 外接程序中使用的对象、方法、属性和事件。

## <a name="whats-new-in-18"></a>1.8 中有哪些新增功能？

要求集 1.8 包括[要求集 1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) 的所有功能。 它还添加了下列功能。

- 添加了用于附件、类别、代理访问、增强位置、Internet 标头和发送时阻止功能的新 API。
- 向 Event.completed 添加了可选的 `options` 参数。
- 添加了对 AttachmentsChanged 和 EnhancedLocationsChanged 事件的支持。

### <a name="change-log"></a>更改日志

- 添加了 [AttachmentContent](/javascript/api/outlook/office.attachmentcontent?view=outlook-js-1.8)：新增了一个表示附件内容的对象。
- 添加了 [Categories](/javascript/api/outlook/office.categories?view=outlook-js-1.8)：新增了一个表示项目类别的对象。
- 添加了 [CategoryDetails](/javascript/api/outlook/office.categorydetails?view=outlook-js-1.8)：新增了一个表示类别详细信息（其名称以及对应的颜色）的对象。
- 添加了 [EnhancedLocation](/javascript/api/outlook/office.enhancedlocation?view=outlook-js-1.8)：新增了一个表示约会位置集的对象。
- 添加了 [InternetHeaders](/javascript/api/outlook/office.internetheaders?view=outlook-js-1.8)：新增了一个表示邮件项目的 Internet 标头的对象。 仅限撰写模式。
- 添加了 [LocationDetails](/javascript/api/outlook/office.locationdetails?view=outlook-js-1.8)：新增了一个表示位置的对象。 只读。
- 添加了 [LocationIdentifier](/javascript/api/outlook/office.locationidentifier?view=outlook-js-1.8)：新增了一个表示位置 ID 的对象。
- 添加了 [MasterCategories](/javascript/api/outlook/office.mastercategories?view=outlook-js-1.8)：新增了一个表示邮箱上类别主列表的对象。
- 添加了 [SharedProperties](/javascript/api/outlook/office.sharedproperties?view=outlook-js-1.8)：新增了一个表示共享文件夹、日历或邮箱中约会或邮件项目的属性的对象。
- 添加了 [SupportsSharedFolders 清单元素](../../manifest/supportssharedfolders.md)：添加了 [DesktopFormFactor](../../manifest/desktopformfactor.md) 清单元素的子元素。 它定义了是否可在代理场景中使用加载项。
- 添加了 [Office.context.mailbox.masterCategories](office.context.mailbox.md#properties)：新增了一个表示邮箱上类别主列表的属性。
- 添加了 [Office.context.mailbox.item.categories](office.context.mailbox.item.md#properties)：新增了一个表示项目上类别集的属性。
- 添加了 [Office.context.mailbox.item.addFileAttachmentFromBase64Async](office.context.mailbox.item.md#methods)：新增了一个方法，它可将 base64 编码字符串形式的文件附加到邮件或约会。
- 添加了 [Office.context.mailbox.item.enhancedLocation](office.context.mailbox.item.md#properties)：新增了一个表示约会位置集的属性。
- 添加了 [Office。 context. getAllInternetHeadersAsync](office.context.mailbox.item.md#methods)：新增了一个为邮件项目获取所有 Internet 标头的方法。 仅限阅读模式。
- 添加了 [Office.context.mailbox.item.getAttachmentContentAsync](office.context.mailbox.item.md#methods)：新增了一个方法，用于获取特定附件的内容。
- 添加了 [Office.context.mailbox.item.getAttachmentsAsync](office.context.mailbox.item.md#methods)：新增了一个可在撰写模式下获取邮件附件的方法。
- 添加了 [Office。 context. getItemIdAsync](office.context.mailbox.item.md#methods)：新增了一个可获取已保存的约会或邮件项目的 ID 的方法。
- 添加了 [Office.context.mailbox.item.getSharedPropertiesAsync](office.context.mailbox.item.md#methods)：新增了一个方法，它可获取显示约会或邮件项目的 sharedProperties 的对象。
- 添加了 [Office.context.mailbox.item.internetHeaders](office.context.mailbox.item.md#properties)：新增了一个可显示邮件项目上的 Internet 标头的属性。 仅限撰写模式。
- 修改了 [Event.completed](/javascript/api/office/office.addincommands.event#completed-options-)：添加了一个新的可选参数 `options`，它是具有一个有效值 `allowEvent` 的字典。 此值可用于取消执行事件。
- 添加了 [Office.MailboxEnums.AttachmentContentFormat](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.8)：新增了一个指定应用于附件内容的格式设置的枚举。
- 添加了 [Office.MailboxEnums.AttachmentStatus](/javascript/api/outlook/office.mailboxenums.attachmentstatus?view=outlook-js-1.8)：新增了一个指定是添加附件还是从邮件中删除附件的枚举。
- 添加了 [Office.MailboxEnums.CategoryColor](/javascript/api/outlook/office.mailboxenums.categorycolor?view=outlook-js-1.8)：新增了一个指定可用于与类别关联的颜色的枚举。
- 添加了 [Office.MailboxEnums.DelegatePermissions](/javascript/api/outlook/office.mailboxenums.delegatepermissions?view=outlook-js-1.8)：新增了一个指定代理权限的位标记枚举。
- 添加了 [Office.MailboxEnums.LocationType](/javascript/api/outlook/office.mailboxenums.locationtype?view=outlook-js-1.8)：新增了一个指定约会位置的类型的枚举。
- 修改了 [Office.EventType](/javascript/api/office/office.eventtype)：添加对 `AttachmentsChanged` 和 `EnhancedLocationsChanged` 事件的支持。

## <a name="see-also"></a>另请参阅

- [Outlook 加载项](/outlook/add-ins/)
- [Outlook 外接程序代码示例](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [入门](/outlook/add-ins/quick-start)
- [要求集和支持的客户端](../../requirement-sets/outlook-api-requirement-sets.md)
