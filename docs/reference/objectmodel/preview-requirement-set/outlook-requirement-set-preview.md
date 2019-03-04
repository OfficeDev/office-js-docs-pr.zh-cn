---
title: Outlook 外接程序 API 预览要求集
description: ''
ms.date: 02/26/2019
localization_priority: Priority
ms.openlocfilehash: 233bc6770faefaa0e101fd01c353e7ce0df972a1
ms.sourcegitcommit: f7f3d38ae4430e2218bf0abe7bb2976108de3579
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/01/2019
ms.locfileid: "30359245"
---
# <a name="outlook-add-in-api-preview-requirement-set"></a>Outlook 外接程序 API 预览要求集

适用于 Office 的 JavaScript API 的 Outlook 外接程序 API 子集包括可以在 Outlook 外接程序中使用的对象、方法、属性和事件。

> [!NOTE]
> 本文档适用于**预览**[要求集](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)。 此要求集尚未完全实现，客户端不会准确报告对它的支持。 不应在外接程序清单中指定此要求集。 在此要求集中引入的方法和属性应在使用前单独测试其可用性。

预览要求集包括[要求集 1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) 的所有功能。

## <a name="features-in-preview"></a>预览阶段的功能

以下是预览版中的功能。

- [AttachmentContent](/javascript/api/outlook/office.attachmentcontent) - 新增了一个对象，显示附件的内容。
- [EnhancedLocation](/javascript/api/outlook/office.enhancedlocation) - 新增了一个对象，显示约会的位置集。
- [InternetHeaders](/javascript/api/outlook/office.internetheaders) - 新增了一个对象，显示邮件项目的 Internet 标头。
- [LocationDetails](/javascript/api/outlook/office.locationdetails) - 新增了一个对象，显示位置。 只读。
- [LocationIdentifier](/javascript/api/outlook/office.locationidentifier) - 新增了一个对象，显示位置 ID。
- [SharedProperties](/javascript/api/outlook/office.sharedproperties) - 新增了一个对象，显示共享文件夹、日历或邮箱中约会或邮件项目的属性。
- [Event.completed](/javascript/api/office/office.addincommands.event#completed-options-) - 一个新的可选参数 `options`，它是具有一个有效值 `allowEvent` 的字典。该值用于取消事件的执行。
- [Office.context.mailbox.item.addFileAttachmentFromBase64Async](office.context.mailbox.item.md#addfileattachmentfrombase64asyncbase64file-attachmentname-options-callback) - 新增了一个方法，可将 base64 编码字符串形式的文件附加到邮件或约会。
- [Office.context.mailbox.item.enhancedLocation](office.context.mailbox.item.md#enhancedlocation-enhancedlocationjavascriptapioutlookofficeenhancedlocation) - 新增了一个属性，显示约会的位置集。
- [Office.context.mailbox.item.getAttachmentContentAsync](office.context.mailbox.item.md#getattachmentcontentasyncattachmentid-options-callback--attachmentcontentjavascriptapioutlookofficeattachmentcontent) - 新增了一个方法，用于获取特定附件的内容。
- [Office.context.mailbox.item.getAttachmentsAsync](office.context.mailbox.item.md#getattachmentsasyncoptions-callback--arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetails) - 新增了一个方法，可在撰写模式下获取邮件的附件。
- [Office.context.mailbox.item.getInitializationContextAsync](office.context.mailbox.item.md#getinitializationcontextasyncoptions-callback) - 新增了一个函数，当[可操作邮件激活](https://docs.microsoft.com/outlook/actionable-messages/invoke-add-in-from-actionable-message)加载项时，此函数返回传递的初始化数据。
- [Office.context.mailbox.item.getSharedPropertiesAsync](office.context.mailbox.item.md#getsharedpropertiesasyncoptions-callback) - 新增了一个方法，可获取显示约会或邮件项目的 sharedProperties 的对象。
- [Office.context.mailbox.item.internetHeaders](office.context.mailbox.item.md#internetheaders-internetheadersjavascriptapioutlookofficeinternetheaders) - 新增了一个属性，可显示邮件项目上的 Internet 标头。
- [Office.context.auth.getAccessTokenAsync](https://docs.microsoft.com/office/dev/add-ins/develop/sso-in-office-add-ins#sso-api-reference) - 现已开始支持访问 `getAccessTokenAsync`，以便外接程序能够[获取 Microsoft Graph API 的访问令牌](https://docs.microsoft.com/outlook/add-ins/authenticate-a-user-with-an-sso-token)。
- [Office.MailboxEnums.AttachmentContentFormat](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat) - 新增了一个枚举，可指定应用于附件内容的格式设置。
- [Office.MailboxEnums.AttachmentStatus](/javascript/api/outlook/office.mailboxenums.attachmentstatus) - 新增了一个枚举，可指定是添加附件还是从邮件中删除附件。
- [Office.MailboxEnums.DelegatePermissions](/javascript/api/outlook/office.mailboxenums.delegatepermissions) - 新增了一个位标记枚举，可指定代理权限。
- [Office.MailboxEnums.LocationType](/javascript/api/outlook/office.mailboxenums.locationtype) - 新增了一个枚举，可指定约会位置的类型。
- [Office.EventType](/javascript/api/office/office.eventtype) - 通过分别添加 `AttachmentsChanged`、`EnhancedLocationsChanged` 和 `OfficeThemeChanged` 条目对支持 AttachmentsChanged、EnhancedLocationsChanged 和 OfficeThemeChanged 事件进行修改。
- [SupportsSharedFolders manifest element](../../manifest/supportssharedfolders.md) - 添加了 [DesktopFormFactor](../../manifest/desktopformfactor.md) 清单元素的子元素。 它定义外接程序是否在代理应用场景中可用。

## <a name="see-also"></a>另请参阅

- [Outlook 加载项](https://docs.microsoft.com/outlook/add-ins/)
- [Outlook 外接程序代码示例](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [入门](https://docs.microsoft.com/outlook/add-ins/quick-start)
