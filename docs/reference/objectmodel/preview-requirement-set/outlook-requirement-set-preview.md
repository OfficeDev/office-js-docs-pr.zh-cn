---
title: Outlook 外接程序 API 预览要求集
description: ''
ms.date: 03/19/2019
localization_priority: Priority
ms.openlocfilehash: d24c4647116b4af56d85a434f3ece5ccf4662a39
ms.sourcegitcommit: c5daedf017c6dd5ab0c13607589208c3f3627354
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/20/2019
ms.locfileid: "30691165"
---
# <a name="outlook-add-in-api-preview-requirement-set"></a>Outlook 外接程序 API 预览要求集

适用于 Office 的 JavaScript API 的 Outlook 外接程序 API 子集包括可以在 Outlook 外接程序中使用的对象、方法、属性和事件。

> [!NOTE]
> 本文档适用于**预览**[要求集](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)。 此要求集尚未完全实现，客户端不会准确报告对它的支持。 不应在外接程序清单中指定此要求集。 在此要求集中引入的方法和属性应在使用前单独测试其可用性。 此外，你还需要加入 [Office 预览体验成员计划](https://products.office.com/office-insider)。

预览要求集包括[要求集 1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) 的所有功能。

## <a name="features-in-preview"></a>预览阶段的功能

以下是预览版中的功能。

### <a name="add-in-commands"></a>加载项命令

#### <a name="eventcompletedjavascriptapiofficeofficeaddincommandseventcompleted-options-"></a>[Event.completed](/javascript/api/office/office.addincommands.event#completed-options-)

新增了可选参数 `options`，它是有效值为 `allowEvent` 的字典。 此值可用于取消执行事件。

**适用对象**：Outlook 网页版（经典）

### <a name="attachments"></a>附件

#### <a name="attachmentcontentjavascriptapioutlookofficeattachmentcontent"></a>[AttachmentContent](/javascript/api/outlook/office.attachmentcontent)

新增了表示附件内容的对象。

**适用于**：Outlook for Windows (Office 365)

#### <a name="officecontextmailboxitemaddfileattachmentfrombase64asyncofficecontextmailboxitemmdaddfileattachmentfrombase64asyncbase64file-attachmentname-options-callback"></a>[Office.context.mailbox.item.addFileAttachmentFromBase64Async](office.context.mailbox.item.md#addfileattachmentfrombase64asyncbase64file-attachmentname-options-callback)

新增了一个方法，可将 base64 编码字符串形式的文件附加到邮件或约会。

**适用于**：Outlook for Windows (Office 365)

#### <a name="officecontextmailboxitemgetattachmentcontentasyncofficecontextmailboxitemmdgetattachmentcontentasyncattachmentid-options-callback--attachmentcontent"></a>[Office.context.mailbox.item.getAttachmentContentAsync](office.context.mailbox.item.md#getattachmentcontentasyncattachmentid-options-callback--attachmentcontent)

新增了一个方法，可获取特定附件的内容。

**适用于**：Outlook for Windows (Office 365)

#### <a name="officecontextmailboxitemgetattachmentsasyncofficecontextmailboxitemmdgetattachmentsasyncoptions-callback--arrayattachmentdetails"></a>[Office.context.mailbox.item.getAttachmentsAsync](office.context.mailbox.item.md#getattachmentsasyncoptions-callback--arrayattachmentdetails)

新增了一个方法，可在撰写模式下获取项目附件。

**适用于**：Outlook for Windows (Office 365)

#### <a name="officemailboxenumsattachmentcontentformatjavascriptapioutlookofficemailboxenumsattachmentcontentformat"></a>[Office.MailboxEnums.AttachmentContentFormat](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat)

新增了一个枚举，可指定应用于附件内容的格式设置。

**适用于**：Outlook for Windows (Office 365)

#### <a name="officemailboxenumsattachmentstatusjavascriptapioutlookofficemailboxenumsattachmentstatus"></a>[Office.MailboxEnums.AttachmentStatus](/javascript/api/outlook/office.mailboxenums.attachmentstatus)

新增了一个枚举，可指定将附件添加至项目还是从项目中删除附件。

**适用于**：Outlook for Windows (Office 365)

#### <a name="officeeventtypeattachmentschangedjavascriptapiofficeofficeeventtype"></a>[Office.EventType.AttachmentsChanged](/javascript/api/office/office.eventtype)

向 `Item` 中添加了 `AttachmentsChanged` 事件。

**适用于**：Outlook for Windows (Office 365)

### <a name="delegate-access"></a>委托访问

#### <a name="sharedpropertiesjavascriptapioutlookofficesharedproperties"></a>[SharedProperties](/javascript/api/outlook/office.sharedproperties)

新增了一个对象，表示共享文件夹、日历或邮箱中的约会或邮件项目的属性。

**适用于**：Outlook for Windows (Office 365)

#### <a name="officecontextmailboxitemgetsharedpropertiesasyncofficecontextmailboxitemmdgetsharedpropertiesasyncoptions-callback"></a>[Office.context.mailbox.item.getSharedPropertiesAsync](office.context.mailbox.item.md#getsharedpropertiesasyncoptions-callback)

新增了一个对象，用于获取表示约会或邮件项目的 sharedProperties 的对象。

**适用于**：Outlook for Windows (Office 365)

#### <a name="officemailboxenumsdelegatepermissionsjavascriptapioutlookofficemailboxenumsdelegatepermissions"></a>[Office.MailboxEnums.DelegatePermissions](/javascript/api/outlook/office.mailboxenums.delegatepermissions)

新增了一个位标志枚举，可指定委派权限。

**适用于**：Outlook for Windows (Office 365)

#### <a name="supportssharedfolders-manifest-elementmanifestsupportssharedfoldersmd"></a>[SupportsSharedFolders manifest element](../../manifest/supportssharedfolders.md)

向 [DesktopFormFactor](../../manifest/desktopformfactor.md) 清单元素中添加了子元素。 它定义外接程序是否在代理应用场景中可用。

**适用于**：Outlook for Windows (Office 365)

### <a name="enhanced-location"></a>增强位置

#### <a name="enhancedlocationjavascriptapioutlookofficeenhancedlocation"></a>[EnhancedLocation](/javascript/api/outlook/office.enhancedlocation)

新增了一个对象，显示约会的位置。

**适用于**：Outlook for Windows (Office 365)

#### <a name="locationdetailsjavascriptapioutlookofficelocationdetails"></a>[LocationDetails](/javascript/api/outlook/office.locationdetails)

新增了一个表示位置的对象。 只读。

**适用于**：Outlook for Windows (Office 365)

#### <a name="locationidentifierjavascriptapioutlookofficelocationidentifier"></a>[LocationIdentifier](/javascript/api/outlook/office.locationidentifier)

新增了一个表示位置 ID 的对象。

**适用于**：Outlook for Windows (Office 365)

#### <a name="officecontextmailboxitemenhancedlocationofficecontextmailboxitemmdenhancedlocation-enhancedlocation"></a>[Office.context.mailbox.item.enhancedLocation](office.context.mailbox.item.md#enhancedlocation-enhancedlocation)

新增了一个表示约会位置的属性。

**适用于**：Outlook for Windows (Office 365)

#### <a name="officemailboxenumslocationtypejavascriptapioutlookofficemailboxenumslocationtype"></a>[Office.MailboxEnums.LocationType](/javascript/api/outlook/office.mailboxenums.locationtype)

新增了一个用于指定约会位置类型的枚举。

**适用于**：Outlook for Windows (Office 365)

#### <a name="officeeventtypeenhancedlocationschangedjavascriptapiofficeofficeeventtype"></a>[Office.EventType.EnhancedLocationsChanged](/javascript/api/office/office.eventtype)

向 `Item` 中添加了 `EnhancedLocationsChanged` 事件。

**适用于**：Outlook for Windows (Office 365)

### <a name="integration-with-actionable-messages"></a>与可操作邮件集成

#### <a name="officecontextmailboxitemgetinitializationcontextasyncofficecontextmailboxitemmdgetinitializationcontextasyncoptions-callback"></a>[Office.context.mailbox.item.getInitializationContextAsync](office.context.mailbox.item.md#getinitializationcontextasyncoptions-callback)

新增了一个函数，当外接程序[由可操作邮件激活时](/outlook/actionable-messages/invoke-add-in-from-actionable-message)，返回传递的初始化数据。

**适用于**：Outlook for Windows (Office 365)、Outlook 网页版（经典）

### <a name="internet-headers"></a>Internet 标头：

#### <a name="internetheadersjavascriptapioutlookofficeinternetheaders"></a>[InternetHeaders](/javascript/api/outlook/office.internetheaders)

新增了一个对象，显示邮件项目的 Internet 标头。

**适用于**：Outlook for Windows (Office 365)

#### <a name="officecontextmailboxiteminternetheadersofficecontextmailboxitemmdinternetheaders-internetheaders"></a>[Office.context.mailbox.item.internetHeaders](office.context.mailbox.item.md#internetheaders-internetheaders)

新增了一个属性，显示邮件项目的 Internet 标头。

**适用于**：Outlook for Windows (Office 365)

### <a name="office-theme"></a>Office 主题

#### <a name="officecontextmailboxofficethemejavascriptapiofficeofficeofficetheme"></a>[Office.context.mailbox.officeTheme](/javascript/api/office/office.officetheme)

增加了获取 Office 主题的功能。

**适用于**：Outlook for Windows (Office 365)

#### <a name="officeeventtypeofficethemechangedjavascriptapiofficeofficeeventtype"></a>[Office.EventType.OfficeThemeChanged](/javascript/api/office/office.eventtype)

向 `Mailbox` 中添加了 `OfficeThemeChanged` 事件。

**适用于**：Outlook for Windows (Office 365)

### <a name="sso"></a>SSO

#### <a name="officecontextauthgetaccesstokenasyncofficedevadd-insdevelopsso-in-office-add-inssso-api-reference"></a>[Office.context.auth.getAccessTokenAsync](/office/dev/add-ins/develop/sso-in-office-add-ins#sso-api-reference)

添加了对 `getAccessTokenAsync` 的访问，使外接程序[能够访问](/outlook/add-ins/authenticate-a-user-with-an-sso-token) Microsoft Graph API 的访问令牌。

**适用于**：Outlook for Windows (Office 365)、Outlook for Mac (Office 365)、Outlook 网页版（Office 365 和 Outlook.com）、Outlook 网页版（经典）

## <a name="see-also"></a>另请参阅

- [Outlook 加载项](/outlook/add-ins/)
- [Outlook 外接程序代码示例](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [入门](/outlook/add-ins/quick-start)
