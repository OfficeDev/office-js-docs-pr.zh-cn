---
title: Outlook 外接程序 API 预览要求集
description: ''
ms.date: 08/13/2019
localization_priority: Priority
ms.openlocfilehash: b563d6cfc279a18a6a61f39c33a5ab42e1bd6984
ms.sourcegitcommit: 1c7e555733ee6d5a08e444a3c4c16635d998e032
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 08/14/2019
ms.locfileid: "36395706"
---
# <a name="outlook-add-in-api-preview-requirement-set"></a>Outlook 外接程序 API 预览要求集

适用于 Office 的 JavaScript API 的 Outlook 外接程序 API 子集包括可以在 Outlook 外接程序中使用的对象、方法、属性和事件。

> [!NOTE]
> 本文档适用于**预览**[要求集](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)。 此要求集尚未完全实现，客户端不会准确报告对它的支持。 不应在外接程序清单中指定此要求集。 在此要求集中引入的方法和属性应在使用前单独测试其可用性。
>
> 要使用预览 API：
>
> - 必须参考 CDN 上的 **beta** 库 (https://appsforoffice.microsoft.com/lib/beta/hosted/office.js)。
> - 还可能需要加入 [Office 预览体验计划](https://products.office.com/office-insider)才能访问更新的 Office 版本。

预览要求集包括[要求集 1.7](../requirement-set-1.7/outlook-requirement-set-1.7.md) 的所有功能。

## <a name="features-in-preview"></a>预览阶段的功能

以下是预览版中的功能。

### <a name="attachments"></a>附件

#### <a name="attachmentcontentjavascriptapioutlookofficeattachmentcontent"></a>[AttachmentContent](/javascript/api/outlook/office.attachmentcontent)

新增了表示附件内容的对象。

**适用于**：Windows 版 Outlook（已连接到 Office 365 订阅）、Outlook 网页版（新式）、Mac 版 Outlook（已连接到 Office 365 订阅）

#### <a name="officecontextmailboxitemaddfileattachmentfrombase64asyncofficecontextmailboxitemmdaddfileattachmentfrombase64asyncbase64file-attachmentname-options-callback"></a>[Office.context.mailbox.item.addFileAttachmentFromBase64Async](office.context.mailbox.item.md#addfileattachmentfrombase64asyncbase64file-attachmentname-options-callback)

新增了一个方法，可将 base64 编码字符串形式的文件附加到邮件或约会。

**适用于**：Windows 版 Outlook（已连接到 Office 365 订阅）、Outlook 网页版（新式）、Mac 版 Outlook（已连接到 Office 365 订阅）

#### <a name="officecontextmailboxitemgetattachmentcontentasyncofficecontextmailboxitemmdgetattachmentcontentasyncattachmentid-options-callback--attachmentcontent"></a>[Office.context.mailbox.item.getAttachmentContentAsync](office.context.mailbox.item.md#getattachmentcontentasyncattachmentid-options-callback--attachmentcontent)

新增了一个方法，可获取特定附件的内容。

**适用于**：Windows 版 Outlook（已连接到 Office 365 订阅）、Outlook 网页版（新式）、Mac 版 Outlook（已连接到 Office 365 订阅）

#### <a name="officecontextmailboxitemgetattachmentsasyncofficecontextmailboxitemmdgetattachmentsasyncoptions-callback--arrayattachmentdetails"></a>[Office.context.mailbox.item.getAttachmentsAsync](office.context.mailbox.item.md#getattachmentsasyncoptions-callback--arrayattachmentdetails)

新增了一个方法，可在撰写模式下获取项目附件。

**适用于**：Windows 版 Outlook（已连接到 Office 365 订阅）、Outlook 网页版（新式）、Mac 版 Outlook（已连接到 Office 365 订阅）

#### <a name="officemailboxenumsattachmentcontentformatjavascriptapioutlookofficemailboxenumsattachmentcontentformat"></a>[Office.MailboxEnums.AttachmentContentFormat](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat)

新增了一个枚举，可指定应用于附件内容的格式设置。

**适用于**：Windows 版 Outlook（已连接到 Office 365 订阅）、Outlook 网页版（新式）、Mac 版 Outlook（已连接到 Office 365 订阅）

#### <a name="officemailboxenumsattachmentstatusjavascriptapioutlookofficemailboxenumsattachmentstatus"></a>[Office.MailboxEnums.AttachmentStatus](/javascript/api/outlook/office.mailboxenums.attachmentstatus)

新增了一个枚举，可指定将附件添加至项目还是从项目中删除附件。

**适用于**：Windows 版 Outlook（已连接到 Office 365 订阅）、Outlook 网页版（新式）、Mac 版 Outlook（已连接到 Office 365 订阅）

#### <a name="officeeventtypeattachmentschangedjavascriptapiofficeofficeeventtype"></a>[Office.EventType.AttachmentsChanged](/javascript/api/office/office.eventtype)

向 `Item` 中添加了 `AttachmentsChanged` 事件。

**适用于**：Windows 版 Outlook（已连接到 Office 365 订阅）、Outlook 网页版（新式）、Mac 版 Outlook（已连接到 Office 365 订阅）

---

### <a name="block-on-send"></a>阻止发送

#### <a name="eventcompletedjavascriptapiofficeofficeaddincommandseventcompleted-options-"></a>[Event.completed](/javascript/api/office/office.addincommands.event#completed-options-)

新增了可选参数 `options`，它是有效值为 `allowEvent` 的字典。 此值可用于取消执行事件。

**适用于**：Outlook 网页版（经典）、Windows 版 Outlook（已连接到 Office 365 订阅）、Mac 版 Outlook（已连接到 Office 365 订阅）

---

### <a name="categories"></a>类别

在 Outlook 中，用户可以使用类别对邮件和约会进行颜色编码。 用户在其邮箱的主列表中定义类别。 然后，他们可以将一个或多个类别应用于项目。

> [!NOTE]
> iOS 版 Outlook 或 Android 版 Outlook 不支持此功能。

#### <a name="categoriesjavascriptapioutlookofficecategories"></a>[Categories](/javascript/api/outlook/office.categories)

新增了一个表示项目类别的对象。

**适用于**：Windows 版 Outlook（已连接到 Office 365 订阅）、Mac 版 Outlook（已连接到 Office 365 订阅）

#### <a name="categorydetailsjavascriptapioutlookofficecategorydetails"></a>[CategoryDetails](/javascript/api/outlook/office.categorydetails)

新增了一个表示类别详细信息（其名称以及对应的颜色）的对象。

**适用于**：Windows 版 Outlook（已连接到 Office 365 订阅）、Mac 版 Outlook（已连接到 Office 365 订阅）

#### <a name="mastercategoriesjavascriptapioutlookofficemastercategories"></a>[MasterCategories](/javascript/api/outlook/office.mastercategories)

新增了一个表示邮箱上类别主列表的对象。

**适用于**：Windows 版 Outlook（已连接到 Office 365 订阅）、Mac 版 Outlook（已连接到 Office 365 订阅）

#### <a name="officecontextmailboxmastercategoriesjavascriptapioutlookofficemailboxmastercategories"></a>[Office.context.mailbox.masterCategories](/javascript/api/outlook/office.mailbox#mastercategories)

新增了一个表示邮箱上类别主列表的属性。

**适用于**：Windows 版 Outlook（已连接到 Office 365 订阅）、Mac 版 Outlook（已连接到 Office 365 订阅）

#### <a name="officecontextmailboxitemcategoriesjavascriptapioutlookofficeitemcategories"></a>[Office.context.mailbox.item.categories](/javascript/api/outlook/office.item#categories)

新增了一个表示项目上类别集的属性。

**适用于**：Windows 版 Outlook（已连接到 Office 365 订阅）、Mac 版 Outlook（已连接到 Office 365 订阅）

#### <a name="officemailboxenumscategorycolorjavascriptapioutlookofficemailboxenumscategorycolor"></a>[Office.MailboxEnums.CategoryColor](/javascript/api/outlook/office.mailboxenums.categorycolor)

新增了一个指定可用于与类别关联的颜色的枚举。

**适用于**：Windows 版 Outlook（已连接到 Office 365 订阅）、Mac 版 Outlook（已连接到 Office 365 订阅）

---

### <a name="delegate-access"></a>委托访问

#### <a name="sharedpropertiesjavascriptapioutlookofficesharedproperties"></a>[SharedProperties](/javascript/api/outlook/office.sharedproperties)

新增了一个对象，表示共享文件夹、日历或邮箱中的约会或邮件项目的属性。

**适用于**：Windows 版 Outlook（已连接到 Office 365 订阅）、Outlook 网页版（新式）、Mac 版 Outlook（已连接到 Office 365 订阅）

#### <a name="officecontextmailboxitemgetitemidasyncofficecontextmailboxitemmdgetitemidasyncoptions-callback"></a>[Office.context.mailbox.item.getItemIdAsync](office.context.mailbox.item.md#getitemidasyncoptions-callback)

添加了用于获取已保存约会或邮件项目的 ID 的新方法。

**适用于**：Windows 版 Outlook（已连接到 Office 365 订阅）、Outlook 网页版（新式）、Mac 版 Outlook（已连接到 Office 365 订阅）

#### <a name="officecontextmailboxitemgetsharedpropertiesasyncofficecontextmailboxitemmdgetsharedpropertiesasyncoptions-callback"></a>[Office.context.mailbox.item.getSharedPropertiesAsync](office.context.mailbox.item.md#getsharedpropertiesasyncoptions-callback)

新增了一个对象，用于获取表示约会或邮件项目的 sharedProperties 的对象。

**适用于**：Windows 版 Outlook（已连接到 Office 365 订阅）、Outlook 网页版（新式）、Mac 版 Outlook（已连接到 Office 365 订阅）

#### <a name="officemailboxenumsdelegatepermissionsjavascriptapioutlookofficemailboxenumsdelegatepermissions"></a>[Office.MailboxEnums.DelegatePermissions](/javascript/api/outlook/office.mailboxenums.delegatepermissions)

新增了一个位标志枚举，可指定委派权限。

**适用于**：Windows 版 Outlook（已连接到 Office 365 订阅）、Outlook 网页版（新式）、Mac 版 Outlook（已连接到 Office 365 订阅）

#### <a name="supportssharedfolders-manifest-elementmanifestsupportssharedfoldersmd"></a>[SupportsSharedFolders manifest element](../../manifest/supportssharedfolders.md)

向 [DesktopFormFactor](../../manifest/desktopformfactor.md) 清单元素中添加了子元素。 它定义外接程序是否在代理应用场景中可用。

**适用于**：Windows 版 Outlook（已连接到 Office 365 订阅）、Outlook 网页版（新式）、Mac 版 Outlook（已连接到 Office 365 订阅）

---

### <a name="enhanced-location"></a>增强位置

#### <a name="enhancedlocationjavascriptapioutlookofficeenhancedlocation"></a>[EnhancedLocation](/javascript/api/outlook/office.enhancedlocation)

新增了一个对象，显示约会的位置。

**适用于**：Windows 版 Outlook（已连接到 Office 365 订阅）、Outlook 网页版（新式）、Mac 版 Outlook（已连接到 Office 365 订阅）

#### <a name="locationdetailsjavascriptapioutlookofficelocationdetails"></a>[LocationDetails](/javascript/api/outlook/office.locationdetails)

新增了一个表示位置的对象。 只读。

**适用于**：Windows 版 Outlook（已连接到 Office 365 订阅）、Outlook 网页版（新式）、Mac 版 Outlook（已连接到 Office 365 订阅）

#### <a name="locationidentifierjavascriptapioutlookofficelocationidentifier"></a>[LocationIdentifier](/javascript/api/outlook/office.locationidentifier)

新增了一个表示位置 ID 的对象。

**适用于**：Windows 版 Outlook（已连接到 Office 365 订阅）、Outlook 网页版（新式）、Mac 版 Outlook（已连接到 Office 365 订阅）

#### <a name="officecontextmailboxitemenhancedlocationofficecontextmailboxitemmdenhancedlocation-enhancedlocation"></a>[Office.context.mailbox.item.enhancedLocation](office.context.mailbox.item.md#enhancedlocation-enhancedlocation)

新增了一个表示约会位置的属性。

**适用于**：Windows 版 Outlook（已连接到 Office 365 订阅）、Outlook 网页版（新式）、Mac 版 Outlook（已连接到 Office 365 订阅）

#### <a name="officemailboxenumslocationtypejavascriptapioutlookofficemailboxenumslocationtype"></a>[Office.MailboxEnums.LocationType](/javascript/api/outlook/office.mailboxenums.locationtype)

新增了一个用于指定约会位置类型的枚举。

**适用于**：Windows 版 Outlook（已连接到 Office 365 订阅）、Outlook 网页版（新式）、Mac 版 Outlook（已连接到 Office 365 订阅）

#### <a name="officeeventtypeenhancedlocationschangedjavascriptapiofficeofficeeventtype"></a>[Office.EventType.EnhancedLocationsChanged](/javascript/api/office/office.eventtype)

向 `Item` 中添加了 `EnhancedLocationsChanged` 事件。

**适用于**：Windows 版 Outlook（已连接到 Office 365 订阅）、Outlook 网页版（新式）、Mac 版 Outlook（已连接到 Office 365 订阅）

---

### <a name="integration-with-actionable-messages"></a>与可操作邮件集成

#### <a name="officecontextmailboxitemgetinitializationcontextasyncofficecontextmailboxitemmdgetinitializationcontextasyncoptions-callback"></a>[Office.context.mailbox.item.getInitializationContextAsync](office.context.mailbox.item.md#getinitializationcontextasyncoptions-callback)

新增了一个函数，当外接程序[由可操作邮件激活时](/outlook/actionable-messages/invoke-add-in-from-actionable-message)，返回传递的初始化数据。

**适用于**：Windows 版 Outlook（已连接到 Office 365 订阅）、Outlook 网页版（经典）

---

### <a name="internet-headers"></a>Internet 标头：

#### <a name="internetheadersjavascriptapioutlookofficeinternetheaders"></a>[InternetHeaders](/javascript/api/outlook/office.internetheaders)

添加了一个表示邮件项目的自定义 Internet 标头的新对象。

**适用于**：Windows 版 Outlook（已连接到 Office 365 订阅）、Mac 版 Outlook（已连接到 Office 365 订阅）

#### <a name="officecontextmailboxiteminternetheadersofficecontextmailboxitemmdinternetheaders-internetheaders"></a>[Office.context.mailbox.item.internetHeaders](office.context.mailbox.item.md#internetheaders-internetheaders)

添加了一个表示邮件项目的自定义 Internet 标头的新属性。

**适用于**：Windows 版 Outlook（已连接到 Office 365 订阅）、Mac 版 Outlook（已连接到 Office 365 订阅）

---

### <a name="office-theme"></a>Office 主题

#### <a name="officecontextofficethemejavascriptapiofficeofficecontextofficetheme"></a>[Office.context.officeTheme](/javascript/api/office/office.context#officetheme)

增加了获取 Office 主题的功能。

**适用于**：Windows 版 Outlook（已连接到 Office 365 订阅）

#### <a name="officeeventtypeofficethemechangedjavascriptapiofficeofficeeventtype"></a>[Office.EventType.OfficeThemeChanged](/javascript/api/office/office.eventtype)

向 `Mailbox` 中添加了 `OfficeThemeChanged` 事件。

**适用于**：Windows 版 Outlook（已连接到 Office 365 订阅）

---

### <a name="sso"></a>SSO

#### <a name="officecontextauthgetaccesstokenasyncofficedevadd-insdevelopsso-in-office-add-inssso-api-reference"></a>[Office.context.auth.getAccessTokenAsync](/office/dev/add-ins/develop/sso-in-office-add-ins#sso-api-reference)

添加了对 `getAccessTokenAsync` 的访问，使外接程序[能够访问](/outlook/add-ins/authenticate-a-user-with-an-sso-token) Microsoft Graph API 的访问令牌。

**适用于**：Windows 版 Outlook（已连接到 Office 365 订阅）、Mac 版 Outlook（已连接到 Office 365 订阅）、Outlook 网页版（新式）、Outlook 网页版（经典）

## <a name="see-also"></a>另请参阅

- [Outlook 加载项](/outlook/add-ins/)
- [Outlook 外接程序代码示例](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [入门](/outlook/add-ins/quick-start)
- [要求集和支持的客户端](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)
