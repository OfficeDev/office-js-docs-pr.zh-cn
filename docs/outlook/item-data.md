---
title: 获取或设置 Outlook 加载项中的项目数据
description: 根据加载项是在阅读窗体中激活还是在撰写窗体中激活，项目为加载项提供的属性也有所不同。
ms.date: 12/10/2019
ms.localizationpriority: medium
ms.openlocfilehash: dbd512f45dc9e77fc4a150da4ee8b8924799670a
ms.sourcegitcommit: b66ba72aee8ccb2916cd6012e66316df2130f640
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/26/2022
ms.locfileid: "64483388"
---
# <a name="get-and-set-outlook-item-data-in-read-or-compose-forms"></a>在阅读或撰写窗体中获取和设置 Outlook 项目数据

从 Office 加载项清单架构的版本 1.1 开始，Outlook 可以在用户查看或撰写项目时激活加载项。 根据加载项是在阅读窗体中激活还是在撰写窗体中激活，项目为加载项提供的属性也有所不同。

例如，仅针对已发送项目（随后在阅读窗体中查看项目）定义 [dateTimeCreated](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties) 和 [dateTimeModified](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties) 属性，但（在撰写窗体中）创建项目时不定义这两个属性。 另一个示例是 [bcc](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties) 属性，它仅在（撰写窗体中）撰写邮件时具有意义，并且用户无法在阅读窗体中访问此属性。

## <a name="item-properties-available-in-compose-and-read-forms"></a>撰写和阅读窗体中可用的项目属性

表 1 显示了 Office JavaScript API 中的项目级属性，这些属性在邮件外接程序 (阅读和撰写) 模式下可用。通常，阅读窗体中可用的属性是只读的，撰写窗体中可用的属性是可读/写属性，[itemId](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties)、[conversationId](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties) 和 [itemType](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties) 属性除外，无论如何，这些属性始终为只读。

对于撰写窗体中的其余项目级属性，由于加载项和用户可以同时读取或写入同一属性，在撰写模式下获取或设置这些属性的方法都是异步的，因此这些属性在撰写窗体中和阅读窗体中返回的对象类型可能也有所不同。 有关在撰写模式下使用异步方法获取或设置项目级属性的详细信息，请参阅[在 Outlook 的撰写窗体中获取和设置项目数据](get-and-set-item-data-in-a-compose-form.md)。


**表 1. 撰写窗体和阅读窗体中可用的项目属性**

<br/>

|**项目类型**|**属性**|**阅读窗体中的属性类型**|**撰写窗体中的属性类型**|
|:-----|:-----|:-----|:-----|
|约会和邮件|[dateTimeCreated](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties)|JavaScript **Date** 对象|属性不可用|
|约会和邮件|[dateTimeModified](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties)|JavaScript **Date** 对象|属性不可用|
|约会和邮件|[itemClass](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties)|字符串|属性不可用|
|约会和邮件|[itemId](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties)|字符串|属性不可用|
|约会和邮件|[itemType](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties)|[ItemType](/javascript/api/outlook/office.mailboxenums.itemtype) 枚举中的字符串|[ItemType](/javascript/api/outlook/office.mailboxenums.itemtype) 枚举中的 (只读) |
|约会和邮件|[attachments](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties)|[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)|属性不可用|
|约会和邮件|[body](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties)|[Body](/javascript/api/outlook/office.body)|[Body](/javascript/api/outlook/office.body)|
|约会和邮件|[normalizedSubject](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties)|字符串|属性不可用|
|约会和邮件|[subject](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties)|字符串|[Subject](/javascript/api/outlook/office.subject)|
|约会|[end](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties)|JavaScript **Date** 对象|[Time](/javascript/api/outlook/office.time)|
|约会|[location](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties)|字符串|[位置](/javascript/api/outlook/office.location)|
|约会|[optionalAttendees](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties)|[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[Recipients](/javascript/api/outlook/office.recipients)|
|约会|[organizer](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties)|[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[Organizer](/javascript/api/outlook/office.organizer)|
|约会|[requiredAttendees](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties)|[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[Recipients](/javascript/api/outlook/office.recipients)|
|约会|[start](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties)|JavaScript **Date** 对象|[Time](/javascript/api/outlook/office.time)|
|邮件|[bcc](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties)|属性不可用|[收件人](/javascript/api/outlook/office.recipients)|
|邮件|[cc](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties)|[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[收件人](/javascript/api/outlook/office.recipients)|
|邮件|[conversationId](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties)|字符串|字符串 (只读) |
|邮件|[from](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties)|[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[From](/javascript/api/outlook/office.from)|
|邮件|[internetMessageId](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties)|整数|属性不可用|
|邮件|[sender](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties)|[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|属性不可用|
|邮件|[to](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties)|[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[收件人](/javascript/api/outlook/office.recipients)|

## <a name="use-exchange-server-callback-tokens-from-a-read-add-in"></a>从阅读加载项使用 Exchange Server 回调令牌

如果 Outlook 加载项在读取表单中激活，则可以获取 Exchange 回调令牌。 该令牌可用于服务器端代码，以便通过 Exchange Web 服务 (EWS) 访问完整项目。

通过在加载项清单中指定 **ReadItem** 权限，可以使用 [mailbox.getCallbackTokenAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods) 方法获取 Exchange 回调令牌，使用 [mailbox.ewsUrl](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#properties) 属性获取用户邮箱 EWS 终结点的 URL，以及使用 [item.itemId](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties) 获取所选项目的 EWS ID。 然后，可以将回调令牌、EWS 终结点 URL 和 EWS 项目 ID 传递到服务器端代码，以访问 [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) 操作，从而获取项目的更多属性。


## <a name="access-ews-from-a-read-or-compose-add-in"></a>从阅读或撰写加载项访问 EWS

另外，还可以使用 [mailbox.makeEwsRequestAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods) 方法直接从加载项访问 Exchange Web 服务 (EWS) 操作 [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) 和 [UpdateItem](/exchange/client-developer/web-service-reference/updateitem-operation)。 可以使用这两个操作获取并设置指定项目的多个属性。 无论加载项已在阅读还是撰写窗体中激活，只要在加载项清单中指定了 **ReadWriteMailbox** 权限，Outlook 加载项就可以使用此方法。

有关使用 **makeEwsRequestAsync** 访问 EWS 操作的详细信息，请参阅 [从 Outlook 加载项调用 Web 服务](web-services.md)。


## <a name="see-also"></a>另请参阅

- [在 Outlook 的撰写窗体中获取和设置项目数据](get-and-set-item-data-in-a-compose-form.md)
- [从 Outlook 外接程序调用 Web 服务](web-services.md)
