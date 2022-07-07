---
title: 了解 Outlook 加载项权限
description: Outlook 加载项在清单中指定所需的权限级别，其中包括受限、ReadItem、ReadWriteItem 或 ReadWriteMailbox。
ms.date: 02/19/2020
ms.localizationpriority: medium
ms.openlocfilehash: b515ef470331a513d6b57007f372b3e4dec1d25b
ms.sourcegitcommit: 4ba5f750358c139c93eb2170ff2c97322dfb50df
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 07/06/2022
ms.locfileid: "66660226"
---
# <a name="understanding-outlook-add-in-permissions"></a>了解 Outlook 加载项权限

Outlook 外接程序在清单中指定所需的权限级别。可用级别为 **Restricted**、**ReadItem**、**ReadWriteItem** 或 **ReadWriteMailbox**。这些权限级别可累计：“**Restricted**”是最低的级别，并且每个更高级别包括所有较低级别的权限。“**ReadWriteMailbox**”包含所有受支持的权限。

在从 [AppSource](https://appsource.microsoft.com) 安装邮件加载项之前，你可以查看该邮件加载项所需的权限。你还可以在 Exchange 管理中心中查看已安装加载项所需的权限。

## <a name="restricted-permission"></a>“Restricted”权限

**Restricted** 权限是最基本级别的权限。在清单的 [权限](/javascript/api/manifest/permissions)元素中指定 **Restricted** 可以请求获取此权限。如果外接程序不请求其清单中的将特定权限，在默认情况下，Outlook 会将此权限分配给邮件外接程序。

### <a name="can-do"></a>可以执行的操作

- [仅获取项目主题或正文的特定实体](match-strings-in-an-item-as-well-known-entities.md)（电话号码、地址、URL）。

- 指定[项目激活规则](activation-rules.md#itemis-rule)，此类规则需要阅读或撰写窗体中的当前项目为特定的项目类型，或与选定项目中支持的已知实体（电话号码、地址、URL）的任何较小子集匹配的 [ItemHasKnownEntity rule](match-strings-in-an-item-as-well-known-entities.md) 规则。

- 访问与用户或项目具体信息 **无** 关的任何属性和方法。（请参阅下一部分，了解与用户或项目具体信息相关的属性和方法列表）。

### <a name="cant-do"></a>不能执行的操作

- 对联系人、电子邮件地址、会议建议或任务建议实体使用 [ItemHasKnownEntity](/javascript/api/manifest/rule#itemhasknownentity-rule) 规则。

- 使用 [ItemHasAttachment](/javascript/api/manifest/rule#itemhasattachment-rule) 或 [ItemHasRegularExpressionMatch](/javascript/api/manifest/rule#itemhasregularexpressionmatch-rule) 规则。

- 访问以下列表中与用户或邮件具体信息相关的属性和方法。尝试访问此列表中的成员将返回 **null**，并生成指明 Outlook 要求邮件外接程序具有提升的权限的错误消息。

  - [item.addFileAttachmentAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#methods)
  - [item.addItemAttachmentAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#methods)
  - [item.attachments](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties)
  - [item.bcc](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties)
  - [item.body](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties)
  - [item.cc](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties)
  - [item.from](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties)
  - [item.getRegExMatches](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#methods)
  - [item.getRegExMatchesByName](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#methods)
  - [item.optionalAttendees](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties)
  - [item.organizer](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties)
  - [item.removeAttachmentAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#methods)
  - [item.requiredAttendees](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties)
  - [item.sender](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties)
  - [item.to](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties)
  - [mailbox.getCallbackTokenAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods)
  - [mailbox.getUserIdentityTokenAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods)
  - [mailbox.makeEwsRequestAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods)
  - [mailbox.userProfile](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#properties)
  - [Body](/javascript/api/outlook/office.body) 及其所有子成员
  - [Location](/javascript/api/outlook/office.location) 及其所有子成员
  - [Recipients](/javascript/api/outlook/office.recipients) 及其所有子成员
  - [Subject](/javascript/api/outlook/office.subject) 及其所有子成员
  - [Time](/javascript/api/outlook/office.time) 及其所有子成员

## <a name="readitem-permission"></a>“ReadItem”权限

“ReadItem”**** 权限是权限模型中的下一级别权限。 在清单中的元素中 **\<Permissions\>** 指定 **ReadItem** 以请求此权限。

### <a name="can-do"></a>可以执行的操作

- 在读取或 [撰写窗体](item-data.md)[中读取当前项目的所有属性](get-and-set-item-data-in-a-compose-form.md)，例如阅读窗体中的 [item.to](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties) 和撰写窗体中的 [item.to.getAsync](/javascript/api/outlook/office.recipients#outlook-office-recipients-getasync-member(1))。

- [获取回调令牌](get-attachments-of-an-outlook-item.md)，以使用 Exchange Web 服务 (EWS) 或 [Outlook REST API](use-rest-api.md) 获取邮件附件或整个邮件。

- [编写外接程序在相应邮件上设置的自定义属性](/javascript/api/outlook/office.customproperties)。

- 从该邮件的主题或正文中[获取所有现有已知实体](match-strings-in-an-item-as-well-known-entities.md)，而不仅仅是一个子集。

- 使用 [ItemHasKnownEntity](activation-rules.md#itemhasknownentity-rule) 规则中所有的 [已知实体](/javascript/api/manifest/rule#itemhasknownentity-rule)，或者 [ItemHasRegularExpressionMatch](activation-rules.md#itemhasregularexpressionmatch-rule) 规则中的 [正则表达式](/javascript/api/manifest/rule#itemhasregularexpressionmatch-rule)。 以下示例遵循架构 v1.1。 它显示一个规则，如果在所选消息的主题或正文中找到一个或多个已知实体，则该规则将激活加载项。

  ```XML
    <Permissions>ReadItem</Permissions>
        <Rule xsi:type="RuleCollection" Mode="And">
        <Rule xsi:type="ItemIs" FormType = "Read" ItemType="Message" />
        <Rule xsi:type="RuleCollection" Mode="Or">
            <Rule xsi:type="ItemHasKnownEntity" 
                EntityType="PhoneNumber" />
            <Rule xsi:type="ItemHasKnownEntity" EntityType="Address" />
            <Rule xsi:type="ItemHasKnownEntity" EntityType="Url" />
            <Rule xsi:type="ItemHasKnownEntity" 
                EntityType="MeetingSuggestion" />
            <Rule xsi:type="ItemHasKnownEntity" 
                EntityType="TaskSuggestion" />
            <Rule xsi:type="ItemHasKnownEntity" 
                EntityType="EmailAddress" />
            <Rule xsi:type="ItemHasKnownEntity" EntityType="Contact" />
    </Rule>
  ```

### <a name="cant-do"></a>禁止事项

- **mailbox.getCallbackTokenAsync** 提供的令牌可用于：
  - 使用 Outlook REST API 更新或删除当前邮件，或访问用户邮箱中的其他任何邮件。
  - 使用 Outlook REST API 获取当前日历事件项。

- 使用以下任何 API。
  - [mailbox.makeEwsRequestAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods)
  - [item.addFileAttachmentAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#methods)
  - [item.addItemAttachmentAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#methods)
  - [item.bcc.addAsync](/javascript/api/outlook/office.recipients#outlook-office-recipients-addasync-member(1))
  - [item.bcc.setAsync](/javascript/api/outlook/office.recipients#outlook-office-recipients-setasync-member(1))
  - [item.body.prependAsync](/javascript/api/outlook/office.body#outlook-office-body-prependasync-member(1))
  - [item.body.setAsync](/javascript/api/outlook/office.body#outlook-office-body-setasync-member(1))
  - [item.body.setSelectedDataAsync](/javascript/api/outlook/office.body#outlook-office-body-setselecteddataasync-member(1))
  - [item.cc.addAsync](/javascript/api/outlook/office.recipients#outlook-office-recipients-addasync-member(1))
  - [item.cc.setAsync](/javascript/api/outlook/office.recipients#outlook-office-recipients-setasync-member(1))
  - [item.end.setAsync](/javascript/api/outlook/office.time#outlook-office-time-setasync-member(1))
  - [item.location.setAsync](/javascript/api/outlook/office.location#outlook-office-location-setasync-member(1))
  - [item.optionalAttendees.addAsync](/javascript/api/outlook/office.recipients#outlook-office-recipients-addasync-member(1))
  - [item.optionalAttendees.setAsync](/javascript/api/outlook/office.recipients#outlook-office-recipients-setasync-member(1))
  - [item.removeAttachmentAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#methods)
  - [item.requiredAttendees.addAsync](/javascript/api/outlook/office.recipients#outlook-office-recipients-addasync-member(1))
  - [item.requiredAttendees.setAsync](/javascript/api/outlook/office.recipients#outlook-office-recipients-setasync-member(1))
  - [item.start.setAsync](/javascript/api/outlook/office.time#outlook-office-time-setasync-member(1))
  - [item.subject.setAsync](/javascript/api/outlook/office.subject#outlook-office-subject-setasync-member(1))
  - [item.to.addAsync](/javascript/api/outlook/office.recipients#outlook-office-recipients-addasync-member(1))
  - [item.to.setAsync](/javascript/api/outlook/office.recipients#outlook-office-recipients-setasync-member(1))

## <a name="readwriteitem-permission"></a>ReadWriteItem 权限

在清单中的元素中 **\<Permissions\>** 指定 **ReadWriteItem** 以请求此权限。 在使用撰写方法（例如，**Message.to.addAsync** 或 **Message.to.setAsync**）的撰写窗体中激活的邮件加载项必须使用至少这个等级的权限。

### <a name="can-do"></a>允许事项

- [读取和写入正在 Outlook 中查阅或撰写的邮件的所有项目级别属性](item-data.md)。

- [添加或删除该邮件的附件](add-and-remove-attachments-to-an-item-in-a-compose-form.md)。

- 使用适用于邮件外接程序的 Office JavaScript API 的所有其他成员， **Mailbox.makeEWSRequestAsync** 除外。

### <a name="cant-do"></a>禁止事项

- **mailbox.getCallbackTokenAsync** 提供的令牌可用于：
  - 使用 Outlook REST API 更新或删除当前邮件，或访问用户邮箱中的其他任何邮件。
  - 使用 Outlook REST API 获取当前日历事件项。

- 使用 **mailbox.makeEWSRequestAsync**。

## <a name="readwritemailbox-permission"></a>“ReadWriteMailbox”权限

“ReadWriteMailbox”**** 是最高级别权限。 在清单中的元素中 **\<Permissions\>** 指定 **ReadWriteMailbox** 以请求此权限。

除了可以执行 **ReadWriteItem** 权限支持的操作外，还可以使用 **mailbox.getCallbackTokenAsync** 提供的令牌，通过 Exchange Web 服务 (EWS) 操作或 Outlook REST API 执行以下操作：

- 读取和写入用户邮箱中任何邮件的所有属性。
- 创建、读取和写入该邮箱中的任何文件夹或项目。
- 从用户邮箱发送邮件

通过 **mailbox.makeEWSRequestAsync**，可以访问以下 EWS 操作。

- [CopyItem](/exchange/client-developer/web-service-reference/copyitem-operation)
- [CreateFolder](/exchange/client-developer/web-service-reference/createfolder-operation)
- [CreateItem](/exchange/client-developer/web-service-reference/createitem-operation)
- [FindConversation](/exchange/client-developer/web-service-reference/findconversation-operation)
- [FindFolder](/exchange/client-developer/web-service-reference/findfolder-operation)
- [FindItem](/exchange/client-developer/web-service-reference/finditem-operation)
- [GetConversationItems](/exchange/client-developer/web-service-reference/getconversationitems-operation)
- [GetFolder](/exchange/client-developer/web-service-reference/getfolder-operation)
- [GetItem](/exchange/client-developer/web-service-reference/getitem-operation)
- [MarkAsJunk](/exchange/client-developer/web-service-reference/markasjunk-operation)
- [MoveItem](/exchange/client-developer/web-service-reference/moveitem-operation)
- [SendItem](/exchange/client-developer/web-service-reference/senditem-operation)
- [UpdateFolder](/exchange/client-developer/web-service-reference/updatefolder-operation)
- [UpdateItem](/exchange/client-developer/web-service-reference/updateitem-operation)

尝试执行不受支持的操作会导致错误响应发生。

## <a name="see-also"></a>另请参阅

- [Outlook 加载项的隐私、权限和安全性](../concepts/privacy-and-security.md)
- [将 Outlook 项中的字符串作为已知实体进行匹配](match-strings-in-an-item-as-well-known-entities.md)
