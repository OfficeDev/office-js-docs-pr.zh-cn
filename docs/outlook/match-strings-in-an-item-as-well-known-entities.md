---
title: 将字符串作为 Outlook 加载项中的已知实体进行匹配
description: 使用 Office JavaScript API，可以获取与特定已知实体匹配的字符串以供进一步处理。
ms.date: 07/08/2022
ms.localizationpriority: medium
ms.openlocfilehash: 1e7dea12e19c86ca9db66d48a7a256b4badf8c76
ms.sourcegitcommit: d8ea4b761f44d3227b7f2c73e52f0d2233bf22e2
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 07/11/2022
ms.locfileid: "66713012"
---
# <a name="match-strings-in-an-outlook-item-as-well-known-entities"></a>将 Outlook 项中的字符串作为已知实体进行匹配

发送邮件或会议请求项之前，Exchange Server 将分析项目的内容、标识和标记类似于 Exchange 已知实体的主题和正文中的特定字符串，例如，电子邮件地址、电话号码和 URL。邮件和会议请求通过标有已知实体的 Outlook 收件箱中的 Exchange Server 传递。

使用 Office JavaScript API，可以获取与特定已知实体匹配的这些字符串，以便进一步处理。 还可以在外接程序清单中的某个规则中指定已知实体，以便当用户查看某个包含该实体匹配项的项目时，Outlook 可以激活外接程序。 然后您可以提取实体匹配项并对其执行操作。

能够识别或从所选的邮件或约会中提取此类实例是很方便的。 例如，可以构建一个反向电话查找服务，作为 Outlook 外接程序。 该外接程序可从项目主题或正文中提取类似于电话号码的字符串，执行反向搜索并显示每个电话号码的注册所有者。

本主题将介绍这些已知实体，显示基于已知实体的激活规则示例，以及如何独立使用激活规则中的实体提取实体匹配项。

## <a name="support-for-well-known-entities"></a>支持已知实体

在发件人发送项目之后和 Exchange 将项目传递给收件人之前，Exchange Server 将标记邮件或会议请求项目中的已知实体。因此，只标记在 Exchange 中传输的项目，用户查看此类项目时，Outlook 可以根据这些标记激活外接程序。反之，用户撰写项目或查看“已发送邮件”文件夹中的项目时，由于项目还没有进行传输，Outlook 无法根据已知实体激活外接程序。

同样，无法提取正在撰写的项目中和“已发送邮件”文件夹中的已知实体，因为这些项目尚未进行传输和标记。有关支持激活的项目类型的其他信息，请参阅 [Outlook 外接程序的激活规则](activation-rules.md)。

下表列出 Exchange Server 和 Outlook 支持和识别的实体（因而称作"已知实体"）和每个实体实例的对象类型。将字符串作为某一实体的自然语言识别基于某学习模型，该模型根据大量数据进行训练。因此，该识别具有不确定性。请参阅 [使用已知实体的提示](#tips-for-using-well-known-entities)来了解有关识别条件的详细信息。

**表 1.受支持的实体及其类型**

|实体类型|识别条件|对象类型|
|:-----|:-----|:-----|
|**地址**|美国街道地址；例如：1234 Main Street, Redmond, WA 07722。通常，对于要识别的地址，它应遵循美国邮政地址的结构，包含街道编号、街道名称、城市、州和邮政编码等大部分元素。可在一行或多行中指定地址。|JavaScript **String** 对象|
|**Contact**|对于在自然语言中识别的个人信息的引用。 联系人的识别取决于上下文。 例如，邮件末尾的签名或在以下信息附近出现的人员姓名：电话号码、地址、电子邮件地址和 URL。|[Contact](/javascript/api/outlook/office.contact) 对象|
|**EmailAddress**|SMTP 电子邮件地址。|JavaScript `String` 对象|
|**MeetingSuggestion**|对事件或会议的引用。例如，Exchange 2013 会将以下文本识别为会面建议： _我们明天一起吃午饭吧。_|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion) 对象|
|**PhoneNumber**|美国电话号码；例如：_(235) 555-0110_|[PhoneNumber](/javascript/api/outlook/office.phonenumber) 对象|
|**TaskSuggestion**|电子邮件中的可操作语句。例如：_请更新电子表格。_|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion) 对象|
|**Url**|明确指定了 Web 资源的网络位置和标识符的 Web 地址。 Exchange Server不需要 Web 地址中的访问协议，也不需要将链接文本中嵌入的 URL 识别为实体的`Url`实例。 Exchange Server可以匹配以下示例： `www.youtube.com/user/officevideos` `https://www.youtube.com/user/officevideos` |JavaScript `String` 对象|

下图说明了 Exchange Server 和 Outlook 如何支持加载项的已知实体，以及哪些加载项可以使用已知实体。请参阅[在加载项中检索实体](#retrieving-entities-in-your-add-in)和[根据实体的存在情况激活加载项](#activating-an-add-in-based-on-the-existence-of-an-entity)，详细了解如何使用这些实体。

**Exchange Server、Outlook 和加载项如何支持已知实体**

![支持和使用邮件应用中的已知实体。](../images/well-known-entities-info.png)

## <a name="permissions-to-extract-entities"></a>提取实体的权限

若要提取 JavaScript 代码中的实体，或根据特定已知实体的存在情况激活外接程序，请确保已在外接程序清单中请求了相应的权限。

指定默认受限权限允许外接程序提取或提取`Address``MeetingSuggestion``TaskSuggestion`实体。 若要提取任何其他实体，请指定读取项目、读/写项目或读/写邮箱权限。 若要在清单中执行此操作，请使用 [Permissions](/javascript/api/manifest/permissions) 元素并指定适当的权限&mdash;**Restricted**、 **ReadItem**、 **ReadWriteItem** 或 **ReadWriteMailbox**&mdash;，如以下示例所示。

```xml
<Permissions>ReadItem</Permissions>
```

## <a name="retrieving-entities-in-your-add-in"></a>在外接程序中检索实体

只要用户正在查看的项目的主题或正文包含 Exchange 和 Outlook 可以识别为已知实体的字符串，这些实例就可用于加载项。即使未基于已知实体激活加载项，它们也可用。 使用适当的权限，可以使用 `getEntities` 或 `getEntitiesByType` 方法检索当前消息或约会中存在的已知实体。

该 `getEntities` 方法返回包含项中所有已知实体的 [实体](/javascript/api/outlook/office.entities) 对象数组。

如果对特定类型的实体感兴趣，请使用 `getEntitiesByType`仅返回所需实体数组的方法。 [EntityType](/javascript/api/outlook/office.mailboxenums.entitytype) 枚举表示可以提取的所有已知实体类型。

调用 `getEntities`后，可以使用对象的 `Entities` 相应属性获取实体类型的实例数组。 根据实体的类型，数组中的实例可以只是字符串，也可以映射到特定对象。

作为前面的图中的示例，若要获取该项目中的地址，请访问由 `getEntities().addresses[]` 返回的数组。 该 `Entities.addresses` 属性返回 Outlook 识别为邮政地址的字符串数组。 同样，该 `Entities.contacts` 属性返回 Outlook 识别为联系人信息的 `Contact` 对象数组。 表 1 列出了每个受支持实体的实例的对象类型。

以下示例显示如何检索在邮件中发现的任何地址。

```js
// Get the address entities from the item.
const entities = Office.context.mailbox.item.getEntities();
// Check to make sure that address entities are present.
if (null != entities && null != entities.addresses && undefined != entities.addresses) {
   //Addresses are present, so use them here.
}
```

## <a name="activating-an-add-in-based-on-the-existence-of-an-entity"></a>根据实体的存在情况激活外接程序

使用已知实体的另一种方法是，根据当前查看的项目的主题或正文的一个或多个类型的实体的存在情况，使 Outlook 激活外接程序。 可以通过在外接程序清单中指定 `ItemHasKnownEntity` 规则来执行此操作。 [EntityType](/javascript/api/outlook/office.mailboxenums.entitytype) 简单类型表示规则支持的`ItemHasKnownEntity`不同类型的已知实体。 激活外接程序后，还可以根据需要检索此类实体的实例，如上一节" [在外接程序中检索实体](#retrieving-entities-in-your-add-in)"中所述。

可以选择在规则中 `ItemHasKnownEntity` 应用正则表达式，以便进一步筛选实体的实例，并让 Outlook 仅在实体实例的子集上激活加载项。 例如，可为邮件中包含以"98"开头的华盛顿州邮政编码的街道地址实体指定筛选器。 若要对实体实例应用筛选器，请使用 `RegExFilter` [ItemHasKnownEntity](/javascript/api/manifest/rule#itemhasknownentity-rule) 类型的元素中的`Rule`和`FilterName`属性。

类似于其他激活规则，您可以指定多个规则，为外接程序形成一个规则集合。 以下示例对 2 个规则应用“AND”操作：规则 `ItemIs` 和 `ItemHasKnownEntity` 规则。 只要当前项目为邮件，且 Outlook 识别该项目主题或正文中的地址时，此规则集合就将激活外接程序。

```XML
<Rule xsi:type="RuleCollection" Mode="And">
   <Rule xsi:type="ItemIs" ItemType="Message" />
   <Rule xsi:type="ItemHasKnownEntity" EntityType="Address" />
</Rule>
```

下面的示例使用 `getEntitiesByType` 当前项将变量 `addresses` 设置为上述规则集合的结果。

```js
const addresses = Office.context.mailbox.item.getEntitiesByType(Office.MailboxEnums.EntityType.Address);
```

下面 `ItemHasKnownEntity` 的规则示例在当前项的主题或正文中有 URL 时激活加载项，并且 URL 包含字符串“youtube”，而不管字符串的大小写如何。

```XML
<Rule xsi:type="ItemHasKnownEntity" 
    EntityType="Url" 
    RegExFilter="youtube"
    FilterName="youtube"
    IgnoreCase="true"/>
```

下面的示例使用 `getFilteredEntitiesByName(name)` 当前项设置变量 `videos` ，以获取与前面 `ItemHasKnownEntity` 规则中的正则表达式匹配的结果数组。

```js
const videos = Office.context.mailbox.item.getFilteredEntitiesByName(youtube);
```

## <a name="tips-for-using-well-known-entities"></a>使用已知实体的提示

在外接程序中使用已知实体时，应了解一些事实和限制。 只要在用户读取包含已知实体匹配项的项目时激活外接程序，无论你是否使用 `ItemHasKnownEntity` 规则，都适用以下内容。

- 仅当字符串为英文形式时，您才可以提取已知实体字符串。

- 您可以从项目正文的前 2,000 个字符中提取已知实体，但不能超过此限制。 此大小限制有助于平衡功能和性能之间的需求，因此 Exchange Server 和 Outlook 不会因分析和确定大型邮件和约会中的已知实体实例而停滞。 请注意，此限制与加载项是否指定 `ItemHasKnownEntity` 规则无关。 如果外接程序使用此类规则，还要注意以下项目 2 中针对 Outlook 富客户端的的规则处理限制。

- 您可以从约会（由邮箱所有者之外的人员组织的会议）中提取实体。如果日历项目不是会议或不是由邮箱所有者组织的会议，则不能从中提取实体。

- 可以仅从消息中提取类型的 `MeetingSuggestion` 实体，但不能从约会中提取实体。

- 您可以提取项目正文中明确存在的 URL，但无法提取 HTML 项目正文中内嵌在超链接文本中的 URL。 请考虑改用 `ItemHasRegularExpressionMatch` 规则来获取显式 URL 和嵌入式 URL。 指定 `BodyAsHTML` 为 _PropertyName_，以及将 URL 匹配为  _RegExValue 的_ 正则表达式。

- 不能从"已发送邮件"文件夹中的邮件提取实体。

此外，如果使用 [ItemHasKnownEntity](/javascript/api/manifest/rule#itemhasknownentity-rule) 规则，则会应用以下内容，并且可能会影响其他情况下预期外接程序会被激活的情况。

- 使用规则 `ItemHasKnownEntity` 时，无论清单中指定的默认区域设置如何，Outlook 都希望仅使用英语匹配实体字符串。

- 当你的外接程序在 Outlook 富客户端上运行时，期望 Outlook 将规则应用 `ItemHasKnownEntity` 到项目正文的第一个 MB，而不是应用到超过该限制的正文的其余部分。

- 不能使用 `ItemHasKnownEntity` 规则激活“已发送邮件”文件夹中的项的加载项。

## <a name="see-also"></a>另请参阅

- [创建适用于阅读窗体的 Outlook 加载项](read-scenario.md)
- [从 Outlook 项目中提取实体字符串](extract-entity-strings-from-an-item.md)
- [Outlook 加载项的激活规则](activation-rules.md)
- [使用正则表达式激活规则显示 Outlook 加载项](use-regular-expressions-to-show-an-outlook-add-in.md)
- [了解 Outlook 外接程序权限](understanding-outlook-add-in-permissions.md)
