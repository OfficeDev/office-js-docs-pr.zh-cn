---
title: Outlook 加载项的激活规则
description: 如果用户正在读取或撰写的邮件或约会符合加载项的激活规则，则 Outlook 将激活某些类型的加载项。
ms.date: 09/22/2020
localization_priority: Normal
ms.openlocfilehash: 24f17b7bb3da4665f3f05b23d34ba15bcc4ae729
ms.sourcegitcommit: 42c55a8d8e0447258393979a09f1ddb44c6be884
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/08/2021
ms.locfileid: "58936517"
---
# <a name="activation-rules-for-contextual-outlook-add-ins"></a>上下文 Outlook 加载项的激活规则

如果用户正在读取或撰写的邮件或约会符合外接程序的激活规则，则 Outlook 将激活某些类型的外接程序。这一点对使用 1.1 清单架构的所有外接程序均适用。然后，用户可从 Outlook UI 选择外接程序，以开始将其用于当前项目。

下图显示在“阅读”窗格中的邮件的外接程序栏中激活的 Outlook 外接程序。

![显示已激活阅读邮件应用程序的应用程序栏。](../images/read-form-app-bar.png)


## <a name="specify-activation-rules-in-a-manifest"></a>在清单中指定激活规则


若要Outlook条件激活外接程序，请通过使用下列元素之一在外接程序清单中指定激活 `Rule` 规则。

- [Rule 元素 (MailApp complexType)](../reference/manifest/rule.md) - 指定单个规则。
- [Rule 元素 (RuleCollection complexType)](../reference/manifest/rule.md#rulecollection) - 使用逻辑操作组合多个规则。


 > [!NOTE]
 > `Rule`用于指定单个规则的元素是抽象[Rule](../reference/manifest/rule.md)复杂类型。 以下每种类型的规则扩展了此抽象 `Rule` 复杂类型。 因此当你在清单中指定单个规则时，你必须使用 [xsi:type](https://www.w3.org/TR/xmlschema-1/) 属性来进一步定义某个以下类型的规则。
 > 
 > 例如，以下规则定义 [ItemIs](../reference/manifest/rule.md#itemis-rule) 规则。
 > `<Rule xsi:type="ItemIs" ItemType="Message" />`
 > 
 > `FormType`属性适用于清单 v1.1 中的激活规则，但在 `VersionOverrides` v1.0 中未定义。 因此，当在节点中使用 [ItemIs](../reference/manifest/rule.md#itemis-rule) 时，它不能 `VersionOverrides` 使用。

下表列出了可用的规则类型。你可以在表后面以及[创建适用于阅读窗体的 Outlook 外接程序](read-scenario.md)中指定的文章中查找更多信息。

<br/>

|**规则名称**|**适用的窗体**|**说明**|
|:-----|:-----|:-----|
|[ItemIs](#itemis-rule)|读取，撰写|检查当前项目是否属于指定类型（邮件或约会），另外还可以检查项目类别、窗体类型和（可选）项目邮件类别。|
|[ItemHasAttachment](#itemhasattachment-rule)|读取|检查所选项是否包含附件。|
|[ItemHasKnownEntity](#itemhasknownentity-rule)|读取|检查所选项是否包含一个或多个已知实体。更多信息：[将 Outlook 项中的字符串作为已知实体进行匹配](match-strings-in-an-item-as-well-known-entities.md)。|
|[ItemHasRegularExpressionMatch](#itemhasregularexpressionmatch-rule)|读取|检查发件人的电子邮件地址、所选项的主题和/或所选项的正文是否包含正则表达式的匹配项。更多信息： [使用正则表达式激活规则显示 Outlook 外接程序](use-regular-expressions-to-show-an-outlook-add-in.md)。|
|[RuleCollection](#rulecollection-rule)|读取，撰写|组合一组规则以便形成更复杂的规则。|

## <a name="itemis-rule"></a>ItemIs 规则

**ItemIs** 复杂类型定义一个计算结果为 **true** 的规则（如果当前项与项类型匹配）和（可选）项邮件类别（如果在规则中指明）。

在 ItemIs 规则的 属性中指定 `ItemType` 以下 **项目类型之** 一。 可以在清单中指定多个 **ItemIs** 规则。 ItemType simpleType 定义了支持 Outlook 加载项的 Outlook 项类型。

<br/>

|**Value**|**说明**|
|:-----|:-----|
|**约会**|在 Outlook 日历中指定一个项目。 这包括已获取响应并且具有组织者和参与者的会议项目，或者没有组织者或参与者且仅为日历上的一个项目的约会。 这与 Outlook 中的 IPM.Appointment 邮件类别相对应。|
|**消息**|指定通常在收件箱中收到的以下项目之一。 <ul><li><p>电子邮件。这与 Outlook 中的 IPM.Note 邮件类别相对应。</p></li><li><p>会议请求、响应或取消。 这对应于 Outlook 中的以下邮件Outlook。</p><p>IPM.Schedule.Meeting.Request</p><p>IPM.Schedule.Meeting.Neg</p><p>IPM.Schedule.Meeting.Pos</p><p>IPM.Schedule.Meeting.Tent</p><p>IPM.Schedule.Meeting.Canceled</p></li></ul>|

`FormType`属性用于指定阅读 (撰写) 外接程序应激活的模式。


 > [!NOTE]
 > ItemIs `FormType` 属性在架构 v1.1 及更高版本中定义，但不在 `VersionOverrides` v1.0 中定义。 定义外接程序 `FormType` 命令时，请勿包含 属性。

激活外接程序后，可以使用 [mailbox.item](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md) 属性获取 Outlook 中的当前所选项，以及使用 [item.itemType](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) 属性获取当前项的类型。

可以选择使用 属性指定项目的邮件类，以及使用 属性指定当项目是指定类的子类时规则是否应该 `ItemClass` `IncludeSubClasses` 为true。

若要详细了解邮件类，请参阅[项类型和邮件类](/office/vba/outlook/Concepts/Forms/item-types-and-message-classes)。

下面的示例是 **一个 ItemIs** 规则，它允许用户在阅读邮件时Outlook外接程序栏中查看外接程序。

```xml
<Rule xsi:type="ItemIs" ItemType="Message" FormType="Read" />
```

下面的示例展示了 **ItemIs** 规则，可便于用户在阅读邮件或约会时在 Outlook 加载项栏中看到加载项。

```xml
<Rule xsi:type="RuleCollection" Mode="Or">
  <Rule xsi:type="ItemIs" ItemType="Message" FormType="Read" />
  <Rule xsi:type="ItemIs" ItemType="Appointment" FormType="Read" />
</Rule>
```


## <a name="itemhasattachment-rule"></a>ItemHasAttachment 规则


复杂 `ItemHasAttachment` 类型定义一个规则，用于检查所选项目是否包含附件。

```xml
<Rule xsi:type="ItemHasAttachment" />
```


## <a name="itemhasknownentity-rule"></a>ItemHasKnownEntity 规则

在项对外接程序可用之前，服务器将对其进行检查以确定主题和正文是否包含可能为某个已知实体的任何文本。 如果找到这些实体中的任意一个，则放置在使用该项的 或 方法访问的 `getEntities` 已知 `getEntitiesByType` 实体集合中。

您可以使用 在项中出现指定类型的实体时显示外接程序来 `ItemHasKnownEntity` 指定规则。 可以在规则的 属性中指定以下 `EntityType` 已知 `ItemHasKnownEntity` 实体。

- Address
- Contact
- EmailAddress
- MeetingSuggestion
- PhoneNumber
- TaskSuggestion
- URL

可以选择在 属性中包括正则表达式，以便仅在存在与正则表达式匹配的实体 `RegularExpression` 时显示外接程序。 若要获取规则中指定的正则表达式的匹配项，可以将 或 方法用于当前选定的 `ItemHasKnownEntity` `getRegExMatches` `getFilteredEntitiesByName` Outlook项。

以下示例显示一组元素，这些元素在邮件中出现指定的已知实体之一时 `Rule` 显示外接程序。

```xml
<Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasKnownEntity" EntityType="Address" />
    <Rule xsi:type="ItemHasKnownEntity" EntityType="MeetingSuggestion" />
    <Rule xsi:type="ItemHasKnownEntity" EntityType="TaskSuggestion" />
</Rule>
```

以下示例显示一个包含属性的规则，当邮件中包含单词"contoso"的 URL 时，该规则 `ItemHasKnownEntity` `RegularExpression` 将激活外接程序。


```xml
<Rule xsi:type="ItemHasKnownEntity" EntityType="Url" RegularExpression="contoso" />
```

有关激活规则中的实体的详细信息，请参阅[将 Outlook 项中的字符串作为已知实体进行匹配](match-strings-in-an-item-as-well-known-entities.md)。


## <a name="itemhasregularexpressionmatch-rule"></a>ItemHasRegularExpressionMatch 规则

`ItemHasRegularExpressionMatch`复杂类型定义一个规则，该规则使用正则表达式来匹配项目的指定属性的内容。 如果在项的指定属性中发现与正则表达式匹配的文本，则 Outlook 会激活外接程序栏并显示外接程序。 可以使用表示当前选定项的对象的 或 方法 `getRegExMatches` `getRegExMatchesByName` 获取指定正则表达式的匹配项。

以下示例演示一个 ，当选定项的正文包含"apple"、"apple"或"可能忽略大小写"时，将激活 `ItemHasRegularExpressionMatch` 加载项。

```xml
<Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
```

有关使用该规则的信息 `ItemHasRegularExpressionMatch` ，请参阅使用[正则表达式激活规则Outlook外接程序。](use-regular-expressions-to-show-an-outlook-add-in.md)


## <a name="rulecollection-rule"></a>RuleCollection 规则


复杂 `RuleCollection` 类型将多个规则合并为一个规则。 可以使用 属性指定集合中的规则是应结合逻辑 OR 还是逻辑 `Mode` AND。

指定逻辑 AND 时，项必须与集合中的所有指定规则匹配才能显示外接程序。指定逻辑 OR 时，与集合中的任何指定规则匹配的项都将显示外接程序。

您可以组合 `RuleCollection` 规则以形成复杂的规则。 以下示例在用户查看约会或邮件项目（项目的主题或正文包含地址）时激活外接程序。

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemIs" ItemType="Message" FormType="Read" />
    <Rule xsi:type="ItemIs" ItemType="Appointment" FormType="Read"/>
  </Rule>
  <Rule xsi:type="ItemHasKnownEntity" EntityType="Address" />
</Rule>
```

以下示例在用户撰写邮件时或查看约会（约会的标题或正文包含地址）时激活外接程序。

```xml
<Rule xsi:type="RuleCollection" Mode="Or"> 
  <Rule xsi:type="ItemIs" ItemType="Message" FormType="Edit" /> 
  <Rule xsi:type="RuleCollection" Mode="And">
    <Rule xsi:type="ItemIs" ItemType="Appointment" FormType="Read" />
    <Rule xsi:type="ItemHasKnownEntity" EntityType="Address" />
  </Rule> 
</Rule>
```


## <a name="limits-for-rules-and-regular-expressions"></a>规则和正则表达式的限制


为了提供使用 Outlook 外接程序的满意体验，您应该遵守激活和 API 使用准则。 下表显示了正则表达式和规则的常规限制，但不同的应用程序有特定的规则。 有关详细信息，请参阅 [Outlook 外接程序的激活和 JavaScript API 的限制](limits-for-activation-and-javascript-api-for-outlook-add-ins.md)和 [排查 Outlook 外接程序激活问题](troubleshoot-outlook-add-in-activation.md)。

<br/>

|**外接程序元素**|**准则**|
|:-----|:-----|
|清单大小|不大于 256 KB。|
|规则|不超过 15 条规则。|
|ItemHasKnownEntity|Outlook 富客户端将对正文的前 1 MB 内容应用规则，对正文其余部分则不应用。|
|正则表达式|对于所有应用程序上的 ItemHasKnownEntity 或 ItemHasRegularExpressionMatch Outlook规则：<br><ul><li>在 Outlook 加载项的激活规则中指定不超过 5 个正则表达式。如果超过该限制，则无法安装加载项。</li><li>指定由 <b>getRegExMatches</b> 方法调用在前 50 个匹配项内返回其预期结果的正则表达式。 </li><li>在正则表达式中指定向前断言，但不支持向后 `(?<=text)` 和否定向后 `(?<!text)` 断言。</li><li>指定其匹配不超过下表中的限制的正则表达式。<br/><br/><table><tr><th>正则表达式匹配项的长度限制</th><th>Outlook 富客户端</th><th>iOS 版和 Android 版 Outlook</th></tr><tr><td>项目正文采用纯文本</td><td>1.5 KB</td><td>3 KB</td></tr><tr><td>项目正文采用 HTML</td><td>3 KB</td><td>3KB</td></tr></table>|

## <a name="see-also"></a>另请参阅

- [创建适用于撰写窗体的 Outlook 加载项](compose-scenario.md)
- [Outlook 加载项的激活限制和 JavaScript API](limits-for-activation-and-javascript-api-for-outlook-add-ins.md)
- [使用正则表达式激活规则显示 Outlook 加载项](use-regular-expressions-to-show-an-outlook-add-in.md)
- [将 Outlook 项中的字符串作为已知实体进行匹配](match-strings-in-an-item-as-well-known-entities.md)
    
