---
title: Outlook 加载项的激活规则
description: 如果用户正在读取或撰写的邮件或约会符合加载项的激活规则，则 Outlook 将激活某些类型的加载项。
ms.date: 12/09/2021
ms.localizationpriority: medium
ms.openlocfilehash: af9edf0254156d7bdac13d0553036a614d8c4c39
ms.sourcegitcommit: df7964b6509ee6a807d754fbe895d160bc52c2d3
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 07/20/2022
ms.locfileid: "66889637"
---
# <a name="activation-rules-for-contextual-outlook-add-ins"></a>上下文 Outlook 加载项的激活规则

如果用户正在读取或撰写的邮件或约会符合外接程序的激活规则，则 Outlook 将激活某些类型的外接程序。这一点对使用 1.1 清单架构的所有外接程序均适用。然后，用户可从 Outlook UI 选择外接程序，以开始将其用于当前项目。

下图显示在“阅读”窗格中的邮件的外接程序栏中激活的 Outlook 外接程序。

![显示已激活的读取邮件应用的应用栏。](../images/read-form-app-bar.png)

## <a name="specify-activation-rules-in-a-manifest"></a>在清单中指定激活规则

若要让 Outlook 激活特定条件的外接程序，请使用以下 `Rule` 元素之一在加载项清单中指定激活规则。

- [Rule 元素 (MailApp complexType)](/javascript/api/manifest/rule) - 指定单个规则。
- [Rule 元素 (RuleCollection complexType)](/javascript/api/manifest/rule#rulecollection) - 使用逻辑操作组合多个规则。

 > [!NOTE]
 > `Rule`用于指定单个规则的元素属于抽象[规则](/javascript/api/manifest/rule)复杂类型。 以下每种类型的规则都扩展了此抽象 `Rule` 复杂类型。 因此当你在清单中指定单个规则时，你必须使用 [xsi:type](https://www.w3.org/TR/xmlschema-1/) 属性来进一步定义某个以下类型的规则。
 >
 > 例如，以下规则定义了 [ItemIs](/javascript/api/manifest/rule#itemis-rule) 规则。
 > `<Rule xsi:type="ItemIs" ItemType="Message" />`
 >
 > 该 `FormType` 属性适用于清单 v1.1 中的激活规则，但未在 v1.0 中 `VersionOverrides` 定义。 因此，当在节点中`VersionOverrides`使用 [ItemIs](/javascript/api/manifest/rule#itemis-rule) 时，无法使用它。

下表列出了可用的规则类型。你可以在表后面以及[创建适用于阅读窗体的 Outlook 外接程序](read-scenario.md)中指定的文章中查找更多信息。

|**规则名称**|**适用的窗体**|**说明**|
|:-----|:-----|:-----|
|[ItemIs](#itemis-rule)|读取，撰写|检查当前项目是否属于指定类型（邮件或约会），另外还可以检查项目类别、窗体类型和（可选）项目邮件类别。|
|[ItemHasAttachment](#itemhasattachment-rule)|读取|检查所选项是否包含附件。|
|[ItemHasKnownEntity](#itemhasknownentity-rule)|读取|检查所选项是否包含一个或多个已知实体。更多信息：[将 Outlook 项中的字符串作为已知实体进行匹配](match-strings-in-an-item-as-well-known-entities.md)。|
|[ItemHasRegularExpressionMatch](#itemhasregularexpressionmatch-rule)|读取|检查发件人的电子邮件地址、所选项的主题和/或所选项的正文是否包含正则表达式的匹配项。更多信息： [使用正则表达式激活规则显示 Outlook 外接程序](use-regular-expressions-to-show-an-outlook-add-in.md)。|
|[RuleCollection](#rulecollection-rule)|读取，撰写|组合一组规则以便形成更复杂的规则。|

## <a name="itemis-rule"></a>ItemIs 规则

复杂 `ItemIs` 类型定义一个规则，该规则的计算结果为 `true` 当前项是否与项类型匹配;如果在规则中声明了项消息类，则可选择使用该规则。

在规则的属性中 `ItemType` 指定以下项类型之一 `ItemIs` 。 可以在清单中指定多个 `ItemIs` 规则。 ItemType simpleType 定义了支持 Outlook 加载项的 Outlook 项类型。

|**Value**|**说明**|
|:-----|:-----|
|**约会**|在 Outlook 日历中指定一个项目。 这包括已获取响应并且具有组织者和参与者的会议项目，或者没有组织者或参与者且仅为日历上的一个项目的约会。 这与 Outlook 中的 IPM.Appointment 邮件类别相对应。|
|**邮件**|指定通常在收件箱中收到的以下项之一。 <ul><li><p>电子邮件。 这与 Outlook 中的 IPM.Note 邮件类别相对应。</p></li><li><p>会议请求、响应或取消。 这与 Outlook 中的以下消息类相对应。</p><p>IPM.Schedule.Meeting.Request</p><p>IPM.Schedule.Meeting.Neg</p><p>IPM.Schedule.Meeting.Pos</p><p>IPM.Schedule.Meeting.Tent</p><p>IPM.Schedule.Meeting.Canceled</p></li></ul>|

该 `FormType` 属性用于指定 (读取或撰写加载项应在其中激活的) 模式。

 > [!NOTE]
 > ItemIs `FormType` 属性在架构 v1.1 及更高版本中定义，但不是在 v1.0 中 `VersionOverrides` 定义。 定义加载项命令时不要包含 `FormType` 该属性。

激活外接程序后，可以使用 [mailbox.item](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item) 属性获取 Outlook 中的当前所选项，以及使用 [item.itemType](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties) 属性获取当前项的类型。

可以选择性地使用该 `ItemClass` 属性来指定项的消息类，以及 `IncludeSubClasses` 指定当项是指定类的子类时规则是否应 `true` 为该规则的属性。

若要详细了解邮件类，请参阅[项类型和邮件类](/office/vba/outlook/Concepts/Forms/item-types-and-message-classes)。

下面的示例是一个 `ItemIs` 规则，允许用户在用户读取消息时在 Outlook 加载项栏中查看加载项。

```xml
<Rule xsi:type="ItemIs" ItemType="Message" FormType="Read" />
```

下面的示例是一个 `ItemIs` 规则，允许用户在用户读取消息或约会时在 Outlook 外接程序栏中查看外接程序。

```xml
<Rule xsi:type="RuleCollection" Mode="Or">
  <Rule xsi:type="ItemIs" ItemType="Message" FormType="Read" />
  <Rule xsi:type="ItemIs" ItemType="Appointment" FormType="Read" />
</Rule>
```

## <a name="itemhasattachment-rule"></a>ItemHasAttachment 规则

复杂 `ItemHasAttachment` 类型定义一个规则，用于检查所选项是否包含附件。

```xml
<Rule xsi:type="ItemHasAttachment" />
```

## <a name="itemhasknownentity-rule"></a>ItemHasKnownEntity 规则

在项对外接程序可用之前，服务器将对其进行检查以确定主题和正文是否包含可能为某个已知实体的任何文本。 如果找到这些实体中的任何一个，则会将其放置在使用 `getEntities` 该项或 `getEntitiesByType` 方法访问的已知实体的集合中。

当项目中存在指定类型的实体时，可以使用 `ItemHasKnownEntity` 该规则来指定一个显示加载项的规则。 可以在规则的属性`ItemHasKnownEntity`中`EntityType`指定以下已知实体。

- Address
- Contact
- EmailAddress
- MeetingSuggestion
- PhoneNumber
- TaskSuggestion
- URL

可以选择在属性中 `RegularExpression` 包含正则表达式，以便仅当与当前正则表达式匹配的实体时才显示外接程序。 若要获取与规则中 `ItemHasKnownEntity` 指定的正则表达式的匹配项，可以使用 `getRegExMatches` 当前所选 Outlook 项目的或 `getFilteredEntitiesByName` 方法。

下面的 `Rule` 示例显示了在消息中存在指定的已知实体之一时显示加载项的元素集合。

```xml
<Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasKnownEntity" EntityType="Address" />
    <Rule xsi:type="ItemHasKnownEntity" EntityType="MeetingSuggestion" />
    <Rule xsi:type="ItemHasKnownEntity" EntityType="TaskSuggestion" />
</Rule>
```

以下示例演示一个 `ItemHasKnownEntity` 规则，该规则具有一个 `RegularExpression` 属性，当消息中存在包含“contoso”一词的 URL 时，该规则会激活加载项。

```xml
<Rule xsi:type="ItemHasKnownEntity" EntityType="Url" RegularExpression="contoso" />
```

有关激活规则中的实体的详细信息，请参阅[将 Outlook 项中的字符串作为已知实体进行匹配](match-strings-in-an-item-as-well-known-entities.md)。

## <a name="itemhasregularexpressionmatch-rule"></a>ItemHasRegularExpressionMatch 规则

复杂 `ItemHasRegularExpressionMatch` 类型定义一个规则，该规则使用正则表达式来匹配项的指定属性的内容。 如果在项的指定属性中发现与正则表达式匹配的文本，则 Outlook 会激活外接程序栏并显示外接程序。 可以使用 `getRegExMatches` 表示当前所选项的对象或 `getRegExMatchesByName` 方法来获取指定正则表达式的匹配项。

以下示例演示当 `ItemHasRegularExpressionMatch` 所选项的正文包含“apple”、“banana”或“椰子”（忽略大小写）时激活加载项的示例。

```xml
<Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
```

有关使用规则的 `ItemHasRegularExpressionMatch` 详细信息，请参阅 [使用正则表达式激活规则来显示 Outlook 加载项](use-regular-expressions-to-show-an-outlook-add-in.md)。

## <a name="rulecollection-rule"></a>RuleCollection 规则

复杂 `RuleCollection` 类型将多个规则合并到单个规则中。 可以使用 `Mode` 该属性指定集合中的规则是应与逻辑 OR 还是逻辑 AND 结合使用。

指定逻辑 AND 时，项必须与集合中的所有指定规则匹配才能显示外接程序。指定逻辑 OR 时，与集合中的任何指定规则匹配的项都将显示外接程序。

可以组合 `RuleCollection` 规则来形成复杂的规则。 以下示例在用户查看约会或邮件项目（项目的主题或正文包含地址）时激活外接程序。

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

为了提供使用 Outlook 外接程序的满意体验，您应该遵守激活和 API 使用准则。 下表显示了正则表达式和规则的一般限制，但不同应用程序有特定的规则。 有关详细信息，请参阅 [Outlook 外接程序的激活和 JavaScript API 的限制](limits-for-activation-and-javascript-api-for-outlook-add-ins.md)和 [排查 Outlook 外接程序激活问题](troubleshoot-outlook-add-in-activation.md)。

|**外接程序元素**|**准则**|
|:-----|:-----|
|清单大小|不大于 256 KB。|
|规则|不超过 15 条规则。|
|ItemHasKnownEntity|Outlook 富客户端将对正文的前 1 MB 内容应用规则，对正文其余部分则不应用。|
|正则表达式|对于所有 Outlook 应用程序的 ItemHasKnownEntity 或 ItemHasRegularExpressionMatch 规则：<br><ul><li>在 Outlook 加载项的激活规则中指定不超过 5 个正则表达式。如果超过该限制，则无法安装加载项。</li><li>指定由 <b>getRegExMatches</b> 方法调用在前 50 个匹配项内返回其预期结果的正则表达式。 </li><li>**重要** 提示：文本基于匹配正则表达式后产生的字符串突出显示。 不过，突出显示的出现可能与实际正则表达式断言的结果不完全匹配，例如负前 `(?!text)`观、后 `(?<=text)`看和负面观望 `(?<!text)`。 例如，如果在“Like under、under score 和 underscore”上使用正则表达式 `under(?!score)` ，则字符串“under”将突出显示所有匹配项，而不只是前两个匹配项。</li><li>指定匹配项不超过下表中的限制的正则表达式。<br/><br/><table><tr><th>正则表达式匹配项的长度限制</th><th>Outlook 富客户端</th><th>iOS 版和 Android 版 Outlook</th></tr><tr><td>项目正文采用纯文本</td><td>1.5 KB</td><td>3 KB</td></tr><tr><td>项目正文采用 HTML</td><td>3 KB</td><td>3KB</td></tr></table>|

## <a name="see-also"></a>另请参阅

- [创建适用于撰写窗体的 Outlook 加载项](compose-scenario.md)
- [Outlook 加载项的激活限制和 JavaScript API](limits-for-activation-and-javascript-api-for-outlook-add-ins.md)
- [使用正则表达式激活规则显示 Outlook 加载项](use-regular-expressions-to-show-an-outlook-add-in.md)
- [将 Outlook 项中的字符串作为已知实体进行匹配](match-strings-in-an-item-as-well-known-entities.md)
