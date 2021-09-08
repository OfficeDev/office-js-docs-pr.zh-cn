---
title: 使用正则表达式激活规则显示加载项
description: 了解如何为 Outlook 上下文加载项使用正则表达式激活规则。
ms.date: 07/28/2020
localization_priority: Normal
ms.openlocfilehash: d334ba6b2e0f044fc8d876cd6edd218743ccb390
ms.sourcegitcommit: 42c55a8d8e0447258393979a09f1ddb44c6be884
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/08/2021
ms.locfileid: "58937009"
---
# <a name="use-regular-expression-activation-rules-to-show-an-outlook-add-in"></a>使用正则表达式激活规则显示 Outlook 外接程序

可以将正则表达式规则指定为在邮件的特定字段中找到匹配项时激活[上下文外接程序](contextual-outlook-add-ins.md)。 上下文外接程序仅在阅读模式下激活，Outlook 不会在用户撰写某个项目时激活上下文外接程序。 还有其他一些情况，Outlook无法激活外接程序，例如，数字签名项目。 有关详细信息，请参阅 [Outlook 外接程序的激活规则](activation-rules.md)。

你可以将正则表达式指定为外接程序 XML 清单中的 [ItemHasRegularExpressionMatch](../reference/manifest/rule.md#itemhasregularexpressionmatch-rule) 规则或 [ItemHasKnownEntity](../reference/manifest/rule.md#itemhasknownentity-rule) 规则的一部分。 在 [DetectedEntity](../reference/manifest/extensionpoint.md#detectedentity) 扩展点中指定了这些规则。

Outlook 基于客户端计算机上浏览器所使用的 JavaScript 解释器的规则计算正则表达式。 Outlook 支持所有 XML 处理器也支持的相同特殊字符列表。 下表列出了这些特殊字符。 你可以通过为相应字符指定转义序列以在正则表达式中使用这些字符，如下表中所述。

<br/>

|字符|说明|要使用的转义序列|
|:-----|:-----|:-----|
|`"`|双引号|`&quot;`|
|`&`|与号|`&amp;`|
|`'`|撇号|`&apos;`|
|`<`|小于号|`&lt;`|
|`>`|大于号|`&gt;`|

## <a name="itemhasregularexpressionmatch-rule"></a>ItemHasRegularExpressionMatch 规则

`ItemHasRegularExpressionMatch` 规则对于基于受支持属性的特定值控制外接程序的激活很有用。 `ItemHasRegularExpressionMatch` 规则具有以下属性。

<br/>

|属性名|说明|
|:-----|:-----|
|`RegExName`|指定正则表达式的名称，以便能够在外接程序的代码中引用该表达式。|
|`RegExValue`|指定将对其求值的正则表达式，以确定是否应显示外接程序。|
|`PropertyName`|指定正则表达式进行计算所依据的属性名称。 允许的值为 `BodyAsHTML`、`BodyAsPlaintext`、`SenderSMTPAddress` 和 `Subject`。<br/><br/>如果指定 `BodyAsHTML`，则 Outlook 只会在项目正文为 HTML 时应用正则表达式。 否则，Outlook 将不会返回该正则表达式的匹配项。<br/><br/>如果指定 `BodyAsPlaintext`，则 Outlook 将始终对项目正文应用正则表达式。<br/><br/>**注释：** 如果指定 `Rule` 元素的 `Highlight` 属性，则必须将 `PropertyName` 属性设为 `BodyAsPlaintext`。|
|`IgnoreCase`|指定当匹配由 `RegExName` 指定的正则表达式时是否忽略大小写。|
| `Highlight` | 指定客户端应如何突出显示匹配的文本。 此元素仅适用于 `ExtensionPoint` 元素中的 `Rule` 元素。 可以是以下值之一：`all` 或 `none`。 如果未指定，则默认值为 `all`。<br/><br/>**注释：** 如果指定 `Rule` 元素的 `Highlight` 属性，则必须将 `PropertyName` 属性设为 `BodyAsPlaintext`。 |

### <a name="best-practices-for-using-regular-expressions-in-rules"></a>在规则中使用正则表达式的最佳做法

在使用正则表达式时，请特别注意以下几点。

- 如果在项目的正文中指定 `ItemHasRegularExpressionMatch` 规则，则正则表达式应进一步筛选正文，不应尝试返回该项目的整个正文。 使用正则表达式（如 `.*`）来尝试获取项目的整个正文并不总是返回预期的结果。
- 一个浏览器上返回的纯文本正文与另一个浏览器上返回的纯文本正文可能略有不同。 如果使用含有 `BodyAsPlaintext` 的 `ItemHasRegularExpressionMatch` 规则作为 `PropertyName` 属性，请在你的外接程序支持的所有浏览器上测试正则表达式。

    因为不同的浏览器获取所选项目的文本正文的方法不同，所以应确保你的正则表达式支持正文文本部分所返回的细微差异。 例如，一些浏览器（如 Internet Explorer 9）使用 DOM 的 `innerText` 属性，而其他浏览器（如 Firefox）使用.`.textContent()` 方法来获取项目的文本正文。 同样，不同浏览器所返回的换行符也可能不同：在 Internet Explorer 上返回的换行符为 `\r\n`，而在 Firefox 和 Chrome 上返回的换行符为 `\n`。 有关详细信息，请参阅 [W3C DOM 兼容性 - HTML](https://quirksmode.org/dom/html/)。

- Outlook 富客户端与 Outlook 网页版或 Outlook Mobile 之间的项目的 HTML 正文略有不同。 请仔细定义正则表达式。

- 根据 Outlook 客户端、设备类型或要应用正则表达式的属性，在设计正则表达式作为激活规则时，您应该了解每个客户端的其他最佳实践和限制。 有关详细信息，请参阅 [Outlook 外接程序的激活和 JavaScript API 的限制](limits-for-activation-and-javascript-api-for-outlook-add-ins.md)。

### <a name="examples"></a>示例

以下 `ItemHasRegularExpressionMatch` 规则将在发件人的 SMTP 电子邮件地址与 `@contoso` 匹配（不管是大写还是小写字符）时激活外接程序。

```XML
<Rule xsi:type="ItemHasRegularExpressionMatch"
    RegExName="addressMatches"
    RegExValue="@[cC][oO][nN][tT][oO][sS][oO]"
    PropertyName="SenderSMTPAddress"
/>
```

<br/>

以下是使用 `IgnoreCase` 属性指定同一正则表达式的另一种方式。

```XML
<Rule xsi:type="ItemHasRegularExpressionMatch"
    RegExName="addressMatches"
    RegExValue="@contoso"
    PropertyName="SenderSMTPAddress"
    IgnoreCase="true"
/>
```

<br/>

以下 `ItemHasRegularExpressionMatch` 规则将在股票代号包含在当前项目的正文中时激活外接程序。

```XML
<Rule xsi:type="ItemHasRegularExpressionMatch"
    PropertyName="BodyAsPlaintext"
    RegExName="TickerSymbols"
    RegExValue="\b(NYSE|NASDAQ|AMEX):\s*[A-Za-z]+\b"/>

```

## <a name="itemhasknownentity-rule"></a>ItemHasKnownEntity 规则

`ItemHasKnownEntity` 规则根据所选项目的主题或正文中是否存在实体来激活外接程序。 [EntityType](/javascript/api/outlook/office.mailboxenums.entitytype) 类型定义受支持的实体。 在 `ItemHasKnownEntity` 规则中应用正则表达式，可为基于实体（例如，一组特定的 URL，或含有某个区号的电话号码）的值子集进行的激活提供便利。

> [!NOTE]
> Outlook 只能提取用英语编写的实体字符串，无论清单中指定的默认区域设置如何。 仅邮件支持 `MeetingSuggestion` 实体类型；约会不支持该类型。 你无法从“已发送邮件”文件夹的邮件中提取实体，也不能使用 `ItemHasKnownEntity` 规则来激活“已发送邮件”文件夹中邮件的外接程序。

`ItemHasKnownEntity` 规则支持下表中的属性。 请注意，尽管在 `ItemHasKnownEntity` 规则中指定正则表达式是可选项，如果选择使用正则表达式作为实体筛选器，则必须同时指定 `RegExFilter` 和 `FilterName` 属性。

<br/>

|属性名|说明|
|:-----|:-----|
|`EntityType`|指定若想规则计算结果为 `true` 而必须存在的实体类型。 请使用多个规则来指定多个类型的实体。|
|`RegExFilter`|指定用于进一步筛选由 `EntityType` 指定的实体实例的正则表达式。|
|`FilterName`|指定由 `RegExFilter` 指定的正则表达式的名称，以便稍后可通过代码引用它。|
|`IgnoreCase`|指定当匹配由 `RegExFilter` 指定的正则表达式时是否忽略大小写。|

### <a name="examples"></a>示例

下面的 `ItemHasKnownEntity` 规则将在当前项目的主题或正文中存在 URL 且该 URL 包含字符串 `youtube` 时激活外接程序，而不考虑字符串的大小写。

```XML
<Rule xsi:type="ItemHasKnownEntity"
    EntityType="Url"
    RegExFilter="youtube"
    FilterName="youtube"
    IgnoreCase="true"/>
```

## <a name="using-regular-expression-results-in-code"></a>在代码中使用正则表达式结果

可以通过对当前项使用下列方法获取正则表达式的匹配项。

- [getRegExMatches](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods) 为在外接程序的 `ItemHasRegularExpressionMatch` 和 `ItemHasKnownEntity` 规则中指定的所有正则表达式返回当前项目中的匹配项。

- [getRegExMatchesByName](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods) 为外接程序的 `ItemHasRegularExpressionMatch` 规则中指定的已标识正则表达式返回当前项目中的匹配项。

- [getFilteredEntitiesByName](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods) 对于包含在外接程序的 `ItemHasKnownEntity` 规则中指定的已标识正则表达式匹配项的实体，将返回完整实例。

计算正则表达式时，匹配项将以数组对象的形式返回到你的外接程序。 对于 `getRegExMatches`，该对象具有正则表达式名称的标识符。

> [!NOTE]
> Outlook 不会在数组中以任何特定顺序返回匹配项。 另外，即使在同一邮箱中的同一项目上的每个客户端运行相同的外接程序，也不应假定匹配项返回的顺序与数组中返回的顺序相同。

### <a name="examples"></a>示例

以下是包含 `ItemHasRegularExpressionMatch` 规则且具有名为 `videoURL` 的正则表达式的规则集合示例。

```XML
<Rule xsi:type="RuleCollection" Mode="And">
    <Rule xsi:type="ItemIs" ItemType="Message"/>
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="videoURL" RegExValue="http://www\.youtube\.com/watch\?v=[a-zA-Z0-9_-]{11}" PropertyName="BodyAsPlaintext"/>
</Rule>
```

<br/>

以下示例使用当前项目的 `getRegExMatches` 将变量 `videos` 设置为上一个 `ItemHasRegularExpressionMatch` 规则的结果。

```js
var videos = Office.context.mailbox.item.getRegExMatches().videoURL;
```

<br/>

多个匹配项将作为数组元素存储在该对象中。以下代码示例说明如何对名为  `reg1` 的正则表达式循环访问匹配项以生成将显示为 HTML 的字符串。

```js
function initDialer()
{
    var myEntities;
    var myString;
    var myCell;
    myEntities = Office.context.mailbox.item.getRegExMatches();

    myString = "";
    myCell = document.getElementById('dialerholder');
    // Loop over the myEntities collection.
    for (var i in myEntities.reg1) {
        myString += "<p><a href='callto:tel:" + myEntities.reg1[i] + "'>" + myEntities.reg1[i] + "</a></p>";
    }

    myCell.innerHTML = myString;
}
```

<br/>

以下是指定 `MeetingSuggestion` 实体和名为 `CampSuggestion` 的正则表达式的 `ItemHasKnownEntity` 规则的示例。 Outlook 在检测到当前所选项目包含会议建议，并且主题或正文包含术语 `WonderCamp` 时将激活外接程序。

```XML
<Rule xsi:type="ItemHasKnownEntity"
    EntityType="MeetingSuggestion"
    RegExFilter="WonderCamp"
    FilterName="CampSuggestion"
    IgnoreCase="false"/>
```

<br/>

以下代码示例使用当前项目中的 `getFilteredEntitiesByName` 设置变量 `suggestions`，以获取针对上一个 `ItemHasKnownEntity` 规则检测到的一组会议建议。

```js
var suggestions = Office.context.mailbox.item.getFilteredEntitiesByName("CampSuggestion");
```

## <a name="see-also"></a>另请参阅

- [Outlook 外接程序：Contoso 订单编号](https://github.com/OfficeDev/Outlook-Add-In-Contextual-Regex) - 基于正则表达式匹配项激活的示例上下文外接程序。
- [创建适用于阅读窗体的 Outlook 外接程序](read-scenario.md)
- [Outlook 外接程序的激活规则](activation-rules.md)
- [Outlook 外接程序的激活和 JavaScript API 限制](limits-for-activation-and-javascript-api-for-outlook-add-ins.md)
- [将 Outlook 项中的字符串作为已知实体进行匹配](match-strings-in-an-item-as-well-known-entities.md)
- [.NET Framework 中的正则表达式的最佳做法](/dotnet/standard/base-types/best-practices)
